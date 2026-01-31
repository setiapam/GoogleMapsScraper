import requests
import time
import pandas as pd
import io
import datetime
import os
import sys
from fuzzywuzzy import fuzz

# --- KONFIGURASI ---
FILE_DATA = "data.xlsx"
KOLOM_ID = "perusahaan_id"
KOLOM_CEK = "hasilgc"
KOLOM_RESUME_ALT = "input_id"
KOLOM_SKOR = "perbandingan"
BASE_URL = "http://localhost:8080/api/v1"
TIMEOUT_WORKING = 300

# Mapping kolom hasil: {Nama_di_CSV_Google : Nama_Kolom_di_Excel}
MAPPING_KOLOM = {
    "input_id": "input_id",
    "link": "link",
    "title": "title",
    "category": "category",
    "address": "address",
    "phone": "phone",
    "plus_code": "plus_code",
    "latitude": "lat",
    "longitude": "lon",
}


def log(msg):
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}")


def hitung_kemiripan(row):
    """Menghitung tingkat kemiripan nama dan alamat dalam persen"""
    try:
        score_nama = fuzz.token_sort_ratio(
            str(row.get("nama_usaha", "")), str(row.get("title", ""))
        )
        score_alamat = fuzz.token_sort_ratio(
            str(row.get("alamat_usaha", "")), str(row.get("address", ""))
        )
        return (score_nama + score_alamat) / 2
    except:
        return 0


def run_final_power_updater():
    log("=== MEMULAI POWER UPDATER (ULTIMATE STABILITY VERSION) ===")

    if not os.path.exists(FILE_DATA):
        log(f"ERROR: File {FILE_DATA} tidak ditemukan!")
        return

    try:
        # 1. Load Data & Inisialisasi
        df = pd.read_excel(FILE_DATA)
        target_cols = list(MAPPING_KOLOM.values()) + [
            KOLOM_SKOR,
            KOLOM_CEK,
            KOLOM_RESUME_ALT,
        ]
        for col in target_cols:
            if col not in df.columns:
                df[col] = None
            df[col] = df[col].astype(object)  # Fix Dtype Float64 Error

        # ---------------------------------------------------------
        # TAHAP 1: UPDATE SKOR LOKAL (Efisiensi Tanpa API)
        # ---------------------------------------------------------
        mask_skor_kosong = (
            (df[KOLOM_RESUME_ALT].notna())
            & (df[KOLOM_RESUME_ALT] != "")
            & (
                ~df[KOLOM_RESUME_ALT]
                .astype(str)
                .str.contains("TIDAK DITEMUKAN", na=False)
            )
            & (df[KOLOM_SKOR].isna() | (df[KOLOM_SKOR] == ""))
        )
        df_lokal = df[mask_skor_kosong].copy()

        if len(df_lokal) > 0:
            log(
                f"Ditemukan {len(df_lokal)} data lama tanpa skor. Menghitung secara lokal..."
            )
            for index, row in df_lokal.iterrows():
                df.at[index, KOLOM_SKOR] = hitung_kemiripan(row)
            df.to_excel(FILE_DATA, index=False)
            log("Update skor lokal selesai.")

        # ---------------------------------------------------------
        # TAHAP 2: SCRAPING DENGAN DUAL-RESUME LOGIC
        # ---------------------------------------------------------
        mask_todo = (df[KOLOM_CEK].isna() | (df[KOLOM_CEK] == "")) & (
            df[KOLOM_RESUME_ALT].isna() | (df[KOLOM_RESUME_ALT] == "")
        )
        df_todo = df[mask_todo].copy()

        log(f"Sisa baris yang perlu di-scrape: {len(df_todo)}")

        if len(df_todo) == 0:
            log("Semua data sudah terisi. Selesai.")
            return

        for index, row_data in df_todo.iterrows():
            current_id = row_data[KOLOM_ID]

            # Susun Keyword Dinamis
            parts = [
                str(row_data.get("nama_usaha", "")),
                str(row_data.get("nmdesa", "")),
                str(row_data.get("nmkec", "")),
                str(row_data.get("nmkab", "")),
            ]
            kw = " ".join([p.strip() for p in parts if p.strip()])

            log(f"Scraping ID {current_id} | Keyword: {kw}")

            # Persiapkan Payload API
            payload = {
                "name": f"job_{int(time.time())}",
                "keywords": [kw],
                "lang": "id",
                "zoom": 15,
                "radius": 1500,
                "depth": 1,
                "fast_mode": False,
                "max_time": 3600,
            }

            try:
                # A. Membuat Job
                res = requests.post(f"{BASE_URL}/jobs", json=payload)
                job_id = res.json().get("id")
                start_time = time.time()
                success = False

                # B. Polling Status (Dual Case Handle)
                while True:
                    job_data = requests.get(f"{BASE_URL}/jobs/{job_id}").json()
                    status = (
                        job_data.get("Status") or job_data.get("status", "")
                    ).lower()

                    if status == "ok":
                        success = True
                        break

                    if (time.time() - start_time > TIMEOUT_WORKING) or status in [
                        "error",
                        "failed",
                    ]:
                        log(
                            f"   [!] Gagal/Timeout di Server. Menandai TIDAK DITEMUKAN."
                        )
                        df.at[index, KOLOM_RESUME_ALT] = "TIDAK DITEMUKAN"
                        df.at[index, KOLOM_CEK] = "TIDAK DITEMUKAN"
                        break
                    time.sleep(10)

                # C. Download & Extract
                if success:
                    csv_res = requests.get(f"{BASE_URL}/jobs/{job_id}/download")
                    if (
                        csv_res.status_code == 200
                        and not csv_res.text.strip().startswith("<!DOCTYPE")
                    ):
                        try:
                            # Proteksi No Columns to Parse Error
                            df_temp = pd.read_csv(
                                io.StringIO(csv_res.text), on_bad_lines="skip"
                            )
                            if not df_temp.empty:
                                best_match = df_temp.iloc[0]
                                for col_csv, col_excel in MAPPING_KOLOM.items():
                                    df.at[index, col_excel] = str(
                                        best_match.get(col_csv, "")
                                    )

                                df.at[index, KOLOM_SKOR] = hitung_kemiripan(
                                    df.loc[index]
                                )
                                log(f"   [DONE] Match: {df.at[index, KOLOM_SKOR]}%")
                            else:
                                raise ValueError("CSV Empty Content")
                        except Exception as e:
                            log(f"   [!] Error CSV: {e}. Menandai TIDAK DITEMUKAN.")
                            df.at[index, KOLOM_RESUME_ALT] = "TIDAK DITEMUKAN"
                            df.at[index, KOLOM_CEK] = "TIDAK DITEMUKAN"
                    else:
                        log("   [!] File Rusak/HTML. Menandai TIDAK DITEMUKAN.")
                        df.at[index, KOLOM_RESUME_ALT] = "TIDAK DITEMUKAN"
                        df.at[index, KOLOM_CEK] = "TIDAK DITEMUKAN"

                    requests.delete(f"{BASE_URL}/jobs/{job_id}")

                # Simpan Bertahap per Baris
                df.to_excel(FILE_DATA, index=False)

            except Exception as e:
                log(
                    f"   [!] Error Fatal pada ID {current_id}: {e}. Menandai untuk lewati."
                )
                df.at[index, KOLOM_RESUME_ALT] = "TIDAK DITEMUKAN"
                df.at[index, KOLOM_CEK] = "TIDAK DITEMUKAN"
                df.to_excel(FILE_DATA, index=False)

    except KeyboardInterrupt:
        log("\n[STOP] Dihentikan paksa (Ctrl+C). Progres aman di data.xlsx.")
        sys.exit(0)

    log(f"=== PROSES SELESAI. Cek file {FILE_DATA} ===")


if __name__ == "__main__":
    run_final_power_updater()
