import pandas as pd
import os
import gdown
import time

# --- PENGATURAN ---
CSV_FILE_PATH = 'dataset_magang.csv'
DOWNLOAD_FOLDER = 'cv'
COLUMN_NAME = 'CV'
# Nama file log untuk mencatat link yang sudah di-download
LOG_FILE_NAME = '_download_log.txt'
# --------------------

def load_downloaded_links(log_path):
    """Membaca file log dan mengembalikan set berisi link."""
    downloaded = set()
    if os.path.exists(log_path):
        try:
            with open(log_path, 'r') as f:
                for line in f:
                    downloaded.add(line.strip())
            print(f"Berhasil memuat {len(downloaded)} link dari log.")
        except Exception as e:
            print(f"Peringatan: Gagal membaca file log: {e}")
    return downloaded

def add_to_log(link, log_path):
    """Menambahkan satu link ke file log."""
    try:
        with open(log_path, 'a') as f:
            f.write(link + '\n')
    except Exception as e:
        print(f"Peringatan: Gagal menulis ke file log: {e}")

def main():
    print(f"Memulai proses download...")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_path = os.path.join(script_dir, CSV_FILE_PATH)
    download_dir = os.path.join(script_dir, DOWNLOAD_FOLDER)
    log_file_path = os.path.join(download_dir, LOG_FILE_NAME)

    # 1. Buat folder download
    os.makedirs(download_dir, exist_ok=True)
    print(f"Hasil download akan disimpan di: {download_dir}")

    # 2. Muat daftar link yang sudah di-download
    downloaded_links = load_downloaded_links(log_file_path)

    # 3. Baca file CSV
    try:
        df = pd.read_csv(csv_path, skiprows=1)
        print(f"Berhasil membaca file: {CSV_FILE_PATH}")
    except Exception as e:
        print(f"ERROR saat membaca CSV: {e}")
        return

    # 4. Dapatkan link unik
    if COLUMN_NAME not in df.columns:
        print(f"ERROR: Kolom '{COLUMN_NAME}' tidak ditemukan.")
        return
        
    links = df[COLUMN_NAME].dropna().unique()
    print(f"Ditemukan total {len(links)} link unik.")

    # 5. Pindah ke folder download
    original_cwd = os.getcwd()
    os.chdir(download_dir)
    
    print("\n--- Mulai Memeriksa dan Mengunduh File (Logika Baru) ---")

    # 6. Loop dan download
    success_count = 0
    fail_count = 0
    skipped_count = 0

    for i, link in enumerate(links):
        link_str = str(link).strip() # Pastikan bersih dari spasi
        
        if not link_str or 'drive.google.com' not in link_str:
            print(f"\n({i+1}/{len(links)}) Melewatkan: '{link_str[:50]}...' (Link tidak valid)")
            continue

        print(f"\n({i+1}/{len(links)}) Memeriksa: {link_str}")
        
        # === LOGIKA BARU: Cek berdasarkan file log ===
        if link_str in downloaded_links:
            print("... SUDAH ADA DI LOG. MELEWATKAN.")
            skipped_count += 1
            success_count += 1 # Anggap sukses karena sudah ada
            continue
            
        # Jika belum ada di log, coba download
        print("... Link baru, mencoba mengunduh...")
        try:
            # Panggil gdown.download
            file_name = gdown.download(link_str, quiet=False, fuzzy=True)
            
            # Jika download berhasil (gdown mengembalikan nama file)
            if file_name:
                print(f"... BERHASIL diunduh sebagai: '{file_name}'")
                add_to_log(link_str, log_file_path) # Catat ke log
                success_count += 1
            else:
                # Ini terjadi jika gdown gagal tapi tidak error (jarang)
                print("... GAGAL (gdown tidak mengembalikan nama file).")
                fail_count += 1

        except Exception as e:
            # Ini akan menangkap error "Cannot retrieve..." (kena limit)
            # atau error "Permission denied"
            error_message = str(e).split('\n')[0] # Ambil baris pertama error
            print(f"... GAGAL: {error_message}...")
            fail_count += 1
        
        # Beri jeda 1 detik
        time.sleep(1)

    # 7. Kembali ke folder semula
    os.chdir(original_cwd)

    print("\n--- Selesai ---")
    print(f"Total link unik: {len(links)}")
    print(f"Berhasil (termasuk yg dilewati): {success_count}")
    print(f"Dilewati (karena sudah di log): {skipped_count}")
    print(f"Gagal diunduh (kena limit/error): {fail_count}")
    
    if fail_count > 0:
        print("\nCATATAN: Jika masih ada yg Gagal, jalankan lagi skrip ini nanti.")
        print("Skrip akan otomatis melewatkan yg sudah berhasil dan hanya mencoba lagi yg gagal.")

if __name__ == "__main__":
    main()