import pdfplumber
import pandas as pd
import re
import os

def ekstrak_teks_dari_pdf(pdf_path):
    """
    Membuka file PDF dan mengekstrak seluruh teksnya.
    """
    teks_lengkap = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for halaman in pdf.pages:
                teks_halaman = halaman.extract_text()
                if teks_halaman:
                    teks_lengkap += teks_halaman + "\n"
        return teks_lengkap
    except Exception as e:
        print(f"Error membaca {pdf_path}: {e}")
        return None

def parse_data_cv(teks):
    """
    Menggunakan Regex untuk mencari dan mengekstrak data spesifik dari teks CV.
    Versi ini telah DITINGKATKAN berdasarkan 5 sampel CV.
    """
    data = {
        "IPK": None,
        "Jurusan": None,
        "Semester": None,
        "Keterampilan": None,
        "Sertifikat": None
    }
    
    # --- Pola Regex yang Diperbarui ---
    # Kita akan menggunakan re.IGNORECASE (tidak peduli huruf besar/kecil)
    # dan re.DOTALL (titik '.' bisa mencocokkan baris baru)
    
    # --- Pola IPK (Mencari 2 format) ---
    try:
        # Pola 1: Mencari "IPK: 3.75" atau "GPA : 3,75"
        pola_ipk_1 = r"(IPK|GPA)\s*[:\s]*([\d\.,]+)"
        match = re.search(pola_ipk_1, teks, re.IGNORECASE)
        if match:
            data["IPK"] = match.group(2).replace(',', '.')
        else:
            # Pola 2: Mencari "3.5/4.00"
            pola_ipk_2 = r"([\d\.,]+)\s*/\s*4[.,]0{0,2}"
            match_2 = re.search(pola_ipk_2, teks, re.IGNORECASE)
            if match_2:
                data["IPK"] = match_2.group(1).replace(',', '.')
    except Exception as e:
        print(f"  [Parse Error] IPK: {e}")

    # --- Pola Jurusan (Mencari keyword di dekat "Universitas Dinamika" atau keyword spesifik) ---
    try:
        # Pola 1: Mencari S1/D3/Diploma/Sarjana + Nama Jurusan
        pola_jurusan_1 = r"((S1|D3|Diploma|Bachelor|Undergraduate)[\s,]+(Sistem Informasi|Information System|Desain Komunikasi Visual|Public Relations|DKV|SI))"
        match = re.search(pola_jurusan_1, teks, re.IGNORECASE)
        if match:
            data["Jurusan"] = match.group(1).strip() # Mengambil seluruh frasa (misal: "S1 Sistem Informasi")
        else:
             # Pola 2: Mengambil baris setelah "Universitas Dinamika" (jika Pola 1 gagal)
             pola_jurusan_2 = r"(Universitas Dinamika|Dinamika University).*?[\n\s]+(.*?)\n"
             match_2 = re.search(pola_jurusan_2, teks, re.IGNORECASE)
             if match_2:
                 # Membersihkan dari teks yang tidak relevan seperti tanggal
                 jurusan_raw = match_2.group(2).strip()
                 if not re.search(r'\d{4}', jurusan_raw): # Jika bukan tahun
                     data["Jurusan"] = jurusan_raw.split(',')[0] # Ambil bagian pertama
    except Exception as e:
        print(f"  [Parse Error] Jurusan: {e}")

    # --- Pola Semester ---
    try:
        # Mencari kata "semester" diikuti angka
        pola_semester = r"semester\s*(\d+)"
        match = re.search(pola_semester, teks, re.IGNORECASE)
        if match:
            data["Semester"] = match.group(1).strip()
    except Exception as e:
        print(f"  [Parse Error] Semester: {e}")

    # --- Pola Keterampilan (Skills) ---
    try:
        # Menggunakan daftar judul bagian (header) yang lebih LENGKAP
        # dan daftar kata berhenti (stop_word) yang lebih BANYAK
        headers = r"(Keahlian|Keterampilan|Skill|Skills|Cakap dalam)"
        stop_words = r"(Sertifikat|Pendidikan|Pengalaman|Organisasi|Edukasi|EDUCATION|PROJECT|ACHIEVEMENTS|Leadership|Social Media|Bahasa|Language)"
        
        pola_keterampilan = fr"{headers}\s*[:\n](.*?)(?={stop_words})"
        match = re.search(pola_keterampilan, teks, re.IGNORECASE | re.DOTALL)
        
        if match:
            keterampilan_raw = match.group(2)
            # Membersihkan teks (hapus bullet points, strip spasi ekstra, hapus "Soft Skill", "Hard Skill")
            keterampilan_raw = re.sub(r"(Soft Skill|Hard Skill|Technical Skill|Tools)s*:.*?\n", "\n", keterampilan_raw, flags=re.IGNORECASE)
            keterampilan_list = [skill.strip() for skill in re.split(r"[\n\•\*-]", keterampilan_raw) if skill.strip() and len(skill.strip()) > 1]
            data["Keterampilan"] = ", ".join(keterampilan_list)
    except Exception as e:
        print(f"  [Parse Error] Keterampilan: {e}")

    # --- Pola Sertifikat ---
    try:
        # Menggunakan "ACHIEVEMENTS" sebagai kata kunci utama
        headers = r"(Sertifikat|Certificates|ACHIEVEMENTS|Penghargaan)"
        stop_words = r"(Pendidikan|Pengalaman|Organisasi|Hobi|Social Media|PROJECT|Language|Leadership)"

        pola_sertifikat = fr"{headers}\s*[:\n](.*?)(?={stop_words})"
        match = re.search(pola_sertifikat, teks, re.IGNORECASE | re.DOTALL)
        
        if match:
            sertifikat_raw = match.group(2)
            sertifikat_list = [cert.strip() for cert in re.split(r"[\n\•\*-]", sertifikat_raw) if cert.strip() and len(cert.strip()) > 5] # Filter baris pendek
            data["Sertifikat"] = ", ".join(sertifikat_list)
    except Exception as e:
        print(f"  [Parse Error] Sertifikat: {e}")

    return data

def main():
    """
    Fungsi utama untuk menjalankan seluruh proses.
    """
    folder_path = "list_CV" # Nama folder tempat Anda menyimpan CV
    file_excel_output = "hasil_ekstraksi_cv.xlsx"
    if file_excel_output.exists():
        os.remove(file_excel_output)
    
    # Daftar untuk menampung semua data dari semua CV
    data_semua_cv = []
    
    print(f"Mulai memproses file di folder: {folder_path}")

    # Loop melalui setiap file di dalam folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            print(f"Memproses: {filename}...")
            
            # 1. Ekstrak teks
            teks_cv = ekstrak_teks_dari_pdf(pdf_path)
            
            if teks_cv:
                # 2. Parse data
                data_hasil = parse_data_cv(teks_cv)
                # Tambahkan nama file agar kita tahu datanya dari CV mana
                data_hasil["Nama File"] = filename
                
                # 3. Tambahkan ke daftar utama
                data_semua_cv.append(data_hasil)
            else:
                print(f"Gagal mengekstrak teks dari {filename}.")

    # 4. Simpan ke Excel
    if data_semua_cv:
        df = pd.DataFrame(data_semua_cv)
        
        # Mengatur urutan kolom agar lebih rapi
        kolom_utama = ["Nama File", "IPK", "Jurusan", "Semester", "Keterampilan", "Sertifikat"]
        df = df[kolom_utama]
        
        df.to_excel(file_excel_output, index=False)
        print(f"\nSelesai! Data telah disimpan di {file_excel_output}")
    else:
        print("Tidak ada data CV yang berhasil diproses.")

# Menjalankan fungsi utama
if __name__ == "__main__":
    main()