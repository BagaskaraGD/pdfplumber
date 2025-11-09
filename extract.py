import os
import re
import pandas as pd
import pdfplumber
from typing import Dict, List, Optional, Any
import logging
from datetime import datetime

# ==============================
# Logging
# ==============================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ==============================
# Regex precompiled (lebih efisien)
# ==============================
RE_PREFIX_NUMBER_DOT   = re.compile(r'^\s*\d+\s*\.\s*')                     # "12." di awal
RE_CV_ANY               = re.compile(r'(?i)\bcurriculum\s+vitae\b|\bcv\b')  # CV/curriculum vitae
RE_SPLIT_AND_COLON      = re.compile(r'\s+(?:&|:)\s+')                      # pisah "Nama & ..."
RE_SPLIT_DASH           = re.compile(r'\s+-\s+')                            # pisah "Nama - ... "
RE_PARENS_NUM           = re.compile(r'\(\s*\d+\s*\)')                      # "(1)"
RE_NUMBERS              = re.compile(r'\d+')                                # angka sisa
RE_SPACES               = re.compile(r'\s+')                                # spasi ganda
RE_NON_ALPHA_SPACE      = re.compile(r'[^A-Za-z\s]')                        # hanya huruf & spasi
RE_IPK_LIST = [
    re.compile(p, re.IGNORECASE) for p in [
        r'\bIPK\s*[:\.\-]?\s*([0-4]\.[0-9]{2,3})\b',
        r'\bIpk\s*[:\.\-]?\s*([0-4]\.[0-9]{2,3})\b',
        r'\bGPA\s*[:\.\-]?\s*([0-4]\.[0-9]{2,3})\b',
        r'\bIndeks Prestasi\s*[:\.\-]?\s*([0-4]\.[0-9]{2,3})\b',
        r'\b([0-4]\.[0-9]{2,3})\s*(?:/|dari|out of)\s*4\b',
        r'\b([3-4]\.[0-9]{2})\b',
        r'\b[0-4]\.[0-9]{2}\b',
        r'IPK[^0-9]*([0-4]\.[0-9]{2})',
        r'GPA[^0-9]*([0-4]\.[0-9]{2})',
    ]
]
RE_JURUSAN_LIST = [
    re.compile(p, re.IGNORECASE) for p in [
        r'\b(S1|S2|S3|D1|D2|D3|D4)\s+([A-Za-zÀ-ÖØ-öø-ÿ\s&-]{3,})\b',
        r'\b(?:Jurusan|Program Studi|Prodi|Study Program|Department)[:\s]*([A-Za-zÀ-ÖØ-öø-ÿ\s&-]{3,}?)(?:\n|,|\.|$)',
        r'\b(?:Fakultas|Faculty)[:\s]*([A-Za-zÀ-ÖØ-öø-ÿ\s&-]{3,}?)(?:\n|,|\.|$)',
    ]
]
RE_SEMESTER_LIST = [
    re.compile(p, re.IGNORECASE) for p in [
        r'\bSemester\s*[:\.\-]?\s*([1-9]|1[0-2])\b',
        r'\bSem\s*[:\.\-]?\s*([1-9]|1[0-2])\b',
        r'\b(?:tahun|year)\s*[:\.\-]?\s*([1-9])\b',
        r'\b([1-9]|1[0-2])\s*(?:semester|sem)\b',
        r'\b([1-9])\s*tahun\b',
    ]
]

# ==============================
# Util: ambil nama dari filename
# ==============================
def clean_person_name_from_filename(filename: str) -> str:
    """
    Ambil nama dari pola filename CV dan bersihkan:
    - Hapus .pdf, angka urutan depan "1.", token 'CV/curriculum vitae', titik/underscore/dash.
    - Potong pada pemisah umum (' & ', ':', ' - ') → ambil sisi kiri.
    - Bersihkan angka/kurung sisa, rapikan spasi, Title Case.
    """
    s = os.path.splitext(os.path.basename(filename))[0]

    s = RE_PREFIX_NUMBER_DOT.sub('', s)
    s = s.replace('_', ' ').replace('-', ' ').replace('.', ' ')
    s = RE_CV_ANY.sub(' ', s)

    s = RE_SPLIT_AND_COLON.split(s)[0]
    s = RE_SPLIT_DASH.split(s)[0]

    s = RE_PARENS_NUM.sub(' ', s)
    s = RE_NUMBERS.sub(' ', s)

    s = RE_NON_ALPHA_SPACE.sub(' ', s)
    s = RE_SPACES.sub(' ', s).strip()

    tokens = s.split()
    if not tokens:
        return ""
    if len(tokens) > 6:   # batasi agar tidak kepanjangan
        tokens = tokens[:6]

    return ' '.join(w.capitalize() for w in tokens)

# ==============================
# Ekstraktor bidang selain nama
# ==============================
class CVDataExtractor:
    def __init__(self):
        # lookup cepat untuk skills
        self.skill_keywords: set = {
            'python','java','javascript','typescript','html','css','php','ruby','swift',
            'kotlin','c++','c#','go','react','vue','angular','nodejs','express','django','flask','laravel',
            'mysql','postgresql','mongodb','sqlite','oracle','redis',
            'flutter','react native','android studio','xcode',
            'aws','azure','gcp','docker','kubernetes',
            'photoshop','illustrator','figma','adobe xd','canva','coreldraw',
            'tableau','power bi','google analytics','spss',
            'linux','windows server','git','graphql','rest','microservices',
            'agile','scrum','jira','trello','notion',
        }

    # ---------- PDF ----------
    def extract_text_from_pdf(self, pdf_path: str) -> str:
        text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"

                    # tables → baris teks (opsional)
                    try:
                        for table in (page.extract_tables() or []):
                            for row in (table or []):
                                if row:
                                    text += " | ".join([str(c) if c else "" for c in row]) + "\n"
                    except Exception:
                        pass

                    # fallback halaman pertama
                    if not page_text and page_num == 0:
                        try:
                            alt = page.extract_text(x_tolerance=1, y_tolerance=1)
                            if alt: text += alt + "\n"
                        except Exception:
                            pass
                        try:
                            words = page.extract_words()
                            if words: text += " ".join([w["text"] for w in words]) + "\n"
                        except Exception:
                            pass
        except Exception as e:
            logger.error(f"Error reading PDF {pdf_path}: {e}")
        return text

    # ---------- Bidang lain ----------
    def extract_ipk(self, text: str) -> Optional[float]:
        for rex in RE_IPK_LIST:
            for match in rex.findall(text):
                try:
                    v = float(match)
                    if 2.0 <= v <= 4.0:
                        return v
                except Exception:
                    continue
        return None

    def extract_jurusan(self, text: str) -> Optional[str]:
        for rex in RE_JURUSAN_LIST:
            for match in rex.findall(text):
                if isinstance(match, tuple):
                    if len(match) >= 2:
                        edu = match[0].upper().strip()
                        major = re.sub(r'\s+', ' ', match[1]).strip()
                        if len(major) >= 3:
                            major_tc = ' '.join(w.capitalize() for w in major.split())
                            return f"{edu} {major_tc}"
                    else:
                        val = re.sub(r'\s+', ' ', match[0]).strip()
                        if len(val) >= 3:
                            return ' '.join(w.capitalize() for w in val.split())
                else:
                    val = re.sub(r'\s+', ' ', match).strip()
                    if len(val) >= 3:
                        return ' '.join(w.capitalize() for w in val.split())
        return None

    def extract_semester(self, text: str) -> Optional[int]:
        for rex in RE_SEMESTER_LIST:
            for match in rex.findall(text):
                try:
                    v = int(match)
                    if 1 <= v <= 12:
                        return v
                except Exception:
                    continue
        return None

    def extract_skills(self, text: str) -> List[str]:
        low = text.lower()
        found = [kw.title() for kw in self.skill_keywords if kw in low]
        return sorted(found)

    # ---------- Pipeline ----------
    def extract_from_cv(self, pdf_path: str) -> Dict[str, Any]:
        # nama dari filename (kita tidak keluarkan kolom filename lagi)
        nama = clean_person_name_from_filename(pdf_path)

        text = self.extract_text_from_pdf(pdf_path)
        if not text:
            logger.warning(f"No text extracted from {os.path.basename(pdf_path)}")
            return {
                "nama": nama,
                "ipk": None,
                "jurusan": None,
                "semester": None,
                "skills": "",
                "skill_count": 0,
                "extraction_status": "Failed - No text extracted",
                "extraction_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

        ipk = self.extract_ipk(text)
        jurusan = self.extract_jurusan(text)
        semester = self.extract_semester(text)
        skills = self.extract_skills(text)

        status = "Success"
        missing = []
        if ipk is None:     missing.append("IPK")
        if jurusan is None: missing.append("Jurusan")
        if missing:
            status = f"Partial - Missing: {', '.join(missing)}"

        return {
            "nama": nama,  # diambil dari filename
            "ipk": ipk,
            "jurusan": jurusan,
            "semester": semester,
            "skills": ", ".join(skills) if skills else "",
            "skill_count": len(skills),
            "extraction_status": status,
            "extraction_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

    def process_cv_folder(self, folder_path: str, output_excel: str = None) -> pd.DataFrame:
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"Folder {folder_path} not found")

        pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
        if not pdf_files:
            logger.warning(f"No PDF files found in {folder_path}")
            return pd.DataFrame()

        logger.info(f"Found {len(pdf_files)} PDF files to process")
        results = []
        for i, pdf_file in enumerate(pdf_files, 1):
            pdf_path = os.path.join(folder_path, pdf_file)
            logger.info(f"Processing {i}/{len(pdf_files)}: {pdf_file}")
            try:
                results.append(self.extract_from_cv(pdf_path))
            except Exception as e:
                logger.error(f"Error processing {pdf_file}: {e}")
                results.append({
                    "nama": clean_person_name_from_filename(pdf_file),
                    "ipk": None,
                    "jurusan": None,
                    "semester": None,
                    "skills": "",
                    "skill_count": 0,
                    "extraction_status": f"Error: {str(e)}",
                    "extraction_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                })

        df = pd.DataFrame(results)

        # Kolom final TANPA 'filename'
        column_order = [
            "nama", "ipk", "jurusan", "semester",
            "skills", "skill_count", "extraction_status", "extraction_date",
        ]
        for col in column_order:
            if col not in df.columns:
                df[col] = pd.NA
        df = df.reindex(columns=column_order)

        if output_excel:
            df.to_excel(output_excel, index=False, engine="openpyxl")
            logger.info(f"Results saved to {output_excel}")

        return df

# ==============================
# Main
# ==============================
def main():
    cv_folder = "cv"
    output_excel = "cv_extracted_results_improved.xlsx"

    extractor = CVDataExtractor()
    try:
        logger.info("Starting CV data extraction...")
        df = extractor.process_cv_folder(cv_folder, output_excel)

        print("\n" + "="*50)
        print("EXTRACTION SUMMARY")
        print("="*50)
        print(f"Total CVs processed: {len(df)}")
        print(f"Successful: {len(df[df['extraction_status'] == 'Success'])}")
        print(f"Partial: {len(df[df['extraction_status'].str.contains('Partial', na=False)])}")
        print(f"Failed/Errors: {len(df[df['extraction_status'].str.contains('Failed|Error', na=False)])}")

        print("\nSample Results:")
        sample_cols = ["nama","ipk","jurusan","semester","skill_count"]
        print(df[sample_cols].head())

        print(f"\nResults saved to: {output_excel}")
    except Exception as e:
        logger.error(f"Error in main execution: {e}")
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
