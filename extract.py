# extract.py
import os, re, logging
import pandas as pd
import pdfplumber
from typing import Dict, List, Optional, Any
from datetime import datetime

# ==============================
# KONFIGURASI
# ==============================
USE_OCR = True  # aktifkan OCR fallback untuk PDF gambar
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # ubah bila perlu

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ==============================
# KAMUS / REGEX
# ==============================
MAJOR_NORMALIZE = {
    'information system': 'Sistem Informasi',
    'information systems': 'Sistem Informasi',
    'system information': 'Sistem Informasi',
    'sistem informasi': 'Sistem Informasi',

    'informatics engineering': 'Teknik Informatika',
    'teknik informatika': 'Teknik Informatika',
    'informatika': 'Teknik Informatika',

    'computer engineering': 'Teknik Komputer',
    'teknik komputer': 'Teknik Komputer',

    'dkv': 'Desain Komunikasi Visual',
    'desain komunikasi visual': 'Desain Komunikasi Visual',
    'visual communication design': 'Desain Komunikasi Visual',

    'desain produk': 'Desain Produk',
    'product design': 'Desain Produk',

    'manajemen bisnis': 'Manajemen Bisnis',
    'business management': 'Manajemen Bisnis',

    'akuntansi': 'Akuntansi',
    'akuntasi': 'Akuntansi',
    'accounting': 'Akuntansi',
}

ABBREV_MAP = {'si':'Sistem Informasi','ti':'Teknik Informatika','tk':'Teknik Komputer','dkv':'Desain Komunikasi Visual'}

RE_DEGREE = re.compile(r'(?i)\b(S1|S2|S3|D1|D2|D3|D4)\b')
RE_IPK_NEAR = [
    re.compile(r'(?i)\bIPK\s*[:\-]?\s*([0-4][\.,][0-9]{2,3})'),
    re.compile(r'(?i)\bGPA\s*[:\-]?\s*([0-4][\.,][0-9]{2,3})'),
    re.compile(r'(?i)\bIndeks\s+Prestasi(?:\s*Kumulatif)?\s*[:\-]?\s*([0-4][\.,][0-9]{2,3})'),
    re.compile(r'(?i)([0-4][\.,][0-9]{2,3})\s*(?:/|dari|out of)\s*4'),
]
RE_SEM_NEAR = [
    re.compile(r'(?i)\bsemester\s*[:\-]?\s*(1[0-2]|[1-9])\b'),
    re.compile(r'(?i)\bsem\s*[:\-]?\s*(1[0-2]|[1-9])\b'),
    re.compile(r'(?i)\b(1[0-2]|[1-9])\b\s*(?:th|tahun)?\s*(?:semester|sem)\b'),
]

# nama dari filename (tanpa fallback jurusan)
RE_PREFIX_NUMBER_DOT   = re.compile(r'^\s*\d+\s*\.\s*')
RE_CV_ANY              = re.compile(r'(?i)\bcurriculum\s+vitae\b|\bcv\b')
RE_SPACES              = re.compile(r'\s+')
def clean_person_name_from_filename(filename: str) -> str:
    s = os.path.splitext(os.path.basename(filename))[0]
    s = RE_PREFIX_NUMBER_DOT.sub('', s)
    s = s.replace('_',' ').replace('-',' ').replace('.',' ')
    s = RE_CV_ANY.sub(' ', s)
    s = re.sub(r'\(\s*\d+\s*\)', ' ', s)
    s = re.sub(r'\d+', ' ', s)
    s = re.sub(r'[^A-Za-z\s]', ' ', s)
    s = RE_SPACES.sub(' ', s).strip()
    toks = s.split()
    if not toks: return ""
    if len(toks) > 6: toks = toks[:6]
    return ' '.join(w.capitalize() for w in toks)

# Skills
SKILL_SYNONYM = {'rest': 'REST','rest api': 'REST','node': 'Nodejs','node.js': 'Nodejs',
                 'adobe xd': 'Adobe Xd','c plus plus': 'C++','c sharp': 'C#'}
SKILL_SET = {
    'python','java','javascript','typescript','html','css','php','ruby','swift',
    'kotlin','c++','c#','go','react','vue','angular','nodejs','express','django','flask','laravel',
    'mysql','postgresql','mongodb','sqlite','oracle','redis','flutter','react native','android studio','xcode',
    'aws','azure','gcp','docker','kubernetes','jenkins','gitlab','github','bitbucket','terraform','ansible',
    'photoshop','illustrator','figma','adobe xd','canva','coreldraw','premiere','after effects',
    'tableau','power bi','google analytics','spss','linux','windows server','macos','git','graphql','rest',
    'microservices','agile','scrum','jira','trello','notion'
}
SKILL_SECTIONS = [
    re.compile(r'(?is)(?:keahlian|keterampilan|kemampuan|skill[s]?|technical\s*skill[s]?)[\s:]*([\s\S]{0,600})'),
    re.compile(r'(?is)(?:software|tool[s]?|teknologi|bahasa\s*pemrograman|programming\s*language[s]?)[\s:]*([\s\S]{0,400})'),
]

# ==============================
# OCR fallback
# ==============================
def try_ocr_pdf_pages(pdf_path: str) -> str:
    if not USE_OCR:
        return ""
    try:
        import pytesseract
        from PIL import Image
        if TESSERACT_CMD:
            import pytesseract as _pt
            _pt.pytesseract.tesseract_cmd = TESSERACT_CMD
    except Exception:
        return ""
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                try:
                    img = page.to_image(resolution=300).original
                    t = pytesseract.image_to_string(img, lang='eng+ind')
                    if t: text += t + "\n"
                except Exception:
                    continue
    except Exception:
        return ""
    return text

# ==============================
# Ekstraktor
# ==============================
class CVDataExtractor:
    # ---- text ----
    def extract_text_from_pdf(self, pdf_path: str) -> str:
        text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    t = page.extract_text()
                    if t: text += t + "\n"
                    try:
                        for table in (page.extract_tables() or []):
                            for row in (table or []):
                                if row:
                                    text += " | ".join([str(c) if c else "" for c in row]) + "\n"
                    except Exception:
                        pass
                    if not t and i == 0:
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

        if not text:
            ocr_text = try_ocr_pdf_pages(pdf_path)
            if ocr_text:
                logger.info(f"OCR used for {os.path.basename(pdf_path)}")
                text = ocr_text
        return text

    # ---- ipk ----
    def extract_ipk(self, text: str) -> Optional[float]:
        for rex in RE_IPK_NEAR:
            for m in rex.findall(text):
                val = m if isinstance(m, str) else m[0]
                val = val.replace(',', '.')
                try:
                    v = float(val)
                    if 2.0 <= v <= 4.0:
                        return round(v, 2)
                except Exception:
                    continue
        return None

    # ---- jurusan (HANYA dari MAJOR_NORMALIZE di TEKS) ----
    def extract_jurusan(self, text: str) -> Optional[str]:
        low = text.lower()

        # 1) degree +/- 40 char dari keyword jurusan
        for key, target in MAJOR_NORMALIZE.items():
            # degree sebelum keyword
            if re.search(rf'(?i)\b(S1|S2|S3|D1|D2|D3|D4)\b.{0,40}\b{re.escape(key)}\b', low):
                deg = re.search(r'(?i)\b(S1|S2|S3|D1|D2|D3|D4)\b.{0,40}\b'+re.escape(key)+r'\b', low).group(1).upper()
                return f"{deg} {target}"
            # degree sesudah keyword
            if re.search(rf'(?i)\b{re.escape(key)}\b.{0,40}\b(S1|S2|S3|D1|D2|D3|D4)\b', low):
                deg = re.search(rf'(?i)\b{re.escape(key)}\b.{0,40}\b(S1|S2|S3|D1|D2|D3|D4)\b', low).group(1).upper()
                return f"{deg} {target}"

        # 2) singkatan degree + singkatan mayor (S1 SI, D3 TI, dll)
        m = re.search(r'(?i)\b(S1|S2|S3|D1|D2|D3|D4)\s+(SI|TI|TK|DKV)\b', low)
        if m:
            return f"{m.group(1).upper()} {ABBREV_MAP[m.group(2).lower()]}"

        # 3) keyword jurusan saja (tanpa degree) → tetap kembalikan normalized,
        #    karena keyword-nya memang ada di TEKS (sesuai permintaan)
        for key, target in MAJOR_NORMALIZE.items():
            if re.search(rf'(?i)\b{re.escape(key)}\b', low):
                # opsional: tambahkan default S1; kalau tak ingin, kembalikan target saja
                return f"S1 {target}"

        # 4) tidak ditemukan satupun keyword → kosong
        return None

    # ---- semester ----
    def extract_semester(self, text: str) -> Optional[int]:
        for rex in RE_SEM_NEAR:
            for m in rex.findall(text):
                try:
                    v = int(m if isinstance(m, str) else m[0])
                    if 1 <= v <= 12: return v
                except Exception:
                    continue
        return None

    # ---- skills ----
    def extract_skills(self, text: str) -> List[str]:
        low = text.lower()
        zone = ""
        for rex in SKILL_SECTIONS:
            for m in rex.findall(text):
                zone += " " + m
        scan = zone.lower() if zone.strip() else low

        for syn, canon in SKILL_SYNONYM.items():
            scan = scan.replace(syn, canon.lower())

        found = []
        for kw in SKILL_SET:
            if kw in scan:
                name = kw.upper() if kw in {'c++','c#','aws','gcp','css','html','sql'} else kw.title()
                if name.lower() == 'rest': name = 'REST'
                found.append(name)
        return sorted(set(found))

    # ---- pipeline ----
    def extract_from_cv(self, pdf_path: str) -> Dict[str, Any]:
        filename = os.path.basename(pdf_path)
        nama = clean_person_name_from_filename(filename)
        text = self.extract_text_from_pdf(pdf_path)

        if not text:
            logger.warning(f"No text extracted from {filename}")
            return {
                "nama": nama,
                "ipk": None,
                "jurusan": None,          # <- tidak isi dari filename
                "semester": None,
                "skills": "",
                "skill_count": 0,
                "extraction_status": "Failed - No text extracted",
                "extraction_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

        ipk = self.extract_ipk(text)
        jurusan = self.extract_jurusan(text)   # <- hanya dari teks
        semester = self.extract_semester(text)
        skills = self.extract_skills(text)

        status = "Success"
        miss = []
        if ipk is None: miss.append("IPK")
        if jurusan is None: miss.append("Jurusan")
        if miss: status = f"Partial - Missing: {', '.join(miss)}"

        return {
            "nama": nama,
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
        pdfs = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
        if not pdfs:
            logger.warning(f"No PDF files found in {folder_path}")
            return pd.DataFrame()

        logger.info(f"Found {len(pdfs)} PDF files to process")
        rows = []
        for i, f in enumerate(pdfs, 1):
            p = os.path.join(folder_path, f)
            logger.info(f"Processing {i}/{len(pdfs)}: {f}")
            try:
                rows.append(self.extract_from_cv(p))
            except Exception as e:
                logger.error(f"Error processing {f}: {e}")
                rows.append({
                    "nama": clean_person_name_from_filename(f),
                    "ipk": None, "jurusan": None, "semester": None,
                    "skills": "", "skill_count": 0,
                    "extraction_status": f"Error: {str(e)}",
                    "extraction_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                })

        df = pd.DataFrame(rows)
        cols = ["nama","ipk","jurusan","semester","skills","skill_count","extraction_status","extraction_date"]
        for c in cols:
            if c not in df.columns: df[c] = pd.NA
        df = df.reindex(columns=cols)

        if output_excel:
            df.to_excel(output_excel, index=False, engine="openpyxl")
            logger.info(f"Results saved to {output_excel}")
        return df

# ==============================
# MAIN
# ==============================
def main():
    cv_folder = "cv"
    output_excel = "cv_extracted_results_improved.xlsx"
    extractor = CVDataExtractor()
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
    print(df[['nama','ipk','jurusan','semester','skills','skill_count']].head())
    print(f"\nResults saved to: {output_excel}")

if __name__ == "__main__":
    main()
