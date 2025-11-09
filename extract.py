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
# Regex precompiled & kamus
# ==============================
RE_PREFIX_NUMBER_DOT   = re.compile(r'^\s*\d+\s*\.\s*')
RE_CV_ANY              = re.compile(r'(?i)\bcurriculum\s+vitae\b|\bcv\b')
RE_SPLIT_AND_COLON     = re.compile(r'\s+(?:&|:)\s+')
RE_SPLIT_DASH          = re.compile(r'\s+-\s+')
RE_PARENS_NUM          = re.compile(r'\(\s*\d+\s*\)')
RE_NUMBERS             = re.compile(r'\d+')
RE_SPACES              = re.compile(r'\s+')
RE_NON_ALPHA_SPACE     = re.compile(r'[^A-Za-z\s]')

# IPK: terima titik atau koma; wajib dekat kata kunci
RE_IPK_NEAR = [
    re.compile(r'(?i)\bIPK\s*[:\-]?\s*([0-4][\.,][0-9]{2,3})'),
    re.compile(r'(?i)\bGPA\s*[:\-]?\s*([0-4][\.,][0-9]{2,3})'),
    re.compile(r'(?i)\bIndeks\s+Prestasi(?:\s*Kumulatif)?\s*[:\-]?\s*([0-4][\.,][0-9]{2,3})'),
    re.compile(r'(?i)([0-4][\.,][0-9]{2,3})\s*(?:/|dari|out of)\s*4'),
]

# Jurusan
RE_JURUSAN_LABEL = [
    re.compile(r'(?i)\b(?:Jurusan|Program\s*Studi|Prodi|Study\s*Program|Department)\s*[:\-]?\s*([^\n,\.]{3,100})'),
    re.compile(r'(?i)\b(S1|S2|S3|D1|D2|D3|D4)\s+([A-Za-zÀ-ÖØ-öø-ÿ\s&\-/]{3,100})'),
]
STOP_TOKENS_TAIL = {
    'universitas','university','dinamika','surabaya','semester','agustus','juli','juni','mei','oktober',
    'november','desember','september','maret','april','tahun','dengan','minat','kuat','bidang','yang',
    'sekarang','sedang','melanjutkan','studi','jujur','juara','disiplin','smk','smkn'
}
# peta sinonim jurusan
MAJOR_NORMALIZE = {
    'information system': 'Sistem Informasi',
    'information systems': 'Sistem Informasi',
    'sistem informasi': 'Sistem Informasi',
    'teknik informatika': 'Teknik Informatika',
    'informatika': 'Teknik Informatika',
    'computer engineering': 'Teknik Komputer',
    'teknik komputer': 'Teknik Komputer',
    'dkv': 'Desain Komunikasi Visual',
    'desain komunikasi visual': 'Desain Komunikasi Visual',
    'manajemen bisnis': 'Manajemen Bisnis',
    'teknik elektro': 'Teknik Elektro',
    'teknik mesin': 'Teknik Mesin',
    'teknik sipil': 'Teknik Sipil',
    'arsitektur': 'Arsitektur',
}
# kata SMA yang sering muncul—abaikan sebagai jurusan S1
SMA_TRACK = {'ipa','ips','bahasa'}

# Semester: angka 1–12 di dekat kata “semester/sem”
RE_SEM_NEAR = [
    re.compile(r'(?i)\bsemester\s*[:\-]?\s*(1[0-2]|[1-9])\b'),
    re.compile(r'(?i)\bsem\s*[:\-]?\s*(1[0-2]|[1-9])\b'),
    # dua arah: "6 (semester)" atau "semester ke 6"
    re.compile(r'(?i)\b(1[0-2]|[1-9])\b\s*(?:th|tahun)?\s*(?:semester|sem)\b'),
]

# Skills
SKILL_SYNONYM = {
    'rest': 'REST',
    'rest api': 'REST',
    'node': 'Nodejs',
    'node.js': 'Nodejs',
    'adobe xd': 'Adobe Xd',
    'c plus plus': 'C++',
    'c sharp': 'C#',
}
SKILL_SET = {
    'python','java','javascript','typescript','html','css','php','ruby','swift',
    'kotlin','c++','c#','go','react','vue','angular','nodejs','express','django','flask','laravel',
    'mysql','postgresql','mongodb','sqlite','oracle','redis',
    'flutter','react native','android studio','xcode',
    'aws','azure','gcp','docker','kubernetes','jenkins',
    'gitlab','github','bitbucket','terraform','ansible',
    'photoshop','illustrator','figma','adobe xd','canva','coreldraw','premiere','after effects',
    'tableau','power bi','google analytics','spss',
    'linux','windows server','macos','git','graphql','rest','microservices',
    'agile','scrum','jira','trello','notion'
}
SKILL_SECTIONS = [
    re.compile(r'(?is)(?:keahlian|keterampilan|kemampuan|skill[s]?|technical\s*skill[s]?)[\s:]*([\s\S]{0,600})'),
    re.compile(r'(?is)(?:software|tool[s]?|teknologi|bahasa\s*pemrograman|programming\s*language[s]?)[\s:]*([\s\S]{0,400})'),
]

# ==============================
# Util: ambil nama dari filename
# ==============================
def clean_person_name_from_filename(filename: str) -> str:
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
    if len(tokens) > 6:
        tokens = tokens[:6]
    return ' '.join(w.capitalize() for w in tokens)

def _title(s: str) -> str:
    return ' '.join(w.capitalize() for w in s.split())

# ==============================
# Ekstraktor
# ==============================
class CVDataExtractor:
    def __init__(self):
        self.skill_set = set(SKILL_SET)

    # ---- PDF text ----
    def extract_text_from_pdf(self, pdf_path: str) -> str:
        text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    t = page.extract_text()
                    if t:
                        text += t + "\n"
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
        return text

    # ---- IPK ----
    def extract_ipk(self, text: str) -> Optional[float]:
        for rex in RE_IPK_NEAR:
            for m in rex.findall(text):
                val = m if isinstance(m, str) else m[0]
                val = val.replace(',', '.')
                try:
                    ipk = float(val)
                    if 2.0 <= ipk <= 4.0:
                        return round(ipk, 2)
                except Exception:
                    continue
        return None

    # ---- Jurusan ----
    def _normalize_major_phrase(self, phrase: str) -> Optional[str]:
        # bersihkan ekor kalimat setelah stop token
        toks = []
        for w in phrase.split():
            lw = w.lower()
            if lw in STOP_TOKENS_TAIL:
                break
            toks.append(w)
        phrase = ' '.join(toks).strip()
        if not phrase:
            return None
        # ambil hanya 2–6 token pertama supaya tidak kepanjangan
        ptoks = phrase.split()
        ptoks = ptoks[:6]
        phrase = ' '.join(ptoks).strip()

        low = phrase.lower()
        # ganti sinonim
        for k, v in MAJOR_NORMALIZE.items():
            if k in low:
                return v

        # singkatan umum
        if low in MAJOR_NORMALIZE:
            return MAJOR_NORMALIZE[low]

        # “Sistem Informasi”, “Teknik Komputer”, dst (title case aman)
        phrase_tc = _title(phrase)
        # buang jejak SMA
        if phrase_tc.lower() in SMA_TRACK:
            return None
        return phrase_tc

    def extract_jurusan(self, text: str) -> Optional[str]:
        for rex in RE_JURUSAN_LABEL:
            for match in rex.findall(text):
                if isinstance(match, tuple):
                    jenjang, major = match[0], match[1]
                    norm = self._normalize_major_phrase(major)
                    if norm:
                        return f"{jenjang.upper()} {norm}"
                else:
                    norm = self._normalize_major_phrase(match)
                    if norm:
                        return norm
        return None

    # ---- Semester ----
    def extract_semester(self, text: str) -> Optional[int]:
        for rex in RE_SEM_NEAR:
            for match in rex.findall(text):
                try:
                    sem = int(match if isinstance(match, str) else match[0])
                    if 1 <= sem <= 12:
                        return sem
                except Exception:
                    continue
        return None

    # ---- Skills ----
    def _collect_skills_from_text(self, text: str) -> List[str]:
        low = text.lower()

        # 1) coba ambil dari section khusus
        section_text = ""
        for rex in SKILL_SECTIONS:
            for m in rex.findall(text):
                section_text += " " + m

        scan_zone = section_text if section_text.strip() else low

        found = set()
        # normalisasi sinonim ringan
        z = scan_zone.lower()
        for syn, canon in SKILL_SYNONYM.items():
            z = z.replace(syn, canon.lower())

        for kw in self.skill_set:
            if kw in z:
                # kapitalisasi yang pas
                name = kw.upper() if kw in {'c++','c#','aws','gcp','css','html','sql'} else kw.title()
                if name.lower() == 'rest':
                    name = 'REST'
                found.add(name)

        return sorted(found)

    def extract_skills(self, text: str) -> List[str]:
        return self._collect_skills_from_text(text)

    # ---- Pipeline ----
    def extract_from_cv(self, pdf_path: str) -> Dict[str, Any]:
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
        print(df[['nama','ipk','jurusan','semester','skills','skill_count']].head())

        print(f"\nResults saved to: {output_excel}")
    except Exception as e:
        logger.error(f"Error in main execution: {e}")
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
