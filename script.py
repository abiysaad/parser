from pathlib import Path
import pdfplumber
import pandas as pd
import re
from tqdm import tqdm
from datetime import datetime
from openpyxl.styles import numbers

# === SETUP ===
BASE_DIR = Path(__file__).resolve().parent
PDF_DIR = BASE_DIR / 'pdfs'
OUTPUT_DIR = BASE_DIR / 'output'
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
TIMESTAMP = datetime.now().strftime('%Y-%m-%d_%H-%M')
OUTPUT_FILE = OUTPUT_DIR / f'data_{TIMESTAMP}.xlsx'

# === BL PARSER ===
def parse_bl(pdf_path, subfolder):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ' '.join(page.extract_text() or '' for page in pdf.pages)
    except:
        print(f"⚠️  Gagal baca teks dari {pdf_path.name}")
        return None

    text = text.replace('\n', ' ')

    no_bl = re.search(r'\b\d{3}-\d{2}/[A-Z]+-[A-Z]+/[A-Z]+/\d{4}\b', text)
    tanggal = re.search(
        r'\b\d{1,2} (January|February|March|April|May|June|July|August|September|October|November|December|'
        r'Januari|Februari|Maret|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember) \d{4}\b',
        text, re.IGNORECASE
    )

    qty_matches = re.findall(r'([\d.,]+)\s*WMT', text, re.IGNORECASE)
    qty = qty_matches[-1] if qty_matches else ''
    qty = qty.replace('WMT', '').replace('%', '').strip()

    return {
        "Subfolder": subfolder,
        "File": pdf_path.name,
        "Doc_Type": "BL",
        "No_BL": no_bl.group(0) if no_bl else '',
        "Tanggal_BL": tanggal.group(0) if tanggal else '',
        "Qty_WMT": qty,
        "No_SI": "",
        "Nama_TB_BG": "",
        "Laycan": ""
    }

# === SI PARSER ===
def parse_si(pdf_path, subfolder):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ' '.join(page.extract_text() or '' for page in pdf.pages)
    except:
        print(f"⚠️  Gagal baca teks dari {pdf_path.name}")
        return None

    text = text.replace('\n', ' ')

    no_si = re.search(r'SHIPPING INSTRUCTION\s+([A-Z0-9/\-]+)', text)
    laycan = re.search(r'LAYCAN DATE\s*:\s*(\d{1,2} ?[-–] ?\d{1,2} [A-Z]+ \d{4})', text, re.IGNORECASE)

    filename = pdf_path.stem
    nama_tb_bg = ''
    match_filename = re.search(r'(TB\..+?)(?=\s+\d{1,2}\s)', filename)
    if match_filename:
        nama_tb_bg = match_filename.group(1).strip()

    return {
        "Subfolder": subfolder,
        "File": pdf_path.name,
        "Doc_Type": "SI",
        "No_BL": "",
        "Tanggal_BL": "",
        "Qty_WMT": "",
        "No_SI": no_si.group(1) if no_si else '',
        "Nama_TB_BG": nama_tb_bg,
        "Laycan": laycan.group(1).strip() if laycan else ''
    }

# === SCAN SUBFOLDERS ===
def process_all():
    data = []
    for subfolder in PDF_DIR.iterdir():
        if subfolder.is_dir():
            print(f"\n📂 Processing: {subfolder.name}")
            for pdf_path in subfolder.glob("*.pdf"):
                name_lower = pdf_path.name.lower()

                # skip draught survey (file like SI. 27.pdf, SI.1.pdf, etc)
                if re.search(r'\bsi\.\s*\d+', name_lower):
                    print(f"⏭️  Skip draught survey: {pdf_path.name}")
                    continue

                if "bl" in name_lower:
                    parsed = parse_bl(pdf_path, subfolder.name)
                elif "si" in name_lower or "shipping" in name_lower:
                    parsed = parse_si(pdf_path, subfolder.name)
                else:
                    print(f"⚠️  Skip (unknown type): {pdf_path.name}")
                    parsed = None

                if parsed:
                    data.append(parsed)
    return pd.DataFrame(data)

# === MAIN ===
def main():
    print("📦 SHIPPING PARSER (BL + SI, Multi-Subfolder Mode, Skip Draught Survey)")
    df = process_all()

    if df.empty:
        print("\n❌ No data found in any subfolder.")
        return

    df['Tanggal_BL'] = pd.to_datetime(df['Tanggal_BL'], dayfirst=True, errors='coerce')
    df = df.sort_values(by=['Subfolder', 'Doc_Type'])

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n✅ Output file created: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
