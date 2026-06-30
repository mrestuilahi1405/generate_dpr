import pandas as pd
from docxtpl import DocxTemplate
import os
import subprocess
import sys
import re

def get_libreoffice_command():
    if sys.platform == 'win32':
        return r"C:\Program Files\LibreOffice\program\soffice.exe"
    elif sys.platform == 'darwin':
        return "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    else:
        return "libreoffice"

def generate_documents_from_excel(template_path: str, excel_path: str, output_dir: str, 
                                  filename_column: str, file_prefix: str, 
                                  data_sheet: str, meta_sheet: str = None):
    """
    Core engine universal yang mendukung nama sheet dinamis dan jumlah sheet banyak.
    """
    if not os.path.exists(template_path) or not os.path.exists(excel_path):
        raise FileNotFoundError("Template atau file Excel tidak ditemukan.")

    os.makedirs(output_dir, exist_ok=True)
    libreoffice_cmd = get_libreoffice_command()
    
    if sys.platform == 'win32' and not os.path.exists(libreoffice_cmd):
        raise FileNotFoundError(f"LibreOffice tidak ditemukan di path: {libreoffice_cmd}")

    # 1. Baca Sheet Metadata (Jika dipilih oleh user)
    metadata = {}
    if meta_sheet and meta_sheet != "Tidak Ada (Hanya Data Baris)":
        try:
            df_meta = pd.read_excel(excel_path, sheet_name=meta_sheet)
            if df_meta.shape[1] >= 2:
                # Ambil kolom 0 sebagai kunci, kolom 1 sebagai nilai (posisional, nama kolom bebas)
                metadata = dict(zip(df_meta.iloc[:, 0], df_meta.iloc[:, 1]))
            else:
                raise ValueError("Sheet Metadata minimal harus memiliki 2 kolom (Kunci & Nilai).")
        except Exception as e:
            raise ValueError(f"Gagal memproses sheet metadata '{meta_sheet}': {e}")

    # 2. Baca Sheet Data Utama
    try:
        df_data = pd.read_excel(excel_path, sheet_name=data_sheet).fillna("")
        records = df_data.to_dict(orient='records')
    except Exception as e:
        raise ValueError(f"Gagal membaca sheet data '{data_sheet}': {e}")

    doc = DocxTemplate(template_path)
    berkas_sukses = 0
    
    # 3. Looping Pengerjaan
    for index, record in enumerate(records):
        try:
            # Gabungkan metadata global (jaku ada) dengan data baris
            context = {**metadata, **record}
            doc.render(context)
            
            raw_identifier = str(record.get(filename_column, f"berkas_{index}")).strip()
            clean_identifier = re.sub(r'[\\/*?:"<>|]', "", raw_identifier).replace(" ", "_")
            filename_base = f"{file_prefix}_{clean_identifier}" if file_prefix else clean_identifier
            
            temp_docx_path = os.path.join(output_dir, f"{filename_base}.docx")
            output_pdf_path = os.path.join(output_dir, f"{filename_base}.pdf")
            
            doc.save(temp_docx_path)
            
            subprocess.run([
                libreoffice_cmd, '--headless', '--convert-to', 'pdf',
                temp_docx_path, '--outdir', output_dir
            ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            if os.path.exists(temp_docx_path):
                os.remove(temp_docx_path)
            
            berkas_sukses += 1
            
        except Exception as e:
            print(f"[ERROR] Gagal memproses baris {index + 1}: {e}")

    return berkas_sukses