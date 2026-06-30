import pandas as pd
from docxtpl import DocxTemplate
import os
import subprocess
import sys

def get_libreoffice_command():
    """
    Fungsi pembantu untuk mendapatkan command eksekusi LibreOffice
    berdasarkan sistem operasi yang sedang digunakan.
    """
    if sys.platform == 'win32':
        # Path instalasi standar LibreOffice di Windows. 
        # (Ubah path ini jika LibreOffice Anda terinstal di direktori lain)
        return r"C:\Program Files\LibreOffice\program\soffice.exe"
    elif sys.platform == 'darwin':
        # Path instalasi standar di MacOS
        return "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    else:
        # Command standar untuk Linux (dan nantinya untuk di dalam Docker/Hosting)
        return "libreoffice"

def generate_documents_from_excel(template_path: str, excel_path: str, output_dir: str):
    """
    Membaca data dari Excel (Opsi 1: Sheet Metadata & Peserta) 
    dan menghasilkan dokumen PDF menggunakan LibreOffice Headless.
    """
    # Validasi file
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template tidak ditemukan di: {template_path}")
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"File data tidak ditemukan di: {excel_path}")

    # Buat folder output jika belum ada
    os.makedirs(output_dir, exist_ok=True)

    # 1. Baca Sheet 'Metadata' (Untuk data statis: nomor_spd, tanggal_spd, dll)
    try:
        df_meta = pd.read_excel(excel_path, sheet_name="Metadata")
        # Mengubah kolom 'kunci' dan 'nilai' menjadi dictionary
        metadata = dict(zip(df_meta['Key'], df_meta['Value']))
    except Exception as e:
        raise ValueError(f"Gagal membaca sheet 'Metadata'. Pastikan kolom 'kunci' dan 'nilai' ada. Error: {e}")

    # 2. Baca Sheet 'Peserta' (Untuk data tabular per baris)
    try:
        df_peserta = pd.read_excel(excel_path, sheet_name="Peserta").fillna("")
        records = df_peserta.to_dict(orient='records')
    except Exception as e:
        raise ValueError(f"Gagal membaca sheet 'Peserta'. Error: {e}")

    # Load template Word
    doc = DocxTemplate(template_path)
    berkas_sukses = 0
    libreoffice_cmd = get_libreoffice_command()
    
    # Validasi keberadaan LibreOffice (khusus Windows)
    if sys.platform == 'win32' and not os.path.exists(libreoffice_cmd):
        raise FileNotFoundError(f"LibreOffice tidak ditemukan di path: {libreoffice_cmd}. Pastikan LibreOffice terinstal.")

    # 3. Proses looping per baris data peserta
    for index, record in enumerate(records):
        try:
            # Gabungkan metadata global dengan data spesifik peserta ini
            context = {**metadata, **record}
            
            # Render context ke dalam template Word
            doc.render(context)
            
            # Gunakan kolom 'nama_pelaksana_nip' dari sheet Peserta sebagai nama file
            nama_pelaksana_nip = str(record.get('nama_pelaksana_nip', f'berkas_{index}')).strip()
            identifier = nama_pelaksana_nip.rstrip(".")
            
            temp_docx_path = os.path.join(output_dir, f"DPR_{identifier}.docx")
            output_pdf_path = os.path.join(output_dir, f"DPR_{identifier}.pdf")
            
            # Simpan hasil render ke format Word terlebih dahulu
            doc.save(temp_docx_path)
            
            # Eksekusi konversi ke PDF menggunakan LibreOffice Headless
            subprocess.run([
                libreoffice_cmd,
                '--headless',
                '--convert-to',
                'pdf',
                temp_docx_path,
                '--outdir',
                output_dir
            ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            # Hapus file Word sementara
            if os.path.exists(temp_docx_path):
                os.remove(temp_docx_path)
            
            berkas_sukses += 1
            print(f"[SUCCESS] Berhasil membuat: {output_pdf_path}")
            
        except subprocess.CalledProcessError as e:
            print(f"[ERROR] Gagal mengonversi baris {index + 1} ke PDF. Pastikan LibreOffice tidak sedang terbuka. Error: {e}")
        except Exception as e:
            print(f"[ERROR] Gagal memproses baris {index + 1} pada sheet Peserta: {e}")

    print(f"\nSelesai! {berkas_sukses} dari {len(records)} dokumen PDF berhasil dibuat.")

# Blok eksekusi
if __name__ == "__main__":
    TEMPLATE_FILE = "template_dpr.docx"  # Sesuaikan dengan nama template Word Anda
    EXCEL_FILE = "data_dpr_juni_2026.xlsx" # Sesuaikan dengan nama file Excel Anda
    OUTPUT_FOLDER = "hasil_generate_juni_2026"
    
    generate_documents_from_excel(TEMPLATE_FILE, EXCEL_FILE, OUTPUT_FOLDER)