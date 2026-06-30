import streamlit as st
import os
import tempfile
import shutil
import pandas as pd
import subprocess
from core_engine import generate_documents_from_excel

# ==========================================
# FUNGSI INJEKSI FONT KE SERVER LINUX
# ==========================================
def setup_custom_fonts():
    # Folder tempat font disimpan di repositori GitHub
    font_source = "fonts"
    # Direktori font global untuk user di Linux
    font_dest = os.path.expanduser("~/.fonts")
    
    # Cek apakah folder 'fonts' ada di repositori
    if os.path.exists(font_source):
        os.makedirs(font_dest, exist_ok=True)
        # Cek apakah Arial sudah terpasang untuk menghindari proses berulang
        if not os.path.exists(os.path.join(font_dest, "arial.ttf")):
            for font_file in os.listdir(font_source):
                if font_file.lower().endswith(".ttf"):
                    shutil.copy(os.path.join(font_source, font_file), font_dest)
            
            # Beritahu OS Linux untuk merefresh daftar font
            try:
                subprocess.run(['fc-cache', '-f', '-v'], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                print("Custom fonts (Arial) berhasil diinstal di server!")
            except Exception as e:
                print(f"Gagal refresh font cache: {e}")

# Eksekusi injeksi sebelum Streamlit merender halaman
setup_custom_fonts()
# ==========================================

# Sisa kode Streamlit Anda di bawah ini
st.set_page_config(page_title="Ultra Universal Doc Gen", layout="centered")

st.title("Generator Berkas Otomatis (Multi-Sheet Mode)")
st.markdown("Unggah template `.docx` dan berkas `.xlsx` dengan struktur sheet mana saja.")

template_file = st.file_uploader("1. Unggah Template Word (.docx)", type=["docx"])
excel_file = st.file_uploader("2. Unggah Data Excel (.xlsx)", type=["xlsx"])

if excel_file is not None:
    try:
        # Ekstrak seluruh nama sheet yang ada di dalam file Excel
        excel_inspector = pd.ExcelFile(excel_file)
        all_sheets = excel_inspector.sheet_names
        
        st.success(f"Berhasil mendeteksi {len(all_sheets)} sheet di dalam berkas Excel.")
        
        # Kolom Pengaturan Antarmuka
        st.subheader("Konfigurasi Pemetaan Sheet & Output")
        
        # 1. Pilih Sheet Data
        selected_data_sheet = st.selectbox(
            "Pilih Sheet yang Berisi Data Utama (Looping Baris):",
            options=all_sheets
        )
        
        # 2. Pilih Sheet Metadata (Opsional)
        meta_options = ["Tidak Ada (Hanya Data Baris)"] + [s for s in all_sheets]
        selected_meta_sheet = st.selectbox(
            "Pilih Sheet yang Berisi Variabel Global/Metadata (Opsional):",
            options=meta_options,
            help="Pilih jika ada variabel statis yang diletakkan di sheet terpisah (Kolom 1=Kunci, Kolom 2=Nilai)."
        )
        
        # Pembacaan kolom secara dinamis berdasarkan sheet data yang dipilih
        df_headers = pd.read_excel(excel_file, sheet_name=selected_data_sheet, nrows=0)
        columns_options = df_headers.columns.tolist()
        
        selected_column = st.selectbox(
            "Pilih Kolom Acuan Nama File PDF:", 
            options=columns_options
        )
        
        file_prefix_input = st.text_input("Awalan (Prefix) Nama File PDF:", value="Dokumen")
        
        # Tombol Eksekusi di dalam kondisi jika excel terisi
        if st.button("Mulai Proses Engine", type="primary"):
            if template_file is None:
                st.error("Template Word belum diunggah.")
            else:
                with st.spinner("Sedang memproses dokumen massal via LibreOffice..."):
                    with tempfile.TemporaryDirectory() as temp_dir:
                        temp_template_path = os.path.join(temp_dir, "template.docx")
                        temp_excel_path = os.path.join(temp_dir, "data.xlsx")
                        temp_output_dir = os.path.join(temp_dir, "output_pdf")
                        zip_base_path = os.path.join(temp_dir, "Hasil_Output_Dokumen")
                        
                        with open(temp_template_path, "wb") as f:
                            f.write(template_file.getbuffer())
                        with open(temp_excel_path, "wb") as f:
                            f.write(excel_file.getbuffer())
                        
                        try:
                            sukses = generate_documents_from_excel(
                                template_path=temp_template_path,
                                excel_path=temp_excel_path,
                                output_dir=temp_output_dir,
                                filename_column=selected_column,
                                file_prefix=file_prefix_input.strip(),
                                data_sheet=selected_data_sheet,
                                meta_sheet=selected_meta_sheet
                            )
                            
                            generated_files = os.listdir(temp_output_dir)
                            if sukses == 0 or len(generated_files) == 0:
                                st.error("Proses selesai, namun tidak ada file berhasil dibuat. Periksa kecocokan tag template Anda.")
                            else:
                                shutil.make_archive(zip_base_path, 'zip', temp_output_dir)
                                zip_file_path = f"{zip_base_path}.zip"
                                
                                st.success(f"Sukses memproses {sukses} dokumen PDF!")
                                with open(zip_file_path, "rb") as f:
                                    st.download_button(
                                        label="⬇️ Unduh Semua Dokumen PDF (.zip)",
                                        data=f,
                                        file_name="Output_Surat_Massal.zip",
                                        mime="application/zip"
                                    )
                        except Exception as e:
                            st.error(f"Terjadi kesalahan pada internal engine: {e}")
                            
    except Exception as e:
        st.error(f"Gagal membedah struktur file Excel. Error: {e}")