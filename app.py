import streamlit as st
import os
import tempfile
import shutil
from core_engine import generate_documents_from_excel

# Konfigurasi Halaman
st.set_page_config(page_title="Automasi Surat & Berkas", layout="centered")

st.title("Generator Surat & Berkas Otomatis")
st.markdown("Unggah **Template Word (.docx)** dan **Data Excel (.xlsx)** (wajib memiliki sheet `Metadata` & `Peserta`) untuk memproduksi dokumen massal.")

# Form Upload
template_file = st.file_uploader("1. Unggah Template Word", type=["docx"])
excel_file = st.file_uploader("2. Unggah Data Excel", type=["xlsx"])

if st.button("Generate Dokumen", type="primary"):
    if template_file is None or excel_file is None:
        st.error("Error: Harap unggah kedua dokumen (Template dan Excel) sebelum memproses.")
    else:
        with st.spinner("Mesin sedang memproses dokumen..."):
            # Isolasi proses menggunakan direktori sementara (Temporary Directory)
            with tempfile.TemporaryDirectory() as temp_dir:
                # Definisikan path sementara
                temp_template_path = os.path.join(temp_dir, "template.docx")
                temp_excel_path = os.path.join(temp_dir, "data.xlsx")
                temp_output_dir = os.path.join(temp_dir, "output_berkas")
                zip_base_path = os.path.join(temp_dir, "Hasil_Dokumen")
                
                # Tulis file dari RAM (buffer) ke direktori sementara
                with open(temp_template_path, "wb") as f:
                    f.write(template_file.getbuffer())
                with open(temp_excel_path, "wb") as f:
                    f.write(excel_file.getbuffer())
                
                try:
                    # Panggil logika utama dari core_engine.py
                    generate_documents_from_excel(temp_template_path, temp_excel_path, temp_output_dir)
                    
                    # Kompresi folder output menjadi .zip untuk diunduh user
                    shutil.make_archive(zip_base_path, 'zip', temp_output_dir)
                    zip_file_path = f"{zip_base_path}.zip"
                    
                    # Jika proses zip berhasil, tampilkan tombol unduh
                    if os.path.exists(zip_file_path):
                        st.success("Dokumen berhasil dirender!")
                        with open(zip_file_path, "rb") as f:
                            st.download_button(
                                label="⬇️ Unduh Semua Dokumen (.zip)",
                                data=f,
                                file_name="Hasil_Generate_Surat.zip",
                                mime="application/zip"
                            )
                except Exception as e:
                    # Menangkap error dari pandas atau docxtpl yang dilempar oleh core_engine
                    st.error(f"Kegagalan Eksekusi: {e}")