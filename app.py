import streamlit as st
import pandas as pd
from fuzzywuzzy import process
import os
from PIL import Image

# Judul aplikasi dengan Markdown
st.markdown("""
# 🚀 Aplikasi Pencocokan Data dan Pengecekan Duplikat
Aplikasi ini digunakan untuk mencocokkan data dan mengecek duplikat.
""")

# Menambahkan gambar logo
image = Image.open("custody.png")
st.image(image, caption="Logo Aplikasi", width=200)

# Upload file Excel
st.sidebar.header("📤 Upload File Excel")
file_resi = st.sidebar.file_uploader("Upload File Resi (resi.xlsx)", type=["xlsx"])
file_laporan = st.sidebar.file_uploader("Upload File Laporan (laporan.xlsx)", type=["xlsx"])

if file_resi and file_laporan:
    # Baca data dari file Excel
    df_resi = pd.read_excel(file_resi)
    df_laporan = pd.read_excel(file_laporan)

    # Bersihkan nama kolom
    df_resi.columns = df_resi.columns.str.strip()
    df_laporan.columns = df_laporan.columns.str.strip()

    # Tampilkan data dalam 2 kolom
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📄 Data Resi")
        st.write(df_resi)

    with col2:
        st.subheader("📊 Data Laporan")
        st.write(df_laporan)

    # Pencocokan Fuzzy Matching
    st.subheader("🔍 Pencocokan Nama Debitur dan Nama Penerima")
    if st.button("🚀 Cocokkan Data"):
        # Pastikan kolom 'Nama Debitur' dan 'Nama Penerima' bertipe string
        df_laporan['Nama Debitur'] = df_laporan['Nama Debitur'].astype(str)
        df_resi['Nama Penerima'] = df_resi['Nama Penerima'].astype(str)

        # Hapus baris yang memiliki nilai NaN atau string kosong
        df_laporan = df_laporan[df_laporan['Nama Debitur'].str.strip() != '']
        df_resi = df_resi[df_resi['Nama Penerima'].str.strip() != '']

        # Fungsi untuk mencocokkan nama dengan fuzzy matching
        def fuzzy_match(name, list_names, min_score=80):
            if not isinstance(name, str):
                return None, 0
            match, score = process.extractOne(name, list_names)
            if score >= min_score:
                return match, score
            return None, 0

        # Buat list nama penerima dari file resi
        list_nama_penerima = df_resi['Nama Penerima'].tolist()

        # Loop melalui setiap baris di file laporan
        for index, row in df_laporan.iterrows():
            nama_debitur = row['Nama Debitur']
            nama_cocok, similarity_score = fuzzy_match(nama_debitur, list_nama_penerima)
            if nama_cocok:
                data_resi = df_resi[df_resi['Nama Penerima'] == nama_cocok]
                if not data_resi.empty:
                    data_resi = data_resi.iloc[0]
                    df_laporan.at[index, 'Nomor Resi Pengiriman'] = str(data_resi['Nomor Resi'])
                    df_laporan.at[index, 'Nama Penerima Somasi'] = str(data_resi['Nama Penerima'])
                    df_laporan.at[index, 'Similarity Score'] = similarity_score
                else:
                    df_laporan.at[index, 'Similarity Score'] = 0
            else:
                df_laporan.at[index, 'Similarity Score'] = 0

        # Tampilkan hasil pencocokan
        st.subheader("✅ Hasil Pencocokan")
        st.write(df_laporan)

        # Simpan hasil ke file Excel
        output_file = "laporan_terupdate.xlsx"
        df_laporan.to_excel(output_file, index=False)

        # Tombol unduh file Excel
        with open(output_file, "rb") as file:
            btn = st.download_button(
                label="📥 Download Hasil Pencocokan (Excel)",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # Hapus file sementara setelah diunduh
        if os.path.exists(output_file):
            os.remove(output_file)

    # Pengecekan Duplikat
    st.subheader("🔎 Pengecekan Duplikat")
    if st.button("🔍 Cek Duplikat"):
        # Fungsi untuk mengecek duplikat
        def cek_duplikat(df, kolom, nama_file):
            duplikat = df[df.duplicated(kolom, keep=False)]
            if not duplikat.empty:
                st.write(f"Duplikat ditemukan di kolom '{kolom}':")
                st.write(duplikat)
                duplikat.to_excel(nama_file, index=False)

                # Tombol unduh file Excel
                with open(nama_file, "rb") as file:
                    btn = st.download_button(
                        label=f"📥 Download Duplikat {kolom} (Excel)",
                        data=file,
                        file_name=nama_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # Hapus file sementara setelah diunduh
                if os.path.exists(nama_file):
                    os.remove(nama_file)
            else:
                st.write(f"Tidak ada duplikat di kolom '{kolom}'.")

        # Cek duplikat nama penerima di file resi
        cek_duplikat(df_resi, 'Nama Penerima', 'duplikat_nama_penerima.xlsx')

        # Cek duplikat nama debitur di file laporan
        cek_duplikat(df_laporan, 'Nama Debitur', 'duplikat_nama_debitur.xlsx')