import streamlit as st
import pandas as pd
from fuzzywuzzy import process
from io import BytesIO
from PIL import Image
import re
import pdfplumber  # <-- tambahan untuk baca PDF

# =========================
# Konfigurasi halaman
# =========================
st.set_page_config(page_title="Pencocokan Somasi", layout="wide")

# =========================
# Normalisasi & Pemetaan Kolom
# =========================
CANON_MAP = {
    # laporan & resi
    r'^nama\s*debitur$': 'Nama Debitur',
    r'^nama\s*penerima$': 'Nama Penerima',
    r'^nama\s*penerima\s*somasi$': 'Nama Penerima Somasi',
    r'^nomor\s*resi$': 'Nomor Resi',
    r'^nomor\s*resi\s*pengiriman$': 'Nomor Resi Pengiriman',
    r'^(no|nomor|no\.)\s*surat\s*somasi$': 'Nomor Surat Somasi',
    r'^kode$': 'Nomor Resi',
    r'^nomor\s*surat\s*somasi$': 'Nomor Surat Somasi',
    r'^no\s*surat\s*somasi$': 'Nomor Surat Somasi',
    r'^nosurat\s*somasi$': 'Nomor Surat Somasi',
    r'^surat\s*somasi$': 'Nomor Surat Somasi',
    r'^(no|nomor|no\.)\s*surat\s*kuasa$': 'Nomor Surat Kuasa',
    r'^nomor\s*surat\s*kuasa$': 'Nomor Surat Kuasa',
    r'^(no|nomor|no\.)\s*kontrak$': 'Nomor Kontrak',
    r'^nomor\s*kontrak$': 'Nomor Kontrak',
    # bahan
    r'^nama\s*konsumen$': 'Nama Konsumen',
}

def strip_enumeration_prefix(s: str) -> str:
    # "3. NOMOR SURAT SOMASI" -> "NOMOR SURAT SOMASI"
    return re.sub(r'^\s*\d+\s*[\.\)]\s*', '', s or "")

def canonicalize_columns(cols):
    """Bersihkan dan petakan nama kolom ke bentuk kanonik TANPA dobel-append."""
    out = []
    for c in cols:
        c = (c or "")
        c = c.replace("\ufeff", "").replace("\xa0", " ")
        c = strip_enumeration_prefix(c)
        c = re.sub(r'\s+', ' ', c).strip()
        lc = c.lower()

        mapped = None
        for pat, target in CANON_MAP.items():
            if re.match(pat, lc):
                mapped = target
                break

        out.append(mapped if mapped is not None else c)
    return out

def normalize_somasi(s: str) -> str:
    if not isinstance(s, str):
        s = str(s) if pd.notna(s) else ""
    s = s.upper().strip().replace("\xa0", " ")
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'[^A-Z0-9/\-]', '', s)  # pertahankan A-Z, 0-9, '/', '-'
    return s

def normalize_kontrak(s: str) -> str:
    # penting untuk hindari notasi ilmiah → paksa string, buang spasi & leading '
    s = "" if pd.isna(s) else str(s)
    s = s.strip().lstrip("'").replace(" ", "")
    return s

def normalize_nama(s: str) -> str:
    s = "" if pd.isna(s) else str(s)
    s = re.sub(r'\s+', ' ', s.strip()).upper()
    return s

# =========================
# Loader Fleksibel: xlsx/csv/tsv (umum)
# =========================
def load_file_flex(uploaded_file):
    if uploaded_file.name.lower().endswith(".xlsx"):
        df = pd.read_excel(uploaded_file)
    else:
        uploaded_file.seek(0)
        try:
            df = pd.read_csv(uploaded_file, sep=None, engine="python", encoding="utf-8-sig")
        except Exception:
            uploaded_file.seek(0)
            try:
                df = pd.read_csv(uploaded_file, sep=None, engine="python", encoding="latin1")
            except Exception:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, sep="\t", engine="python", encoding_errors="ignore")
    df.columns = canonicalize_columns([str(c) for c in df.columns])
    return df

# =========================
# Loader khusus Resi: + dukung PDF
# =========================
def load_resi_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        # Ekstrak tabel dari PDF menggunakan pdfplumber
        tables = []
        header = None
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                if not page_tables:
                    continue
                for table in page_tables:
                    if not table:
                        continue
                    if header is None:
                        header = table[0]
                        rows = table[1:]
                    else:
                        # kalau header sama, skip baris header
                        if table[0] == header:
                            rows = table[1:]
                        else:
                            rows = table
                    for row in rows:
                        # kadang row panjangnya beda, amankan
                        if len(row) < len(header):
                            row = row + [None] * (len(header) - len(row))
                        elif len(row) > len(header):
                            row = row[:len(header)]
                        tables.append(row)

        if header is None:
            st.error("❌ Tidak menemukan tabel di PDF Resi. Pastikan resi berupa tabel (bukan hanya gambar).")
            return pd.DataFrame(columns=["Nama Penerima", "Nomor Resi"])

        df = pd.DataFrame(tables, columns=canonicalize_columns([str(c) for c in header]))
        return df
    else:
        # fallback ke loader biasa (xlsx/csv/tsv)
        return load_file_flex(uploaded_file)

# =========================
# Streamlit-safe display helper
# =========================
def make_streamlit_safe(df: pd.DataFrame) -> pd.DataFrame:
    safe = df.copy()
    for col in safe.columns:
        s = safe[col]
        try:
            coerced = pd.to_numeric(s, errors='coerce')
            if coerced.notna().any() and coerced.isna().any():
                safe[col] = s.astype('string')
        except Exception:
            safe[col] = s.astype('string')
    return safe

# =========================
# Writer Excel (openpyxl → fallback xlsxwriter)
# =========================
def write_excel_download(df: pd.DataFrame, default_filename: str):
    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Hasil")
    except Exception:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Hasil")
    output.seek(0)
    st.download_button(
        label=f"📥 Download {default_filename}",
        data=output,
        file_name=default_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

def get_str(val) -> str:
    return "" if pd.isna(val) else str(val)

def normalize_resi(s: str) -> str:
    if pd.isna(s):
        return ""
    if isinstance(s, float) and s.is_integer():
        s = str(int(s))
    else:
        s = str(s)
    s = s.strip().lstrip("'").replace("\xa0", " ")
    s = re.sub(r'\s+', '', s).upper()
    if re.match(r'^\d+\.0$', s):
        s = s[:-2]
    return s

def normalize_column_key(name: str) -> str:
    name = strip_enumeration_prefix(str(name or ""))
    name = name.replace("\ufeff", "").replace("\xa0", " ")
    name = re.sub(r'\s+', ' ', name).strip().lower()
    return re.sub(r'[^a-z0-9]+', '', name)

def find_column(df: pd.DataFrame, *candidates: str):
    columns_by_key = {normalize_column_key(c): c for c in df.columns}
    for candidate in candidates:
        if candidate in df.columns:
            return candidate
        key = normalize_column_key(candidate)
        if key in columns_by_key:
            return columns_by_key[key]
    return None

def format_date_for_report(val) -> str:
    if pd.isna(val):
        return ""
    if hasattr(val, "strftime") and not isinstance(val, str):
        return val.strftime("%d/%m/%Y")
    return str(val).strip()

def write_csv_download(df: pd.DataFrame, default_filename: str):
    data = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label=f"📥 Download {default_filename}",
        data=data,
        file_name=default_filename,
        mime="text/csv",
    )

# =========================
# UI
# =========================
st.sidebar.header("📌 Menu")
active_tab = st.sidebar.radio(
    "Pilih tab",
    ["Pencocokan Somasi", "Pencocokan No Resi dan Detail"],
    label_visibility="collapsed",
)
st.sidebar.markdown("---")

if active_tab == "Pencocokan No Resi dan Detail":
    st.markdown("# 📦 Pencocokan No Resi dan Detail")

    try:
        image = Image.open("custody.png")
        st.image(image, caption="Logo Aplikasi", width=200)
    except Exception:
        pass

    st.sidebar.header("📤 Upload Berkas")
    file_laporan_detail = st.sidebar.file_uploader(
        "Laporan CSV",
        type=["csv"],
        key="laporan_detail_resi",
    )
    file_detail_resi = st.sidebar.file_uploader(
        "Detail Resi XLSX",
        type=["xlsx"],
        key="detail_resi_xlsx",
    )

    if file_laporan_detail and file_detail_resi:
        df_laporan_detail = load_file_flex(file_laporan_detail)
        df_detail_resi = load_file_flex(file_detail_resi)

        detail_no_resi_col = find_column(df_detail_resi, "No Resi", "Nomor Resi", "Nomor Resi Pengiriman")
        detail_penerima_paket_col = find_column(df_detail_resi, "Penerima Paket")
        detail_status_penerima_col = find_column(df_detail_resi, "Status Penerima")
        detail_tanggal_col = find_column(df_detail_resi, "Tanggal")

        laporan_no_resi_col = find_column(df_laporan_detail, "Nomor Resi Pengiriman", "No Resi", "Nomor Resi")
        laporan_penerima_col = find_column(df_laporan_detail, "NAMA PENERIMA EKSPEDISI SOMASI")
        laporan_status_col = find_column(df_laporan_detail, "STATUS PENERIMA EKSPEDISI SOMASI")
        laporan_tanggal_col = find_column(df_laporan_detail, "Tanggal Somasi Diterima")

        created_columns = []
        if laporan_penerima_col is None:
            laporan_penerima_col = "NAMA PENERIMA EKSPEDISI SOMASI"
            df_laporan_detail[laporan_penerima_col] = ""
            created_columns.append(laporan_penerima_col)
        if laporan_status_col is None:
            laporan_status_col = "STATUS PENERIMA EKSPEDISI SOMASI"
            df_laporan_detail[laporan_status_col] = ""
            created_columns.append(laporan_status_col)
        if laporan_tanggal_col is None:
            laporan_tanggal_col = "Tanggal Somasi Diterima"
            df_laporan_detail[laporan_tanggal_col] = ""
            created_columns.append(laporan_tanggal_col)

        with st.expander("🔧 Kolom terdeteksi"):
            st.write("Laporan:", list(df_laporan_detail.columns))
            st.write("Detail Resi:", list(df_detail_resi.columns))
            if created_columns:
                st.write("Kolom laporan yang dibuat:", created_columns)

        missing_cols = []
        if laporan_no_resi_col is None:
            missing_cols.append("Laporan: Nomor Resi Pengiriman")
        if detail_no_resi_col is None:
            missing_cols.append("Detail Resi: No Resi")
        if detail_penerima_paket_col is None:
            missing_cols.append("Detail Resi: Penerima Paket")
        if detail_status_penerima_col is None:
            missing_cols.append("Detail Resi: Status Penerima")
        if detail_tanggal_col is None:
            missing_cols.append("Detail Resi: Tanggal")

        if missing_cols:
            st.error(f"Kolom wajib tidak ditemukan: {missing_cols}")
            st.stop()

        col_lap, col_detail = st.columns(2)
        with col_lap:
            st.subheader("📊 Laporan")
            st.dataframe(make_streamlit_safe(df_laporan_detail))
        with col_detail:
            st.subheader("📦 Detail Resi")
            st.dataframe(make_streamlit_safe(df_detail_resi))

        if st.button("🚀 Cocokkan No Resi dan Isi Detail"):
            df_laporan_out = df_laporan_detail.copy()
            df_detail_work = df_detail_resi.copy()

            df_laporan_out["__resi_norm"] = df_laporan_out[laporan_no_resi_col].apply(normalize_resi)
            df_detail_work["__resi_norm"] = df_detail_work[detail_no_resi_col].apply(normalize_resi)
            df_detail_work = df_detail_work[df_detail_work["__resi_norm"] != ""].copy()

            duplicate_detail = df_detail_work[df_detail_work.duplicated("__resi_norm", keep=False)].copy()
            detail_lookup = (
                df_detail_work
                .drop_duplicates("__resi_norm", keep="first")
                .set_index("__resi_norm", drop=False)
            )

            matched_count = 0
            empty_resi_count = 0
            unmatched_idx = []

            for idx, row in df_laporan_out.iterrows():
                resi_norm = row["__resi_norm"]
                if not resi_norm:
                    empty_resi_count += 1
                    continue
                if resi_norm not in detail_lookup.index:
                    unmatched_idx.append(idx)
                    continue

                detail_row = detail_lookup.loc[resi_norm]
                df_laporan_out.at[idx, laporan_penerima_col] = get_str(detail_row.get(detail_penerima_paket_col))
                df_laporan_out.at[idx, laporan_status_col] = get_str(detail_row.get(detail_status_penerima_col))
                df_laporan_out.at[idx, laporan_tanggal_col] = format_date_for_report(detail_row.get(detail_tanggal_col))
                matched_count += 1

            df_laporan_out = df_laporan_out.drop(columns=["__resi_norm"])

            st.subheader("✅ Hasil Pencocokan No Resi dan Detail")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Baris Laporan", len(df_laporan_detail))
            m2.metric("Berhasil Cocok", matched_count)
            m3.metric("Resi Tidak Ditemukan", len(unmatched_idx))
            m4.metric("No Resi Kosong", empty_resi_count)

            if not duplicate_detail.empty:
                st.warning("Ada No Resi duplikat di Detail Resi. Aplikasi memakai baris pertama untuk pengisian laporan.")
                duplicate_cols = [
                    c for c in [
                        detail_no_resi_col,
                        detail_penerima_paket_col,
                        detail_status_penerima_col,
                        detail_tanggal_col,
                    ] if c in duplicate_detail.columns
                ]
                st.dataframe(make_streamlit_safe(duplicate_detail[duplicate_cols]))

            if unmatched_idx:
                st.warning("Beberapa Nomor Resi Pengiriman di laporan tidak ditemukan pada Detail Resi.")
                st.dataframe(make_streamlit_safe(df_laporan_detail.loc[unmatched_idx, [laporan_no_resi_col]]))

            st.dataframe(make_streamlit_safe(df_laporan_out))
            write_excel_download(df_laporan_out, "laporan_detail_resi_terupdate.xlsx")
            write_csv_download(df_laporan_out, "laporan_detail_resi_terupdate.csv")
    else:
        st.info("📎 Upload **Laporan CSV** dan **Detail Resi XLSX** untuk mulai mencocokkan No Resi.")

    st.stop()

st.markdown("""
# 🚀 Pencocokan Somasi
1) Fuzzy **Nama Debitur ↔ Nama Penerima** (opsional)  
2) **Isi/Cek Nomor Surat Somasi** dari **Bahan** (berdasar **No. Kontrak** → fallback fuzzy **Nama**)
""")

try:
    image = Image.open("custody.png")
    st.image(image, caption="Logo Aplikasi", width=200)
except Exception:
    pass

st.sidebar.header("📤 Upload Berkas")
file_resi    = st.sidebar.file_uploader("Resi (xlsx/csv/tsv/pdf)",    type=["xlsx", "csv", "tsv", "pdf"])
file_laporan = st.sidebar.file_uploader("Laporan (xlsx/csv/tsv)",     type=["xlsx", "csv", "tsv"])
files_bahan  = st.sidebar.file_uploader(
    "Bahan (boleh banyak, opsional) – berisi '3. NOMOR SURAT SOMASI' & '4. NAMA KONSUMEN'",
    type=["xlsx","csv","tsv"], accept_multiple_files=True
)

st.sidebar.markdown("---")
only_somasi_mode = st.sidebar.checkbox("🔖 Hanya cocokkan **Nomor Surat Somasi** (berdasar Nama / No. Kontrak)")
somasi_priority = st.sidebar.radio(
    "Prioritas pencocokan Somasi (saat mengisi dari Bahan):",
    options=["Utamakan No. Kontrak", "Utamakan Nama"],
    horizontal=False,
    index=0
)

threshold_name        = st.sidebar.slider("Ambang Similarity Nama (0-100)",   0, 100, 80, 1)
threshold_somasi      = st.sidebar.slider("Ambang Similarity Somasi (0-100)", 0, 100, 90, 1)
threshold_nama_bahan  = st.sidebar.slider("Ambang Fuzzy Nama (isi Somasi dari Bahan)", 0, 100, 85, 1)

# =========================================
# Jalankan minimal dengan Resi + Laporan
# =========================================
if file_resi and file_laporan:
    df_resi = load_resi_file(file_resi)       # <--- pakai loader baru
    df_laporan = load_file_flex(file_laporan)

    # Siapkan df_bahan (mungkin kosong)
    bahan_list, bahan_cols_debug = [], []
    if files_bahan:
        for f in files_bahan:
            df_b = load_file_flex(f)
            bahan_list.append(df_b)
            bahan_cols_debug.append((f.name, list(df_b.columns)))
    df_bahan = pd.concat(bahan_list, ignore_index=True) if bahan_list else pd.DataFrame()

    # --- Debug kolom
    with st.expander("🔧 Kolom terdeteksi"):
        st.write("Resi:", list(df_resi.columns))
        st.write("Laporan:", list(df_laporan.columns))
        if not df_bahan.empty:
            for nm, cols in bahan_cols_debug:
                st.write(f"Bahan ({nm}):", cols)
            st.write("Bahan (Gabungan):", list(df_bahan.columns))
        else:
            st.write("Bahan: (tidak diupload)")

    # --- Validasi minimum untuk mode tanpa bahan
    required_laporan_min = ["Nama Debitur"]
    required_resi_min    = ["Nama Penerima"]

    missing_min = [c for c in required_laporan_min if c not in df_laporan.columns] \
                + [c for c in required_resi_min    if c not in df_resi.columns]
    if missing_min:
        st.error(f"Kolom wajib tidak ditemukan (mode minimal): {missing_min}")
        st.stop()

    # --- Cast dasar
    df_laporan['Nama Debitur'] = df_laporan['Nama Debitur'].astype('string').fillna("")
    df_resi['Nama Penerima']   = df_resi['Nama Penerima'].astype('string').fillna("")
    if 'Nomor Resi' in df_resi.columns:
        df_resi['Nomor Resi'] = df_resi['Nomor Resi'].astype('string').fillna("")

    # --- Jika ada bahan, siapkan kolom & normalisasi
    bahan_mode = not df_bahan.empty
    if bahan_mode:
        # kolom tambahan yang dipakai dalam mode bahan
        for c in ["Nomor Surat Somasi", "Nomor Kontrak"]:
            if c in df_laporan.columns:
                df_laporan[c] = df_laporan[c].astype('string').fillna("")
        if "Nomor Surat Somasi" not in df_laporan.columns:
            df_laporan["Nomor Surat Somasi"] = pd.Series("", index=df_laporan.index, dtype='string')

        # bahan wajib minimal punya kolom somasi
        if "Nomor Surat Somasi" not in df_bahan.columns:
            st.warning("File Bahan diupload tapi kolom 'Nomor Surat Somasi' tidak ada. Mode Bahan dinonaktifkan.")
            bahan_mode = False
        else:
            df_bahan['Nomor Surat Somasi'] = df_bahan['Nomor Surat Somasi'].astype('string').fillna("")
            if 'Nama Konsumen' in df_bahan.columns:
                df_bahan['Nama Konsumen'] = df_bahan['Nama Konsumen'].astype('string').fillna("")
            if 'Nomor Kontrak' in df_bahan.columns:
                df_bahan['Nomor Kontrak'] = df_bahan['Nomor Kontrak'].astype('string').fillna("")

            # normalisasi kunci
            df_laporan['__nama_norm']    = df_laporan['Nama Debitur'].apply(normalize_nama)
            df_laporan['__kontrak_norm'] = df_laporan.get('Nomor Kontrak', "").apply(normalize_kontrak) \
                                            if 'Nomor Kontrak' in df_laporan.columns else ""
            df_laporan['__somasi_norm']  = df_laporan['Nomor Surat Somasi'].apply(normalize_somasi)

            df_bahan['__somasi_norm'] = df_bahan['Nomor Surat Somasi'].apply(normalize_somasi)
            if 'Nama Konsumen' in df_bahan.columns:
                df_bahan['__nama_norm'] = df_bahan['Nama Konsumen'].apply(normalize_nama)
            else:
                df_bahan['__nama_norm'] = ""
            if 'Nomor Kontrak' in df_bahan.columns:
                df_bahan['__kontrak_norm'] = df_bahan['Nomor Kontrak'].apply(normalize_kontrak)
            else:
                df_bahan['__kontrak_norm'] = ""

            bahan_by_kontrak = df_bahan[df_bahan['__kontrak_norm'] != ""].set_index('__kontrak_norm')

    # --- Tampilkan data
    cols = st.columns(3)
    with cols[0]:
        st.subheader("📄 Resi")
        st.dataframe(make_streamlit_safe(df_resi))
    with cols[1]:
        st.subheader("📊 Laporan")
        st.dataframe(make_streamlit_safe(df_laporan))
    with cols[2]:
        st.subheader("🧾 Bahan")
        if bahan_mode:
            st.dataframe(make_streamlit_safe(df_bahan))
        else:
            st.caption("Tidak ada / dinonaktifkan")

    # =========================
    # Pencocokan
    # =========================
    st.subheader("🔍 Pencocokan Nama & (Opsional) Nomor Surat Somasi")
    if st.button("🚀 Cocokkan Data"):
        df_laporan_out = df_laporan.copy()

        # Kolom hasil (selalu ada)
        def ensure_col(df, name, dtype, default):
            if name not in df.columns:
                df[name] = pd.Series(default, index=df.index, dtype=dtype)
            else:
                if dtype=='string':
                    df[name] = df[name].astype('string').fillna(default)
                elif dtype=='int64':
                    df[name] = pd.to_numeric(df[name], errors='coerce').fillna(default).astype('int64')
                elif dtype=='bool':
                    df[name] = df[name].astype('bool').fillna(default)
            return df

        ensure_col(df_laporan_out, 'Nomor Resi Pengiriman', 'string', "")
        ensure_col(df_laporan_out, 'Nama Penerima Somasi',  'string', "")
        ensure_col(df_laporan_out, 'Similarity Score (Nama)','int64', 0)
        ensure_col(df_laporan_out, 'Somasi (Bahan)',        'string', "")
        ensure_col(df_laporan_out, 'Somasi Match Score',    'int64', 0)
        ensure_col(df_laporan_out, 'Somasi Match',          'bool',  False)

        # === MODE A: FULL MATCH (fuzzy Nama Debitur ↔ Nama Penerima dari Resi) ===
        # Pencocokan 1:1 — setiap baris resi hanya boleh dipakai SEKALI
        if not only_somasi_mode:
            # Buat daftar kandidat resi dengan index asli, agar bisa ditandai "sudah dipakai"
            resi_candidates = []
            for resi_idx, resi_row in df_resi.iterrows():
                resi_candidates.append({
                    'resi_idx': resi_idx,
                    'nama': str(resi_row.get('Nama Penerima', '') or '').strip(),
                    'used': False,
                    'row': resi_row,
                })

            def fuzzy_match_name_1to1(query, candidates, min_score):
                """Cari kandidat resi terbaik yang BELUM dipakai."""
                query = query.strip()
                if not query:
                    return None, 0, -1
                available = [(c['nama'], i) for i, c in enumerate(candidates) if not c['used'] and c['nama']]
                if not available:
                    return None, 0, -1
                names_only = [a[0] for a in available]
                res = process.extractOne(query, names_only)
                if not res:
                    return None, 0, -1
                match_name, score = res
                if score < min_score:
                    return None, 0, -1
                # Cari index kandidat yang cocok (ambil yang pertama ditemukan)
                for name, cand_i in available:
                    if name == match_name:
                        return match_name, score, cand_i
                return None, 0, -1

            for idx, row in df_laporan_out.iterrows():
                nm = str(row['Nama Debitur'])
                m, sc, cand_i = fuzzy_match_name_1to1(nm, resi_candidates, threshold_name)
                if m and cand_i >= 0:
                    # Tandai resi ini sudah dipakai (tidak bisa dipakai baris lain)
                    resi_candidates[cand_i]['used'] = True
                    res_row = resi_candidates[cand_i]['row']

                    # isi nama penerima
                    df_laporan_out.at[idx, 'Nama Penerima Somasi'] = get_str(res_row.get('Nama Penerima'))

                    # ambil nomor resi dari dua kemungkinan kolom
                    resi_val = ""
                    if 'Nomor Resi Pengiriman' in df_resi.columns and pd.notna(res_row.get('Nomor Resi Pengiriman')):
                        resi_val = get_str(res_row.get('Nomor Resi Pengiriman'))
                    elif 'Nomor Resi' in df_resi.columns and pd.notna(res_row.get('Nomor Resi')):
                        resi_val = get_str(res_row.get('Nomor Resi'))

                    if resi_val:
                        df_laporan_out.at[idx, 'Nomor Resi Pengiriman'] = resi_val

                    df_laporan_out.at[idx, 'Similarity Score (Nama)'] = int(sc)

        else:
            st.info("🧩 Mode ringan aktif: hanya mencocokkan Nomor Surat Somasi berdasarkan Nama / No. Kontrak.")

        # === B) (Opsional) ISI/CEK Nomor Surat Somasi dari Bahan ===
        if bahan_mode:
            # default: isi Somasi (Bahan) = nilai Laporan (biar tidak kosong)
            if 'Nomor Surat Somasi' not in df_laporan_out.columns:
                df_laporan_out['Nomor Surat Somasi'] = pd.Series("", index=df_laporan_out.index, dtype='string')
            df_laporan_out['Somasi (Bahan)'] = df_laporan_out['Nomor Surat Somasi'].astype('string').fillna("")

            # util pencari dari bahan by kontrak / by nama
            def find_somasi_from_bahan_by_kontrak(kontrak_norm):
                if not kontrak_norm or '__kontrak_norm' not in df_bahan.columns:
                    return ""
                hit = df_bahan[df_bahan['__kontrak_norm'] == kontrak_norm].head(1)
                if not hit.empty:
                    return get_str(hit.iloc[0]['Nomor Surat Somasi'])
                return ""

            def find_somasi_from_bahan_by_nama(nama_norm):
                if not nama_norm or '__nama_norm' not in df_bahan.columns:
                    return "", 0
                res = process.extractOne(nama_norm, df_bahan['__nama_norm'].tolist())
                if not res: return ("", 0)
                mname, score = res
                if score >= threshold_nama_bahan:
                    b = df_bahan[df_bahan['__nama_norm'] == mname].head(1)
                    if not b.empty:
                        return (get_str(b.iloc[0]['Nomor Surat Somasi']), int(score))
                return ("", 0)

            # urutan prioritas sesuai pilihan
            prefer_kontrak = (somasi_priority == "Utamakan No. Kontrak")

            for idx, row in df_laporan_out.iterrows():
                somasi_lap = str(row.get('Nomor Surat Somasi', "") or "")
                if somasi_lap.strip():
                    continue  # sudah ada, nanti tetap diverifikasi

                kontrak_norm = normalize_kontrak(row.get('Nomor Kontrak', ""))
                nama_norm    = normalize_nama(row.get('Nama Debitur', ""))

                def try_fill_by_kontrak_then_nama():
                    # kontrak → nama
                    if kontrak_norm:
                        s = find_somasi_from_bahan_by_kontrak(kontrak_norm)
                        if s:
                            return s, 100
                    s, sc = find_somasi_from_bahan_by_nama(nama_norm)
                    return s, sc

                def try_fill_by_nama_then_kontrak():
                    # nama → kontrak
                    s, sc = find_somasi_from_bahan_by_nama(nama_norm)
                    if s:
                        return s, sc
                    if kontrak_norm:
                        s2 = find_somasi_from_bahan_by_kontrak(kontrak_norm)
                        return (s2, 100 if s2 else 0)
                    return ("", 0)

                s_final, sc_final = (try_fill_by_kontrak_then_nama() if prefer_kontrak
                                     else try_fill_by_nama_then_kontrak())

                if s_final:
                    df_laporan_out.at[idx, 'Nomor Surat Somasi'] = s_final
                    df_laporan_out.at[idx, 'Somasi (Bahan)'] = s_final
                    df_laporan_out.at[idx, 'Somasi Match'] = True
                    df_laporan_out.at[idx, 'Somasi Match Score'] = int(max(sc_final, df_laporan_out.at[idx, 'Somasi Match Score']))

            # verifikasi: jika Laporan punya Somasi, cocokkan dengan Bahan (exact→fuzzy)
            bahan_norm_set = set(df_bahan['__somasi_norm'].tolist())
            daftar_bahan_norm = df_bahan['__somasi_norm'].tolist()
            def verify_somasi(val):
                if not val: return False, 0, ""
                vnorm = normalize_somasi(val)
                if vnorm in bahan_norm_set: return True, 100, val
                res = process.extractOne(vnorm, daftar_bahan_norm)
                if not res: return False, 0, ""
                match, sc = res
                if sc >= threshold_somasi:
                    asli = df_bahan[df_bahan['__somasi_norm'] == match].head(1)
                    return True, int(sc), get_str(asli.iloc[0]['Nomor Surat Somasi']) if not asli.empty else val
                return False, 0, ""

            for idx, row in df_laporan_out.iterrows():
                cur = get_str(row.get('Nomor Surat Somasi', ""))
                ok, sc, from_bahan = verify_somasi(cur)
                if ok:
                    df_laporan_out.at[idx, 'Somasi Match'] = True
                    df_laporan_out.at[idx, 'Somasi Match Score'] = int(max(sc, df_laporan_out.at[idx, 'Somasi Match Score']))
                    if from_bahan:
                        df_laporan_out.at[idx, 'Somasi (Bahan)'] = from_bahan
        else:
            st.info("ℹ️ File Bahan tidak diupload → pencocokan Nomor Surat Somasi dilewati.")

        # --- Tampilkan & unduh
        st.subheader("✅ Hasil Pencocokan")
        st.dataframe(make_streamlit_safe(df_laporan_out))
        write_excel_download(df_laporan_out, "laporan_terupdate.xlsx")

    # =========================
    # Duplikat
    # =========================
    st.subheader("🔎 Pengecekan Duplikat")
    if st.button("🔍 Cek Duplikat"):
        def cek_duplikat(df, kolom, default_filename):
            if kolom not in df.columns:
                st.warning(f"Kolom '{kolom}' tidak ditemukan.")
                return
            duplikat = df[df.duplicated(kolom, keep=False)]
            if not duplikat.empty:
                st.write(f"Duplikat ditemukan di kolom '{kolom}':")
                st.dataframe(make_streamlit_safe(duplikat))
                out = duplikat.copy()
                for c in out.columns:
                    if pd.api.types.is_object_dtype(out[c]) or str(out[c].dtype) == 'string':
                        out[c] = out[c].astype('string').fillna("")
                write_excel_download(out, default_filename)
            else:
                st.write(f"Tidak ada duplikat di kolom '{kolom}'.")
        colA, colB, colC = st.columns(3)
        with colA: st.caption("Resi");   cek_duplikat(df_resi,  'Nama Penerima',       'duplikat_nama_penerima_resi.xlsx')
        with colB: st.caption("Laporan");cek_duplikat(df_laporan,'Nama Debitur',        'duplikat_nama_debitur_laporan.xlsx')
        with colC:
            st.caption("Bahan")
            if not df_bahan.empty:
                cek_duplikat(df_bahan, 'Nomor Surat Somasi',  'duplikat_nomor_surat_somasi_bahan.xlsx')
            else:
                st.write("—")

else:
    st.info("📎 Upload minimal **Resi** dan **Laporan**. Tambahkan **Bahan** bila ingin aktifkan cek Somasi.")
