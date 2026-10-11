import streamlit as st
import pandas as pd
from fuzzywuzzy import process
from io import BytesIO
from PIL import Image
import re

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
        try:
            import pdfplumber
        except (ImportError, OSError):
            st.error(
                "❌ Pembaca PDF tidak dapat dimuat. Gunakan file XLSX/CSV/TSV "
                "atau perbaiki instalasi pdfplumber/cryptography."
            )
            return pd.DataFrame(columns=["Nama Penerima", "Nomor Resi"])

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

def build_name_candidates(df: pd.DataFrame, name_col: str):
    """Siapkan kandidat nama yang dapat dipakai satu kali saat fuzzy matching."""
    candidates = []
    for source_idx, row in df.iterrows():
        display_name = get_str(row.get(name_col)).strip()
        candidates.append({
            "source_idx": source_idx,
            "nama": display_name,
            "nama_norm": normalize_nama(display_name),
            "used": False,
            "row": row,
        })
    return candidates

def fuzzy_match_name_1to1(query, candidates, min_score):
    """Cari kandidat nama terbaik yang belum dipakai."""
    query_norm = normalize_nama(query)
    if not query_norm:
        return None, 0, -1

    available = [
        (candidate["nama_norm"], candidate_idx)
        for candidate_idx, candidate in enumerate(candidates)
        if not candidate["used"] and candidate["nama_norm"]
    ]
    if not available:
        return None, 0, -1

    result = process.extractOne(query_norm, [name for name, _ in available])
    if not result:
        return None, 0, -1

    matched_name, score = result
    if score < min_score:
        return None, int(score), -1

    for name, candidate_idx in available:
        if name == matched_name:
            return candidates[candidate_idx]["nama"], int(score), candidate_idx
    return None, 0, -1

def format_date_for_report(val) -> str:
    if pd.isna(val):
        return ""
    if hasattr(val, "strftime") and not isinstance(val, str):
        return val.strftime("%d/%m/%Y")
    return str(val).strip()

def match_laporan_with_detail(
    df_laporan,
    df_detail,
    *,
    laporan_no_resi_col,
    laporan_nama_debitur_col,
    laporan_penerima_col,
    laporan_status_col,
    laporan_tanggal_col,
    laporan_similarity_col,
    detail_no_resi_col,
    detail_nama_penerima_col,
    detail_penerima_paket_col,
    detail_status_penerima_col,
    detail_tanggal_col,
    threshold_name,
    excluded_statuses=None,
    overwrite_detail=True,
):
    """Isi resi/detail: prioritaskan resi yang ada, lalu fuzzy match nama 1:1."""
    laporan_out = df_laporan.copy()
    detail_work = df_detail.copy()

    # Status dibandingkan tanpa membedakan kapital/spasi agar pilihan filter konsisten.
    excluded_statuses = {
        re.sub(r"\s+", " ", str(status).strip()).upper()
        for status in (excluded_statuses or [])
        if str(status).strip()
    }
    excluded_resi = set()
    excluded_detail_count = 0
    if excluded_statuses:
        detail_status_norm = detail_work[detail_status_penerima_col].apply(
            lambda value: re.sub(r"\s+", " ", get_str(value)).upper()
        )
        excluded_rows = detail_work[detail_status_norm.isin(excluded_statuses)]
        excluded_resi = set(excluded_rows[detail_no_resi_col].apply(normalize_resi)) - {""}
        excluded_detail_count = len(excluded_rows)
        detail_work = detail_work[~detail_status_norm.isin(excluded_statuses)].copy()

    laporan_out["__resi_norm"] = laporan_out[laporan_no_resi_col].apply(normalize_resi)
    detail_work["__resi_norm"] = detail_work[detail_no_resi_col].apply(normalize_resi)
    detail_work = detail_work[detail_work["__resi_norm"] != ""].copy()

    duplicate_detail = detail_work[detail_work.duplicated("__resi_norm", keep=False)].copy()
    detail_unique = detail_work.drop_duplicates("__resi_norm", keep="first").copy()
    detail_lookup = {}
    for detail_idx, detail_row in detail_unique.iterrows():
        detail_lookup.setdefault(detail_row["__resi_norm"], (detail_idx, detail_row))

    needs_name_matching = laporan_out["__resi_norm"].eq("").any()
    detail_candidates = (
        build_name_candidates(detail_unique, detail_nama_penerima_col)
        if needs_name_matching else []
    )
    candidate_by_source_idx = {
        candidate["source_idx"]: candidate_idx
        for candidate_idx, candidate in enumerate(detail_candidates)
    }

    laporan_out[laporan_similarity_col] = pd.to_numeric(
        laporan_out[laporan_similarity_col], errors="coerce"
    ).fillna(0).astype("int64")

    matched_by_resi = 0
    matched_by_name = 0
    unmatched_records = []
    row_number_by_idx = {idx: position + 2 for position, idx in enumerate(laporan_out.index)}

    def fill_detail(report_idx, detail_row, similarity_score=None, fill_resi=False):
        if fill_resi:
            laporan_out.at[report_idx, laporan_no_resi_col] = get_str(
                detail_row.get(detail_no_resi_col)
            )
        detail_values = {
            laporan_penerima_col: get_str(detail_row.get(detail_penerima_paket_col)),
            laporan_status_col: get_str(detail_row.get(detail_status_penerima_col)),
            laporan_tanggal_col: format_date_for_report(detail_row.get(detail_tanggal_col)),
        }
        for laporan_col, detail_value in detail_values.items():
            if overwrite_detail or not get_str(laporan_out.at[report_idx, laporan_col]):
                laporan_out.at[report_idx, laporan_col] = detail_value
        if similarity_score is not None:
            if overwrite_detail or not get_str(laporan_out.at[report_idx, laporan_similarity_col]):
                laporan_out.at[report_idx, laporan_similarity_col] = int(similarity_score)

    # Pass 1: nomor resi yang sudah tersedia selalu mendapat prioritas.
    for idx, row in laporan_out.iterrows():
        resi_norm = row["__resi_norm"]
        if not resi_norm:
            continue
        if resi_norm not in detail_lookup:
            unmatched_records.append({
                "Baris Laporan": row_number_by_idx[idx],
                "Nama Debitur": (
                    get_str(row.get(laporan_nama_debitur_col))
                    if laporan_nama_debitur_col else ""
                ),
                "Nomor Resi": get_str(row.get(laporan_no_resi_col)),
                "Alasan": (
                    "Status Detail Resi dikecualikan oleh filter"
                    if resi_norm in excluded_resi
                    else "Nomor resi tidak ditemukan di Detail Resi"
                ),
            })
            continue

        detail_idx, detail_row = detail_lookup[resi_norm]
        candidate_idx = candidate_by_source_idx.get(detail_idx)
        if candidate_idx is not None:
            detail_candidates[candidate_idx]["used"] = True
        fill_detail(idx, detail_row)
        matched_by_resi += 1

    # Pass 2: nomor resi kosong diisi lewat fuzzy matching nama secara 1:1.
    for idx, row in laporan_out.iterrows():
        if row["__resi_norm"]:
            continue

        nama_debitur = get_str(row.get(laporan_nama_debitur_col))
        matched_name, score, candidate_idx = fuzzy_match_name_1to1(
            nama_debitur,
            detail_candidates,
            threshold_name,
        )
        if not matched_name or candidate_idx < 0:
            unmatched_records.append({
                "Baris Laporan": row_number_by_idx[idx],
                "Nama Debitur": nama_debitur,
                "Nomor Resi": "",
                "Alasan": (
                    f"Nama tidak mencapai ambang similarity {threshold_name}"
                    if nama_debitur else "Nama Debitur kosong"
                ),
            })
            continue

        detail_candidates[candidate_idx]["used"] = True
        fill_detail(
            idx,
            detail_candidates[candidate_idx]["row"],
            similarity_score=score,
            fill_resi=True,
        )
        matched_by_name += 1

    laporan_out = laporan_out.drop(columns=["__resi_norm"])
    return {
        "laporan": laporan_out,
        "duplicate_detail": duplicate_detail,
        "matched_by_resi": matched_by_resi,
        "matched_by_name": matched_by_name,
        "unmatched_records": unmatched_records,
        "excluded_detail_count": excluded_detail_count,
    }

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
    threshold_detail_name = st.sidebar.slider(
        "Ambang Similarity Nama Debitur ↔ Nama Penerima",
        0,
        100,
        80,
        1,
        key="threshold_detail_name",
        help="Dipakai hanya untuk baris laporan yang nomor resinya masih kosong.",
    )

    if file_laporan_detail and file_detail_resi:
        df_laporan_detail = load_file_flex(file_laporan_detail)
        df_detail_resi = load_file_flex(file_detail_resi)

        detail_no_resi_col = find_column(df_detail_resi, "No Resi", "Nomor Resi", "Nomor Resi Pengiriman")
        detail_nama_penerima_col = find_column(
            df_detail_resi,
            "Nama Penerima",
            "Nama Penerima Somasi",
            "Nama Debitur",
            "Nama Konsumen",
        )
        detail_penerima_paket_col = find_column(
            df_detail_resi,
            "Penerima Paket",
            "Penerima",
            "Nama Penerima Paket",
            "Nama Penerima",
        )
        detail_status_penerima_col = find_column(
            df_detail_resi,
            "Status Penerima",
            "Status Pengiriman",
            "Status Resi",
            "Status",
        )
        detail_tanggal_col = find_column(
            df_detail_resi,
            "Tanggal",
            "Tanggal Terima",
            "Tanggal Diterima",
            "Tanggal Pengiriman",
        )

        # Jika judul kolom dari ekspedisi berbeda, pengguna tetap dapat
        # memetakan kolomnya sendiri tanpa perlu mengubah file Excel.
        def choose_detail_column(label, detected_column, widget_key):
            options = [None] + list(df_detail_resi.columns)
            selected_index = options.index(detected_column) if detected_column in options else 0
            return st.sidebar.selectbox(
                label,
                options=options,
                index=selected_index,
                format_func=lambda column: "-- Pilih kolom --" if column is None else column,
                key=widget_key,
            )

        detail_penerima_paket_col = choose_detail_column(
            "Kolom Detail: Penerima Paket",
            detail_penerima_paket_col,
            "detail_penerima_paket_column",
        )
        detail_status_penerima_col = choose_detail_column(
            "Kolom Detail: Status Penerima",
            detail_status_penerima_col,
            "detail_status_penerima_column",
        )
        detail_tanggal_col = choose_detail_column(
            "Kolom Detail: Tanggal",
            detail_tanggal_col,
            "detail_tanggal_column",
        )

        laporan_no_resi_col = find_column(df_laporan_detail, "Nomor Resi Pengiriman", "No Resi", "Nomor Resi")
        laporan_nama_debitur_col = find_column(
            df_laporan_detail,
            "Nama Debitur",
            "Nama Konsumen",
        )
        laporan_penerima_col = find_column(df_laporan_detail, "NAMA PENERIMA EKSPEDISI SOMASI")
        laporan_status_col = find_column(df_laporan_detail, "STATUS PENERIMA EKSPEDISI SOMASI")
        laporan_tanggal_col = find_column(df_laporan_detail, "Tanggal Somasi Diterima")
        laporan_similarity_col = find_column(df_laporan_detail, "Similarity Score (Nama Resi)")

        created_columns = []
        if laporan_no_resi_col is None:
            laporan_no_resi_col = "Nomor Resi Pengiriman"
            df_laporan_detail[laporan_no_resi_col] = ""
            created_columns.append(laporan_no_resi_col)
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
        if laporan_similarity_col is None:
            laporan_similarity_col = "Similarity Score (Nama Resi)"
            df_laporan_detail[laporan_similarity_col] = 0
            created_columns.append(laporan_similarity_col)

        needs_name_matching = df_laporan_detail[laporan_no_resi_col].apply(normalize_resi).eq("").any()

        with st.expander("🔧 Kolom terdeteksi"):
            st.write("Laporan:", list(df_laporan_detail.columns))
            st.write("Detail Resi:", list(df_detail_resi.columns))
            if created_columns:
                st.write("Kolom laporan yang dibuat:", created_columns)

        missing_cols = []
        if detail_no_resi_col is None:
            missing_cols.append("Detail Resi: No Resi")
        if detail_penerima_paket_col is None:
            missing_cols.append("Detail Resi: Penerima Paket")
        if detail_status_penerima_col is None:
            missing_cols.append("Detail Resi: Status Penerima")
        if detail_tanggal_col is None:
            missing_cols.append("Detail Resi: Tanggal")
        if needs_name_matching and laporan_nama_debitur_col is None:
            missing_cols.append("Laporan: Nama Debitur (diperlukan karena ada No Resi kosong)")
        if needs_name_matching and detail_nama_penerima_col is None:
            missing_cols.append("Detail Resi: Nama Penerima (diperlukan karena ada No Resi kosong)")

        if missing_cols:
            st.error(f"Kolom wajib tidak ditemukan: {missing_cols}")
            st.stop()

        status_options = sorted({
            get_str(status)
            for status in df_detail_resi[detail_status_penerima_col]
            if get_str(status)
        }, key=str.upper)
        on_process_statuses = [
            status for status in status_options
            if re.sub(r"\s+", " ", status.strip()).upper() == "ON PROCESS"
        ]
        st.sidebar.subheader("⚙️ Opsi Pembaruan Detail")
        include_on_process = st.sidebar.toggle(
            "Masukkan resi berstatus ON PROCESS",
            value=False,
            help=(
                "Aktifkan bila resi dengan Status Penerima ON PROCESS tetap boleh "
                "dipakai untuk pencocokan dan pembaruan laporan."
            ),
            key="include_on_process",
        )
        other_status_options = [
            status for status in status_options if status not in on_process_statuses
        ]
        other_excluded_statuses = st.sidebar.multiselect(
            "Status lain yang tidak diproses",
            options=other_status_options,
            default=[],
            help=(
                "Resi dengan status terpilih tidak dipakai untuk pencocokan atau "
                "pengisian laporan."
            ),
            key="other_excluded_detail_statuses",
        )
        excluded_detail_statuses = list(other_excluded_statuses)
        if not include_on_process:
            excluded_detail_statuses.extend(on_process_statuses)
        update_mode = st.sidebar.radio(
            "Cara mengisi detail di laporan",
            options=["Isi hanya kolom yang kosong", "Timpa data laporan yang sudah ada"],
            index=0,
            help="Berlaku untuk Nama Penerima, Status Penerima, dan Tanggal Diterima.",
            key="detail_update_mode",
        )
        overwrite_detail = update_mode == "Timpa data laporan yang sudah ada"

        col_lap, col_detail = st.columns(2)
        with col_lap:
            st.subheader("📊 Laporan")
            st.dataframe(make_streamlit_safe(df_laporan_detail))
        with col_detail:
            st.subheader("📦 Detail Resi")
            st.dataframe(make_streamlit_safe(df_detail_resi))

        if needs_name_matching:
            st.info(
                "Baris laporan dengan nomor resi kosong akan dicocokkan 1:1 melalui "
                "**Nama Debitur ↔ Nama Penerima**. Nomor resi yang sudah terisi tetap "
                "diprioritaskan untuk pencocokan langsung."
            )

        if st.button("🚀 Cocokkan No Resi dan Isi Detail"):
            match_result = match_laporan_with_detail(
                df_laporan_detail,
                df_detail_resi,
                laporan_no_resi_col=laporan_no_resi_col,
                laporan_nama_debitur_col=laporan_nama_debitur_col,
                laporan_penerima_col=laporan_penerima_col,
                laporan_status_col=laporan_status_col,
                laporan_tanggal_col=laporan_tanggal_col,
                laporan_similarity_col=laporan_similarity_col,
                detail_no_resi_col=detail_no_resi_col,
                detail_nama_penerima_col=detail_nama_penerima_col,
                detail_penerima_paket_col=detail_penerima_paket_col,
                detail_status_penerima_col=detail_status_penerima_col,
                detail_tanggal_col=detail_tanggal_col,
                threshold_name=threshold_detail_name,
                excluded_statuses=excluded_detail_statuses,
                overwrite_detail=overwrite_detail,
            )
            df_laporan_out = match_result["laporan"]
            duplicate_detail = match_result["duplicate_detail"]
            matched_by_resi = match_result["matched_by_resi"]
            matched_by_name = match_result["matched_by_name"]
            unmatched_records = match_result["unmatched_records"]
            excluded_detail_count = match_result["excluded_detail_count"]
            matched_count = matched_by_resi + matched_by_name

            st.subheader("✅ Hasil Pencocokan No Resi dan Detail")
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Baris Laporan", len(df_laporan_detail))
            m2.metric("Berhasil Cocok", matched_count)
            m3.metric("Via No Resi", matched_by_resi)
            m4.metric("Via Nama", matched_by_name)
            m5.metric("Tidak Cocok", len(unmatched_records))

            if excluded_detail_count:
                st.info(
                    f"{excluded_detail_count} baris Detail Resi tidak diproses sesuai filter status: "
                    f"{', '.join(excluded_detail_statuses)}."
                )

            if not duplicate_detail.empty:
                st.warning("Ada No Resi duplikat di Detail Resi. Aplikasi memakai baris pertama untuk pengisian laporan.")
                duplicate_cols = [
                    c for c in [
                        detail_no_resi_col,
                        detail_nama_penerima_col,
                        detail_penerima_paket_col,
                        detail_status_penerima_col,
                        detail_tanggal_col,
                    ] if c in duplicate_detail.columns
                ]
                st.dataframe(make_streamlit_safe(duplicate_detail[duplicate_cols]))

            if unmatched_records:
                st.warning("Beberapa baris laporan belum dapat dicocokkan dengan Detail Resi.")
                st.dataframe(pd.DataFrame(unmatched_records), hide_index=True)

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

st.sidebar.header("⚙️ Mode Pencocokan")
only_somasi_mode = st.sidebar.checkbox("🔖 Hanya cocokkan **Nomor Surat Somasi** (berdasar Nama / No. Kontrak)")

st.sidebar.header("📤 Upload Berkas")
if only_somasi_mode:
    file_resi = None
    st.sidebar.caption("File Resi tidak diperlukan pada mode Hanya Somasi.")
else:
    file_resi = st.sidebar.file_uploader(
        "Resi (xlsx/csv/tsv/pdf)",
        type=["xlsx", "csv", "tsv", "pdf"],
    )
file_laporan = st.sidebar.file_uploader("Laporan (xlsx/csv/tsv)",     type=["xlsx", "csv", "tsv"])
files_bahan  = st.sidebar.file_uploader(
    "Bahan (boleh banyak, opsional) – berisi '3. NOMOR SURAT SOMASI' & '4. NAMA KONSUMEN'",
    type=["xlsx","csv","tsv"], accept_multiple_files=True
)

st.sidebar.markdown("---")
somasi_priority = st.sidebar.radio(
    "Prioritas pencocokan Somasi (saat mengisi dari Bahan):",
    options=["Utamakan No. Kontrak", "Utamakan Nama"],
    horizontal=False,
    index=0
)

threshold_name = 80
if not only_somasi_mode:
    threshold_name = st.sidebar.slider("Ambang Similarity Nama (0-100)", 0, 100, 80, 1)
threshold_somasi      = st.sidebar.slider("Ambang Similarity Somasi (0-100)", 0, 100, 90, 1)
threshold_nama_bahan  = st.sidebar.slider("Ambang Fuzzy Nama (isi Somasi dari Bahan)", 0, 100, 85, 1)

# =========================================
# Mode penuh: Resi + Laporan. Mode Hanya Somasi: cukup Laporan.
# =========================================
if file_laporan and (file_resi or only_somasi_mode):
    df_resi = load_resi_file(file_resi) if file_resi else pd.DataFrame()
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
        if only_somasi_mode:
            st.write("Resi: dilewati pada mode Hanya Somasi")
        else:
            st.write("Resi:", list(df_resi.columns))
        st.write("Laporan:", list(df_laporan.columns))
        if not df_bahan.empty:
            for nm, cols in bahan_cols_debug:
                st.write(f"Bahan ({nm}):", cols)
            st.write("Bahan (Gabungan):", list(df_bahan.columns))
        else:
            st.write("Bahan: (tidak diupload)")

    # --- Validasi minimum sesuai mode
    missing_min = []
    if only_somasi_mode:
        if "Nama Debitur" not in df_laporan.columns and "Nomor Kontrak" not in df_laporan.columns:
            missing_min.append("Laporan: Nama Debitur atau Nomor Kontrak")
    else:
        if "Nama Debitur" not in df_laporan.columns:
            missing_min.append("Laporan: Nama Debitur")
        if "Nama Penerima" not in df_resi.columns:
            missing_min.append("Resi: Nama Penerima")

    if missing_min:
        st.error(f"Kolom wajib tidak ditemukan (mode minimal): {missing_min}")
        st.stop()

    # --- Cast dasar
    if 'Nama Debitur' not in df_laporan.columns:
        df_laporan['Nama Debitur'] = pd.Series("", index=df_laporan.index, dtype='string')
    df_laporan['Nama Debitur'] = df_laporan['Nama Debitur'].astype('string').fillna("")
    if not only_somasi_mode:
        df_resi['Nama Penerima'] = df_resi['Nama Penerima'].astype('string').fillna("")
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
        if only_somasi_mode:
            st.caption("Dilewati pada mode Hanya Somasi")
        else:
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
            resi_candidates = build_name_candidates(df_resi, 'Nama Penerima')

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
        with colA:
            st.caption("Resi")
            if only_somasi_mode:
                st.write("—")
            else:
                cek_duplikat(df_resi, 'Nama Penerima', 'duplikat_nama_penerima_resi.xlsx')
        with colB: st.caption("Laporan");cek_duplikat(df_laporan,'Nama Debitur',        'duplikat_nama_debitur_laporan.xlsx')
        with colC:
            st.caption("Bahan")
            if not df_bahan.empty:
                cek_duplikat(df_bahan, 'Nomor Surat Somasi',  'duplikat_nomor_surat_somasi_bahan.xlsx')
            else:
                st.write("—")

else:
    if only_somasi_mode:
        st.info("📎 Upload **Laporan** dan **Bahan** untuk mencocokkan Nomor Surat Somasi. File Resi tidak diperlukan.")
    else:
        st.info("📎 Upload minimal **Resi** dan **Laporan**. Tambahkan **Bahan** bila ingin aktifkan cek Somasi.")
