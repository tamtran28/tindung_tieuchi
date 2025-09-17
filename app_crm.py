# -------------------------------------------------------------
# Streamlit: Báo cáo Phân tích Tín dụng (CRM4 / CRM32)
# Tác giả: ChatGPT — chuyển đổi từ script notebook sang web app
# Yêu cầu môi trường: streamlit, pandas, numpy, openpyxl, xlrd (đọc .xls)
# Chạy:  streamlit run app_streamlit_crm_dashboard.py
# -------------------------------------------------------------

import io
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# ============================ UI & LAYOUT ============================ #
st.set_page_config(
    page_title="Báo cáo Phân tích Tín dụng",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Báo cáo Phân tích Tín dụng — CRM4 / CRM32")
st.caption(
    "Tải dữ liệu → Chọn chi nhánh → Nhấn **Chạy phân tích** → Xem bảng và **Xuất Excel**.\n"
    "Hỗ trợ file .xls / .xlsx. Một số bảng là **tuỳ chọn** (Mục 17, 55/56, 57, Giải ngân tiền mặt).")

# ============================ HELPERS ============================ #

def read_excel_any(file) -> pd.DataFrame:
    """Đọc Excel từ streamlit uploader (hỗ trợ .xls/.xlsx)."""
    if file is None:
        return pd.DataFrame()
    name = getattr(file, "name", "").lower()
    data = file.read()
    bio = io.BytesIO(data)
    try:
        if name.endswith(".xls"):
            # pandas>=2 cần xlrd để đọc .xls
            return pd.read_excel(bio, engine="xlrd")
        return pd.read_excel(bio)
    except Exception as e:
        st.error(f"Không đọc được file **{name}**: {e}")
        return pd.DataFrame()


def read_excel_multi(files: List) -> List[pd.DataFrame]:
    out = []
    for f in files or []:
        df = read_excel_any(f)
        if not df.empty:
            out.append(df)
    return out


def safe_str(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip()


def to_str_intlike(series: pd.Series) -> pd.Series:
    """Coerce số dạng object/float → int → str (giữ nguyên NaN)."""
    s = pd.to_numeric(series, errors="coerce")
    s = s.dropna().astype("int64").astype(str)
    return s


def ensure_cols(df: pd.DataFrame, cols: List[str]) -> bool:
    missing = [c for c in cols if c not in df.columns]
    if missing:
        st.warning(f"Thiếu cột: {', '.join(missing)}")
        return False
    return True

# ============================ SIDEBAR ============================ #
with st.sidebar:
    st.header("⚙️ Cài đặt & Tải tệp")
    crm4_files = st.file_uploader(
        "CRM4 – Dư nợ theo tài sản đảm bảo (có thể nhiều file)",
        type=["xls", "xlsx"], accept_multiple_files=True,
        help="Ví dụ: CRM4_Du_no_theo_tai_san_dam_bao_ALL*.xls/.xlsx",
    )
    crm32_files = st.file_uploader(
        "CRM32 – RPT_CRM_32 (có thể nhiều file)",
        type=["xls", "xlsx"], accept_multiple_files=True,
    )

    st.markdown("**Bảng mã (bắt buộc/khuyến nghị):**")
    df_muc_dich_file = st.file_uploader("CODE_MDSDV4.xlsx (mã mục đích vay)", type=["xls", "xlsx"])
    df_code_tsbd_file = st.file_uploader("CODE_LOAI TSBD.xlsx (mã loại TSBĐ)", type=["xls", "xlsx"])

    st.markdown("**Các bảng bổ sung (tuỳ chọn):**")
    file_giai_ngan_tm = st.file_uploader("Giai_ngan_tien_mat_1_ty.xls/xlsx", type=["xls", "xlsx"])
    file_muc17 = st.file_uploader("MUC17.xlsx (Tài sản – địa bàn)", type=["xls", "xlsx"])
    file_muc55 = st.file_uploader("Muc55_*.xlsx (Tất toán)", type=["xls", "xlsx"])
    file_muc56 = st.file_uploader("Muc56_*.xlsx (Giải ngân)", type=["xls", "xlsx"])
    file_muc57 = st.file_uploader("Muc57_*.xlsx (Chậm trả)", type=["xls", "xlsx"])

    st.divider()
    chi_nhanh = st.text_input(
        "Nhập tên chi nhánh hoặc mã SOL để lọc (ví dụ: HANOI hoặc 001)",
        value="",
    ).upper().strip()

    ngay_danh_gia = st.date_input("Ngày đánh giá", value=pd.to_datetime("2025-08-31").date())

    dia_ban_kt_input = st.text_area(
        "Tên tỉnh/thành của đơn vị đang kiểm toán (phân cách dấu phẩy)",
        value="Hồ Chí Minh, Long An",
        help="Dùng cho kiểm tra TSBĐ khác địa bàn (Mục 17)",
    )
    dia_ban_kt = [t.strip().lower() for t in dia_ban_kt_input.split(',') if t.strip()]

    run = st.button("🚀 Chạy phân tích", use_container_width=True, type="primary")

# ============================ CORE LOGIC ============================ #
@st.cache_data(show_spinner=False)
def load_and_concat(files: List) -> pd.DataFrame:
    dfs = read_excel_multi(files)
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True)


def add_loai_ts(df_crm4: pd.DataFrame, df_code_tsbd: pd.DataFrame) -> pd.DataFrame:
    if df_crm4.empty:
        return df_crm4
    # Chuẩn hoá & ghép mã loại TSBĐ
    if not df_code_tsbd.empty:
        use_cols = []
        # Chuẩn tên cột theo yêu cầu script gốc: 'CODE CAP 2' -> 'CAP_2', 'CODE' -> 'LOAI_TS'
        if "CODE CAP 2" in df_code_tsbd.columns and "CODE" in df_code_tsbd.columns:
            tmp = df_code_tsbd[["CODE CAP 2", "CODE"]].copy()
            tmp.columns = ["CAP_2", "LOAI_TS"]
            use_cols = ["CAP_2", "LOAI_TS"]
        elif "CAP_2" in df_code_tsbd.columns and "LOAI_TS" in df_code_tsbd.columns:
            tmp = df_code_tsbd[["CAP_2", "LOAI_TS"]].copy()
            use_cols = ["CAP_2", "LOAI_TS"]
        else:
            st.warning("Bảng mã TSBĐ không có cột 'CODE CAP 2'/'CAP_2' và 'CODE'/'LOAI_TS'. Bỏ qua ánh xạ.")
            tmp = pd.DataFrame()

        if not tmp.empty:
            df_crm4 = df_crm4.merge(tmp.drop_duplicates(), how="left", on="CAP_2")
            # Gán 'Không TS' nếu thiếu mã
            df_crm4["LOAI_TS"] = df_crm4.apply(
                lambda r: "Không TS" if pd.isna(r.get("CAP_2")) or str(r.get("CAP_2", "")).strip() == "" else r.get("LOAI_TS"),
                axis=1,
            )
            # Ghi chú 'MỚI' nếu có CAP_2 nhưng không tìm thấy loại TS
            df_crm4["GHI_CHU_TSBD"] = df_crm4.apply(
                lambda r: "MỚI" if str(r.get("CAP_2", "")).strip() != "" and pd.isna(r.get("LOAI_TS")) else "",
                axis=1,
            )
    return df_crm4


def add_muc_dich_crm32(df_crm32: pd.DataFrame, df_muc_dich: pd.DataFrame) -> pd.DataFrame:
    if df_crm32.empty:
        return df_crm32
    if not df_muc_dich.empty:
        # Chuẩn tên: CODE_MDSDV4 -> MUC_DICH_VAY_CAP_4, GROUP -> MUC DICH
        cols_ok = [c in df_muc_dich.columns for c in ["CODE_MDSDV4", "GROUP"]]
        if all(cols_ok):
            md = df_muc_dich[["CODE_MDSDV4", "GROUP"]].drop_duplicates().copy()
            md.columns = ["MUC_DICH_VAY_CAP_4", "MUC DICH"]
            df_crm32 = df_crm32.merge(md, how="left", on="MUC_DICH_VAY_CAP_4")
            df_crm32["MUC DICH"] = df_crm32["MUC DICH"].fillna("(blank)")
        else:
            st.warning("Bảng CODE_MDSDV4 thiếu cột 'CODE_MDSDV4'/'GROUP'. Bỏ qua ánh xạ mục đích vay.")
    return df_crm32


def build_pivots(df_crm4: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Trả về: (pivot_ts, pivot_no, pivot_merge, pivot_final)."""
    required = ["CIF_KH_VAY", "LOAI", "LOAI_TS", "TS_KW_VND", "DU_NO_PHAN_BO_QUY_DOI"]
    if not ensure_cols(df_crm4, required):
        return (pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame())

    df_vay = df_crm4[~df_crm4["LOAI"].isin(["Bao lanh", "LC"])].copy()

    pivot_ts = (
        df_vay.pivot_table(
            index="CIF_KH_VAY",
            columns="LOAI_TS",
            values="TS_KW_VND",
            aggfunc="sum",
            fill_value=0,
        )
        .add_suffix(" (Giá trị TS)")
        .reset_index()
    )

    pivot_no = (
        df_vay.pivot_table(
            index="CIF_KH_VAY",
            columns="LOAI_TS",
            values="DU_NO_PHAN_BO_QUY_DOI",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    pivot_merge = pivot_no.merge(pivot_ts, on="CIF_KH_VAY", how="left")

    # Tính tổng DƯ NỢ & GIÁ TRỊ TS theo cột
    debt_cols = [c for c in pivot_no.columns if c != "CIF_KH_VAY"]
    ts_cols = [c for c in pivot_ts.columns if c != "CIF_KH_VAY"]
    pivot_merge["DƯ NỢ"] = pivot_merge[debt_cols].sum(axis=1) if debt_cols else 0
    pivot_merge["GIÁ TRỊ TS"] = pivot_merge[ts_cols].sum(axis=1) if ts_cols else 0

    # Thêm info khách hàng (nếu có)
    info_cols = [c for c in ["CIF_KH_VAY", "TEN_KH_VAY", "CUSTTPCD", "NHOM_NO"] if c in df_crm4.columns]
    df_info = df_crm4[info_cols].drop_duplicates(subset="CIF_KH_VAY") if info_cols else pd.DataFrame({"CIF_KH_VAY": pivot_merge["CIF_KH_VAY"]})
    pivot_final = df_info.merge(pivot_merge, on="CIF_KH_VAY", how="left")
    pivot_final = pivot_final.reset_index(drop=True)
    pivot_final.insert(0, "STT", np.arange(1, len(pivot_final) + 1))

    # Sắp xếp cột hiển thị
    debt_only = sorted([c for c in debt_cols if "(Giá trị TS)" not in c])
    ts_only = sorted(ts_cols)
    ordered = (["STT"] + [c for c in ["CUSTTPCD", "CIF_KH_VAY", "TEN_KH_VAY", "NHOM_NO"] if c in pivot_final.columns]
               + debt_only + ts_only + ["DƯ NỢ", "GIÁ TRỊ TS"])
    pivot_final = pivot_final[[c for c in ordered if c in pivot_final.columns]]

    return pivot_ts, pivot_no, pivot_merge, pivot_final


def enrich_crm32(df_crm32: pd.DataFrame) -> Tuple[pd.DataFrame, np.ndarray, np.ndarray]:
    if df_crm32.empty:
        return df_crm32, np.array([]), np.array([])

    df_crm32 = df_crm32.copy()
    if "CAP_PHE_DUYET" in df_crm32.columns:
        df_crm32["MA_PHE_DUYET"] = safe_str(df_crm32["CAP_PHE_DUYET"]).str.split("-").str[0].str.zfill(2)
    else:
        df_crm32["MA_PHE_DUYET"] = ""

    if "CUSTSEQLN" in df_crm32.columns:
        df_crm32["CUSTSEQLN"] = safe_str(df_crm32["CUSTSEQLN"])  # normalize

    ma_cap_c = [f"{i:02d}" for i in range(1, 8)] + [f"{i:02d}" for i in range(28, 32)]
    list_cif_cap_c = df_crm32[df_crm32["MA_PHE_DUYET"].isin(ma_cap_c)]["CUSTSEQLN"].unique()

    list_co_cau = [
        "ACOV1", "ACOV3", "ATT01", "ATT02", "ATT03", "ATT04",
        "BCOV1", "BCOV2", "BTT01", "BTT02", "BTT03",
        "CCOV2", "CCOV3", "CTT03", "RCOV3", "RTT03",
    ]
    if "SCHEME_CODE" in df_crm32.columns:
        cif_co_cau = df_crm32[df_crm32["SCHEME_CODE"].isin(list_co_cau)]["CUSTSEQLN"].unique()
    else:
        cif_co_cau = np.array([])

    return df_crm32, list_cif_cap_c, cif_co_cau


def pivot_muc_dich(df_crm32: pd.DataFrame) -> pd.DataFrame:
    if df_crm32.empty:
        return pd.DataFrame()
    if not ensure_cols(df_crm32, ["CUSTSEQLN", "MUC DICH", "DU_NO_QUY_DOI"]):
        return pd.DataFrame()
    p = (
        df_crm32.pivot_table(index="CUSTSEQLN", columns="MUC DICH", values="DU_NO_QUY_DOI", aggfunc="sum", fill_value=0)
        .reset_index()
    )
    p["DƯ NỢ CRM32"] = p.drop(columns=["CUSTSEQLN"]).sum(axis=1)
    return p


def add_flags_and_joins(
    pivot_final: pd.DataFrame,
    pivot_crm32_by_mucdich: pd.DataFrame,
    df_crm4_filtered: pd.DataFrame,
    df_crm32_filtered: pd.DataFrame,
    list_cif_cap_c: np.ndarray,
    cif_co_cau: np.ndarray,
    giai_ngan_tm: Optional[pd.DataFrame],
    ngay_danh_gia: pd.Timestamp,
    df_muc17: Optional[pd.DataFrame],
    dia_ban_kt: List[str],
    df_muc55: Optional[pd.DataFrame],
    df_muc56: Optional[pd.DataFrame],
    df_muc57: Optional[pd.DataFrame],
) -> Tuple[pd.DataFrame, dict]:
    """Bổ sung các cờ & ghép các bảng phụ, trả về pivot_full và dict[kpi]."""
    if pivot_final.empty:
        return pivot_final, {}

    piv = pivot_final.copy()

    # Ghép CRM32 theo mục đích vay
    if not pivot_crm32_by_mucdich.empty:
        p32 = pivot_crm32_by_mucdich.rename(columns={"CUSTSEQLN": "CIF_KH_VAY"})
        piv = piv.merge(p32, on="CIF_KH_VAY", how="left").fillna(0)

    # Lệch dư nợ & bổ sung (blank) từ CRM4 (không gồm Cho vay/Bảo lãnh/LC)
    if "DƯ NỢ" in piv.columns and "DƯ NỢ CRM32" in piv.columns:
        piv["LECH"] = piv["DƯ NỢ"] - piv["DƯ NỢ CRM32"]
    else:
        piv["LECH"] = 0

    df_blank = df_crm4_filtered[~df_crm4_filtered["LOAI"].isin(["Cho vay", "Bao lanh", "LC"])].copy()
    if not df_blank.empty and "DU_NO_PHAN_BO_QUY_DOI" in df_blank.columns:
        du_no_bosung = (
            df_blank.groupby("CIF_KH_VAY", as_index=False)["DU_NO_PHAN_BO_QUY_DOI"].sum().rename(columns={"DU_NO_PHAN_BO_QUY_DOI": "(blank)"})
        )
        piv = piv.merge(du_no_bosung, on="CIF_KH_VAY", how="left")
        piv["(blank)"] = piv["(blank)"].fillna(0)
        if "DƯ NỢ CRM32" in piv.columns:
            piv["DƯ NỢ CRM32"] = piv["DƯ NỢ CRM32"] + piv["(blank)"]
        piv["LECH"] = piv.get("DƯ NỢ", 0) - piv.get("DƯ NỢ CRM32", 0)

    # Nợ nhóm 2 / Nợ xấu
    if "NHOM_NO" in piv.columns:
        piv["Nợ nhóm 2"] = piv["NHOM_NO"].apply(lambda x: "x" if str(x).strip() == "2" else "")
        piv["Nợ xấu"] = piv["NHOM_NO"].apply(lambda x: "x" if str(x).strip() in ["3", "4", "5"] else "")

    # Phê duyệt cấp C / Cơ cấu
    cif_in_cap_c = set(map(str, list_cif_cap_c or []))
    cif_in_cocau = set(map(str, cif_co_cau or []))
    piv["Chuyên gia PD cấp C duyệt"] = piv["CIF_KH_VAY"].astype(str).apply(lambda x: "x" if x in cif_in_cap_c else "")
    piv["NỢ CƠ_CẤU"] = piv["CIF_KH_VAY"].astype(str).apply(lambda x: "x" if x in cif_in_cocau else "")

    # Dư nợ Bảo lãnh & LC
    def _sum_by_loai(loai: str, newcol: str):
        tmp = df_crm4_filtered[df_crm4_filtered["LOAI"] == loai]
        if not tmp.empty:
            s = tmp.groupby("CIF_KH_VAY", as_index=False)["DU_NO_PHAN_BO_QUY_DOI"].sum().rename(columns={"DU_NO_PHAN_BO_QUY_DOI": newcol})
            return s
        return pd.DataFrame({"CIF_KH_VAY": [], newcol: []})

    piv = piv.merge(_sum_by_loai("Bao lanh", "DƯ_NỢ_BẢO_LÃNH"), on="CIF_KH_VAY", how="left")
    piv = piv.merge(_sum_by_loai("LC", "DƯ_NỢ_LC"), on="CIF_KH_VAY", how="left")
    for c in ["DƯ_NỢ_BẢO_LÃNH", "DƯ_NỢ_LC"]:
        if c in piv.columns:
            piv[c] = piv[c].fillna(0)

    # Giải ngân tiền mặt 1 tỷ (tuỳ chọn)
    if giai_ngan_tm is not None and not giai_ngan_tm.empty and "FORACID" in giai_ngan_tm.columns:
        df_crm32_filtered = df_crm32_filtered.copy()
        for c in ["KHE_UOC", "CUSTSEQLN"]:
            if c in df_crm32_filtered.columns:
                df_crm32_filtered[c] = safe_str(df_crm32_filtered[c])
        giai_ngan_tm["FORACID"] = safe_str(giai_ngan_tm["FORACID"])  # chuẩn mã
        ds_cif_tm = df_crm32_filtered[df_crm32_filtered.get("KHE_UOC", "").isin(giai_ngan_tm["FORACID"])]["CUSTSEQLN"].unique()
        ds_cif_tm = set(map(str, ds_cif_tm))
        piv["GIẢI_NGÂN_TIEN_MAT"] = piv["CIF_KH_VAY"].astype(str).apply(lambda x: "x" if x in ds_cif_tm else "")

    # Cầm cố tại TCTD khác (CAP_2 chứa 'TCTD')
    cc_flag = set(
        df_crm4_filtered[df_crm4_filtered.get("CAP_2", "").astype(str).str.contains("TCTD", case=False, na=False)][
            "CIF_KH_VAY"
        ].astype(str)
    )
    piv["Cầm cố tại TCTD khác"] = piv["CIF_KH_VAY"].astype(str).apply(lambda x: "x" if x in cc_flag else "")

    # Top 10 KHCN & KHDN theo DƯ NỢ (nếu có CUSTTPCD)
    if "CUSTTPCD" in piv.columns and "DƯ NỢ" in piv.columns:
        top_khcn = set(piv[piv["CUSTTPCD"] == "Ca nhan"].nlargest(10, "DƯ NỢ")["CIF_KH_VAY"].astype(str))
        top_khdn = set(piv[piv["CUSTTPCD"] == "Doanh nghiep"].nlargest(10, "DƯ NỢ")["CIF_KH_VAY"].astype(str))
        piv["Top 10 dư nợ KHCN"] = piv["CIF_KH_VAY"].astype(str).apply(lambda x: "x" if x in top_khcn else "")
        piv["Top 10 dư nợ KHDN"] = piv["CIF_KH_VAY"].astype(str).apply(lambda x: "x" if x in top_khdn else "")

    # Quá hạn định giá R34 (BĐS/MMTB/PTVT)
    ndg = pd.to_datetime(pd.Timestamp(ngay_danh_gia))
    r34 = ["BĐS", "MMTB", "PTVT"]
    df_r = df_crm4_filtered.copy()
    if "VALUATION_DATE" in df_r.columns:
        df_r["VALUATION_DATE"] = pd.to_datetime(df_r["VALUATION_DATE"], errors="coerce")
        mask_r34 = df_r.get("LOAI_TS", "").isin(r34)
        df_r.loc[mask_r34, "SO_NGAY_QUA_HAN"] = (ndg - df_r.loc[mask_r34, "VALUATION_DATE"]).dt.days - 365
        df_r.loc[df_r["LOAI_TS"] == "BĐS", "SO_THANG_QUA_HAN"] = (
            ((ndg - df_r.loc[df_r["LOAI_TS"] == "BĐS", "VALUATION_DATE"]).dt.days / 31) - 18
        )
        df_r.loc[df_r["LOAI_TS"].isin(["MMTB", "PTVT"]), "SO_THANG_QUA_HAN"] = (
            ((ndg - df_r.loc[df_r["LOAI_TS"].isin(["MMTB", "PTVT"]), "VALUATION_DATE"]).dt.days / 31) - 12
        )
        cif_quahan = set(df_r[df_r.get("SO_NGAY_QUA_HAN", 0) > 30]["CIF_KH_VAY"].astype(str))
        piv["KH có TSBĐ quá hạn định giá"] = piv["CIF_KH_VAY"].astype(str).apply(lambda x: "X" if x in cif_quahan else "")

    # Mục 17 – cảnh báo TSBĐ khác địa bàn
    if df_muc17 is not None and not df_muc17.empty:
        if all(c in df_muc17.columns for c in ["C01", "C02", "C19"]) and "SECU_SRL_NUM" in df_crm4_filtered.columns:
            ds_secu = df_crm4_filtered["SECU_SRL_NUM"].dropna().unique()
            df_17 = df_muc17[df_muc17["C01"].isin(ds_secu)]
            df_bds = df_17[df_17["C02"].astype(str).str.strip() == "Bat dong san"].copy()

            def extract_tinh_thanh(addr: str) -> str:
                if not isinstance(addr, str) or not addr:
                    return ""
                parts = addr.split(',')
                return parts[-1].strip().lower() if parts else ""

            df_bds["TINH_TP_TSBD"] = df_bds["C19"].apply(extract_tinh_thanh)
            df_bds["CANH_BAO_TS_KHAC_DIABAN"] = df_bds["TINH_TP_TSBD"].apply(
                lambda x: "x" if x and x.strip().lower() not in dia_ban_kt else ""
            )
            ma_ts_canh_bao = set(df_bds[df_bds["CANH_BAO_TS_KHAC_DIABAN"] == "x"]["C01"].unique())
            cif_canh_bao = set(
                df_crm4_filtered[df_crm4_filtered["SECU_SRL_NUM"].isin(ma_ts_canh_bao)]["CIF_KH_VAY"].astype(str).unique()
            )
            piv["KH có TSBĐ khác địa bàn"] = piv["CIF_KH_VAY"].astype(str).apply(lambda x: "x" if x in cif_canh_bao else "")
        else:
            st.info("Mục 17: thiếu các cột bắt buộc (C01, C02, C19) hoặc CRM4 thiếu SECU_SRL_NUM — bỏ qua kiểm tra địa bàn.")

    # Tiêu chí 3 – cùng ngày có cả Giải ngân và Tất toán (Mục 55/56)
    if df_muc55 is not None and not df_muc55.empty:
        # Chuẩn Mục 55 (TT)
        need55 = ["CUSTSEQLN", "NMLOC", "KHE_UOC", "SOTIENGIAINGAN", "NGAYGN", "NGAYDH", "NGAY_TT", "LOAITIEN"]
        if all(c in df_muc55.columns for c in need55):
            df_tt = df_muc55[need55].copy()
            df_tt.columns = [
                "CIF", "TEN_KHACH_HANG", "KHE_UOC", "SO_TIEN_GIAI_NGAN_VND",
                "NGAY_GIAI_NGAN", "NGAY_DAO_HAN", "NGAY_TT", "LOAI_TIEN_HD",
            ]
            df_tt["GIAI_NGAN_TT"] = "Tất toán"
            df_tt["NGAY"] = pd.to_datetime(df_tt["NGAY_TT"], errors="coerce")
        else:
            df_tt = pd.DataFrame(columns=["CIF", "GIAI_NGAN_TT", "NGAY"])  # rỗng an toàn
    else:
        df_tt = pd.DataFrame(columns=["CIF", "GIAI_NGAN_TT", "NGAY"])  # rỗng an toàn

    if df_muc56 is not None and not df_muc56.empty and all(c in df_muc56.columns for c in ["CIF", "TEN_KHACH_HANG", "KHE_UOC", "SO_TIEN_GIAI_NGAN_VND", "NGAY_GIAI_NGAN", "NGAY_DAO_HAN", "LOAI_TIEN_HD"]):
        df_gn = df_muc56[["CIF", "TEN_KHACH_HANG", "KHE_UOC", "SO_TIEN_GIAI_NGAN_VND", "NGAY_GIAI_NGAN", "NGAY_DAO_HAN", "LOAI_TIEN_HD"]].copy()
        df_gn["GIAI_NGAN_TT"] = "Giải ngân"
        df_gn["NGAY_GIAI_NGAN"] = pd.to_datetime(df_gn["NGAY_GIAI_NGAN"], errors="coerce")
        df_gn["NGAY_DAO_HAN"] = pd.to_datetime(df_gn["NGAY_DAO_HAN"], errors="coerce")
        df_gn["NGAY"] = df_gn["NGAY_GIAI_NGAN"]
    else:
        df_gn = pd.DataFrame(columns=["CIF", "GIAI_NGAN_TT", "NGAY"])  # rỗng an toàn

    df_gop = pd.concat([df_tt, df_gn], ignore_index=True)
    df_gop = df_gop[df_gop["NGAY"].notna()].copy()

    if not df_gop.empty:
        df_count = df_gop.groupby(["CIF", "NGAY", "GIAI_NGAN_TT"]).size().unstack(fill_value=0).reset_index()
        df_count["CO_CA_GN_VA_TT"] = ((df_count.get("Giải ngân", 0) > 0) & (df_count.get("Tất toán", 0) > 0)).astype(int)
        ds_ca_gn_tt = set(df_count[df_count["CO_CA_GN_VA_TT"] == 1]["CIF"].astype(str).unique())
        piv["KH có cả GNG và TT trong 1 ngày"] = piv["CIF_KH_VAY"].astype(str).apply(lambda x: "x" if x in ds_ca_gn_tt else "")
    else:
        df_count = pd.DataFrame()

    # Tiêu chí 4 – Chậm trả (Mục 57)
    if df_muc57 is not None and not df_muc57.empty and all(c in df_muc57.columns for c in ["CIF_ID", "NGAY_DEN_HAN_TT", "NGAY_THANH_TOAN"]):
        d = df_muc57.copy()
        d["NGAY_DEN_HAN_TT"] = pd.to_datetime(d["NGAY_DEN_HAN_TT"], errors="coerce")
        d["NGAY_THANH_TOAN"] = pd.to_datetime(d["NGAY_THANH_TOAN"], errors="coerce")
        d = d[d["NGAY_DEN_HAN_TT"].dt.year.between(2023, 2025)]
        d["NGAY_THANH_TOAN_FILL"] = d["NGAY_THANH_TOAN"].fillna(pd.to_datetime(ngay_danh_gia))
        d["SO_NGAY_CHAM_TRA"] = (d["NGAY_THANH_TOAN_FILL"] - d["NGAY_DEN_HAN_TT"]).dt.days

        piv2 = piv.rename(columns={"CIF_KH_VAY": "CIF_ID"})
        d["CIF_ID"] = safe_str(d["CIF_ID"])  # chuẩn kiểu
        piv2["CIF_ID"] = safe_str(piv2["CIF_ID"])  # chuẩn kiểu
        d = d.merge(piv2[["CIF_ID", "DƯ NỢ", "NHOM_NO"]], on="CIF_ID", how="left")
        d = d[d["NHOM_NO"] == 1].copy() if "NHOM_NO" in d.columns else d

        def cap_cham_tra(days: Optional[float]):
            if pd.isna(days):
                return None
            if days >= 10:
                return ">=10"
            if days >= 4:
                return "4-9"
            if days > 0:
                return "<4"
            return None

        d["CAP_CHAM_TRA"] = d["SO_NGAY_CHAM_TRA"].apply(cap_cham_tra)
        d = d.dropna(subset=["CAP_CHAM_TRA"]).copy()
        d["NGAY"] = d["NGAY_DEN_HAN_TT"].dt.date
        d.sort_values(["CIF_ID", "NGAY", "CAP_CHAM_TRA"], key=lambda s: s.map({">=10": 0, "4-9": 1, "<4": 2}), inplace=True)
        d_unique = d.drop_duplicates(subset=["CIF_ID", "NGAY"], keep="first").copy()
        dem = d_unique.groupby(["CIF_ID", "CAP_CHAM_TRA"]).size().unstack(fill_value=0)
        dem["KH Phát sinh chậm trả > 10 ngày"] = np.where(dem.get(">=10", 0) > 0, "x", "")
        dem["KH Phát sinh chậm trả 4-9 ngày"] = np.where((dem.get(">=10", 0) == 0) & (dem.get("4-9", 0) > 0), "x", "")
        piv = piv.merge(dem[["KH Phát sinh chậm trả > 10 ngày", "KH Phát sinh chậm trả 4-9 ngày"]], left_on="CIF_KH_VAY", right_index=True, how="left")
        piv["KH Phát sinh chậm trả > 10 ngày"] = piv["KH Phát sinh chậm trả > 10 ngày"].fillna("")
        piv["KH Phát sinh chậm trả 4-9 ngày"] = piv["KH Phát sinh chậm trả 4-9 ngày"].fillna("")
        df_delay = d  # để xuất Excel
    else:
        df_delay = pd.DataFrame()

    # KPIs nhanh
    kpi = {
        "Số KH": int(piv.shape[0]),
        "Tổng dư nợ": float(piv.get("DƯ NỢ", pd.Series(dtype=float)).sum()) if "DƯ NỢ" in piv.columns else 0.0,
        "Lệch dương (count)": int((piv.get("LECH", 0) > 0).sum()) if "LECH" in piv.columns else 0,
        "Nợ xấu (count)": int((piv.get("Nợ xấu", "") == "x").sum()) if "Nợ xấu" in piv.columns else 0,
    }

    # Thu thập các bảng trung gian để xuất Excel
    extras = {
        "df_gop_tieu_chi_3": df_gop,
        "df_count_tieu_chi_3": df_count,
        "df_delay_tieu_chi_4": df_delay,
    }

    return piv, {**kpi, **extras}


# ============================ RUN ============================ #
if run:
    with st.spinner("Đang tải & xử lý dữ liệu..."):
        df_crm4 = load_and_concat(crm4_files)
        df_crm32 = load_and_concat(crm32_files)
        df_muc_dich = read_excel_any(df_muc_dich_file)
        df_code_tsbd = read_excel_any(df_code_tsbd_file)

        # Lọc chi nhánh
        if chi_nhanh:
            if "BRANCH_VAY" in df_crm4.columns:
                df_crm4 = df_crm4[df_crm4["BRANCH_VAY"].astype(str).str.upper().str.contains(chi_nhanh)]
            if "BRCD" in df_crm32.columns:
                df_crm32 = df_crm32[df_crm32["BRCD"].astype(str).str.upper().str.contains(chi_nhanh)]

        # Chuẩn CIF & keys chính
        if "CIF_KH_VAY" in df_crm4.columns:
            try:
                s = to_str_intlike(df_crm4["CIF_KH_VAY"])  # int-like → str
                df_crm4["CIF_KH_VAY"] = df_crm4["CIF_KH_VAY"].astype(str).str.strip()
                df_crm4.loc[df_crm4.index.isin(s.index), "CIF_KH_VAY"] = s
            except Exception:
                df_crm4["CIF_KH_VAY"] = df_crm4["CIF_KH_VAY"].astype(str).str.strip()

        if "CUSTSEQLN" in df_crm32.columns:
            try:
                s2 = to_str_intlike(df_crm32["CUSTSEQLN"])  # int-like → str
                df_crm32["CUSTSEQLN"] = df_crm32["CUSTSEQLN"].astype(str).str.strip()
                df_crm32.loc[df_crm32.index.isin(s2.index), "CUSTSEQLN"] = s2
            except Exception:
                df_crm32["CUSTSEQLN"] = df_crm32["CUSTSEQLN"].astype(str).str.strip()

        # Ánh xạ loại TSBĐ & mục đích vay
        df_crm4 = add_loai_ts(df_crm4, df_code_tsbd)
        df_crm32 = add_muc_dich_crm32(df_crm32, df_muc_dich)

        # Pivots CRM4
        pivot_ts, pivot_no, pivot_merge, pivot_final = build_pivots(df_crm4)

        # CRM32 – cấp C & cơ cấu
        df_crm32_filtered, list_cif_cap_c, cif_co_cau = enrich_crm32(df_crm32)

        # Pivot theo mục đích CRM32
        p_mucdich = pivot_muc_dich(df_crm32_filtered)

        # Bảng phụ (tuỳ chọn)
        df_tm = read_excel_any(file_giai_ngan_tm)
        df_m17 = read_excel_any(file_muc17)
        df_55 = read_excel_any(file_muc55)
        df_56 = read_excel_any(file_muc56)
        df_57 = read_excel_any(file_muc57)

        pivot_full, kpi = add_flags_and_joins(
            pivot_final,
            p_mucdich,
            df_crm4,
            df_crm32_filtered,
            list_cif_cap_c,
            cif_co_cau,
            df_tm,
            pd.to_datetime(ngay_danh_gia),
            df_m17,
            dia_ban_kt,
            df_55,
            df_56,
            df_57,
        )

    # ======================== OUTPUT UI ======================== #
    if pivot_full.empty:
        st.error("Không có dữ liệu sau khi xử lý. Vui lòng kiểm tra file và tham số lọc.")
    else:
        # KPIs
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Số KH", f"{kpi.get('Số KH', 0):,}")
        c2.metric("Tổng dư nợ", f"{kpi.get('Tổng dư nợ', 0):,.0f}")
        c3.metric("Lệch dương (count)", f"{kpi.get('Lệch dương (count)', 0):,}")
        c4.metric("Nợ xấu (count)", f"{kpi.get('Nợ xấu (count)', 0):,}")

        with st.expander("🔎 Pivot CRM4 (chi tiết)", expanded=False):
            st.dataframe(pivot_merge, use_container_width=True, height=360)
        with st.expander("🎯 Pivot CRM32 theo mục đích", expanded=False):
            st.dataframe(p_mucdich, use_container_width=True, height=360)

        st.subheader("📋 Kết quả tổng hợp theo CIF")
        st.dataframe(pivot_full, use_container_width=True, height=520)

        # Xuất Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            if not df_crm4.empty:
                df_crm4.to_excel(writer, sheet_name="df_crm4_LOAI_TS", index=False)
            if not pivot_final.empty:
                pivot_final.to_excel(writer, sheet_name="KQ_CRM4", index=False)
            if not pivot_merge.empty:
                pivot_merge.to_excel(writer, sheet_name="Pivot_crm4", index=False)
            if not df_crm32_filtered.empty:
                df_crm32_filtered.to_excel(writer, sheet_name="df_crm32_LOAI_TS", index=False)
            if not pivot_full.empty:
                pivot_full.to_excel(writer, sheet_name="KQ_KH", index=False)
            if not p_mucdich.empty:
                p_mucdich.to_excel(writer, sheet_name="Pivot_crm32", index=False)

            # Các sheet tiêu chí
            if isinstance(kpi.get("df_delay_tieu_chi_4"), pd.DataFrame) and not kpi["df_delay_tieu_chi_4"].empty:
                kpi["df_delay_tieu_chi_4"].to_excel(writer, sheet_name="tieu chi 4", index=False)
            if isinstance(kpi.get("df_gop_tieu_chi_3"), pd.DataFrame) and not kpi["df_gop_tieu_chi_3"].empty:
                kpi["df_gop_tieu_chi_3"].to_excel(writer, sheet_name="tieu chi 3_dot3", index=False)
            if isinstance(kpi.get("df_count_tieu_chi_3"), pd.DataFrame) and not kpi["df_count_tieu_chi_3"].empty:
                kpi["df_count_tieu_chi_3"].to_excel(writer, sheet_name="tieu chi 3_dot3_1", index=False)

        st.download_button(
            label="💾 Tải Excel kết quả",
            data=buffer.getvalue(),
            file_name="KQ_phan_tich_CRM.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

else:
    st.info("⬅️ Tải file và điền tham số ở thanh bên, sau đó nhấn **Chạy phân tích**.")
