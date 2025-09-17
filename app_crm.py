import io
import re
import os
import sys
import json
import time
import zipfile
import requests
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime

############################
# Helpers: IO & Excel
############################

EXT_ENGINE = {
    ".xls": "xlrd",          # needs xlrd >= 2.0.1
    ".xlsx": "openpyxl",     # needs openpyxl
    ".xlsm": "openpyxl",
}

@st.cache_data(show_spinner=False)
def fetch_url(url: str) -> bytes:
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    return r.content


def choose_engine_by_ext(name: str) -> str | None:
    name = (name or "").lower()
    for ext, eng in EXT_ENGINE.items():
        if name.endswith(ext):
            return eng
    # Fallback: let pandas decide (usually openpyxl for .xlsx)
    return None


def read_excel_smart(src, filename: str | None = None, sheet_name=None) -> pd.DataFrame:
    """Read Excel from:
    - Uploaded file (st.file_uploader -> src is UploadedFile)
    - bytes (src is bytes)
    - local path (str path)
    - URL (http/https)
    Uses engine based on file extension to avoid ImportError.
    """
    # Determine bytes and a display name
    data_bytes = None
    display_name = filename

    if hasattr(src, "read") and hasattr(src, "name"):  # UploadedFile
        data_bytes = src.read()
        display_name = src.name
    elif isinstance(src, (bytes, bytearray)):
        data_bytes = src
    elif isinstance(src, str):
        if src.startswith("http://") or src.startswith("https://"):
            data_bytes = fetch_url(src)
            display_name = src
        else:
            display_name = os.path.basename(src)
            with open(src, "rb") as f:
                data_bytes = f.read()
    else:
        raise ValueError("Unsupported source for read_excel_smart")

    engine = choose_engine_by_ext(display_name or "")
    bio = io.BytesIO(data_bytes)

    try:
        if engine:
            return pd.read_excel(bio, sheet_name=sheet_name, engine=engine)
        else:
            return pd.read_excel(bio, sheet_name=sheet_name)
    except ImportError as e:
        # Give a clearer message in the UI
        raise ImportError(
            f"\nKhông đọc được file: {display_name} – {e}.\n"
            "📦 Cần cài: `xlrd>=2.0.1` (đọc .xls) và `openpyxl` (đọc .xlsx/.xlsm).\n"
        )


############################
# UI
############################

st.set_page_config(page_title="CRM4/CRM32 Analyzer", layout="wide")
st.title("Chuyển đổi script phân tích CRM4/CRM32 – Upload & GitHub Raw URLs")
st.caption("Hỗ trợ đọc cả .xls (xlrd) và .xlsx (openpyxl). Xuất một file Excel nhiều sheet.")

with st.expander("1) Upload file hoặc dán GitHub Raw URLs", expanded=True):
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Upload files")
        up_crm4 = st.file_uploader("CRM4_Du_no_theo_tai_san_dam_bao_ALL*.xls/xlsx (nhiều file)", type=["xls", "xlsx"], accept_multiple_files=True)
        up_crm32 = st.file_uploader("RPT_CRM_32*.xls/xlsx (nhiều file)", type=["xls", "xlsx"], accept_multiple_files=True)
        up_mdsd = st.file_uploader("CODE_MDSDV4.xlsx", type=["xlsx"])  # bảng mục đích
        up_tsbd = st.file_uploader("CODE_LOAI TSBD.xlsx", type=["xlsx"])  # bảng mã TSBĐ
        up_giaingan1ty = st.file_uploader("Giai_ngan_tien_mat_1_ty.xls/xlsx", type=["xls", "xlsx"])  
        up_muc17 = st.file_uploader("MUC17.xlsx", type=["xlsx"])  
        up_muc55 = st.file_uploader("Muc55_1710.xlsx", type=["xlsx"])  
        up_muc56 = st.file_uploader("Muc56_1710.xlsx", type=["xlsx"])  
        up_muc57 = st.file_uploader("Muc57_1710.xlsx", type=["xlsx"])  

    with c2:
        st.subheader("Hoặc dán URL (GitHub Raw càng tốt)")
        url_crm4 = st.text_area("Danh sách URL CRM4 (mỗi dòng 1 URL)")
        url_crm32 = st.text_area("Danh sách URL CRM32 (mỗi dòng 1 URL)")
        url_mdsd = st.text_input("URL CODE_MDSDV4.xlsx")
        url_tsbd = st.text_input("URL CODE_LOAI TSBD.xlsx")
        url_giaingan1ty = st.text_input("URL Giai_ngan_tien_mat_1_ty.xls/xlsx")
        url_muc17 = st.text_input("URL MUC17.xlsx")
        url_muc55 = st.text_input("URL Muc55_1710.xlsx")
        url_muc56 = st.text_input("URL Muc56_1710.xlsx")
        url_muc57 = st.text_input("URL Muc57_1710.xlsx")

# Nhập chi nhánh & địa bàn kiểm toán
c3, c4 = st.columns(2)
with c3:
    chi_nhanh = st.text_input("Nhập tên chi nhánh hoặc mã SOL (ví dụ: HANOI hoặc 001)", "").strip().upper()
with c4:
    dia_ban_kt_input = st.text_input("Nhập tỉnh/thành của đơn vị đang KT (phân cách bằng dấu phẩy)", "").strip()

run = st.button("▶️ Chạy phân tích")

############################
# Load all inputs
############################

def load_many(label: str, uploads, urls_text: str) -> list[pd.DataFrame]:
    dfs: list[pd.DataFrame] = []
    if uploads:
        for f in uploads:
            try:
                dfs.append(read_excel_smart(f))
            except Exception as e:
                st.error(str(e))
    urls = [u.strip() for u in (urls_text or "").splitlines() if u.strip()]
    for u in urls:
        try:
            dfs.append(read_excel_smart(u))
        except Exception as e:
            st.error(str(e))
    if not dfs:
        st.warning(f"Chưa có dữ liệu cho {label}.")
    return dfs


if run:
    try:
        # --- 1) Đọc danh mục/bảng mã ---
        st.markdown("**1) Upload danh mục/bảng mã**")
        if up_mdsd:
            df_muc_dich_file = read_excel_smart(up_mdsd)
        elif url_mdsd:
            df_muc_dich_file = read_excel_smart(url_mdsd)
        else:
            st.stop()

        if up_tsbd:
            df_code_tsbd_file = read_excel_smart(up_tsbd)
        elif url_tsbd:
            df_code_tsbd_file = read_excel_smart(url_tsbd)
        else:
            st.stop()

        # --- 2) Đọc CRM4/CRM32 ---
        st.markdown("**2) Đọc CRM4/CRM32**")
        df_crm4_ghep = load_many("CRM4", up_crm4, url_crm4)
        df_crm32_ghep = load_many("CRM32", up_crm32, url_crm32)
        if not df_crm4_ghep or not df_crm32_ghep:
            st.stop()
        df_crm4 = pd.concat(df_crm4_ghep, ignore_index=True)
        df_crm32 = pd.concat(df_crm32_ghep, ignore_index=True)

        # --- 3) Chuẩn hóa ID ---
        st.markdown("**3) Chuẩn hóa khóa định danh**")
        if 'CIF_KH_VAY' in df_crm4.columns:
            df_crm4['CIF_KH_VAY'] = pd.to_numeric(df_crm4['CIF_KH_VAY'], errors='coerce')
            df_crm4['CIF_KH_VAY'] = df_crm4['CIF_KH_VAY'].dropna().astype('int64').astype(str)
        if 'CUSTSEQLN' in df_crm32.columns:
            df_crm32['CUSTSEQLN'] = pd.to_numeric(df_crm32['CUSTSEQLN'], errors='coerce')
            df_crm32['CUSTSEQLN'] = df_crm32['CUSTSEQLN'].dropna().astype('int64').astype(str)

        df_muc_dich = df_muc_dich_file.copy()
        df_code_tsbd = df_code_tsbd_file.copy()

        # --- 4) Lọc chi nhánh ---
        st.markdown("**4) Lọc theo chi nhánh**")
        if not chi_nhanh:
            st.warning("Bạn chưa nhập chi nhánh/SOL – đang giữ nguyên toàn bộ dữ liệu.")
            df_crm4_filtered = df_crm4.copy()
            df_crm32_filtered = df_crm32.copy()
        else:
            df_crm4_filtered = df_crm4[df_crm4['BRANCH_VAY'].astype(str).str.upper().str.contains(chi_nhanh)]
            df_crm32_filtered = df_crm32[df_crm32['BRCD'].astype(str).str.upper().str.contains(chi_nhanh)]
        st.write("CRM4 rows:", len(df_crm4_filtered), "| CRM32 rows:", len(df_crm32_filtered))

        # --- 5) Mapping loại tài sản ---
        st.markdown("**5) Mapping loại tài sản (TSBD)**")
        df_code_tsbd = df_code_tsbd[["CODE CAP 2", "CODE"]]
        df_code_tsbd.columns = ["CAP_2", "LOAI_TS"]
        df_tsbd_code = df_code_tsbd[["CAP_2", "LOAI_TS"]].drop_duplicates()
        df_crm4_filtered = df_crm4_filtered.merge(df_tsbd_code, how='left', on='CAP_2')
        df_crm4_filtered['LOAI_TS'] = df_crm4_filtered.apply(
            lambda r: 'Không TS' if pd.isna(r['CAP_2']) or str(r['CAP_2']).strip()=='' else r['LOAI_TS'], axis=1
        )
        df_crm4_filtered['GHI_CHU_TSBD'] = df_crm4_filtered.apply(
            lambda r: 'MỚI' if str(r['CAP_2']).strip() != '' and pd.isna(r['LOAI_TS']) else '', axis=1
        )

        # --- 6) Pivot theo TS & Dư nợ ---
        st.markdown("**6) Tổng hợp theo loại TS và dư nợ**")
        df_vay = df_crm4_filtered[~df_crm4_filtered['LOAI'].isin(['Bao lanh','LC'])]
        pivot_ts = df_vay.pivot_table(index='CIF_KH_VAY', columns='LOAI_TS', values='TS_KW_VND', aggfunc='sum', fill_value=0).add_suffix(' (Giá trị TS)').reset_index()
        pivot_no = df_vay.pivot_table(index='CIF_KH_VAY', columns='LOAI_TS', values='DU_NO_PHAN_BO_QUY_DOI', aggfunc='sum', fill_value=0).reset_index()
        pivot_merge = pivot_no.merge(pivot_ts, on='CIF_KH_VAY', how='left')
        # tổng theo hàng
        if 'CIF_KH_VAY' in pivot_ts.columns:
            pivot_merge['GIÁ TRỊ TS'] = pivot_ts.drop(columns='CIF_KH_VAY').sum(axis=1)
        if 'CIF_KH_VAY' in pivot_no.columns:
            pivot_merge['DƯ NỢ'] = pivot_no.drop(columns='CIF_KH_VAY').sum(axis=1)

        df_info = df_crm4_filtered[['CIF_KH_VAY','TEN_KH_VAY','CUSTTPCD','NHOM_NO']].drop_duplicates(subset='CIF_KH_VAY')
        pivot_final = df_info.merge(pivot_merge, on='CIF_KH_VAY', how='left')
        pivot_final = pivot_final.reset_index().rename(columns={'index': 'STT'})
        pivot_final['STT'] = pivot_final['STT'] + 1

        # --- 7) CRM32: cấp phê duyệt, mục đích vay ---
        st.markdown("**7) CRM32: mục đích vay & cấp phê duyệt**")
        df_crm32_filtered = df_crm32_filtered.copy()
        if 'CAP_PHE_DUYET' in df_crm32_filtered.columns:
            df_crm32_filtered['MA_PHE_DUYET'] = df_crm32_filtered['CAP_PHE_DUYET'].astype(str).str.split('-').str[0].str.strip().str.zfill(2)
        ma_cap_c = [f"{i:02d}" for i in range(1,8)] + [f"{i:02d}" for i in range(28,32)]
        list_cif_cap_c = df_crm32_filtered[df_crm32_filtered['MA_PHE_DUYET'].isin(ma_cap_c)]['CUSTSEQLN'].unique()

        list_co_cau = ['ACOV1','ACOV3','ATT01','ATT02','ATT03','ATT04','BCOV1','BCOV2','BTT01','BTT02','BTT03','CCOV2','CCOV3','CTT03','RCOV3','RTT03']
        cif_co_cau = df_crm32_filtered[df_crm32_filtered['SCHEME_CODE'].isin(list_co_cau)]['CUSTSEQLN'].unique()

        df_muc_dich_vay = df_muc_dich[['CODE_MDSDV4','GROUP']].copy()
        df_muc_dich_vay.columns = ['MUC_DICH_VAY_CAP_4', 'MUC DICH']
        df_crm32_filtered = df_crm32_filtered.merge(df_muc_dich_vay, how='left', on='MUC_DICH_VAY_CAP_4')
        df_crm32_filtered['MUC DICH'] = df_crm32_filtered['MUC DICH'].fillna('(blank)')

        pivot_mucdich = df_crm32_filtered.pivot_table(index='CUSTSEQLN', columns='MUC DICH', values='DU_NO_QUY_DOI', aggfunc='sum', fill_value=0).reset_index()
        pivot_mucdich['DƯ NỢ CRM32'] = pivot_mucdich.drop(columns='CUSTSEQLN').sum(axis=1)
        pivot_final_CRM32 = pivot_mucdich.rename(columns={'CUSTSEQLN':'CIF_KH_VAY'})
        pivot_full = pivot_final.merge(pivot_final_CRM32, on='CIF_KH_VAY', how='left')
        pivot_full = pivot_full.fillna(0)
        pivot_full['LECH'] = pivot_full['DƯ NỢ'] - pivot_full['DƯ NỢ CRM32']

        # (blank) bổ sung từ CRM4 non-loan
        df_crm4_blank = df_crm4_filtered[~df_crm4_filtered['LOAI'].isin(['Cho vay','Bao lanh','LC'])].copy()
        cif_lech = pivot_full[pivot_full['LECH'] != 0]['CIF_KH_VAY'].unique()
        du_no_bosung = (
            df_crm4_blank[df_crm4_blank['CIF_KH_VAY'].isin(cif_lech)]
            .groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI']
            .sum()
            .rename(columns={'DU_NO_PHAN_BO_QUY_DOI':'(blank)'})
        )
        pivot_full = pivot_full.merge(du_no_bosung, on='CIF_KH_VAY', how='left')
        pivot_full['(blank)'] = pivot_full['(blank)'].fillna(0)
        pivot_full['DƯ NỢ CRM32'] = pivot_full['DƯ NỢ CRM32'] + pivot_full['(blank)']

        # Cờ nhóm nợ, cơ cấu, chuyên gia C
        pivot_full['Nợ nhóm 2'] = pivot_full['NHOM_NO'].apply(lambda x: 'x' if str(x).strip()=='2' else '')
        pivot_full['Nợ xấu'] = pivot_full['NHOM_NO'].apply(lambda x: 'x' if str(x).strip() in ['3','4','5'] else '')
        pivot_full['Chuyên gia PD cấp C duyệt'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in list_cif_cap_c else '')
        pivot_full['NỢ CƠ_CẤU'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in cif_co_cau else '')

        # Bảo lãnh & LC
        df_baolanh = df_crm4_filtered[df_crm4_filtered['LOAI']=='Bao lanh']
        df_lc = df_crm4_filtered[df_crm4_filtered['LOAI']=='LC']
        df_baolanh_sum = df_baolanh.groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI':'DƯ_NỢ_BẢO_LÃNH'})
        df_lc_sum = df_lc.groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI':'DƯ_NỢ_LC'})
        pivot_full = pivot_full.merge(df_baolanh_sum, on='CIF_KH_VAY', how='left').merge(df_lc_sum, on='CIF_KH_VAY', how='left')
        pivot_full[['DƯ_NỢ_BẢO_LÃNH','DƯ_NỢ_LC']] = pivot_full[['DƯ_NỢ_BẢO_LÃNH','DƯ_NỢ_LC']].fillna(0)

        # Giải ngân tiền mặt 1 tỷ
        st.markdown("**8) Các tiêu chí & cảnh báo**")
        if up_giaingan1ty:
            df_giai_ngan = read_excel_smart(up_giaingan1ty)
        elif url_giaingan1ty:
            df_giai_ngan = read_excel_smart(url_giaingan1ty)
        else:
            df_giai_ngan = pd.DataFrame(columns=['FORACID'])
        if not df_giai_ngan.empty and 'FORACID' in df_giai_ngan.columns:
            df_crm32_filtered['KHE_UOC'] = df_crm32_filtered['KHE_UOC'].astype(str).str.strip()
            df_crm32_filtered['CUSTSEQLN'] = df_crm32_filtered['CUSTSEQLN'].astype(str).str.strip()
            df_giai_ngan['FORACID'] = df_giai_ngan['FORACID'].astype(str).str.strip()
            pivot_full['CIF_KH_VAY'] = pivot_full['CIF_KH_VAY'].astype(str).str.strip()
            df_match = df_crm32_filtered[df_crm32_filtered['KHE_UOC'].isin(df_giai_ngan['FORACID'])].copy()
            ds_cif_tien_mat = df_match['CUSTSEQLN'].unique()
            pivot_full['GIẢI_NGÂN_TIEN_MAT'] = pivot_full['CIF_KH_VAY'].isin(ds_cif_tien_mat).map({True:'x', False:''})

        # Cầm cố tại TCTD khác
        if 'CAP_2' in df_crm4_filtered.columns:
            df_cc_tctd = df_crm4_filtered[df_crm4_filtered['CAP_2'].astype(str).str.contains('TCTD', case=False, na=False)]
            df_cc_flag = df_cc_tctd[['CIF_KH_VAY']].drop_duplicates()
            df_cc_flag['Cầm cố tại TCTD khác'] = 'x'
            pivot_full = pivot_full.merge(df_cc_flag, on='CIF_KH_VAY', how='left')
            pivot_full['Cầm cố tại TCTD khác'] = pivot_full['Cầm cố tại TCTD khác'].fillna('')

        # Top 10 theo nhóm KH
        if 'CUSTTPCD' in pivot_full.columns:
            top_khcn = pivot_full[pivot_full['CUSTTPCD']=='Ca nhan'].nlargest(10, 'DƯ NỢ')['CIF_KH_VAY']
            pivot_full['Top 10 dư nợ KHCN'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in top_khcn.values else '')
            top_khdn = pivot_full[pivot_full['CUSTTPCD']=='Doanh nghiep'].nlargest(10, 'DƯ NỢ')['CIF_KH_VAY']
            pivot_full['Top 10 dư nợ KHDN'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in top_khdn.values else '')

        # Quá hạn định giá TSBD (R34)
        ngay_danh_gia = pd.to_datetime("2025-08-31")
        loai_ts_r34 = ['BĐS','MMTB','PTVT']
        if 'LOAI_TS' in df_crm4_filtered.columns and 'VALUATION_DATE' in df_crm4_filtered.columns:
            mask_r34 = df_crm4_filtered['LOAI_TS'].isin(loai_ts_r34)
            df_crm4_filtered['VALUATION_DATE'] = pd.to_datetime(df_crm4_filtered['VALUATION_DATE'], errors='coerce')
            df_tmp = df_crm4_filtered.loc[mask_r34].copy()
            df_tmp['SO_NGAY_QUA_HAN'] = (ngay_danh_gia - df_tmp['VALUATION_DATE']).dt.days - 365
            cif_quahan = df_tmp[df_tmp['SO_NGAY_QUA_HAN']>30]['CIF_KH_VAY'].unique()
            pivot_full['KH có TSBĐ quá hạn định giá'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'X' if x in cif_quahan else '')

        # Mục 17 – TS khác địa bàn
        if up_muc17 or url_muc17:
            df_sol = read_excel_smart(up_muc17 if up_muc17 else url_muc17)
            ds_secu = df_crm4_filtered['SECU_SRL_NUM'].dropna().unique() if 'SECU_SRL_NUM' in df_crm4_filtered.columns else []
            if 'C01' in df_sol.columns:
                df_17_filtered = df_sol[df_sol['C01'].isin(ds_secu)]
                df_bds = df_17_filtered[df_17_filtered['C02'].astype(str).str.strip()=='Bat dong san'].copy() if 'C02' in df_17_filtered.columns else pd.DataFrame()
                if not df_bds.empty and 'C19' in df_bds.columns:
                    def extract_tinh_thanh(diachi):
                        if pd.isna(diachi):
                            return ''
                        parts = str(diachi).split(',')
                        return parts[-1].strip().lower() if parts else ''
                    df_bds['TINH_TP_TSBD'] = df_bds['C19'].apply(extract_tinh_thanh)
                    dia_ban_kt = [t.strip().lower() for t in (dia_ban_kt_input or '').split(',') if t.strip()]
                    df_bds['CANH_BAO_TS_KHAC_DIABAN'] = df_bds['TINH_TP_TSBD'].apply(lambda x: 'x' if x and x not in dia_ban_kt else '')
                    ma_ts_canh_bao = df_bds[df_bds['CANH_BAO_TS_KHAC_DIABAN']=='x']['C01'].unique()
                    if 'SECU_SRL_NUM' in df_crm4_filtered.columns:
                        cif_canh_bao = df_crm4_filtered[df_crm4_filtered['SECU_SRL_NUM'].isin(ma_ts_canh_bao)]['CIF_KH_VAY'].dropna().unique()
                        pivot_full['KH có TSBĐ khác địa bàn'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in cif_canh_bao else '')
                else:
                    st.info("MUC17 không có cột C19 hoặc không có bản ghi 'Bat dong san'.")

        # Tiêu chí 3 – GN/TT trong 1 ngày
        def load_df(name, up, url):
            if up:
                return read_excel_smart(up)
            if url:
                return read_excel_smart(url)
            return pd.DataFrame()
        df_55 = load_df('Muc55', up_muc55, url_muc55)
        df_56 = load_df('Muc56', up_muc56, url_muc56)
        if not df_55.empty and not df_56.empty:
            df_tt = df_55[['CUSTSEQLN','NMLOC','KHE_UOC','SOTIENGIAINGAN','NGAYGN','NGAYDH','NGAY_TT','LOAITIEN']].copy()
            df_tt.columns = ['CIF','TEN_KHACH_HANG','KHE_UOC','SO_TIEN_GIAI_NGAN_VND','NGAY_GIAI_NGAN','NGAY_DAO_HAN','NGAY_TT','LOAI_TIEN_HD']
            df_tt['GIAI_NGAN_TT'] = 'Tất toán'
            df_tt['NGAY'] = pd.to_datetime(df_tt['NGAY_TT'], errors='coerce')

            df_gn = df_56[['CIF','TEN_KHACH_HANG','KHE_UOC','SO_TIEN_GIAI_NGAN_VND','NGAY_GIAI_NGAN','NGAY_DAO_HAN','LOAI_TIEN_HD']].copy()
            df_gn['GIAI_NGAN_TT'] = 'Giải ngân'
            df_gn['NGAY_GIAI_NGAN'] = pd.to_datetime(df_gn['NGAY_GIAI_NGAN'], format='%Y%m%d', errors='coerce')
            df_gn['NGAY_DAO_HAN'] = pd.to_datetime(df_gn['NGAY_DAO_HAN'], format='%Y%m%d', errors='coerce')
            df_gn['NGAY'] = df_gn['NGAY_GIAI_NGAN']

            df_gop = pd.concat([df_tt, df_gn], ignore_index=True)
            df_gop = df_gop[df_gop['NGAY'].notna()].sort_values(['CIF','NGAY','GIAI_NGAN_TT'])
            df_count = df_gop.groupby(['CIF','NGAY','GIAI_NGAN_TT']).size().unstack(fill_value=0).reset_index()
            df_count['CO_CA_GN_VA_TT'] = ((df_count.get('Giải ngân',0)>0) & (df_count.get('Tất toán',0)>0)).astype(int)
            ds_ca_gn_tt = df_count[df_count['CO_CA_GN_VA_TT']==1]['CIF'].astype(str).unique()
            pivot_full['CIF_KH_VAY'] = pivot_full['CIF_KH_VAY'].astype(str)
            pivot_full['KH có cả GNG và TT trong 1 ngày'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in ds_ca_gn_tt else '')
        else:
            df_gop = pd.DataFrame(); df_count = pd.DataFrame()

        # Tiêu chí 4 – Chậm trả
        df_delay = load_df('Muc57', up_muc57, url_muc57)
        if not df_delay.empty:
            df_delay['NGAY_DEN_HAN_TT'] = pd.to_datetime(df_delay['NGAY_DEN_HAN_TT'], errors='coerce')
            df_delay['NGAY_THANH_TOAN'] = pd.to_datetime(df_delay['NGAY_THANH_TOAN'], errors='coerce')
            ngay_danh_gia = pd.to_datetime("2025-08-31")
            df_delay['NGAY_THANH_TOAN_FILL'] = df_delay['NGAY_THANH_TOAN'].fillna(ngay_danh_gia)
            df_delay['SO_NGAY_CHAM_TRA'] = (df_delay['NGAY_THANH_TOAN_FILL'] - df_delay['NGAY_DEN_HAN_TT']).dt.days
            mask_period = df_delay['NGAY_DEN_HAN_TT'].dt.year.between(2023, 2025)
            df_delay = df_delay[mask_period].copy()

            df_crm32_for_join = pivot_full.rename(columns={'CIF_KH_VAY':'CIF_ID'})[['CIF_ID','DƯ NỢ','NHOM_NO']].copy()
            if 'CIF_ID' in df_delay.columns:
                df_delay['CIF_ID'] = df_delay['CIF_ID'].astype(str)
                df_crm32_for_join['CIF_ID'] = df_crm32_for_join['CIF_ID'].astype(str)
                df_delay = df_delay.merge(df_crm32_for_join, on='CIF_ID', how='left')
                df_delay = df_delay[df_delay['NHOM_NO']==1].copy()

                def cap_cham_tra(days):
                    if pd.isna(days):
                        return None
                    elif days >= 10:
                        return '>=10'
                    elif days >= 4:
                        return '4-9'
                    elif days > 0:
                        return '<4'
                    else:
                        return None
                df_delay['CAP_CHAM_TRA'] = df_delay['SO_NGAY_CHAM_TRA'].apply(cap_cham_tra)
                df_delay = df_delay.dropna(subset=['CAP_CHAM_TRA']).copy()
                df_delay['NGAY'] = df_delay['NGAY_DEN_HAN_TT'].dt.date
                df_delay.sort_values(['CIF_ID','NGAY','CAP_CHAM_TRA'], key=lambda s: s.map({'>=10':0,'4-9':1,'<4':2}), inplace=True)
                df_unique = df_delay.drop_duplicates(subset=['CIF_ID','NGAY'], keep='first').copy()
                df_dem = df_unique.groupby(['CIF_ID','CAP_CHAM_TRA']).size().unstack(fill_value=0)
                cols_to_merge = []
                df_dem['KH Phát sinh chậm trả > 10 ngày'] = np.where(df_dem.get('>=10',0)>0, 'x','')
                df_dem['KH Phát sinh chậm trả 4-9 ngày'] = np.where((df_dem.get('>=10',0)==0) & (df_dem.get('4-9',0)>0), 'x','')
                cols_to_merge = ['KH Phát sinh chậm trả > 10 ngày','KH Phát sinh chậm trả 4-9 ngày']
                pivot_full = pivot_full.merge(df_dem[cols_to_merge], left_on='CIF_KH_VAY', right_index=True, how='left')
                for c in cols_to_merge:
                    pivot_full[c] = pivot_full[c].fillna('')
        # Hiển thị
        st.success("✔️ Hoàn tất xử lý")
        st.dataframe(pivot_full.head(50))

        # Xuất Excel nhiều sheet
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            df_crm4_filtered.to_excel(writer, sheet_name='df_crm4_LOAI_TS', index=False)
            pivot_final.to_excel(writer, sheet_name='KQ_CRM4', index=False)
            pivot_merge.to_excel(writer, sheet_name='Pivot_crm4', index=False)
            df_crm32_filtered.to_excel(writer, sheet_name='df_crm32_LOAI_TS', index=False)
            pivot_full.to_excel(writer, sheet_name='KQ_KH', index=False)
            pivot_mucdich.to_excel(writer, sheet_name='Pivot_crm32', index=False)
            if 'df_delay' in locals() and not df_delay.empty:
                df_delay.to_excel(writer, sheet_name='tieu chi 4', index=False)
            if 'df_gop' in locals() and not df_gop.empty:
                df_gop.to_excel(writer, sheet_name='tieu chi 3_dot3', index=False)
            if 'df_count' in locals() and not df_count.empty:
                df_count.to_excel(writer, sheet_name='tieu chi 3_dot3_1', index=False)
            if 'df_bds' in locals() and not df_bds.empty:
                df_bds.to_excel(writer, sheet_name='tieu chi 2_dot3', index=False)
        st.download_button(
            "⬇️ Tải Excel kết quả",
            data=out.getvalue(),
            file_name="KQ_1710_.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"❌ Lỗi: {e}")
        st.stop()
