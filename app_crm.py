import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import requests
from datetime import datetime

st.set_page_config(page_title="CRM Audit Dashboard", layout="wide")
st.title("📊 CRM Audit Dashboard")
st.caption("Chuyển đổi script phân tích CRM4/CRM32 sang Streamlit – hỗ trợ Upload hoặc GitHub Raw URLs")

# ======================= Helper functions =======================
def read_excel_smart(file) -> pd.DataFrame:
    """Đọc Excel từ UploadedFile/bytes/tuple(name, bytes). Chọn engine theo phần mở rộng.
    Hỗ trợ .xls (xlrd) & .xlsx (openpyxl). Trả về DataFrame (hoặc rỗng nếu None).
    """
    if file is None:
        return pd.DataFrame()

    # Cho phép truyền tuple (name, bytes) khi lấy từ URL
    if isinstance(file, tuple) and len(file) == 2:
        name, data = file
    else:
        name = getattr(file, "name", "uploaded.xlsx")
        data = file.read() if hasattr(file, "read") else file

    ext = name.lower().rsplit(".", 1)[-1] if "." in name else "xlsx"
    bio = io.BytesIO(data)
    try:
        if ext == "xls":
            df = pd.read_excel(bio, engine="xlrd")
        else:
            df = pd.read_excel(bio, engine="openpyxl")
    finally:
        try:
            bio.seek(0)
        except Exception:
            pass
    # Chuẩn hoá tên cột: strip, gộp khoảng trắng
    df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
    return df


def fetch_url_excel(url: str):
    """Tải file Excel từ URL (ví dụ GitHub Raw) và trả về tuple (name, bytes)."""
    if not url:
        return None
    u = url.strip()
    resp = requests.get(u, timeout=60)
    resp.raise_for_status()
    name = u.split("/")[-1] or "download.xlsx"
    return (name, resp.content)


def load_multiple(files_or_urls):
    dfs = []
    for f in files_or_urls:
        if f is None:
            continue
        try:
            dfs.append(read_excel_smart(f))
        except Exception as e:
            st.warning(f"Không đọc được file: {getattr(f, 'name', str(f))} – {e}")
    return dfs

# ======================= Sidebar – Data inputs =======================
with st.sidebar:
    st.header("⚙️ Thiết lập & Upload dữ liệu")
    DATA_SOURCE = st.radio("Nguồn dữ liệu", ["Upload", "GitHub URLs"], index=0, horizontal=True)

    if DATA_SOURCE == "Upload":
        st.markdown("**1) Upload danh mục/bảng mã**")
        col1, col2 = st.columns(2)
        with col1:
            f_mdsd = st.file_uploader("CODE_MDSDV4.xlsx", type=["xls","xlsx"], help="Bảng mã mục đích vay")
        with col2:
            f_loaits = st.file_uploader("CODE_LOAI TSBD.xlsx", type=["xls","xlsx"], help="Bảng mã loại TSBĐ")

        st.markdown("**2) Upload danh sách CRM4/CRM32**")
        files_crm4 = st.file_uploader("CRM4 (*.xls/x)", type=["xls","xlsx"], accept_multiple_files=True)
        files_crm32 = st.file_uploader("CRM32 (*.xls/x)", type=["xls","xlsx"], accept_multiple_files=True)

        st.markdown("**3) File bổ sung (tuỳ chọn)**")
        f_giaingan_1ty = st.file_uploader("Giai_ngan_tien_mat_1_ty", type=["xls","xlsx"])
        f_muc17 = st.file_uploader("MUC17.xlsx", type=["xls","xlsx"])
        f_muc55 = st.file_uploader("Muc55_1710.xlsx", type=["xls","xlsx"])
        f_muc56 = st.file_uploader("Muc56_1710.xlsx", type=["xls","xlsx"])
        f_muc57 = st.file_uploader("Muc57_1710.xlsx", type=["xls","xlsx"])

    else:
        st.markdown("**Nhập GitHub Raw URLs** (mỗi dòng một URL cho danh sách)")
        url_mdsd = st.text_input("URL CODE_MDSDV4.xlsx")
        url_loaits = st.text_input("URL CODE_LOAI TSBD.xlsx")
        urls_crm4_text = st.text_area("URLs CRM4 (*.xls/x)")
        urls_crm32_text = st.text_area("URLs CRM32 (*.xls/x)")
        url_giaingan_1ty = st.text_input("URL Giai_ngan_tien_mat_1_ty")
        url_muc17 = st.text_input("URL MUC17.xlsx")
        url_muc55 = st.text_input("URL Muc55_1710.xlsx")
        url_muc56 = st.text_input("URL Muc56_1710.xlsx")
        url_muc57 = st.text_input("URL Muc57_1710.xlsx")

        # Tải về thành (name, bytes)
        f_mdsd = fetch_url_excel(url_mdsd) if url_mdsd else None
        f_loaits = fetch_url_excel(url_loaits) if url_loaits else None
        files_crm4 = [fetch_url_excel(u) for u in urls_crm4_text.splitlines() if u.strip()]
        files_crm32 = [fetch_url_excel(u) for u in urls_crm32_text.splitlines() if u.strip()]
        f_giaingan_1ty = fetch_url_excel(url_giaingan_1ty) if url_giaingan_1ty else None
        f_muc17 = fetch_url_excel(url_muc17) if url_muc17 else None
        f_muc55 = fetch_url_excel(url_muc55) if url_muc55 else None
        f_muc56 = fetch_url_excel(url_muc56) if url_muc56 else None
        f_muc57 = fetch_url_excel(url_muc57) if url_muc57 else None

    st.divider()
    st.markdown("**Tham số chạy**")
    chi_nhanh = st.text_input("Nhập tên chi nhánh hoặc mã SOL (vd: HANOI hoặc 001)", value="")
    ngay_danh_gia_str = st.text_input("Ngày đánh giá (YYYY-MM-DD)", value="2025-08-31")
    try:
        ngay_danh_gia = pd.to_datetime(ngay_danh_gia_str)
    except Exception:
        ngay_danh_gia = pd.to_datetime("2025-08-31")
    dia_ban_kt_text = st.text_input("Địa bàn kiểm toán (cách nhau bằng dấu phẩy)", value="")
    dia_ban_kt = [t.strip().lower() for t in dia_ban_kt_text.split(',') if t.strip()]

    run_btn = st.button("🚀 Chạy xử lý")

# ======================= Processing =======================
if run_btn:
    try:
        # 1) Load tất cả file
        df_crm4_list = load_multiple(files_crm4 or [])
        df_crm32_list = load_multiple(files_crm32 or [])
        df_muc_dich_file = read_excel_smart(f_mdsd)
        df_code_tsbd_file = read_excel_smart(f_loaits)

        if not df_crm4_list or not df_crm32_list:
            st.error("Thiếu file CRM4 hoặc CRM32. Vui lòng upload/nhập URL.")
            st.stop()

        df_crm4 = pd.concat(df_crm4_list, ignore_index=True)
        df_crm32 = pd.concat(df_crm32_list, ignore_index=True)
        df_muc_dich = df_muc_dich_file.copy()
        df_code_tsbd = df_code_tsbd_file.copy()

        # 2) Chuẩn hoá CIF/CUSTSEQLN dạng str
        if 'CIF_KH_VAY' in df_crm4.columns:
            df_crm4['CIF_KH_VAY'] = pd.to_numeric(df_crm4['CIF_KH_VAY'], errors='coerce')
            df_crm4['CIF_KH_VAY'] = df_crm4['CIF_KH_VAY'].dropna().astype('int64').astype(str)
        if 'CUSTSEQLN' in df_crm32.columns:
            df_crm32['CUSTSEQLN'] = pd.to_numeric(df_crm32['CUSTSEQLN'], errors='coerce')
            df_crm32['CUSTSEQLN'] = df_crm32['CUSTSEQLN'].dropna().astype('int64').astype(str)

        # 3) Lọc theo chi nhánh (contains, case-insensitive)
        df_crm4_filtered = df_crm4.copy()
        df_crm32_filtered = df_crm32.copy()
        if chi_nhanh:
            df_crm4_filtered = df_crm4[df_crm4['BRANCH_VAY'].astype(str).str.upper().str.contains(chi_nhanh.strip().upper(), na=False)]
            df_crm32_filtered = df_crm32[df_crm32['BRCD'].astype(str).str.upper().str.contains(chi_nhanh.strip().upper(), na=False)]
        st.success(f"Số dòng CRM4 sau lọc: {len(df_crm4_filtered):,}")

        # 4) Map loại TSBĐ từ CODE
        if not df_code_tsbd.empty:
            df_code_tsbd = df_code_tsbd[['CODE CAP 2', 'CODE']].rename(columns={'CODE CAP 2':'CAP_2','CODE':'LOAI_TS'})
            df_tsbd_code = df_code_tsbd[['CAP_2','LOAI_TS']].drop_duplicates()
            df_crm4_filtered = df_crm4_filtered.merge(df_tsbd_code, how='left', on='CAP_2')
            df_crm4_filtered['LOAI_TS'] = df_crm4_filtered.apply(
                lambda r: 'Không TS' if pd.isna(r.get('CAP_2')) or str(r.get('CAP_2')).strip()=='' else r.get('LOAI_TS'), axis=1
            )
            df_crm4_filtered['GHI_CHU_TSBD'] = df_crm4_filtered.apply(
                lambda r: 'MỚI' if str(r.get('CAP_2','')).strip()!='' and pd.isna(r.get('LOAI_TS')) else '', axis=1
            )

        # 5) Pivots ts & dư nợ (loại trừ Bao lanh/LC)
        df_vay = df_crm4_filtered[~df_crm4_filtered['LOAI'].isin(['Bao lanh','LC'])].copy()
        pivot_ts = df_vay.pivot_table(index='CIF_KH_VAY', columns='LOAI_TS', values='TS_KW_VND', aggfunc='sum', fill_value=0).add_suffix(' (Giá trị TS)').reset_index()
        pivot_no = df_vay.pivot_table(index='CIF_KH_VAY', columns='LOAI_TS', values='DU_NO_PHAN_BO_QUY_DOI', aggfunc='sum', fill_value=0).reset_index()
        pivot_merge = pivot_no.merge(pivot_ts, on='CIF_KH_VAY', how='left')
        if not pivot_ts.empty:
            pivot_merge['GIÁ TRỊ TS'] = pivot_ts.drop(columns='CIF_KH_VAY').sum(axis=1)
        else:
            pivot_merge['GIÁ TRỊ TS'] = 0
        pivot_merge['DƯ NỢ'] = pivot_no.drop(columns='CIF_KH_VAY').sum(axis=1)

        df_info = df_crm4_filtered[['CIF_KH_VAY','TEN_KH_VAY','CUSTTPCD','NHOM_NO']].drop_duplicates(subset='CIF_KH_VAY')
        pivot_final = df_info.merge(pivot_merge, on='CIF_KH_VAY', how='left').reset_index().rename(columns={'index':'STT'})
        pivot_final['STT'] = pivot_final['STT'] + 1

        cols_order = ['STT','CUSTTPCD','CIF_KH_VAY','TEN_KH_VAY','NHOM_NO'] \
            + sorted([c for c in pivot_merge.columns if c not in ['CIF_KH_VAY','GIÁ TRỊ TS','DƯ NỢ'] and '(Giá trị TS)' not in c]) \
            + sorted([c for c in pivot_merge.columns if '(Giá trị TS)' in c]) \
            + ['DƯ NỢ','GIÁ TRỊ TS']
        pivot_final = pivot_final[[c for c in cols_order if c in pivot_final.columns]]

        # 6) Phê duyệt cấp C và cơ cấu
        df_crm32_filtered = df_crm32_filtered.copy()
        if 'CAP_PHE_DUYET' in df_crm32_filtered.columns:
            df_crm32_filtered['MA_PHE_DUYET'] = df_crm32_filtered['CAP_PHE_DUYET'].astype(str).str.split('-').str[0].str.strip().str.zfill(2)
        ma_cap_c = [f"{i:02d}" for i in range(1,8)] + [f"{i:02d}" for i in range(28,32)]
        list_cif_cap_c = df_crm32_filtered[df_crm32_filtered['MA_PHE_DUYET'].isin(ma_cap_c)]['CUSTSEQLN'].unique() if 'MA_PHE_DUYET' in df_crm32_filtered else []
        list_co_cau = ['ACOV1','ACOV3','ATT01','ATT02','ATT03','ATT04','BCOV1','BCOV2','BTT01','BTT02','BTT03','CCOV2','CCOV3','CTT03','RCOV3','RTT03']
        cif_co_cau = df_crm32_filtered[df_crm32_filtered['SCHEME_CODE'].isin(list_co_cau)]['CUSTSEQLN'].unique() if 'SCHEME_CODE' in df_crm32_filtered else []

        # 7) Mục đích vay (group)
        if not df_muc_dich.empty and 'CODE_MDSDV4' in df_muc_dich.columns:
            df_muc_dich_vay = df_muc_dich[['CODE_MDSDV4','GROUP']].rename(columns={'CODE_MDSDV4':'MUC_DICH_VAY_CAP_4','GROUP':'MUC DICH'})
            df_crm32_filtered = df_crm32_filtered.merge(df_muc_dich_vay, how='left', on='MUC_DICH_VAY_CAP_4')
            df_crm32_filtered['MUC DICH'] = df_crm32_filtered['MUC DICH'].fillna('(blank)')
            df_crm32_filtered['GHI_CHU_TSBD'] = df_crm32_filtered.apply(
                lambda r: 'MỚI' if str(r.get('MUC_DICH_VAY_CAP_4','')).strip()!='' and pd.isna(r.get('MUC DICH')) else '', axis=1
            )

        pivot_mucdich = pd.DataFrame()
        if 'CUSTSEQLN' in df_crm32_filtered.columns and 'MUC DICH' in df_crm32_filtered.columns and 'DU_NO_QUY_DOI' in df_crm32_filtered.columns:
            pivot_mucdich = df_crm32_filtered.pivot_table(index='CUSTSEQLN', columns='MUC DICH', values='DU_NO_QUY_DOI', aggfunc='sum', fill_value=0).reset_index()
            pivot_mucdich['DƯ NỢ CRM32'] = pivot_mucdich.drop(columns='CUSTSEQLN').sum(axis=1)
            pivot_final_CRM32 = pivot_mucdich.rename(columns={'CUSTSEQLN':'CIF_KH_VAY'})
        else:
            pivot_final_CRM32 = pd.DataFrame(columns=['CIF_KH_VAY','DƯ NỢ CRM32'])

        pivot_full = pivot_final.merge(pivot_final_CRM32, on='CIF_KH_VAY', how='left')
        pivot_full.fillna(0, inplace=True)
        if 'DƯ NỢ' in pivot_full.columns and 'DƯ NỢ CRM32' in pivot_full.columns:
            pivot_full['LECH'] = pivot_full['DƯ NỢ'] - pivot_full['DƯ NỢ CRM32']
        else:
            pivot_full['LECH'] = 0

        cif_lech = pivot_full[pivot_full['LECH'] != 0]['CIF_KH_VAY'].unique()
        df_crm4_blank = df_crm4_filtered[~df_crm4_filtered['LOAI'].isin(['Cho vay','Bao lanh','LC'])].copy()
        if not df_crm4_blank.empty:
            du_no_bosung = df_crm4_blank[df_crm4_blank['CIF_KH_VAY'].isin(cif_lech)].groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI':'(blank)'} )
            pivot_full = pivot_full.merge(du_no_bosung, on='CIF_KH_VAY', how='left')
            pivot_full['(blank)'] = pivot_full['(blank)'].fillna(0)
            pivot_full['DƯ NỢ CRM32'] = pivot_full['DƯ NỢ CRM32'] + pivot_full['(blank)']

        # Cờ nhóm nợ, phê duyệt C, cơ cấu
        pivot_full['Nợ nhóm 2'] = pivot_full['NHOM_NO'].apply(lambda x: 'x' if str(x).strip()=='2' else '') if 'NHOM_NO' in pivot_full else ''
        pivot_full['Nợ xấu'] = pivot_full['NHOM_NO'].apply(lambda x: 'x' if str(x).strip() in ['3','4','5'] else '') if 'NHOM_NO' in pivot_full else ''
        pivot_full['Chuyên gia PD cấp C duyệt'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in list_cif_cap_c else '')
        pivot_full['NỢ CƠ_CẤU'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in cif_co_cau else '')

        # Bảo lãnh & LC
        df_baolanh = df_crm4_filtered[df_crm4_filtered['LOAI']=='Bao lanh']
        df_lc = df_crm4_filtered[df_crm4_filtered['LOAI']=='LC']
        if not df_baolanh.empty:
            df_baolanh_sum = df_baolanh.groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI':'DƯ_NỢ_BẢO_LÃNH'})
            pivot_full = pivot_full.merge(df_baolanh_sum, on='CIF_KH_VAY', how='left')
        if not df_lc.empty:
            df_lc_sum = df_lc.groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI':'DƯ_NỢ_LC'})
            pivot_full = pivot_full.merge(df_lc_sum, on='CIF_KH_VAY', how='left')
        pivot_full['DƯ_NỢ_BẢO_LÃNH'] = pivot_full.get('DƯ_NỢ_BẢO_LÃNH', 0).fillna(0)
        pivot_full['DƯ_NỢ_LC'] = pivot_full.get('DƯ_NỢ_LC', 0).fillna(0)

        # Giải ngân tiền mặt 1 tỷ
        df_giai_ngan = read_excel_smart(f_giaingan_1ty)
        if not df_giai_ngan.empty and 'FORACID' in df_giai_ngan.columns and {'KHE_UOC','CUSTSEQLN'}.issubset(df_crm32_filtered.columns):
            df_crm32_filtered['KHE_UOC'] = df_crm32_filtered['KHE_UOC'].astype(str).str.strip()
            df_crm32_filtered['CUSTSEQLN'] = df_crm32_filtered['CUSTSEQLN'].astype(str).str.strip()
            df_giai_ngan['FORACID'] = df_giai_ngan['FORACID'].astype(str).str.strip()
            pivot_full['CIF_KH_VAY'] = pivot_full['CIF_KH_VAY'].astype(str).str.strip()
            df_match = df_crm32_filtered[df_crm32_filtered['KHE_UOC'].isin(df_giai_ngan['FORACID'])].copy()
            ds_cif_tien_mat = df_match['CUSTSEQLN'].unique()
            pivot_full['GIẢI_NGÂN_TIEN_MAT'] = pivot_full['CIF_KH_VAY'].isin(ds_cif_tien_mat).map({True:'x', False:''})
        else:
            pivot_full['GIẢI_NGÂN_TIEN_MAT'] = ''

        # Cầm cố tại TCTD khác
        if 'CAP_2' in df_crm4_filtered.columns:
            df_cc_tctd = df_crm4_filtered[df_crm4_filtered['CAP_2'].astype(str).str.contains('TCTD', case=False, na=False)]
            df_cc_flag = df_cc_tctd[['CIF_KH_VAY']].drop_duplicates()
            df_cc_flag['Cầm cố tại TCTD khác'] = 'x'
            pivot_full = pivot_full.merge(df_cc_flag, on='CIF_KH_VAY', how='left')
            pivot_full['Cầm cố tại TCTD khác'] = pivot_full['Cầm cố tại TCTD khác'].fillna('')

        # Top 10 KHCN / KHDN
        if {'CUSTTPCD','DƯ NỢ','CIF_KH_VAY'}.issubset(pivot_full.columns):
            top_khcn = pivot_full[pivot_full['CUSTTPCD']=='Ca nhan'].nlargest(10,'DƯ NỢ')['CIF_KH_VAY']
            top_khdn = pivot_full[pivot_full['CUSTTPCD']=='Doanh nghiep'].nlargest(10,'DƯ NỢ')['CIF_KH_VAY']
            pivot_full['Top 10 dư nợ KHCN'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in set(top_khcn.values) else '')
            pivot_full['Top 10 dư nợ KHDN'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in set(top_khdn.values) else '')

        # Quá hạn định giá TSBĐ
        if 'LOAI_TS' in df_crm4_filtered.columns and 'VALUATION_DATE' in df_crm4_filtered.columns:
            loai_ts_r34 = ['BĐS','MMTB','PTVT']
            mask_r34 = df_crm4_filtered['LOAI_TS'].isin(loai_ts_r34)
            df_crm4_filtered['VALUATION_DATE'] = pd.to_datetime(df_crm4_filtered['VALUATION_DATE'], errors='coerce')
            df_crm4_filtered.loc[mask_r34, 'SO_NGAY_QUA_HAN'] = (ngay_danh_gia - df_crm4_filtered.loc[mask_r34,'VALUATION_DATE']).dt.days - 365
            df_crm4_filtered.loc[df_crm4_filtered['LOAI_TS']=='BĐS','SO_THANG_QUA_HAN'] = ((ngay_danh_gia - df_crm4_filtered.loc[df_crm4_filtered['LOAI_TS']=='BĐS','VALUATION_DATE']).dt.days/31) - 18
            df_crm4_filtered.loc[df_crm4_filtered['LOAI_TS'].isin(['MMTB','PTVT']),'SO_THANG_QUA_HAN'] = ((ngay_danh_gia - df_crm4_filtered.loc[df_crm4_filtered['LOAI_TS'].isin(['MMTB','PTVT']),'VALUATION_DATE']).dt.days/31) - 12
            cif_quahan = df_crm4_filtered[df_crm4_filtered['SO_NGAY_QUA_HAN']>30]['CIF_KH_VAY'].unique()
            pivot_full['KH có TSBĐ quá hạn định giá'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'X' if x in cif_quahan else '')

        # Mục 17 – BĐS khác địa bàn
        df_sol = read_excel_smart(f_muc17)
        if not df_sol.empty and 'C01' in df_sol.columns:
            ds_secu = df_crm4_filtered.get('SECU_SRL_NUM', pd.Series(dtype=object)).dropna().unique()
            df_17_filtered = df_sol[df_sol['C01'].isin(ds_secu)]
            df_bds = df_17_filtered[df_17_filtered['C02'].astype(str).str.strip()=='Bat dong san'].copy()
            if 'SECU_SRL_NUM' in df_crm4.columns:
                df_bds_matched = df_bds[df_bds['C01'].isin(df_crm4['SECU_SRL_NUM'])].copy()
            else:
                df_bds_matched = df_bds.copy()
            def extract_tinh_thanh(diachi):
                if pd.isna(diachi):
                    return ''
                parts = str(diachi).split(',')
                return parts[-1].strip().lower() if parts else ''
            if 'C19' in df_bds_matched.columns:
                df_bds_matched['TINH_TP_TSBD'] = df_bds_matched['C19'].apply(extract_tinh_thanh)
                df_bds_matched['CANH_BAO_TS_KHAC_DIABAN'] = df_bds_matched['TINH_TP_TSBD'].apply(
                    lambda x: 'x' if x and x.strip().lower() not in dia_ban_kt else ''
                )
                ma_ts_canh_bao = df_bds_matched[df_bds_matched['CANH_BAO_TS_KHAC_DIABAN']=='x']['C01'].unique()
                cif_canh_bao = df_crm4[df_crm4['SECU_SRL_NUM'].isin(ma_ts_canh_bao)]['CIF_KH_VAY'].dropna().unique() if 'SECU_SRL_NUM' in df_crm4.columns else []
                pivot_full['KH có TSBĐ khác địa bàn'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in cif_canh_bao else '')
        else:
            df_bds_matched = pd.DataFrame()

        # Tiêu chí 3 – Mục 55/56: Giải ngân/Tất toán cùng ngày
        df_55 = read_excel_smart(f_muc55)
        df_56 = read_excel_smart(f_muc56)
        df_gop = pd.DataFrame(); df_count = pd.DataFrame()
        if not df_55.empty:
            df_tt = df_55[['CUSTSEQLN','NMLOC','KHE_UOC','SOTIENGIAINGAN','NGAYGN','NGAYDH','NGAY_TT','LOAITIEN']].copy()
            df_tt.columns = ['CIF','TEN_KHACH_HANG','KHE_UOC','SO_TIEN_GIAI_NGAN_VND','NGAY_GIAI_NGAN','NGAY_DAO_HAN','NGAY_TT','LOAI_TIEN_HD']
            df_tt['GIAI_NGAN_TT'] = 'Tất toán'
            df_tt['NGAY'] = pd.to_datetime(df_tt['NGAY_TT'], errors='coerce')
        else:
            df_tt = pd.DataFrame(columns=['CIF','TEN_KHACH_HANG','KHE_UOC','SO_TIEN_GIAI_NGAN_VND','NGAY_GIAI_NGAN','NGAY_DAO_HAN','NGAY_TT','LOAI_TIEN_HD','GIAI_NGAN_TT','NGAY'])
        if not df_56.empty:
            df_gn = df_56[['CIF','TEN_KHACH_HANG','KHE_UOC','SO_TIEN_GIAI_NGAN_VND','NGAY_GIAI_NGAN','NGAY_DAO_HAN','LOAI_TIEN_HD']].copy()
            df_gn['GIAI_NGAN_TT'] = 'Giải ngân'
            df_gn['NGAY_GIAI_NGAN'] = pd.to_datetime(df_gn['NGAY_GIAI_NGAN'], format='%Y%m%d', errors='coerce')
            df_gn['NGAY_DAO_HAN'] = pd.to_datetime(df_gn['NGAY_DAO_HAN'], format='%Y%m%d', errors='coerce')
            df_gn['NGAY'] = df_gn['NGAY_GIAI_NGAN']
        else:
            df_gn = pd.DataFrame(columns=['CIF','TEN_KHACH_HANG','KHE_UOC','SO_TIEN_GIAI_NGAN_VND','NGAY_GIAI_NGAN','NGAY_DAO_HAN','LOAI_TIEN_HD','GIAI_NGAN_TT','NGAY'])
        df_gop = pd.concat([df_tt, df_gn], ignore_index=True)
        df_gop = df_gop[df_gop['NGAY'].notna()]
        if not df_gop.empty:
            df_count = df_gop.groupby(['CIF','NGAY','GIAI_NGAN_TT']).size().unstack(fill_value=0).reset_index()
            df_count['CO_CA_GN_VA_TT'] = ((df_count.get('Giải ngân',0)>0) & (df_count.get('Tất toán',0)>0)).astype(int)
            ds_ca_gn_tt = df_count[df_count['CO_CA_GN_VA_TT']==1]['CIF'].astype(str).unique()
            pivot_full['CIF_KH_VAY'] = pivot_full['CIF_KH_VAY'].astype(str)
            pivot_full['KH có cả GNG và TT trong 1 ngày'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in ds_ca_gn_tt else '')
        else:
            pivot_full['KH có cả GNG và TT trong 1 ngày'] = ''

        # Tiêu chí 4 – Mục 57: Chậm trả
        df_delay = read_excel_smart(f_muc57)
        if not df_delay.empty and {'NGAY_DEN_HAN_TT','NGAY_THANH_TOAN'}.issubset(df_delay.columns):
            df_delay['NGAY_DEN_HAN_TT'] = pd.to_datetime(df_delay['NGAY_DEN_HAN_TT'], errors='coerce')
            df_delay['NGAY_THANH_TOAN'] = pd.to_datetime(df_delay['NGAY_THANH_TOAN'], errors='coerce')
            df_delay['NGAY_THANH_TOAN_FILL'] = df_delay['NGAY_THANH_TOAN'].fillna(ngay_danh_gia)
            df_delay['SO_NGAY_CHAM_TRA'] = (df_delay['NGAY_THANH_TOAN_FILL'] - df_delay['NGAY_DEN_HAN_TT']).dt.days
            mask_period = df_delay['NGAY_DEN_HAN_TT'].dt.year.between(2023, 2025)
            df_delay = df_delay[mask_period].copy()

            df_crm32_tmp = pivot_full.copy().rename(columns={'CIF_KH_VAY':'CIF_ID'})
            df_crm32_tmp['CIF_ID'] = df_crm32_tmp['CIF_ID'].astype(str)
            if 'CIF_ID' in df_delay.columns:
                df_delay['CIF_ID'] = df_delay['CIF_ID'].astype(str)
            else:
                # nếu file không có CIF_ID, cố gắng suy luận từ cột tên gần đúng
                if 'CIF' in df_delay.columns:
                    df_delay = df_delay.rename(columns={'CIF':'CIF_ID'})
                    df_delay['CIF_ID'] = df_delay['CIF_ID'].astype(str)
                else:
                    df_delay['CIF_ID'] = ''

            df_delay = df_delay.merge(df_crm32_tmp[['CIF_ID','DƯ NỢ','NHOM_NO']], on='CIF_ID', how='left')
            df_delay = df_delay[df_delay['NHOM_NO']==1].copy()

            def cap_cham_tra(days):
                if pd.isna(days): return None
                if days >= 10: return '>=10'
                if days >= 4: return '4-9'
                if days > 0: return '<4'
                return None
            df_delay['CAP_CHAM_TRA'] = df_delay['SO_NGAY_CHAM_TRA'].apply(cap_cham_tra)
            df_delay = df_delay.dropna(subset=['CAP_CHAM_TRA']).copy()
            df_delay['NGAY'] = df_delay['NGAY_DEN_HAN_TT'].dt.date
            df_delay.sort_values(['CIF_ID','NGAY','CAP_CHAM_TRA'], key=lambda s: s.map({'>=10':0,'4-9':1,'<4':2}), inplace=True)
            df_unique = df_delay.drop_duplicates(subset=['CIF_ID','NGAY'], keep='first').copy()
            df_dem = df_unique.groupby(['CIF_ID','CAP_CHAM_TRA']).size().unstack(fill_value=0)
            df_dem['KH Phát sinh chậm trả > 10 ngày'] = np.where(df_dem.get('>=10',0)>0, 'x','')
            df_dem['KH Phát sinh chậm trả 4-9 ngày'] = np.where((df_dem.get('>=10',0)==0) & (df_dem.get('4-9',0)>0), 'x','')
            pivot_full = pivot_full.merge(df_dem[['KH Phát sinh chậm trả > 10 ngày','KH Phát sinh chậm trả 4-9 ngày']], left_on='CIF_KH_VAY', right_index=True, how='left')
            for col in ['KH Phát sinh chậm trả > 10 ngày','KH Phát sinh chậm trả 4-9 ngày']:
                if col in pivot_full.columns:
                    pivot_full[col] = pivot_full[col].fillna('')
        else:
            df_unique = pd.DataFrame(); df_dem = pd.DataFrame()

        # ======================= Outputs =======================
        st.subheader("✅ Kết quả tổng hợp")
        t1, t2 = st.tabs(["Bảng khách hàng (pivot_full)", "CRM4/CRM32 đã lọc"])
        with t1:
            st.dataframe(pivot_full.head(500), use_container_width=True)
        with t2:
            st.markdown("**CRM4 filtered**")
            st.dataframe(df_crm4_filtered.head(200), use_container_width=True)
            st.markdown("**CRM32 filtered**")
            st.dataframe(df_crm32_filtered.head(200), use_container_width=True)

        # Tạo file Excel nhiều sheet để tải xuống
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_crm4_filtered.to_excel(writer, sheet_name='df_crm4_LOAI_TS', index=False)
            pivot_final.to_excel(writer, sheet_name='KQ_CRM4', index=False)
            pivot_merge.to_excel(writer, sheet_name='Pivot_crm4', index=False)
            df_crm32_filtered.to_excel(writer, sheet_name='df_crm32_LOAI_TS', index=False)
            pivot_full.to_excel(writer, sheet_name='KQ_KH', index=False)
            if not pivot_mucdich.empty:
                pivot_mucdich.to_excel(writer, sheet_name='Pivot_crm32', index=False)
            # Tiêu chí
            if not df_delay.empty:
                df_delay.to_excel(writer, sheet_name='tieu chi 4', index=False)
            if not df_gop.empty:
                df_gop.to_excel(writer, sheet_name='tieu chi 3_dot3', index=False)
            if not df_count.empty:
                df_count.to_excel(writer, sheet_name='tieu chi 3_dot3_1', index=False)
            if 'df_bds_matched' in locals() and not df_bds_matched.empty:
                df_bds_matched.to_excel(writer, sheet_name='tieu chi 2_dot3', index=False)
        output.seek(0)

        st.download_button(
            label="📥 Tải Excel tổng hợp",
            data=output,
            file_name="KQ_1710_.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.success("Hoàn tất xử lý!")

    except Exception as e:
        st.exception(e)
