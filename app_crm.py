import io
import numpy as np
import pandas as pd
import streamlit as st

    st.set_page_config(page_title="CRM4/CRM32 Audit Toolkit", layout="wide")
    st.title("📊 CRM4/CRM32 Audit Toolkit")
    st.caption("Nhập file Excel, lọc theo chi nhánh, đối chiếu, tạo pivot & xuất báo cáo.")

    # ============== Helpers ==============
    def read_excel_any(file):
        if file is None:
            return None
        try:
            return pd.read_excel(file)
        except Exception as e:
            st.error(f"Lỗi đọc file {getattr(file, 'name', 'uploaded')}: {e}")
            return None

    def to_int_str(series):
        # Chuyển số thành int -> str (giữ không có .0)
        s = pd.to_numeric(series, errors='coerce')
        s = s.dropna().astype('int64').astype(str)
        return s

    def ensure_datetime(series):
        return pd.to_datetime(series, errors='coerce')

    def download_excel_sheets(sheets_dict, default_name="KQ.xlsx"):
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine="openpyxl") as writer:
            for sheet_name, df in sheets_dict.items():
                try:
                    df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                except Exception:
                    # fallback: reset index if weird columns
                    df.reset_index(drop=True).to_excel(writer, sheet_name=sheet_name[:31], index=False)
        st.download_button(
            "⬇️ Tải kết quả Excel",
            data=bio.getvalue(),
            file_name=default_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ============== Inputs ==============
    st.subheader("1) Tải lên dữ liệu")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**CRM4 - Dư nợ theo TSBĐ (nhiều file .xls)**")
        files_crm4 = st.file_uploader("Chọn file CRM4 (*.xls)", type=["xls", "xlsx"], accept_multiple_files=True, key="crm4")
        st.markdown("**CRM32 (nhiều file .xls)**")
        files_crm32 = st.file_uploader("Chọn file CRM32 (*.xls)", type=["xls", "xlsx"], accept_multiple_files=True, key="crm32")
        sol_branch = st.text_input("Nhập tên chi nhánh hoặc mã SOL cần lọc (ví dụ: HANOI hoặc 001)", value="").strip()
    with c2:
        st.markdown("**Bảng mã mục đích vay (CODE_MDSDV4.xlsx)**")
        file_mucdich = st.file_uploader("Chọn CODE_MDSDV4.xlsx", type=["xlsx"], key="md")
        st.markdown("**Bảng mã loại TSBĐ (CODE_LOAI TSBD.xlsx)**")
        file_loaits = st.file_uploader("Chọn CODE_LOAI TSBD.xlsx", type=["xlsx"], key="tsbd")
        st.markdown("**Giải ngân tiền mặt 1 tỷ (Giai_ngan_tien_mat_1_ty.xls)**")
        file_giaingan = st.file_uploader("Chọn Giai_ngan_tien_mat_1_ty.xls", type=["xls","xlsx"], key="gn")

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("**Mục 17 (df_sol - TSTC)**")
        file_muc17 = st.file_uploader("Chọn file Mục 17 (Excel)", type=["xls","xlsx"], key="m17")
        provinces = st.text_input("Nhập tỉnh/thành của đơn vị kiểm toán (phân cách dấu phẩy)", value="").strip()
    with c4:
        st.markdown("**Mục 55/56/57**")
        file_55 = st.file_uploader("Mục 55 (xlsx)", type=["xlsx"], key="55")
        file_56 = st.file_uploader("Mục 56 (xlsx)", type=["xlsx"], key="56")
        file_57 = st.file_uploader("Mục 57 (xlsx)", type=["xlsx"], key="57")
        ngay_danh_gia = st.date_input("Ngày đánh giá (R34 & chậm trả)", value=pd.to_datetime("2025-06-30"))

    run = st.button("▶️ Chạy xử lý")

    if run:
        if not files_crm4 or not files_crm32:
            st.warning("Vui lòng tải **ít nhất 1 file CRM4 và 1 file CRM32**.")
            st.stop()

        # ====== Đọc dữ liệu nguồn ======
        df_crm4_list = [read_excel_any(f) for f in files_crm4]
        df_crm4_list = [df for df in df_crm4_list if df is not None]
        df_crm32_list = [read_excel_any(f) for f in files_crm32]
        df_crm32_list = [df for df in df_crm32_list if df is not None]

        if not df_crm4_list or not df_crm32_list:
            st.error("Không đọc được dữ liệu CRM4/CRM32.")
            st.stop()

        df_crm4 = pd.concat(df_crm4_list, ignore_index=True)
        df_crm32 = pd.concat(df_crm32_list, ignore_index=True)

        # CODE files
        df_muc_dich_file = read_excel_any(file_mucdich) if file_mucdich else pd.DataFrame()
        df_code_tsbd_file = read_excel_any(file_loaits) if file_loaits else pd.DataFrame()

        # ====== Chuẩn hoá ID ======
        if 'CIF_KH_VAY' in df_crm4.columns:
            df_crm4['CIF_KH_VAY'] = pd.to_numeric(df_crm4['CIF_KH_VAY'], errors='coerce')
            df_crm4['CIF_KH_VAY'] = df_crm4['CIF_KH_VAY'].dropna().astype('int64').astype(str)

        if 'CUSTSEQLN' in df_crm32.columns:
            df_crm32['CUSTSEQLN'] = pd.to_numeric(df_crm32['CUSTSEQLN'], errors='coerce')
            df_crm32['CUSTSEQLN'] = df_crm32['CUSTSEQLN'].dropna().astype('int64').astype(str)

        # ====== Lọc theo chi nhánh ======
        if sol_branch:
            chi_nhanh = sol_branch.upper()
            if 'BRANCH_VAY' in df_crm4.columns:
                df_crm4_filtered = df_crm4[df_crm4['BRANCH_VAY'].astype(str).str.upper().str.contains(chi_nhanh)]
            else:
                df_crm4_filtered = df_crm4.copy()
            if 'BRCD' in df_crm32.columns:
                df_crm32_filtered = df_crm32[df_crm32['BRCD'].astype(str).str.upper().str.contains(chi_nhanh)]
            else:
                df_crm32_filtered = df_crm32.copy()
        else:
            df_crm4_filtered = df_crm4.copy()
            df_crm32_filtered = df_crm32.copy()

        st.info(f"CRM4 sau lọc: {len(df_crm4_filtered):,} dòng | CRM32 sau lọc: {len(df_crm32_filtered):,} dòng")

        # ====== Map loại TSBĐ ======
        if not df_code_tsbd_file.empty and {'CODE CAP 2','CODE'}.issubset(df_code_tsbd_file.columns):
            df_code_tsbd = df_code_tsbd_file[['CODE CAP 2', 'CODE']].copy()
            df_code_tsbd.columns = ['CAP_2','LOAI_TS']
            df_tsbd_code = df_code_tsbd[['CAP_2','LOAI_TS']].drop_duplicates()
            if 'CAP_2' in df_crm4_filtered.columns:
                df_crm4_filtered = df_crm4_filtered.merge(df_tsbd_code, how='left', on='CAP_2')
                df_crm4_filtered['LOAI_TS'] = df_crm4_filtered.apply(
                    lambda row: 'Không TS' if pd.isna(row.get('CAP_2')) or str(row.get('CAP_2')).strip()=='' else row.get('LOAI_TS'),
                    axis=1
                )
                df_crm4_filtered['GHI_CHU_TSBD'] = df_crm4_filtered.apply(
                    lambda row: 'MỚI' if str(row.get('CAP_2')).strip()!='' and pd.isna(row.get('LOAI_TS')) else '',
                    axis=1
                )
        else:
            st.warning("Không có hoặc thiếu cột trong CODE_LOAI TSBD.xlsx -> bỏ qua mapping LOAI_TS.")

        # ====== Loại bỏ loại hình Bao lanh/LC ======
        if 'LOAI' in df_crm4_filtered.columns:
            df_vay_4 = df_crm4_filtered.copy()
            df_vay = df_vay_4[~df_vay_4['LOAI'].isin(['Bao lanh','LC'])].copy()
        else:
            df_vay = df_crm4_filtered.copy()

        # ====== Pivot TSBĐ & Dư nợ ======
        if {'CIF_KH_VAY','LOAI_TS','TS_KW_VND'}.issubset(df_vay.columns):
            pivot_ts = df_vay.pivot_table(
                index='CIF_KH_VAY', columns='LOAI_TS', values='TS_KW_VND', aggfunc='sum', fill_value=0
            ).add_suffix(' (Giá trị TS)').reset_index()
        else:
            pivot_ts = pd.DataFrame()

        if {'CIF_KH_VAY','LOAI_TS','DU_NO_PHAN_BO_QUY_DOI'}.issubset(df_vay.columns):
            pivot_no = df_vay.pivot_table(
                index='CIF_KH_VAY', columns='LOAI_TS', values='DU_NO_PHAN_BO_QUY_DOI', aggfunc='sum', fill_value=0
            ).reset_index()
        else:
            pivot_no = pd.DataFrame()

        if not pivot_no.empty:
            pivot_merge = pivot_no.merge(pivot_ts, on='CIF_KH_VAY', how='left') if not pivot_ts.empty else pivot_no.copy()
            # Tổng
            pivot_merge['GIÁ TRỊ TS'] = pivot_ts.drop(columns='CIF_KH_VAY', errors='ignore').sum(axis=1) if not pivot_ts.empty else 0
            pivot_merge['DƯ NỢ'] = pivot_no.drop(columns='CIF_KH_VAY', errors='ignore').sum(axis=1)
        else:
            pivot_merge = pd.DataFrame()

        if {'CIF_KH_VAY','TEN_KH_VAY','CUSTTPCD','NHOM_NO'}.issubset(df_crm4_filtered.columns) and not pivot_merge.empty:
            df_info = df_crm4_filtered[['CIF_KH_VAY','TEN_KH_VAY','CUSTTPCD','NHOM_NO']].drop_duplicates(subset='CIF_KH_VAY')
            pivot_final = df_info.merge(pivot_merge, on='CIF_KH_VAY', how='left')
            pivot_final = pivot_final.reset_index().rename(columns={'index':'STT'})
            pivot_final['STT'] += 1
        else:
            pivot_final = pd.DataFrame()

        # ====== CRM32: Chuẩn hoá CAP_PHE_DUYET -> MA_PHE_DUYET ======
        if 'CAP_PHE_DUYET' in df_crm32_filtered.columns:
            df_crm32_filtered['MA_PHE_DUYET'] = df_crm32_filtered['CAP_PHE_DUYET'].astype(str).str.split('-').str[0].str.strip().str.zfill(2)

        # Danh sách cấp C
        ma_cap_c = [f"{i:02d}" for i in range(1, 8)] + [f"{i:02d}" for i in range(28, 32)]
        list_cif_cap_c = df_crm32_filtered[df_crm32_filtered.get('MA_PHE_DUYET','').isin(ma_cap_c)]['CUSTSEQLN'].unique() if 'MA_PHE_DUYET' in df_crm32_filtered.columns else []

        # Cơ cấu
        list_co_cau = ['ACOV1','ACOV3','ATT01','ATT02','ATT03','ATT04','BCOV1','BCOV2','BTT01','BTT02','BTT03','CCOV2','CCOV3','CTT03','RCOV3','RTT03']
        cif_co_cau = df_crm32_filtered[df_crm32_filtered.get('SCHEME_CODE','').isin(list_co_cau)]['CUSTSEQLN'].unique() if 'SCHEME_CODE' in df_crm32_filtered.columns else []

        # ====== Mục đích vay (CRM32) ======
        if not df_muc_dich_file.empty and {'CODE_MDSDV4','GROUP'}.issubset(df_muc_dich_file.columns):
            df_muc_dich_vay = df_muc_dich_file[['CODE_MDSDV4','GROUP']].copy()
            df_muc_dich_vay.columns = ['MUC_DICH_VAY_CAP_4','MUC DICH']
            if 'MUC_DICH_VAY_CAP_4' in df_crm32_filtered.columns:
                df_crm32_filtered = df_crm32_filtered.merge(df_muc_dich_vay, how='left', on='MUC_DICH_VAY_CAP_4')
                df_crm32_filtered['MUC DICH'] = df_crm32_filtered['MUC DICH'].fillna('(blank)')
                df_crm32_filtered['GHI_CHU_TSBD'] = df_crm32_filtered.apply(
                    lambda row: 'MỚI' if str(row.get('MUC_DICH_VAY_CAP_4')).strip()!='' and pd.isna(row.get('MUC DICH')) else '',
                    axis=1
                )
        else:
            st.warning("Không có hoặc thiếu cột trong CODE_MDSDV4.xlsx -> bỏ qua mapping mục đích vay.")

        # Pivot CRM32 theo mục đích
        if {'CUSTSEQLN','MUC DICH','DU_NO_QUY_DOI'}.issubset(df_crm32_filtered.columns):
            pivot_mucdich = df_crm32_filtered.pivot_table(
                index='CUSTSEQLN', columns='MUC DICH', values='DU_NO_QUY_DOI',
                aggfunc='sum', fill_value=0
            ).reset_index()
            pivot_mucdich['DƯ NỢ CRM32'] = pivot_mucdich.drop(columns='CUSTSEQLN').sum(axis=1)
            pivot_final_CRM32 = pivot_mucdich.rename(columns={'CUSTSEQLN':'CIF_KH_VAY'})
        else:
            pivot_mucdich = pd.DataFrame()
            pivot_final_CRM32 = pd.DataFrame()

        # ====== Ghép pivot_final & CRM32 ======
        if not pivot_final.empty and not pivot_final_CRM32.empty:
            pivot_full = pivot_final.merge(pivot_final_CRM32, on='CIF_KH_VAY', how='left').fillna(0)
            pivot_full['LECH'] = pivot_full['DƯ NỢ'] - pivot_full['DƯ NỢ CRM32']
        else:
            pivot_full = pivot_final.copy() if not pivot_final.empty else pd.DataFrame()

        # Bổ sung (blank) từ CRM4 (loại 'Cho vay','Bao lanh','LC')
        if not pivot_full.empty and 'LOAI' in df_crm4_filtered.columns:
            df_crm4_blank = df_crm4_filtered[~df_crm4_filtered['LOAI'].isin(['Cho vay','Bao lanh','LC'])].copy()
            if {'CIF_KH_VAY','DU_NO_PHAN_BO_QUY_DOI'}.issubset(df_crm4_blank.columns):
                du_no_bosung = (df_crm4_blank.groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI']
                                .sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI':'(blank)'}))
                pivot_full = pivot_full.merge(du_no_bosung, on='CIF_KH_VAY', how='left')
                pivot_full['(blank)'] = pivot_full['(blank)'].fillna(0)
                if 'DƯ NỢ CRM32' in pivot_full.columns:
                    pivot_full['DƯ NỢ CRM32'] = pivot_full['DƯ NỢ CRM32'] + pivot_full['(blank)']

        # Cờ nhóm nợ / CAP C / Cơ cấu
        if 'NHOM_NO' in pivot_full.columns:
            pivot_full['Nợ nhóm 2'] = pivot_full['NHOM_NO'].apply(lambda x: 'x' if str(x).strip()=='2' else '')
            pivot_full['Nợ xấu'] = pivot_full['NHOM_NO'].apply(lambda x: 'x' if str(x).strip() in ['3','4','5'] else '')
        if not isinstance(list_cif_cap_c, list):
            list_cif_cap_c = list(list_cif_cap_c)
        if not isinstance(cif_co_cau, list):
            cif_co_cau = list(cif_co_cau)
        if 'CIF_KH_VAY' in pivot_full.columns:
            pivot_full['Chuyên gia PD cấp C duyệt'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in list_cif_cap_c else '')
            pivot_full['NỢ CƠ_CẤU'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in cif_co_cau else '')

        # Bảo lãnh / LC
        if 'LOAI' in df_crm4_filtered.columns and {'CIF_KH_VAY','DU_NO_PHAN_BO_QUY_DOI'}.issubset(df_crm4_filtered.columns):
            df_baolanh = df_crm4_filtered[df_crm4_filtered['LOAI']=='Bao lanh']
            df_lc = df_crm4_filtered[df_crm4_filtered['LOAI']=='LC']
            df_baolanh_sum = df_baolanh.groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI':'DƯ_NỢ_BẢO_LÃNH'})
            df_lc_sum = df_lc.groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI':'DƯ_NỢ_LC'})
            if not pivot_full.empty:
                pivot_full = pivot_full.drop(columns=[c for c in ['DƯ_NỢ_BẢO_LÃNH','DƯ_NỢ_LC'] if c in pivot_full.columns], errors='ignore')
                pivot_full = pivot_full.merge(df_baolanh_sum, on='CIF_KH_VAY', how='left').merge(df_lc_sum, on='CIF_KH_VAY', how='left')
                pivot_full['DƯ_NỢ_BẢO_LÃNH'] = pivot_full['DƯ_NỢ_BẢO_LÃNH'].fillna(0)
                pivot_full['DƯ_NỢ_LC'] = pivot_full['DƯ_NỢ_LC'].fillna(0)

        # ====== Giải ngân tiền mặt 1 tỷ ======
        if file_giaingan is not None and not pivot_full.empty and 'KHE_UOC' in df_crm32_filtered.columns and 'CUSTSEQLN' in df_crm32_filtered.columns:
            df_giai_ngan = read_excel_any(file_giaingan)
            if df_giai_ngan is not None and 'FORACID' in df_giai_ngan.columns:
                df_crm32_filtered['KHE_UOC'] = df_crm32_filtered['KHE_UOC'].astype(str).str.strip()
                df_crm32_filtered['CUSTSEQLN'] = df_crm32_filtered['CUSTSEQLN'].astype(str).str.strip()
                df_giai_ngan['FORACID'] = df_giai_ngan['FORACID'].astype(str).str.strip()
                pivot_full['CIF_KH_VAY'] = pivot_full['CIF_KH_VAY'].astype(str).str.strip()

                df_match = df_crm32_filtered[df_crm32_filtered['KHE_UOC'].isin(df_giai_ngan['FORACID'])].copy()
                ds_cif_tien_mat = df_match['CUSTSEQLN'].unique()
                pivot_full['GIẢI_NGÂN_TIEN_MAT'] = pivot_full['CIF_KH_VAY'].isin(ds_cif_tien_mat).map({True:'x', False:''})
            else:
                st.warning("File Giải ngân không có cột FORACID -> bỏ qua cờ GIẢI_NGÂN_TIEN_MAT.")

        # ====== Cầm cố tại TCTD khác ======
        if 'CAP_2' in df_crm4_filtered.columns and 'CIF_KH_VAY' in df_crm4_filtered.columns:
            df_cc_tctd = df_crm4_filtered[df_crm4_filtered['CAP_2'].astype(str).str.contains('TCTD', case=False, na=False)]
            df_cc_flag = df_cc_tctd[['CIF_KH_VAY']].drop_duplicates()
            df_cc_flag['Cầm cố tại TCTD khác'] = 'x'
            if not pivot_full.empty:
                pivot_full = pivot_full.merge(df_cc_flag, on='CIF_KH_VAY', how='left')
                pivot_full['Cầm cố tại TCTD khác'] = pivot_full['Cầm cố tại TCTD khác'].fillna('')

        # ====== Top 10 KHCN/KHDN theo DƯ NỢ ======
        if not pivot_full.empty and {'CUSTTPCD','DƯ NỢ','CIF_KH_VAY'}.issubset(pivot_full.columns):
            top_khcn = pivot_full[pivot_full['CUSTTPCD']=='Ca nhan'].nlargest(10,'DƯ NỢ')['CIF_KH_VAY']
            top_khdn = pivot_full[pivot_full['CUSTTPCD']=='Doanh nghiep'].nlargest(10,'DƯ NỢ')['CIF_KH_VAY']
            pivot_full['Top 10 dư nợ KHCN'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in top_khcn.values else '')
            pivot_full['Top 10 dư nợ KHDN'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in top_khdn.values else '')

        # ====== R34: quá hạn định giá TSBĐ ======
        if 'LOAI_TS' in df_crm4_filtered.columns and 'VALUATION_DATE' in df_crm4_filtered.columns:
            loai_ts_r34 = ['BĐS','MMTB','PTVT']
            mask_r34 = df_crm4_filtered['LOAI_TS'].isin(loai_ts_r34)
            df_crm4_filtered['VALUATION_DATE'] = ensure_datetime(df_crm4_filtered['VALUATION_DATE'])
            ngay_eval = pd.to_datetime(ngay_danh_gia)
            df_crm4_filtered.loc[mask_r34, 'SO_NGAY_QUA_HAN'] = (ngay_eval - df_crm4_filtered.loc[mask_r34, 'VALUATION_DATE']).dt.days - 365
            df_crm4_filtered.loc[mask_r34, 'SO_THANG_QUA_HAN'] = ((ngay_eval - df_crm4_filtered.loc[mask_r34, 'VALUATION_DATE']).dt.days / 31) - 12
            cif_quahan = df_crm4_filtered[df_crm4_filtered['SO_NGAY_QUA_HAN'] > 30]['CIF_KH_VAY'].unique()
            if not pivot_full.empty:
                pivot_full['KH có TSBĐ quá hạn định giá'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'X' if x in cif_quahan else '')

        # ====== Mục 17: TSBĐ khác địa bàn ======
        sheets_m17 = {}
        if file_muc17 is not None:
            df_sol = read_excel_any(file_muc17)
            if df_sol is not None and 'C01' in df_sol.columns:
                ds_secu = df_crm4_filtered.get('SECU_SRL_NUM', pd.Series([], dtype=object)).dropna().unique()
                df_17_filtered = df_sol[df_sol['C01'].isin(ds_secu)] if len(ds_secu)>0 else df_sol.copy()

                # Bất động sản
                if 'C02' in df_17_filtered.columns:
                    df_bds = df_17_filtered[df_17_filtered['C02'].astype(str).str.strip()=='Bat dong san'].copy()
                else:
                    df_bds = pd.DataFrame()

                # Match mã TS với CRM4
                if not df_bds.empty and 'SECU_SRL_NUM' in df_crm4.columns:
                    df_bds_matched = df_bds[df_bds['C01'].isin(df_crm4['SECU_SRL_NUM'])].copy()
                else:
                    df_bds_matched = df_bds.copy()

                # Extract tỉnh/thành từ C19
                def extract_tinh_thanh(diachi):
                    if pd.isna(diachi): return ''
                    parts = str(diachi).split(',')
                    return parts[-1].strip().lower() if parts else ''
                if not df_bds_matched.empty and 'C19' in df_bds_matched.columns:
                    df_bds_matched['TINH_TP_TSBD'] = df_bds_matched['C19'].apply(extract_tinh_thanh)

                    # Provinces input
                    provinces_list = [t.strip().lower() for t in provinces.split(',') if t.strip()]
                    df_bds_matched['CANH_BAO_TS_KHAC_DIABAN'] = df_bds_matched['TINH_TP_TSBD'].apply(
                        lambda x: 'x' if x and (x.strip().lower() not in provinces_list) else ''
                    )
                    ma_ts_canh_bao = df_bds_matched[df_bds_matched['CANH_BAO_TS_KHAC_DIABAN']=='x']['C01'].unique()
                    if 'SECU_SRL_NUM' in df_crm4.columns and not pivot_full.empty:
                        cif_canh_bao = df_crm4[df_crm4['SECU_SRL_NUM'].isin(ma_ts_canh_bao)]['CIF_KH_VAY'].dropna().unique()
                        pivot_full['KH có TSBĐ khác địa bàn'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in cif_canh_bao else '')
                sheets_m17['tieu chi 2_dot3'] = df_bds_matched

        # ====== Mục 55/56: GN/TT cùng ngày ======
        sheets_55_56 = {}
        if file_55 is not None and file_56 is not None:
            df_55 = read_excel_any(file_55)
            df_56 = read_excel_any(file_56)
            if df_55 is not None and df_56 is not None:
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
                df_gop = df_gop[df_gop['NGAY'].notna()]
                df_gop = df_gop.sort_values(by=['CIF','NGAY','GIAI_NGAN_TT'])

                df_count = df_gop.groupby(['CIF','NGAY','GIAI_NGAN_TT']).size().unstack(fill_value=0).reset_index()
                df_count['CO_CA_GN_VA_TT'] = ((df_count.get('Giải ngân',0)>0) & (df_count.get('Tất toán',0)>0)).astype(int)

                ds_nhieu_tt_gn = df_count[(df_count.get('Giải ngân',0)>1) | (df_count.get('Tất toán',0)>1)]['CIF'].astype(str).unique()
                ds_ca_gn_tt = df_count[df_count['CO_CA_GN_VA_TT']==1]['CIF'].astype(str).unique()

                if not pivot_full.empty and 'CIF_KH_VAY' in pivot_full.columns:
                    pivot_full['CIF_KH_VAY'] = pivot_full['CIF_KH_VAY'].astype(str)
                    pivot_full['KH có cả GNG và TT trong 1 ngày'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in ds_ca_gn_tt else '')

                sheets_55_56['tieu chi 3_dot3'] = df_gop
                sheets_55_56['tieu chi 3_dot3_1'] = df_count

        # ====== Mục 57: Chậm trả ======
        df_delay_out = pd.DataFrame()
        if file_57 is not None:
            df_delay = read_excel_any(file_57)
            if df_delay is not None and {'NGAY_DEN_HAN_TT','NGAY_THANH_TOAN','CIF_ID'}.issubset(df_delay.columns):
                df_delay['NGAY_DEN_HAN_TT'] = ensure_datetime(df_delay['NGAY_DEN_HAN_TT'])
                df_delay['NGAY_THANH_TOAN'] = ensure_datetime(df_delay['NGAY_THANH_TOAN'])
                ngay_eval = pd.to_datetime(ngay_danh_gia)
                df_delay['NGAY_THANH_TOAN_FILL'] = df_delay['NGAY_THANH_TOAN'].fillna(ngay_eval)
                df_delay['SO_NGAY_CHAM_TRA'] = (df_delay['NGAY_THANH_TOAN_FILL'] - df_delay['NGAY_DEN_HAN_TT']).dt.days
                mask_period = df_delay['NGAY_DEN_HAN_TT'].dt.year.between(2023, 2025)
                df_delay = df_delay[mask_period].copy()

                # merge dư nợ & nhóm nợ
                if not pivot_full.empty and {'CIF_KH_VAY','DƯ NỢ','NHOM_NO'}.issubset(pivot_full.columns):
                    df_crm32_tmp = pivot_full[['CIF_KH_VAY','DƯ NỢ','NHOM_NO']].copy().rename(columns={'CIF_KH_VAY':'CIF_ID'})
                    df_crm32_tmp['CIF_ID'] = df_crm32_tmp['CIF_ID'].astype(str)
                    df_delay['CIF_ID'] = df_delay['CIF_ID'].astype(str)
                    df_delay = df_delay.merge(df_crm32_tmp, on='CIF_ID', how='left')
                    df_delay = df_delay[df_delay['NHOM_NO']==1].copy()

                def cap_cham_tra(days):
                    if pd.isna(days): return None
                    elif days >= 10: return '>=10'
                    elif days >= 4: return '4-9'
                    elif days > 0: return '<4'
                    else: return None
                df_delay['CAP_CHAM_TRA'] = df_delay['SO_NGAY_CHAM_TRA'].apply(cap_cham_tra)
                df_delay = df_delay.dropna(subset=['CAP_CHAM_TRA']).copy()

                df_delay['NGAY'] = df_delay['NGAY_DEN_HAN_TT'].dt.date
                df_delay.sort_values(['CIF_ID','NGAY','CAP_CHAM_TRA'],
                                     key=lambda s: s.map({'>=10':0, '4-9':1, '<4':2}),
                                     inplace=True)
                df_unique = df_delay.drop_duplicates(subset=['CIF_ID','NGAY'], keep='first').copy()

                df_dem = df_unique.groupby(['CIF_ID','CAP_CHAM_TRA']).size().unstack(fill_value=0)
                # Cờ
                df_dem['KH Phát sinh chậm trả > 10 ngày'] = np.where(df_dem.get('>=10',0) > 0, 'x', '')
                df_dem['KH Phát sinh chậm trả 4-9 ngày'] = np.where((df_dem.get('>=10',0)==0) & (df_dem.get('4-9',0)>0), 'x', '')

                if not pivot_full.empty:
                    pivot_full['CIF_KH_VAY'] = pivot_full['CIF_KH_VAY'].astype(str)
                    cols_to_merge = ['KH Phát sinh chậm trả > 10 ngày','KH Phát sinh chậm trả 4-9 ngày']
                    cols_to_merge_existing = [c for c in cols_to_merge if c in df_dem.columns]
                    if cols_to_merge_existing:
                        pivot_full = pivot_full.merge(df_dem[cols_to_merge_existing], left_on='CIF_KH_VAY', right_index=True, how='left')
                        for col in cols_to_merge_existing:
                            pivot_full[col] = pivot_full[col].fillna('')
                df_delay_out = df_delay.copy()

        # ============== Show & Export ==============
        st.subheader("2) Kết quả")
        sheets = {}
        sheets['df_crm4_LOAI_TS'] = df_crm4_filtered
        if 'pivot_final' in locals(): sheets['KQ_CRM4'] = pivot_final
        if 'pivot_merge' in locals() and isinstance(pivot_merge, pd.DataFrame) and not pivot_merge.empty:
            sheets['Pivot_crm4'] = pivot_merge
        sheets['df_crm32_LOAI_TS'] = df_crm32_filtered
        if 'pivot_full' in locals() and not pivot_full.empty:
            sheets['KQ_KH'] = pivot_full
        if 'pivot_mucdich' in locals() and not pivot_mucdich.empty:
            sheets['Pivot_crm32'] = pivot_mucdich
        if not df_delay_out.empty:
            sheets['tieu chi 4'] = df_delay_out
        if 'sheets_55_56' in locals():
            sheets.update(sheets_55_56)
        if 'sheets_m17' in locals():
            sheets.update(sheets_m17)

        # Display preview
        for name, df in list(sheets.items())[:6]:  # show up to 6 previews
            st.markdown(f"**{name}**  
{len(df):,} dòng")
            st.dataframe(df.head(200))

        download_excel_sheets(sheets, default_name="KQ_2241_streamlit.xlsx")

        st.success("Hoàn thành! ✅")
