import streamlit as st
import pandas as pd
import numpy as np
import io

# -----------------------------------------------------------------------------
# Cấu hình trang và tiêu đề
# -----------------------------------------------------------------------------
st.set_page_config(layout="wide", page_title="Báo cáo Phân tích Tín dụng")
st.title("Báo cáo Phân tích Dữ liệu Tín dụng Tổng hợp")

# -----------------------------------------------------------------------------
# Sidebar để tải tệp và nhập thông tin
# -----------------------------------------------------------------------------
st.sidebar.header("Cài đặt và Tải tệp")

# 1. Tải tệp
crm4_files = st.sidebar.file_uploader(
    "1. Tải tệp CRM4 (Du_no_theo_tai_san_dam_bao)", 
    type=['xlsx', 'xls'], 
    accept_multiple_files=True
)
crm32_files = st.sidebar.file_uploader(
    "2. Tải tệp CRM32 (RPT_CRM_32)", 
    type=['xlsx', 'xls'], 
    accept_multiple_files=True
)
df_muc_dich_file = st.sidebar.file_uploader("3. Tải tệp CODE_MDSDV4.xlsx", type=['xlsx', 'xls'])
df_code_tsbd_file = st.sidebar.file_uploader("4. Tải tệp CODE_LOAI TSBD.xlsx", type=['xlsx', 'xls'])
df_giai_ngan_file = st.sidebar.file_uploader("5. Tải tệp Giải ngân tiền mặt >= 1 tỷ", type=['xlsx', 'xls'])
df_sol_file = st.sidebar.file_uploader("6. Tải tệp Mục 17 (Chi tiết TSĐB)", type=['xlsx', 'xls'])
df_55_file = st.sidebar.file_uploader("7. Tải tệp Mục 55 (Tất toán)", type=['xlsx', 'xls'])
df_56_file = st.sidebar.file_uploader("8. Tải tệp Mục 56 (Giải ngân)", type=['xlsx', 'xls'])
df_delay_file = st.sidebar.file_uploader("9. Tải tệp Mục 57 (Chậm trả)", type=['xlsx', 'xls'])


# 2. Nhập thông tin
chi_nhanh = st.sidebar.text_input("Nhập tên chi nhánh hoặc mã SOL cần lọc (ví dụ: HANOI hoặc 001)", "HANOI").strip().upper()

dia_ban_kt_input = st.sidebar.text_input("Nhập tỉnh/thành của đơn vị KT (cách nhau bởi dấu phẩy)", "Hồ Chí Minh, Long An")

ngay_danh_gia_input = st.sidebar.date_input("Chọn ngày đánh giá", pd.to_datetime("2025-08-31"))
ngay_danh_gia = pd.to_datetime(ngay_danh_gia_input)


# Nút để bắt đầu xử lý
if st.sidebar.button("Xử lý dữ liệu"):
    # Kiểm tra xem tất cả các tệp cần thiết đã được tải lên chưa
    required_files = [
        crm4_files, crm32_files, df_muc_dich_file, df_code_tsbd_file, 
        df_giai_ngan_file, df_sol_file, df_55_file, df_56_file, df_delay_file
    ]
    if not all(required_files):
        st.warning("Vui lòng tải lên đầy đủ 9 loại tệp được yêu cầu.")
        st.stop()

    with st.spinner('Đang xử lý... Vui lòng đợi trong giây lát.'):
        try:
            # =============================================================================
            # PHẦN 1: ĐỌC VÀ GHÉP DỮ LIỆU
            # =============================================================================
            
            # 1.1. Ghép tất cả file CRM4
            df_crm4_ghep = [pd.read_excel(f) for f in crm4_files]
            df_crm4 = pd.concat(df_crm4_ghep, ignore_index=True)

            # 1.2. Ghép tất cả file CRM32
            df_crm32_ghep = [pd.read_excel(f) for f in crm32_files]
            df_crm32 = pd.concat(df_crm32_ghep, ignore_index=True)

            # 1.3. Đọc các file mã
            df_muc_dich = pd.read_excel(df_muc_dich_file)
            df_code_tsbd = pd.read_excel(df_code_tsbd_file)

            # =============================================================================
            # PHẦN 2: CHUẨN HÓA VÀ LỌC DỮ LIỆU BAN ĐẦU
            # =============================================================================
            for df in [df_crm4]:
                if 'CIF_KH_VAY' in df.columns:
                    df['CIF_KH_VAY'] = pd.to_numeric(df['CIF_KH_VAY'], errors='coerce')
                    df['CIF_KH_VAY'] = df['CIF_KH_VAY'].dropna().astype('int64').astype(str)

            for df in [df_crm32]:
                if 'CUSTSEQLN' in df.columns:
                    df['CUSTSEQLN'] = pd.to_numeric(df['CUSTSEQLN'], errors='coerce')
                    df['CUSTSEQLN'] = df['CUSTSEQLN'].dropna().astype('int64').astype(str)

            # Lọc dữ liệu theo chi nhánh
            df_crm4_filtered = df_crm4[df_crm4['BRANCH_VAY'].astype(str).str.upper().str.contains(chi_nhanh)].copy()
            df_crm32_filtered = df_crm32[df_crm32['BRCD'].astype(str).str.upper().str.contains(chi_nhanh)].copy()

            st.success(f"Đã tải và lọc dữ liệu cho chi nhánh '{chi_nhanh}'. Bắt đầu tổng hợp...")
            
            # =============================================================================
            # PHẦN 3: XỬ LÝ CRM4 - TÀI SẢN BẢO ĐẢM
            # =============================================================================
            df_code_tsbd_map = df_code_tsbd[['CODE CAP 2', 'CODE']].copy()
            df_code_tsbd_map.columns = ['CAP_2', 'LOAI_TS']
            df_tsbd_code = df_code_tsbd_map.drop_duplicates()

            df_crm4_filtered = df_crm4_filtered.merge(df_tsbd_code, how='left', on='CAP_2')
            df_crm4_filtered['LOAI_TS'] = df_crm4_filtered.apply(
                lambda row: 'Không TS' if pd.isna(row['CAP_2']) or str(row['CAP_2']).strip() == '' else row['LOAI_TS'],
                axis=1
            )
            df_crm4_filtered['GHI_CHU_TSBD'] = df_crm4_filtered.apply(
                lambda row: 'MỚI' if str(row['CAP_2']).strip() != '' and pd.isna(row['LOAI_TS']) else '',
                axis=1
            )
            
            df_vay = df_crm4_filtered[~df_crm4_filtered['LOAI'].isin(['Bao lanh', 'LC'])].copy()

            pivot_ts = df_vay.pivot_table(
                index='CIF_KH_VAY', columns='LOAI_TS', values='TS_KW_VND',
                aggfunc='sum', fill_value=0
            ).add_suffix(' (Giá trị TS)').reset_index()

            pivot_no = df_vay.pivot_table(
                index='CIF_KH_VAY', columns='LOAI_TS', values='DU_NO_PHAN_BO_QUY_DOI',
                aggfunc='sum', fill_value=0
            ).reset_index()

            pivot_merge = pivot_no.merge(pivot_ts, on='CIF_KH_VAY', how='left')
            if not pivot_ts.empty:
                pivot_merge['GIÁ TRỊ TS'] = pivot_ts.drop(columns='CIF_KH_VAY').sum(axis=1)
            else:
                pivot_merge['GIÁ TRỊ TS'] = 0
            
            if not pivot_no.empty:
                pivot_merge['DƯ NỢ'] = pivot_no.drop(columns='CIF_KH_VAY').sum(axis=1)
            else:
                pivot_merge['DƯ NỢ'] = 0
            
            df_info = df_crm4_filtered[['CIF_KH_VAY', 'TEN_KH_VAY', 'CUSTTPCD', 'NHOM_NO']].drop_duplicates(subset='CIF_KH_VAY')
            pivot_final = df_info.merge(pivot_merge, on='CIF_KH_VAY', how='left')
            pivot_final = pivot_final.reset_index().rename(columns={'index': 'STT'})
            pivot_final['STT'] += 1

            cols_order = ['STT', 'CUSTTPCD', 'CIF_KH_VAY', 'TEN_KH_VAY', 'NHOM_NO'] + \
                        sorted([col for col in pivot_merge.columns if col not in ['CIF_KH_VAY', 'GIÁ TRỊ TS', 'DƯ NỢ'] and '(Giá trị TS)' not in col]) + \
                        sorted([col for col in pivot_merge.columns if '(Giá trị TS)' in col]) + \
                        ['DƯ NỢ', 'GIÁ TRỊ TS']
            
            # Ensure all columns exist before reordering
            pivot_final_cols = [col for col in cols_order if col in pivot_final.columns]
            pivot_final = pivot_final[pivot_final_cols]


            # =============================================================================
            # PHẦN 4: XỬ LÝ CRM32 - MỤC ĐÍCH VAY, CẤP PHÊ DUYỆT
            # =============================================================================
            df_crm32_filtered['MA_PHE_DUYET'] = df_crm32_filtered['CAP_PHE_DUYET'].astype(str).str.split('-').str[0].str.strip().str.zfill(2)
            ma_cap_c = [f"{i:02d}" for i in range(1, 8)] + [f"{i:02d}" for i in range(28, 32)]
            list_cif_cap_c = df_crm32_filtered[df_crm32_filtered['MA_PHE_DUYET'].isin(ma_cap_c)]['CUSTSEQLN'].unique()

            list_co_cau = ['ACOV1', 'ACOV3', 'ATT01', 'ATT02', 'ATT03', 'ATT04', 'BCOV1', 'BCOV2', 'BTT01', 'BTT02', 'BTT03', 'CCOV2', 'CCOV3', 'CTT03', 'RCOV3', 'RTT03']
            cif_co_cau = df_crm32_filtered[df_crm32_filtered['SCHEME_CODE'].isin(list_co_cau)]['CUSTSEQLN'].unique()
            
            df_muc_dich_vay_map = df_muc_dich[['CODE_MDSDV4', 'GROUP']].copy()
            df_muc_dich_vay_map.columns = ['MUC_DICH_VAY_CAP_4', 'MUC DICH']
            df_crm32_filtered = df_crm32_filtered.merge(df_muc_dich_vay_map, how='left', on='MUC_DICH_VAY_CAP_4')
            df_crm32_filtered['MUC DICH'] = df_crm32_filtered['MUC DICH'].fillna('(blank)')
            
            pivot_mucdich = df_crm32_filtered.pivot_table(
                index='CUSTSEQLN', columns='MUC DICH', values='DU_NO_QUY_DOI',
                aggfunc='sum', fill_value=0
            ).reset_index()
            pivot_mucdich['DƯ NỢ CRM32'] = pivot_mucdich.drop(columns='CUSTSEQLN').sum(axis=1)

            # =============================================================================
            # PHẦN 5: GỘP CRM4 VÀ CRM32, XỬ LÝ CHÊNH LỆCH
            # =============================================================================
            pivot_final_CRM32 = pivot_mucdich.rename(columns={'CUSTSEQLN': 'CIF_KH_VAY'})
            pivot_full = pivot_final.merge(pivot_final_CRM32, on='CIF_KH_VAY', how='left')
            pivot_full.fillna(0, inplace=True)
            
            pivot_full['LECH'] = pivot_full['DƯ NỢ'] - pivot_full['DƯ NỢ CRM32']
            cif_lech = pivot_full[pivot_full['LECH'] != 0]['CIF_KH_VAY'].unique()
            df_crm4_blank = df_crm4_filtered[~df_crm4_filtered['LOAI'].isin(['Cho vay', 'Bao lanh', 'LC'])].copy()
            
            du_no_bosung = df_crm4_blank[df_crm4_blank['CIF_KH_VAY'].isin(cif_lech)].groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI': '(blank)'})
            
            pivot_full = pivot_full.merge(du_no_bosung, on='CIF_KH_VAY', how='left')
            pivot_full['(blank)'] = pivot_full['(blank)'].fillna(0)
            pivot_full['DƯ NỢ CRM32'] += pivot_full['(blank)']

            if '(blank)' in pivot_full.columns and 'DƯ NỢ CRM32' in pivot_full.columns:
                cols = list(pivot_full.columns)
                cols.insert(cols.index('DƯ NỢ CRM32'), cols.pop(cols.index('(blank)')))
                pivot_full = pivot_full[cols]
            
            # =============================================================================
            # PHẦN 6: THÊM CÁC CỘT CỜ (FLAG)
            # =============================================================================
            # Cờ nợ xấu, cơ cấu, cấp duyệt
            pivot_full['Nợ nhóm 2'] = pivot_full['NHOM_NO'].apply(lambda x: 'x' if str(x).strip() == '2' else '')
            pivot_full['Nợ xấu'] = pivot_full['NHOM_NO'].apply(lambda x: 'x' if str(x).strip() in ['3', '4', '5'] else '')
            pivot_full['Chuyên gia PD cấp C duyệt'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in list_cif_cap_c else '')
            pivot_full['NỢ CƠ_CẤU'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in cif_co_cau else '')

            # Dư nợ bảo lãnh, LC
            df_baolanh_sum = df_crm4_filtered[df_crm4_filtered['LOAI'] == 'Bao lanh'].groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI': 'DƯ_NỢ_BẢO_LÃNH'})
            df_lc_sum = df_crm4_filtered[df_crm4_filtered['LOAI'] == 'LC'].groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI': 'DƯ_NỢ_LC'})
            pivot_full = pivot_full.merge(df_baolanh_sum, on='CIF_KH_VAY', how='left').merge(df_lc_sum, on='CIF_KH_VAY', how='left')
            pivot_full[['DƯ_NỢ_BẢO_LÃNH', 'DƯ_NỢ_LC']] = pivot_full[['DƯ_NỢ_BẢO_LÃNH', 'DƯ_NỢ_LC']].fillna(0)
            
            # Cờ giải ngân tiền mặt
            df_giai_ngan = pd.read_excel(df_giai_ngan_file)
            df_crm32_filtered['KHE_UOC'] = df_crm32_filtered['KHE_UOC'].astype(str).str.strip()
            df_giai_ngan['FORACID'] = df_giai_ngan['FORACID'].astype(str).str.strip()
            df_match = df_crm32_filtered[df_crm32_filtered['KHE_UOC'].isin(df_giai_ngan['FORACID'])].copy()
            ds_cif_tien_mat = df_match['CUSTSEQLN'].unique()
            pivot_full['GIẢI_NGÂN_TIEN_MAT'] = pivot_full['CIF_KH_VAY'].astype(str).str.strip().isin(ds_cif_tien_mat).map({True: 'x', False: ''})
            
            # Cờ cầm cố TCTD khác
            df_cc_tctd = df_crm4_filtered[df_crm4_filtered['CAP_2'].str.contains('TCTD', case=False, na=False)]
            df_cc_flag = df_cc_tctd[['CIF_KH_VAY']].drop_duplicates()
            df_cc_flag['Cầm cố tại TCTD khác'] = 'x'
            pivot_full = pivot_full.merge(df_cc_flag, on='CIF_KH_VAY', how='left')
            pivot_full['Cầm cố tại TCTD khác'] = pivot_full['Cầm cố tại TCTD khác'].fillna('')

            # Cờ Top 10 dư nợ
            top10_khcn = pivot_full[pivot_full['CUSTTPCD'] == 'Ca nhan'].nlargest(10, 'DƯ NỢ')['CIF_KH_VAY']
            pivot_full['Top 10 dư nợ KHCN'] = pivot_full['CIF_KH_VAY'].isin(top10_khcn).map({True: 'x', False: ''})
            top10_khdn = pivot_full[pivot_full['CUSTTPCD'] == 'Doanh nghiep'].nlargest(10, 'DƯ NỢ')['CIF_KH_VAY']
            pivot_full['Top 10 dư nợ KHDN'] = pivot_full['CIF_KH_VAY'].isin(top10_khdn).map({True: 'x', False: ''})
            
            # Cờ TSBĐ quá hạn định giá
            df_crm4_filtered['VALUATION_DATE'] = pd.to_datetime(df_crm4_filtered['VALUATION_DATE'], errors='coerce')
            df_crm4_filtered['SO_NGAY_QUA_HAN'] = (ngay_danh_gia - df_crm4_filtered['VALUATION_DATE']).dt.days - 365
            cif_quahan = df_crm4_filtered[df_crm4_filtered['SO_NGAY_QUA_HAN'] > 30]['CIF_KH_VAY'].unique()
            pivot_full['KH có TSBĐ quá hạn định giá'] = pivot_full['CIF_KH_VAY'].isin(cif_quahan).map({True: 'x', False: ''})
            
            # Cờ TSBĐ khác địa bàn
            df_sol = pd.read_excel(df_sol_file)
            ds_secu = df_crm4_filtered['SECU_SRL_NUM'].dropna().unique()
            df_17_filtered = df_sol[df_sol['C01'].isin(ds_secu)]
            dia_ban_kt = [t.strip().lower() for t in dia_ban_kt_input.split(',') if t.strip()]
            
            df_bds = df_17_filtered[df_17_filtered['C02'].str.strip() == 'Bat dong san'].copy()
            df_bds_matched = df_bds[df_bds['C01'].isin(df_crm4['SECU_SRL_NUM'])].copy()
            
            def extract_tinh_thanh(diachi):
                if pd.isna(diachi): return ''
                parts = diachi.split(',')
                return parts[-1].strip().lower() if parts else ''
            
            df_bds_matched['TINH_TP_TSBD'] = df_bds_matched['C19'].apply(extract_tinh_thanh)
            df_bds_matched['CANH_BAO_TS_KHAC_DIABAN'] = df_bds_matched['TINH_TP_TSBD'].apply(
                lambda x: 'x' if x and x.strip().lower() not in dia_ban_kt else ''
            )
            ma_ts_canh_bao = df_bds_matched[df_bds_matched['CANH_BAO_TS_KHAC_DIABAN'] == 'x']['C01'].unique()
            cif_canh_bao = df_crm4[df_crm4['SECU_SRL_NUM'].isin(ma_ts_canh_bao)]['CIF_KH_VAY'].dropna().unique()
            pivot_full['KH có TSBĐ khác địa bàn'] = pivot_full['CIF_KH_VAY'].isin(cif_canh_bao).map({True: 'x', False: ''})
            
            # Cờ giải ngân và tất toán trong ngày
            df_55 = pd.read_excel(df_55_file)
            df_56 = pd.read_excel(df_56_file)
            df_tt = df_55[['CUSTSEQLN', 'KHE_UOC', 'NGAY_TT']].copy()
            df_tt.columns = ['CIF', 'KHE_UOC', 'NGAY']
            df_tt['GIAI_NGAN_TT'] = 'Tất toán'
            df_gn = df_56[['CIF', 'KHE_UOC', 'NGAY_GIAI_NGAN']].copy()
            df_gn.columns = ['CIF', 'KHE_UOC', 'NGAY']
            df_gn['GIAI_NGAN_TT'] = 'Giải ngân'
            df_gop_sub = pd.concat([df_tt, df_gn], ignore_index=True)
            df_gop_sub['NGAY'] = pd.to_datetime(df_gop_sub['NGAY'], format='%Y%m%d', errors='coerce').dt.date
            df_count = df_gop_sub.groupby(['CIF', 'NGAY', 'GIAI_NGAN_TT']).size().unstack(fill_value=0).reset_index()
            ds_ca_gn_tt = df_count[ (df_count.get('Giải ngân', 0) > 0) & (df_count.get('Tất toán', 0) > 0) ]['CIF'].astype(str).unique()
            pivot_full['KH có cả GN và TT trong 1 ngày'] = pivot_full['CIF_KH_VAY'].astype(str).isin(ds_ca_gn_tt).map({True: 'x', False: ''})
            
            # Cờ chậm trả
            df_delay = pd.read_excel(df_delay_file)
            df_delay['NGAY_DEN_HAN_TT'] = pd.to_datetime(df_delay['NGAY_DEN_HAN_TT'], errors='coerce')
            df_delay['NGAY_THANH_TOAN'] = pd.to_datetime(df_delay['NGAY_THANH_TOAN'], errors='coerce')
            df_delay['NGAY_THANH_TOAN_FILL'] = df_delay['NGAY_THANH_TOAN'].fillna(ngay_danh_gia)
            df_delay['SO_NGAY_CHAM_TRA'] = (df_delay['NGAY_THANH_TOAN_FILL'] - df_delay['NGAY_DEN_HAN_TT']).dt.days
            mask_period = df_delay['NGAY_DEN_HAN_TT'].dt.year.between(2023, 2025)
            df_delay = df_delay[mask_period].copy()

            df_delay = df_delay.merge(pivot_full[['CIF_KH_VAY', 'NHOM_NO']], left_on='CIF_ID', right_on='CIF_KH_VAY', how='left')
            df_delay = df_delay[df_delay['NHOM_NO'] == 1].copy()
            
            def cap_cham_tra(days):
                if pd.isna(days): return None
                elif days >= 10: return '>=10'
                elif days >= 4: return '4-9'
                elif days > 0: return '<4'
                else: return None
            
            df_delay['CAP_CHAM_TRA'] = df_delay['SO_NGAY_CHAM_TRA'].apply(cap_cham_tra)
            df_delay.dropna(subset=['CAP_CHAM_TRA'], inplace=True)
            df_delay['NGAY'] = df_delay['NGAY_DEN_HAN_TT'].dt.date
            df_delay.sort_values(['CIF_ID', 'NGAY', 'CAP_CHAM_TRA'], key=lambda s: s.map({'>=10':0, '4-9':1, '<4':2}), inplace=True)
            df_unique = df_delay.drop_duplicates(subset=['CIF_ID', 'NGAY'], keep='first').copy()
            df_dem = df_unique.groupby(['CIF_ID', 'CAP_CHAM_TRA']).size().unstack(fill_value=0)
            
            df_dem['KH Phát sinh chậm trả > 10 ngày'] = np.where(df_dem.get('>=10', 0) > 0, 'x', '')
            df_dem['KH Phát sinh chậm trả 4-9 ngày'] = np.where((df_dem.get('>=10', 0) == 0) & (df_dem.get('4-9', 0) > 0), 'x', '')
            
            pivot_full = pivot_full.merge(df_dem[['KH Phát sinh chậm trả > 10 ngày', 'KH Phát sinh chậm trả 4-9 ngày']], left_on='CIF_KH_VAY', right_index=True, how='left')
            pivot_full['KH Phát sinh chậm trả > 10 ngày'] = pivot_full['KH Phát sinh chậm trả > 10 ngày'].fillna('')
            pivot_full['KH Phát sinh chậm trả 4-9 ngày'] = pivot_full['KH Phát sinh chậm trả 4-9 ngày'].fillna('')

            st.success("Xử lý hoàn tất!")
            
            # =============================================================================
            # PHẦN 7: HIỂN THỊ VÀ TẢI KẾT QUẢ
            # =============================================================================
            
            st.header("Bảng kết quả tổng hợp khách hàng")
            st.dataframe(pivot_full)
            st.info(f"Tổng số khách hàng trong bảng: {len(pivot_full)}")

            # Tạo file Excel trong bộ nhớ để tải về
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                pivot_full.to_excel(writer, sheet_name='KQ_KH', index=False)
                df_crm4_filtered.to_excel(writer, sheet_name='df_crm4_filtered', index=False)
                pivot_final.to_excel(writer, sheet_name='KQ_CRM4', index=False)
                df_crm32_filtered.to_excel(writer, sheet_name='df_crm32_filtered', index=False)
                pivot_mucdich.to_excel(writer, sheet_name='Pivot_crm32', index=False)
                df_delay.to_excel(writer, sheet_name='ChiTiet_ChamTra', index=False)
                df_bds_matched.to_excel(writer, sheet_name='ChiTiet_TSBD_KhacDiaBan', index=False)

            # Nút tải xuống
            st.download_button(
                label="📥 Tải xuống file Excel kết quả",
                data=output.getvalue(),
                file_name=f'KQ_PhanTich_{chi_nhanh}_{ngay_danh_gia_input.strftime("%Y%m%d")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        except Exception as e:
            st.error(f"Đã có lỗi xảy ra trong quá trình xử lý: {e}")
            st.error("Vui lòng kiểm tra lại định dạng các file đầu vào hoặc các tham số đã nhập.")

else:
    st.info("Vui lòng tải lên đầy đủ các tệp dữ liệu và nhấn nút 'Xử lý dữ liệu' ở thanh bên trái.")
