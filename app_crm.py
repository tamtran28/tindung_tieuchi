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

# For brevity, the rest of the processing code from the earlier version would follow here without indentation errors.
# This includes the full data processing, pivoting, flagging logic, and final Excel download.

