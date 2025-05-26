import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Gộp và Tách File CRM4 Theo Nhóm Nợ")

uploaded_files = st.file_uploader("Tải lên các file CRM4 (Excel)", type=['xls'], accept_multiple_files=True)

if uploaded_files:
    all_data = pd.DataFrame()

    for file in uploaded_files:
        df = pd.read_excel(file)
        all_data = pd.concat([all_data, df], ignore_index=True)

    st.success(f"Đã gộp {len(uploaded_files)} file với tổng {len(all_data)} dòng.")

    # Lọc theo nhóm nợ
    nhom_1_2 = all_data[all_data['NHOM_NO'].isin([1, 2])]
    nhom_3_4_5 = all_data[all_data['NHOM_NO'].isin([3, 4, 5])]

    # Hiển thị
    st.subheader("Dữ liệu nhóm nợ 1 & 2")
    st.dataframe(nhom_1_2)

    st.subheader("Dữ liệu nhóm nợ 3, 4 & 5")
    st.dataframe(nhom_3_4_5)

    # Tạo file xuất
    def to_excel(nhom_1_2, nhom_3_4_5):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            nhom_1_2.to_excel(writer, index=False, sheet_name='Nhom_no_1_2')
            nhom_3_4_5.to_excel(writer, index=False, sheet_name='Nhom_no_3_4_5')
        output.seek(0)
        return output

    excel_data = to_excel(nhom_1_2, nhom_3_4_5)
    st.download_button("📥 Tải file Excel kết quả", data=excel_data, file_name="Du_no_theo_Nhom_No.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
