import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Gộp và Tách File CRM4 Theo Nhóm Nợ")

# Cho phép tải lên cả file .xls và .xlsx
uploaded_files = st.file_uploader(
    "Tải lên các file CRM4 (Excel)",
    type=['xls', 'xlsx'],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"Đã tải lên {len(uploaded_files)} file. Nhấn 'Xử lý dữ liệu' để tiếp tục.")

    if st.button("Xử lý dữ liệu"):
        all_data = pd.DataFrame()
        error_reading_files = False
        error_files = []

        for file in uploaded_files:
            try:
                df = pd.read_excel(file)
                all_data = pd.concat([all_data, df], ignore_index=True)
            except Exception as e:
                error_reading_files = True
                error_files.append(file.name)
                st.error(f"Lỗi khi đọc file {file.name}: {e}")

        if error_reading_files and all_data.empty:
            st.warning(f"Không thể đọc được bất kỳ file nào: {', '.join(error_files)}. Vui lòng kiểm tra lại.")
        elif error_reading_files:
            st.warning(f"Lỗi khi đọc một số file: {', '.join(error_files)}. Các file đọc thành công vẫn được xử lý.")

        if not all_data.empty:
            st.success(f"Đã gộp {len(uploaded_files) - len(error_files)} file thành công với tổng {len(all_data)} dòng.")

            if 'NHOM_NO' in all_data.columns:
                nhom_1_2 = all_data[all_data['NHOM_NO'].isin([1, 2])]
                nhom_3_4_5 = all_data[all_data['NHOM_NO'].isin([3, 4, 5])]

                st.subheader("Dữ liệu nhóm nợ 1 & 2")
                st.dataframe(nhom_1_2.head()) # Hiển thị một vài dòng đầu

                st.subheader("Dữ liệu nhóm nợ 3, 4 & 5")
                st.dataframe(nhom_3_4_5.head()) # Hiển thị một vài dòng đầu

                def to_excel(df_nhom_1_2, df_nhom_3_4_5):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_nhom_1_2.to_excel(writer, index=False, sheet_name='Nhom_no_1_2')
                        df_nhom_3_4_5.to_excel(writer, index=False, sheet_name='Nhom_no_3_4_5')
                    output.seek(0)
                    return output

                excel_data = to_excel(nhom_1_2, nhom_3_4_5)
                st.download_button(
                    "📥 Tải file Excel kết quả",
                    data=excel_data,
                    file_name="Du_no_theo_Nhom_No.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Không tìm thấy cột 'NHOM_NO' trong các file đã tải lên. Vui lòng kiểm tra lại cấu trúc file.")
        else:
            st.warning("Không có dữ liệu nào được gộp thành công từ các file đã tải lên.")
