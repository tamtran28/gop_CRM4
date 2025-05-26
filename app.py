import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Gộp và Tách File CRM4 Theo Nhóm Nợ")

# Allow uploading both .xls and .xlsx files by specifying their MIME types
# .xls MIME type: application/vnd.ms-excel
# .xlsx MIME type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
uploaded_files = st.file_uploader(
    "Tải lên các file CRM4 (Excel)",
    type=['xls', 'xlsx'], # Allow both .xls and .xlsx extensions
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"Đã tải lên {len(uploaded_files)} file. Nhấn 'Xử lý dữ liệu' để tiếp tục.")

    # Add a button to trigger the processing
    if st.button("Xử lý dữ liệu"):
        all_data = pd.DataFrame()

        # Loop through each uploaded file and concatenate its data
        for file in uploaded_files:
            try:
                # Read the Excel file into a DataFrame
                df = pd.read_excel(file)
                all_data = pd.concat([all_data, df], ignore_index=True)
            except Exception as e:
                st.error(f"Lỗi khi đọc file {file.name}: {e}")
                continue # Skip to the next file if an error occurs

        if not all_data.empty:
            st.success(f"Đã gộp {len(uploaded_files)} file với tổng {len(all_data)} dòng.")

            # Check if 'NHOM_NO' column exists
            if 'NHOM_NO' in all_data.columns:
                # Filter data based on 'NHOM_NO' column
                nhom_1_2 = all_data[all_data['NHOM_NO'].isin([1, 2])]
                nhom_3_4_5 = all_data[all_data['NHOM_NO'].isin([3, 4, 5])]

                # Display filtered dataframes
                st.subheader("Dữ liệu nhóm nợ 1 & 2")
                st.dataframe(nhom_1_2)

                st.subheader("Dữ liệu nhóm nợ 3, 4 & 5")
                st.dataframe(nhom_3_4_5)

                # Function to create an Excel file in memory
                def to_excel(df_nhom_1_2, df_nhom_3_4_5):
                    output = BytesIO()
                    # Use xlsxwriter engine which produces .xlsx files
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_nhom_1_2.to_excel(writer, index=False, sheet_name='Nhom_no_1_2')
                        df_nhom_3_4_5.to_excel(writer, index=False, sheet_name='Nhom_no_3_4_5')
                    output.seek(0) # Rewind the buffer to the beginning
                    return output

                excel_data = to_excel(nhom_1_2, nhom_3_4_5)

                # Provide a download button for the generated Excel file
                # Change file_name to .xlsx to match the xlsxwriter engine output
                st.download_button(
                    "📥 Tải file Excel kết quả",
                    data=excel_data,
                    file_name="Du_no_theo_Nhom_No.xlsx", # Changed to .xlsx
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Không tìm thấy cột 'NHOM_NO' trong các file đã tải lên. Vui lòng kiểm tra lại cấu trúc file.")
        else:
            st.warning("Không có dữ liệu nào được gộp thành công từ các file đã tải lên.")
