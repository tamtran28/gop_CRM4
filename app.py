import streamlit as st
import pandas as pd
from io import BytesIO
import math # Import math module for ceil function

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
                # Hiển thị một vài dòng đầu để tránh lỗi MessageSizeError
                st.dataframe(nhom_1_2.head())
                if len(nhom_1_2) > 5: # Thông báo nếu có nhiều hơn 5 dòng
                    st.info(f"Hiển thị 5 hàng đầu tiên của nhóm nợ 1 & 2. Tổng số hàng: {len(nhom_1_2)}")


                st.subheader("Dữ liệu nhóm nợ 3, 4 & 5")
                # Hiển thị một vài dòng đầu để tránh lỗi MessageSizeError
                st.dataframe(nhom_3_4_5.head())
                if len(nhom_3_4_5) > 5: # Thông báo nếu có nhiều hơn 5 dòng
                    st.info(f"Hiển thị 5 hàng đầu tiên của nhóm nợ 3, 4 & 5. Tổng số hàng: {len(nhom_3_4_5)}")


                # Function to create an Excel file for a single dataframe, splitting into multiple sheets if large
                def create_excel_for_group(df_group, base_sheet_name, chunk_size=1000000):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        if len(df_group) > chunk_size:
                            num_chunks = math.ceil(len(df_group) / chunk_size)
                            for i in range(num_chunks):
                                start_row = i * chunk_size
                                end_row = min((i + 1) * chunk_size, len(df_group))
                                df_chunk = df_group.iloc[start_row:end_row]
                                df_chunk.to_excel(writer, index=False, sheet_name=f'{base_sheet_name}_Part_{i+1}')
                            st.info(f"Dữ liệu {base_sheet_name} đã được chia thành {num_chunks} sheet.")
                        else:
                            df_group.to_excel(writer, index=False, sheet_name=base_sheet_name)
                    output.seek(0)
                    return output

                # Create and provide download button for Nhom_no_1_2
                excel_data_1_2 = create_excel_for_group(nhom_1_2, 'Nhom_no_1_2')
                st.download_button(
                    "📥 Tải file Excel Nhóm nợ 1 & 2",
                    data=excel_data_1_2,
                    file_name="Ket_qua_Nhom_no_1_2.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Create and provide download button for Nhom_no_3_4_5
                excel_data_3_4_5 = create_excel_for_group(nhom_3_4_5, 'Nhom_no_3_4_5')
                st.download_button(
                    "📥 Tải file Excel Nhóm nợ 3, 4 & 5",
                    data=excel_data_3_4_5,
                    file_name="Ket_qua_Nhom_no_3_4_5.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                st.warning("Không tìm thấy cột 'NHOM_NO' trong các file đã tải lên. Vui lòng kiểm tra lại cấu trúc file.")
        else:
            st.warning("Không có dữ liệu nào được gộp thành công từ các file đã tải lên.")
