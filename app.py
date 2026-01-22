import streamlit as st
import pandas as pd

st.set_page_config(page_title="Excel Sheet Value Counter", layout="centered")

st.title("📊 Excel Sheet Column Value Counter By Sm Baqir")

# Upload Excel file
uploaded_file = st.file_uploader(
    "Upload Excel Workbook",
    type=["xlsx", "xls"]
)

if uploaded_file:
    # Read all sheets
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_names = excel_file.sheet_names

    st.success(f"Workbook loaded with {len(sheet_names)} sheets")

    # Read first sheet to get column names
    sample_df = pd.read_excel(uploaded_file, sheet_name=sheet_names[0])

    selected_column = st.selectbox(
        "Select column to count values from",
        sample_df.columns
    )

    # Collect unique values across all sheets
    unique_values = set()
    for sheet in sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)
        if selected_column in df.columns:
            unique_values.update(df[selected_column].dropna().unique())

    selected_values = st.multiselect(
        "Select values to count",
        sorted(unique_values)
    )

    if selected_values:
        summary_data = []

        for sheet in sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet)

            row = {"Sheet Name": sheet}

            for val in selected_values:
                if selected_column in df.columns:
                    row[val] = (df[selected_column] == val).sum()
                else:
                    row[val] = 0

            summary_data.append(row)

        result_df = pd.DataFrame(summary_data)

        st.subheader("📑 Final Summary Report")
        st.dataframe(result_df, use_container_width=True)

        # Download button
        output_file = "sheet_value_summary.xlsx"
        result_df.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                label="⬇️ Download Report",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
