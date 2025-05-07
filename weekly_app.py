from io import BytesIO

import pandas as pd
import streamlit as st
import xlsxwriter

from generate_weekly_report import add_report


def main():
    st.set_page_config(layout="wide")

    st.title("Residential - Over Production - Weekly Report")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.write("Upload IRC Over Production Excel file")
        irc_file = st.file_uploader("", type=["csv"], key="irc_file")
    with col2:
        st.write("Upload EVK Over Production Excel file")
        evk_file = st.file_uploader("", type=["csv"], key="evk_file")
    with col3:
        st.write("Upload UV Over Production Excel file")
        uv_file = st.file_uploader("", type=["csv"], key="uv_file")

    # Ensure all three files are uploaded
    if irc_file and evk_file and uv_file:
        irc_df = pd.read_csv(irc_file)
        evk_df = pd.read_csv(evk_file)
        uv_df = pd.read_csv(uv_file)
        st.success("✅ All files uploaded successfully!")

        # **Generate Report Button**
        if st.button("📥 Generate Report"):
            output = BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            evk_worksheet = workbook.add_worksheet('EVK')
            add_report(evk_df, workbook, evk_worksheet)
            irc_worksheet = workbook.add_worksheet('IRC')
            add_report(irc_df, workbook, irc_worksheet)
            uv_worksheet = workbook.add_worksheet('UV')
            add_report(uv_df, workbook, uv_worksheet)

            workbook.close()
            output.seek(0)

            # Provide the report as a downloadable link
            st.download_button(
                label="📥 Download Report",
                data=output,
                file_name="Weekly_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ Please upload all three files before proceeding.")


if __name__ == '__main__':
    main()
