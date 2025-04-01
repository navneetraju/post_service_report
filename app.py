import os

import pandas as pd
import streamlit as st

from generate_report import generate_report

REMOVE_BARS = True


def resolve_path(path):
    """Resolve the absolute path for a given relative path."""
    return os.path.abspath(os.path.join(os.getcwd(), path))


# Set environment variables for Streamlit configuration
os.environ["STREAMLIT_CONFIG_DIR"] = os.path.join(os.path.dirname(__file__), ".streamlit")


def main():
    # Set Streamlit to full-width mode
    st.set_page_config(layout="wide")

    # Initialize session state variables
    if "grid_update_key" not in st.session_state:
        st.session_state["grid_update_key"] = 0

    # Title and Instructions
    st.title("📊 Residential - Over Production - Monthly Report")
    st.write("Upload Over Production Excel file")

    file = st.file_uploader("Upload Over Production File", type=["xlsx"], key="file")

    # Ensure all three files are uploaded
    if file:
        evk_df = pd.read_excel(file, sheet_name="EVK")
        irc_df = pd.read_excel(file, sheet_name="IRC")
        uv_df = pd.read_excel(file, sheet_name="UV")

        st.success("✅ File uploaded successfully!")

        # **Generate Report Button**
        if st.button("📥 Generate Report"):
            # Generate the final report using the fully updated DataFrames
            buffer = generate_report(
                evk_df,
                irc_df,
                uv_df
            )

            # Provide the report as a downloadable link
            st.download_button(
                label="📥 Download Report",
                data=buffer,
                file_name="Over_Production_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ Please upload the CSV file before proceeding.")


if __name__ == "__main__":
    main()
