import os
import io
import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder
from flag_and_update import flag_rows
from generate_report import generate_report


def resolve_path(path):
    """Resolve the absolute path for a given relative path."""
    return os.path.abspath(os.path.join(os.getcwd(), path))


# Set environment variables for Streamlit configuration
os.environ["STREAMLIT_CONFIG_DIR"] = os.path.join(os.path.dirname(__file__), ".streamlit")


def main():
    # Set Streamlit to full-width mode
    st.set_page_config(layout="wide")

    # Title and Instructions
    st.title("📊 Residential - Over Production - Monthly Report")
    st.write("Upload CSV files for IRC, UV, and EVK. All three files must be uploaded before proceeding.")

    # Custom CSS for file upload boxes
    st.markdown("""
        <style>
            .big-box {
                border: 2px dashed #ccc;
                padding: 30px;
                text-align: center;
                font-size: 20px;
                font-weight: bold;
                color: #333;
                margin-bottom: 20px;
                border-radius: 10px;
                background-color: #f9f9f9;
            }
            .st-emotion-cache-16txtl3 {
                max-width: 100% !important;
            }
        </style>
    """, unsafe_allow_html=True)

    # Upload Boxes for IRC, UV, EVK
    col1, col2, col3 = st.columns(3)

    with col1:
        file_irc = st.file_uploader("Upload IRC", type=["csv"], key="irc")

    with col2:
        file_uv = st.file_uploader("Upload UV", type=["csv"], key="uv")

    with col3:
        file_evk = st.file_uploader("Upload EVK", type=["csv"], key="evk")

    # Ensure all three files are uploaded
    if file_irc and file_uv and file_evk:
        # Load the files into Pandas DataFrames
        df_irc = pd.read_csv(file_irc, header=0, encoding="cp1252")
        df_uv = pd.read_csv(file_uv, header=0, encoding="cp1252")
        df_evk = pd.read_csv(file_evk, header=0, encoding="cp1252")

        st.success("✅ All three files uploaded successfully!")

        # **Flag Rows for Review**
        datasets = {
            "IRC": df_irc,
            "UV": df_uv,
            "EVK": df_evk,
        }
        flagged_dfs = {label: flag_rows(df) for label, df in datasets.items()}

        # Step-by-step navigation for review
        step = st.radio(
            "Step-by-Step Review",
            options=["IRC", "UV", "EVK"],
            index=0,
            format_func=lambda x: f"Review {x} Data",
        )

        # Show flagged rows for the selected dataset
        flagged_rows = flagged_dfs.get(step, None)
        if flagged_rows is not None and not flagged_rows.empty:
            st.subheader(f"📋 Review Flagged Data: {step}")
            st.write(f"Below are the rows flagged for review in {step}:")

            # AgGrid for interactive review and update
            gb = GridOptionsBuilder.from_dataframe(flagged_rows)
            gb.configure_default_column(editable=True, wrapText=True, resizable=True)
            gridOptions = gb.build()

            grid_response = AgGrid(
                flagged_rows,
                gridOptions=gridOptions,
                editable=True,
                height=400,
                theme="streamlit",
                key=f"{step}_grid",
            )

            # Get updated data from AgGrid
            updated_flagged_rows = pd.DataFrame(grid_response["data"])

            # Merge updated rows back into the original DataFrame
            original_df = datasets[step].copy()
            original_df.update(updated_flagged_rows, overwrite=True)  # Update flagged rows
            datasets[step] = original_df  # Update the dataset

            if st.button(f"Confirm {step} Updates", key=f"confirm_{step}"):
                st.success(f"✅ {step} data successfully updated!")
        else:
            st.warning(f"No rows flagged for review in {step}.")

        # Generate Report Button
        if st.button("📥 Generate Report"):
            # Generate the final report using the fully updated DataFrames
            buffer = generate_report(
                datasets["EVK"],
                datasets["IRC"],
                datasets["UV"]
            )

            # Provide the report as a downloadable link
            st.download_button(
                label="📥 Download Report",
                data=buffer,
                file_name="Over_Production_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ Please upload all three CSV files before proceeding.")


if __name__ == "__main__":
    main()
