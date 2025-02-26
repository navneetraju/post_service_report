import os

import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder

from flag_and_update import flag_rows, remove_bars
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
    st.write("Upload CSV files for IRC, UV, and EVK. All three files must be uploaded before proceeding.")

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
        if "datasets" not in st.session_state:
            irc_df, uv_df, evk_df = (
                pd.read_csv(file_irc, header=0, encoding="cp1252"),
                pd.read_csv(file_uv, header=0, encoding="cp1252"),
                pd.read_csv(file_evk, header=0, encoding="cp1252"),
            )
            st.session_state.datasets = {
                "IRC": remove_bars(irc_df) if REMOVE_BARS else irc_df,
                "UV": remove_bars(uv_df) if REMOVE_BARS else uv_df,
                "EVK": remove_bars(evk_df) if REMOVE_BARS else evk_df,
            }

        st.success("✅ All three files uploaded successfully!")

        if "flagged_dfs" not in st.session_state:
            st.session_state.flagged_dfs = {
                label: flag_rows(df) for label, df in st.session_state.datasets.items()
            }

        # Step-by-step navigation for review
        step = st.radio(
            "Step-by-Step Review",
            options=["IRC", "UV", "EVK"],
            index=0,
            format_func=lambda x: f"Review {x} Data",
        )

        # **Force refresh of AgGrid on hall selection**
        if "selected_hall" not in st.session_state or st.session_state["selected_hall"] != step:
            st.session_state["selected_hall"] = step
            st.experimental_rerun()  # Ensure AgGrid reloads when switching halls

        # Show flagged rows for the selected dataset
        flagged_rows = st.session_state.flagged_dfs.get(step, None)
        if flagged_rows is not None and not flagged_rows.empty:
            st.subheader(f"📋 Review Flagged Data: {step}")
            st.write(f"Below are the rows flagged for review in {step}:")

            # **Fix for AgGrid slow rendering**
            gb = GridOptionsBuilder.from_dataframe(flagged_rows)
            gb.configure_default_column(editable=True, wrapText=True, resizable=True)
            gridOptions = gb.build()

            grid_response = AgGrid(
                flagged_rows,
                gridOptions=gridOptions,
                editable=True,
                height=400,
                theme="streamlit",
                key=f"{step}_grid_{st.session_state['grid_update_key']}",
            )

            # **Fix: Ensure index alignment before updating**
            updated_flagged_rows = pd.DataFrame(grid_response["data"])

            # **Fix: Ensure only flagged rows are updated in the original dataset**
            updated_df = st.session_state.datasets[step].copy()

            if not updated_flagged_rows.empty:
                updated_flagged_rows.set_index(flagged_rows.index, inplace=True)  # Align index
                updated_df.update(updated_flagged_rows)  # Apply only updated values

            # **Persist updates to session state**
            if st.button(f"Confirm {step} Updates", key=f"confirm_{step}"):
                st.session_state.datasets[step] = updated_df  # Store updated dataset
                st.session_state.flagged_dfs[step] = flag_rows(updated_df)  # Re-run flagging after update
                st.session_state["grid_update_key"] += 1  # Ensure AgGrid reloads correctly
                st.success(f"✅ {step} data successfully updated!")

                # Debugging: Print updated dataset
                # print(updated_df[['unitid', 'unitname', 'costprice']])

        else:
            st.warning(f"No rows flagged for review in {step}.")

        # **Generate Report Button**
        if st.button("📥 Generate Report"):
            # Generate the final report using the fully updated DataFrames
            buffer = generate_report(
                st.session_state.datasets["EVK"],
                st.session_state.datasets["IRC"],
                st.session_state.datasets["UV"]
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
