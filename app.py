import streamlit as st
import pandas as pd

st.title("Excel Subsidiary Comparison Tool")

display_file = st.file_uploader("Upload Display Sheet", type=["xlsx"])
parse_file = st.file_uploader("Upload Parse Sheet", type=["xlsx"])

if display_file and parse_file:
    display_df = pd.read_excel(display_file)
    parse_df = pd.read_excel(parse_file)

    merged = display_df.merge(
        parse_df,
        on="Subsidiary Name",
        how="outer",
        suffixes=("_display", "_parse"),
        indicator=True
    )

    st.write("Merged Data Preview:")
    st.dataframe(merged)