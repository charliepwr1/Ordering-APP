import streamlit as st
from etl.generate_order import generate_order
import os

st.title("Cannabis Order Generator")

if st.button("Run ETL & Generate Excel"):
    output_file = os.path.join("..", "output", "weekly_order.xlsx")
    with st.spinner("Running ETL & building workbook…"):
        generate_order(output_file)
    st.success("Done! Download below.")
    with open(output_file, "rb") as f:
        st.download_button(
            label="📥 Download Order Excel",
            data=f,
            file_name="weekly_order.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
