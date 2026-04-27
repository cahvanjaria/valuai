import streamlit as st

st.set_page_config(
    page_title="ValuAI - IBBI Compliant Valuation Workbench",
    page_icon="📊",
    layout="wide"
)

st.title("ValuAI")
st.subheader("IBBI-Compliant Valuation Workbench")
st.caption("Designed by RV Harshal Vibhakar Anjaria | IBBI Reg. No: IBBI/RV/03/2026/16120 | Asset Class: Securities or Financial Assets (SFA)")

st.divider()

st.markdown("""
### Welcome to ValuAI

This workbench assists IBBI Registered Valuers (Securities or Financial Assets) in producing 
methodologically rigorous valuation reports compliant with IVS 105, IBBI Valuation Standards, 
and Rule 11UA of the Income Tax Rules.

**Status:** Under Development | Week 1 of 5
""")

st.info("This is the foundation app. Functionality will be added across the next 5 weeks.")