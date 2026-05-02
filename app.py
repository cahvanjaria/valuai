"""
ValuAI - IBBI-Compliant Valuation Workbench
Main entry point. The actual functionality lives in pages/.

Author: RV Harshal Vibhakar Anjaria | IBBI/RV/03/2026/16120
Project: ICAI AI Hackathon 2026
"""

import streamlit as st

st.set_page_config(
    page_title="ValuAI - IBBI Compliant Valuation Workbench",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----- Header -----
st.title("ValuAI")
st.subheader("IBBI-Compliant Valuation Workbench")
st.caption(
    "Designed by RV Harshal Vibhakar Anjaria  |  "
    "IBBI Reg. No: IBBI/RV/03/2026/16120  |  "
    "Asset Class: Securities or Financial Assets (SFA)"
)

st.divider()

# ----- Welcome content -----
st.markdown("""
### Welcome to ValuAI

ValuAI is an **assistive workbench for IBBI Registered Valuers** producing methodologically
rigorous valuation reports compliant with IVS 105, IBBI Valuation Standards, and Rule 11UA
of the Income Tax Rules.

The tool automates the structuring, computation, sensitivity analysis, and report drafting
that surrounds the registered valuer's professional judgment — not the judgment itself.
The valuer retains full authorship, accountability, and signoff on every report.

---

### How to Use ValuAI

1. **Upload a filled-in template** on the *Upload and View* page (sidebar)
2. **Review** all parsed inputs across 7 structured sections
3. **Verify** the WACC, FCFF, and FCFE computations against the input
4. *(Coming in Week 2)* Run DCF / FCFE / NAV / CCM with sensitivity analysis
5. *(Coming in Week 4)* Generate a draft IBBI-format Word report

---

### Methodological Capabilities

ValuAI supports the following methodologies, all with full audit trails:

- **Income Approach** — Discounted Cash Flow (DCF) using FCFF and FCFE in parallel, with
  reconciliation tolerance check
- **Market Approach** — Comparable Companies Method (CCM) with median/mean multiples
  and DLOM adjustments
- **Cost Approach** — Net Asset Value (NAV) cross-check
- **WACC build-up** — CAPM (Rf + β·ERP + CSRP) with Direct Ke override
- **Discounting conventions** — Mid-Year and Year-End
- **Stub period** support for off-cycle valuation dates
- **Pre-money / Post-money** reconciliation for round-based valuations
- **Discount for Lack of Marketability (DLOM)** and **Lack of Control (DLOC)**
- **Method weightages** for blending DCF + CCM + NAV outputs

---

### Project Status

Under Development | ICAI AI Hackathon 2026 Submission
""")

st.info(
    "**Get started:** Click *'Upload and View'* in the left sidebar to upload a "
    "filled-in template and see ValuAI parse your valuation inputs."
)

# ----- Footer -----
st.divider()
st.caption(
    "ValuAI v0.4 (Day 4) — *This tool does not value a company; a Registered Valuer does. "
    "ValuAI automates the surrounding work. Professional judgment remains the product.*"
)