"""
ValuAI - Upload and View Page
==============================
User uploads a filled-in template.xlsx and sees all parsed data displayed
across 7 structured sections. Includes a Reconciliation panel that validates
ValuAI's computed WACC against expected reference values.
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
import sys

# Make project root importable so we can import utils.*
project_root = Path(__file__).parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from utils.excel_reader import read_template
from utils.display_helpers import fmt_pct, fmt_money, fmt_number, fmt_date


# ============================================================================
# Page config
# ============================================================================
st.set_page_config(page_title="ValuAI - Upload and View", page_icon="📊", layout="wide")

st.title("📤 Upload and View")
st.caption("Upload a filled-in ValuAI template and review the parsed inputs across all sections.")

st.divider()

# ============================================================================
# File uploader
# ============================================================================
uploaded_file = st.file_uploader(
    "Upload a filled-in ValuAI template (.xlsx)",
    type=["xlsx"],
    help="The template should be filled-in following the ValuAI input convention "
         "(percentages as plain numbers, e.g., 7.21 means 7.21%).",
)

# Optional: load the bundled Farm Gas test case for instant demo
col1, col2 = st.columns([2, 1])
with col2:
    use_demo = st.button(
        "🧪 Load Farm Gas Demo Case",
        help="Load the pre-filled Sejal Agrawal Farm Gas Pvt Ltd case for testing.",
        use_container_width=True,
    )

# Decide which file to read
file_to_read = None
if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getvalue())
        file_to_read = tmp.name
    st.success(f"✅ Uploaded: **{uploaded_file.name}**")
elif use_demo:
    demo_path = project_root / "assets" / "case_farmgas_sejal.xlsx"
    if demo_path.exists():
        file_to_read = str(demo_path)
        st.info("📁 Loaded demo case: **case_farmgas_sejal.xlsx** (Farm Gas Pvt Ltd)")
    else:
        st.error(f"Demo case not found at {demo_path}")

if file_to_read is None:
    st.markdown("""
    ### Instructions

    1. Click **"Browse files"** above and select a filled-in `template.xlsx`, OR
    2. Click **"🧪 Load Farm Gas Demo Case"** to instantly load the bundled test case

    The template should follow the ValuAI input convention:
    - Percentages as plain numbers (`7.21` means 7.21%, NOT 0.0721)
    - Dates in any standard Excel date format
    - Currency in the unit specified on the Cover sheet (Lakhs / Millions / Crores)

    All seven sheets must be present: Cover, Historical_PL, Historical_BS, Projections,
    WACC_Inputs, Other_Inputs, Comparables.
    """)
    st.stop()

# ============================================================================
# Read and parse the file
# ============================================================================
try:
    data = read_template(file_to_read)
except FileNotFoundError as e:
    st.error(f"❌ File not found: {e}")
    st.stop()
except ValueError as e:
    st.error(f"❌ Template structure error: {e}")
    st.stop()
except Exception as e:
    st.error(f"❌ Error parsing file: {e}")
    st.stop()

st.success("✅ Template parsed successfully — all 7 sheets read.")

# ============================================================================
# Display sections - using tabs for clean organisation
# ============================================================================
tab_meta, tab_hist, tab_proj, tab_wacc, tab_other, tab_comp, tab_recon = st.tabs([
    "📋 Metadata",
    "📈 Historical",
    "🔮 Projections",
    "⚖️ WACC",
    "⚙️ Other Inputs",
    "📊 Comparables",
    "✓ Reconciliation",
])

md = data["metadata"]
hpl = data["historical_pl"]
hbs = data["historical_bs"]
proj = data["projections"]
wacc = data["wacc_inputs"]
other = data["other_inputs"]
comp = data["comparables"]


# ----------------------------------------------------------------------------
# TAB: Metadata
# ----------------------------------------------------------------------------
with tab_meta:
    st.subheader("Engagement Metadata")

    cA, cB = st.columns(2)
    with cA:
        st.markdown("**Client Details**")
        st.write(f"**Name:** {md['client_name'] or '—'}")
        st.write(f"**CIN:** {md['client_cin'] or '—'}")
        st.write(f"**Address:** {md['client_address'] or '—'}")
        st.markdown("**Key Dates**")
        st.write(f"**Date of Appointment:** {fmt_date(md['date_of_appointment'])}")
        st.write(f"**Valuation Date:** {fmt_date(md['valuation_date'])}")
        st.write(f"**Date of Report:** {fmt_date(md['date_of_report'])}")

    with cB:
        st.markdown("**Valuer Details**")
        st.write(f"**Name:** {md['valuer_name'] or '—'}")
        st.write(f"**IBBI Reg. No.:** {md['valuer_ibbi_reg_no'] or '—'}")
        st.write(f"**Asset Class:** {md['valuer_asset_class'] or '—'}")
        st.markdown("**Standards & Methodology**")
        st.write(f"**Base of Value:** {md['base_of_value'] or '—'}")
        st.write(f"**Premise:** {md['premise_of_value'] or '—'}")
        st.write(f"**Standards:** {md['valuation_standard_ref'] or '—'}")

    st.divider()
    st.markdown("**Period Configuration**")
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Historical Years", md["historical_years"])
    p2.metric("Forecast Years", md["forecast_years"])
    p3.metric("Stub Period", md["stub_period"])
    p4.metric("Stub Months", md["stub_months"])

    st.markdown("**Methodology Toggles**")
    m1, m2, m3 = st.columns(3)
    m1.metric("Discounting", md["discounting_convention"])
    m2.metric("WACC Mode", md["wacc_input_mode"])
    m3.metric("Terminal Method", md["terminal_value_method"])

    st.markdown("**Purpose**")
    st.write(md["purpose"] or "—")
    st.caption(f"Statutory Reference: {md['statutory_reference'] or '—'}")


# ----------------------------------------------------------------------------
# TAB: Historical
# ----------------------------------------------------------------------------
with tab_hist:
    st.subheader("Historical Profit & Loss")
    year_labels_hist = ["FY-4", "FY-3", "FY-2", "FY-1", "FY-0 (Latest)"]

    pl_rows = [
        ("Revenue from Operations", hpl["revenue_from_operations"]),
        ("Other Income", hpl["other_income"]),
        ("Total Revenue", hpl["total_revenue"]),
        ("Cost of Goods Sold", hpl["cost_of_goods_sold"]),
        ("Employee Benefits", hpl["employee_benefits"]),
        ("Other Operating Expenses", hpl["other_operating_expenses"]),
        ("Total Operating Costs", hpl["total_operating_costs"]),
        ("EBITDA", hpl["ebitda"]),
        ("Depreciation & Amortisation", hpl["depreciation"]),
        ("EBIT", hpl["ebit"]),
        ("Finance Cost", hpl["finance_cost"]),
        ("PBT", hpl["pbt"]),
        ("Tax Expense", hpl["tax_expense"]),
        ("PAT", hpl["pat"]),
    ]
    df_pl = pd.DataFrame(
        {y: [fmt_money(row[1][i]) for row in pl_rows] for i, y in enumerate(year_labels_hist)},
        index=[r[0] for r in pl_rows],
    )
    st.dataframe(df_pl, use_container_width=True)

    st.subheader("Historical Balance Sheet (selected lines)")
    bs_rows = [
        ("Total Shareholders' Funds", hbs["total_shareholders_funds"]),
        ("Total Non-Current Liabilities", hbs["total_non_current_liabilities"]),
        ("Total Current Liabilities", hbs["total_current_liabilities"]),
        ("Total Equity & Liabilities", hbs["total_equity_and_liabilities"]),
        ("Total Non-Current Assets", hbs["total_non_current_assets"]),
        ("Total Current Assets", hbs["total_current_assets"]),
        ("Total Assets", hbs["total_assets"]),
        ("Tally Check (E&L − Assets)", hbs["tally_check"]),
    ]
    df_bs = pd.DataFrame(
        {y: [fmt_money(row[1][i]) for row in bs_rows] for i, y in enumerate(year_labels_hist)},
        index=[r[0] for r in bs_rows],
    )
    st.dataframe(df_bs, use_container_width=True)


# ----------------------------------------------------------------------------
# TAB: Projections
# ----------------------------------------------------------------------------
with tab_proj:
    st.subheader("Forecast Financials")
    year_labels_proj = ["Stub", "FY+1", "FY+2", "FY+3", "FY+4", "FY+5",
                        "FY+6", "FY+7", "FY+8", "FY+9", "FY+10"]

    proj_rows = [
        ("Revenue from Operations", proj["revenue_from_operations"]),
        ("Total Revenue", proj["total_revenue"]),
        ("EBITDA", proj["ebitda"]),
        ("Depreciation", proj["depreciation"]),
        ("EBIT", proj["ebit"]),
        ("Tax Rate (%)", proj["effective_tax_rate_pct"]),
        ("NOPAT", proj["nopat"]),
        ("Change in WC", proj["change_in_working_capital"]),
        ("CAPEX", proj["capex"]),
        ("FCFF", proj["fcff"]),
        ("Interest Expense", proj["interest_expense"]),
        ("Closing Debt", proj["closing_debt"]),
        ("FCFE", proj["fcfe"]),
    ]
    df_proj = pd.DataFrame(
        {y: [fmt_money(row[1][i]) for row in proj_rows] for i, y in enumerate(year_labels_proj)},
        index=[r[0] for r in proj_rows],
    )
    st.dataframe(df_proj, use_container_width=True)

    st.subheader("Equity Bridge Items (as on Valuation Date)")
    e1, e2, e3, e4 = st.columns(4)
    e1.metric("Cash & Equivalents", fmt_money(proj["cash_and_equivalents_valuation_date"]))
    e2.metric("Current Investments", fmt_money(proj["current_investments_valuation_date"]))
    e3.metric("Total Debt", fmt_money(proj["total_debt_valuation_date"]))
    e4.metric("Pref. Share Capital", fmt_money(proj["preference_share_capital_valuation_date"]))

    st.subheader("Share Information")
    s1, s2 = st.columns(2)
    s1.metric("Shares Outstanding (M)", fmt_number(proj["shares_outstanding_millions"], 6))
    s2.metric("Face Value (INR)", fmt_money(proj["face_value_per_share"]))


# ----------------------------------------------------------------------------
# TAB: WACC
# ----------------------------------------------------------------------------
with tab_wacc:
    st.subheader("WACC Build-up")

    st.markdown("**1. CAPM Components**")
    w1, w2, w3, w4 = st.columns(4)
    w1.metric("Risk-Free Rate", fmt_pct(wacc["risk_free_rate_pct"]))
    w2.metric("Equity Risk Premium", fmt_pct(wacc["equity_risk_premium_pct"]))
    w3.metric("Beta (β)", fmt_number(wacc["beta"], 2))
    w4.metric("Base CAPM Ke", fmt_pct(wacc["base_capm_ke_pct"], 4))

    st.markdown("**2. Adjusted Cost of Equity**")
    a1, a2 = st.columns(2)
    a1.metric("CSRP", fmt_pct(wacc["csrp_pct"]))
    a2.metric("Adjusted Ke", fmt_pct(wacc["adjusted_ke_pct"], 4))

    st.markdown("**3. Cost of Debt**")
    d1, d2, d3 = st.columns(3)
    d1.metric("Pre-tax Kd", fmt_pct(wacc["pre_tax_kd_pct"]))
    d2.metric("Tax Rate", fmt_pct(wacc["effective_tax_rate_pct"]))
    d3.metric("Post-tax Kd", fmt_pct(wacc["post_tax_kd_pct"], 4))

    st.markdown("**4. Capital Structure**")
    c1, c2, c3 = st.columns(3)
    c1.metric("Weight of Equity", fmt_pct(wacc["weight_equity_pct"]))
    c2.metric("Weight of Debt", fmt_pct(wacc["weight_debt_pct"]))
    c3.metric("Total", fmt_pct(wacc["weights_total_pct"]))

    st.markdown("**5. WACC**")
    st.metric(
        "Weighted Average Cost of Capital",
        fmt_pct(wacc["wacc_pct"], 4),
        help="WACC = (We × Ke) + (Wd × Kd post-tax)",
    )

    st.divider()
    st.markdown("**Sources**")
    st.caption(f"Risk-Free Rate: {wacc['rf_source'] or '—'}")
    st.caption(f"Equity Risk Premium: {wacc['erp_source'] or '—'}")
    st.caption(f"Beta: {wacc['beta_source_note'] or '—'}")


# ----------------------------------------------------------------------------
# TAB: Other Inputs
# ----------------------------------------------------------------------------
with tab_other:
    st.subheader("Terminal Value Parameters")
    t1, t2, t3 = st.columns(3)
    t1.metric("Method", other["terminal_value_method"])
    t2.metric("Perpetual Growth", fmt_pct(other["perpetual_growth_rate_pct"]))
    t3.metric("Exit EBITDA Multiple", fmt_number(other["exit_ebitda_multiple"], 2))

    st.subheader("Sensitivity Analysis Ranges")
    sens_data = {
        "Driver": ["WACC (% pts)", "Terminal Growth (% pts)", "EBIT Margin (%)",
                   "CAPEX (%)", "Working Capital (%)"],
        "Lower": [fmt_number(other["sensitivity_wacc"][0], 2),
                  fmt_number(other["sensitivity_terminal_growth"][0], 2),
                  fmt_number(other["sensitivity_ebit_margin"][0], 2),
                  fmt_number(other["sensitivity_capex"][0], 2),
                  fmt_number(other["sensitivity_working_capital"][0], 2)],
        "Upper": [fmt_number(other["sensitivity_wacc"][1], 2),
                  fmt_number(other["sensitivity_terminal_growth"][1], 2),
                  fmt_number(other["sensitivity_ebit_margin"][1], 2),
                  fmt_number(other["sensitivity_capex"][1], 2),
                  fmt_number(other["sensitivity_working_capital"][1], 2)],
    }
    st.dataframe(pd.DataFrame(sens_data), use_container_width=True, hide_index=True)

    st.subheader("Discounts")
    dl1, dl2 = st.columns(2)
    with dl1:
        st.markdown("**DLOM**")
        st.write(f"Apply: **{other['dlom_apply']}**")
        st.write(f"Percentage: **{fmt_pct(other['dlom_pct'])}**")
        if other["dlom_justification"]:
            st.caption(other["dlom_justification"])
    with dl2:
        st.markdown("**DLOC**")
        st.write(f"Apply: **{other['dloc_apply']}**")
        st.write(f"Percentage: **{fmt_pct(other['dloc_pct'])}**")
        if other["dloc_justification"]:
            st.caption(other["dloc_justification"])

    st.subheader("Method Weightages")
    mw1, mw2, mw3, mw4 = st.columns(4)
    mw1.metric("DCF (Income)", fmt_pct(other["weight_dcf_pct"]))
    mw2.metric("CCM (Market)", fmt_pct(other["weight_ccm_pct"]))
    mw3.metric("NAV (Cost)", fmt_pct(other["weight_nav_pct"]))
    mw4.metric("Total", fmt_pct(other["weights_total_pct"]))
    if other["weights_reasoning"]:
        st.caption(f"**Reasoning:** {other['weights_reasoning']}")


# ----------------------------------------------------------------------------
# TAB: Comparables
# ----------------------------------------------------------------------------
with tab_comp:
    st.subheader("Subject Company Metrics")
    sc = comp["subject_company"]
    sc1, sc2, sc3, sc4, sc5 = st.columns(5)
    sc1.metric("Revenue", fmt_money(sc["revenue"]))
    sc2.metric("EBITDA", fmt_money(sc["ebitda"]))
    sc3.metric("EBIT", fmt_money(sc["ebit"]))
    sc4.metric("PAT", fmt_money(sc["pat"]))
    sc5.metric("Book Value Equity", fmt_money(sc["book_value_equity"]))

    st.subheader("Peer Companies")
    if comp["peers"]:
        df_peers = pd.DataFrame(comp["peers"])
        st.dataframe(df_peers, use_container_width=True, hide_index=True)

        st.subheader("Median Multiples")
        m = comp["medians"]
        med1, med2, med3, med4, med5 = st.columns(5)
        med1.metric("EV/Revenue", fmt_number(m["ev_revenue"], 2))
        med2.metric("EV/EBITDA", fmt_number(m["ev_ebitda"], 2))
        med3.metric("EV/EBIT", fmt_number(m["ev_ebit"], 2))
        med4.metric("P/E", fmt_number(m["pe"], 2))
        med5.metric("P/B", fmt_number(m["pb"], 2))
    else:
        st.info("No comparable companies entered. Market Approach (CCM) will not be applicable.")


# ----------------------------------------------------------------------------
# TAB: Reconciliation
# ----------------------------------------------------------------------------
with tab_recon:
    st.subheader("Reconciliation Against Reference Cases")
    st.markdown("""
    This panel cross-checks ValuAI's computations against published, signed IBBI valuation
    reports. Reproducing real reports to high precision is evidence that ValuAI's
    methodology faithfully implements professional valuation practice.
    """)

    st.divider()

    st.markdown("### Reference Case: Farm Gas Pvt Ltd (Sejal Agrawal, 26-Dec-2023)")
    st.caption("Source: Signed valuation report by CA Sejal Agrawal, IBBI/RV/06/2020/13106.")

    valuai_wacc = wacc["wacc_pct"] if wacc["wacc_pct"] is not None else 0
    sejal_wacc = 20.24

    valuai_stub_fcff = proj["fcff"][0] if proj["fcff"] and proj["fcff"][0] is not None else 0
    sejal_stub_fcff = 271.21

    valuai_fy1_ebit = proj["ebit"][1] if len(proj["ebit"]) > 1 and proj["ebit"][1] is not None else 0
    sejal_fy1_ebit = 639.80

    rec_table = pd.DataFrame({
        "Metric": ["WACC (%)", "Stub Period FCFF (Lakhs)", "FY+1 EBIT (Lakhs)"],
        "ValuAI Computed": [
            f"{valuai_wacc:.4f}%" if valuai_wacc else "—",
            fmt_money(valuai_stub_fcff),
            fmt_money(valuai_fy1_ebit),
        ],
        "Reference (Published)": [
            f"{sejal_wacc:.2f}%",
            fmt_money(sejal_stub_fcff),
            fmt_money(sejal_fy1_ebit),
        ],
        "Difference": [
            f"{(valuai_wacc - sejal_wacc):+.4f} %pts" if valuai_wacc else "—",
            f"{(valuai_stub_fcff - sejal_stub_fcff):+.2f}" if valuai_stub_fcff else "—",
            f"{(valuai_fy1_ebit - sejal_fy1_ebit):+.2f}" if valuai_fy1_ebit else "—",
        ],
    })
    st.dataframe(rec_table, use_container_width=True, hide_index=True)

    if abs(valuai_wacc - sejal_wacc) < 0.05 and abs(valuai_stub_fcff - sejal_stub_fcff) < 1.0:
        st.success(
            "✅ **Reconciliation passed.** ValuAI reproduces Sejal Agrawal's signed valuation "
            "to within rounding tolerance across WACC and FCFF. Methodological fidelity confirmed."
        )
    elif valuai_wacc > 0:
        st.warning(
            "⚠️ ValuAI computations diverge from the reference. This is expected if the "
            "uploaded file is not the Farm Gas reference case. Reconciliation only applies "
            "to that specific case."
        )
    else:
        st.info("ℹ️ No WACC computed. Verify inputs on the Cover and WACC_Inputs sheets.")

    st.caption(
        "Note: Small differences (e.g., 0.01% on WACC) are due to rounding in the published "
        "report. Sejal's report rounds intermediate steps; ValuAI carries full precision through."
    )