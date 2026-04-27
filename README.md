# ValuAI

**IBBI-Compliant Valuation Workbench**

An assistive workbench for IBBI Registered Valuers (Securities or Financial Assets) producing methodologically rigorous valuation reports compliant with IVS 105, IBBI Valuation Standards, and Rule 11UA of the Income Tax Rules.

---

## Author

**RV Harshal Vibhakar Anjaria**
IBBI Registered Valuer (Asset Class: Securities or Financial Assets)
Registration No: IBBI/RV/03/2026/16120

## Project Status

Under Development | Submission for ICAI AI Hackathon 2026

## Scope

- **Phase 1 (Demo):** DCF + CCM + NAV with sensitivity analysis and IBBI-format report generation, focused on Section 56(2)(x) / Rule 11UA fair value of unlisted equity shares.
- **Phase 2 (Roadmap):** Section 50CA, FEMA pricing, ESOP under Rule 3(8)(iii), IBC liquidation/fair value, comparable transaction database integration.

## Important Disclaimer

ValuAI is an *assistive* workbench. It does not value a company; a Registered Valuer does. The tool automates structuring, computation, sensitivity analysis, and report drafting. Professional judgment of the Registered Valuer remains the product. The valuer retains authorship, accountability, and signoff.

## Technology Stack

- Python 3.14
- Streamlit (web interface)
- Anthropic Claude API (narrative reasoning)
- Plotly (visualisations)
- python-docx (report generation)
- pandas, numpy (computations)