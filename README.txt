GSTR-3B PDF to Excel Analyzer
1. Install Python 3.11+.
2. Open terminal in this folder.
3. Run: python -m pip install -r requirements.txt
4. Run: python -m streamlit run app.py
5. Upload a ZIP containing filed GSTR-3B PDFs.
6. Download GSTR3B_Analyzed.xlsx.

Workbook sheets:
- GSTR-3B Summary: month-wise consolidated summary
- GSTR-3B Detailed: every mapped head with tax components
- 6.1 Payment of Tax: detailed payment matrix
- Reconciliation: tax liability breakup
- Extraction Audit: PDF/page/section/head trace
- Read Me: methodology and coverage

The parser is mapped to the supplied filed GSTR-3B reference and should be validated against additional filed-copy layouts before production use.
