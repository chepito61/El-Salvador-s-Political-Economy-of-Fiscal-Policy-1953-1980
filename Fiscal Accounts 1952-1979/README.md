# Fiscal Accounts 1952–1979

Combined fiscal workbook: revenue, expenditure, interest and debt, from the Memorias de Hacienda (via the verified yearly notes).

- `El_Salvador_Fiscal_Accounts_Memorias_1952_1979.xlsx`: the revenue sheets plus Fiscal_Accounts, Exp_Ministry_Pillars, Exp_Economic, Periods_Fiscal, Analytics_Cycle and Charts_Fiscal.
- `Note_Fiscal_Accounts_Expenditure_First_Results.md/.docx`: design, first results and limits.
- Rebuild: rebuild the revenue workbook first (see Revenue Data 1952-1990/README.md); then `python3 build_fiscal_workbook.py`; recalculate (LibreOffice); `python3 build_fiscal_charts_png.py`.
- `charts/`: images used in the note.
