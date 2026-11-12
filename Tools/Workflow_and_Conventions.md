# Workflow and Conventions — Fiscal Notes Series, 1952–1980

**Handoff note for a new working thread (11 November 2026).** Read this first, then continue with the next year (1976).

---

## 1. The project in one paragraph

The research question is whether the military governments' fiscal and monetary management of the coffee cycle (1952–1980) deepened the political polarization that led to the civil war. The author is an economist (formerly a fiscal specialist at the IDB), and the notes are written for policy makers and economists. The evidence base is the annual Memoria de Hacienda (*Informe Complementario Constitucional*), read year by year. Secondary sources give context.

## 2. How the author works

- **Iterative.** The author gives the next step in a short message (for example, "1976"); that is the go-ahead. Do not jump ahead or start the next year unprompted.
- **"Thinking out loud" messages are not instructions.** Discuss; do not edit the notes unless asked.
- **English first;** Spanish and Portuguese translations come at the end of the series.
- **Deliverables are Word files.** Each note is also kept as Markdown in the repo and in the claude.ai project.

## 3. What each year needs

Two notes per year:

1. **Note_MH_YYYY.md**, the working note. Sections:
   - Political Context (dated event table with sources), then a **"Political economy reading"**;
   - Introduction: what the volume contains;
   - 1. Revenue (Estado I-4 itemized: income-tax split, coffee duty and other export duties, tax burden and elasticity from the Memoria);
   - 2. Expenditure: execution; economic classification; ministries on credits used; security vs development; **unit lines** (teachers' Enseñanza Básica, University subsidy and scholarships, Fuerza Armada, Presidency and Defensa *Administración Superior* lump sums in non-personal services, IRA);
   - 3. Commitments and arrears;
   - 4. Budget balance and financial position (with the General Fund balance sheet);
   - 5. Public debt;
   - 6. Loan-financed investment (loan register);
   - 7. **The Cost of the War** (required in every note since 1969);
   - 8. Comparative perspective (multi-year table);
   - 9. Preliminary observations;
   - 10. Reading the year against the hypotheses (table + "The story in one paragraph").
2. **Technical_Note_Fiscal_Diagnostic_YYYY.md**, the IMF-style diagnostic: Executive Summary; Fiscal Accounts Summary; revenue and expenditure performance; revenue concentration; progressivity; rigidity; debt sustainability with thresholds; pillars; Cost of the War; vulnerability matrix; dashboard; assessment; data gaps.

Use the most recent notes (1974 and 1975 in "Policy and Technical Notes/") as the models.

### Hypotheses and tracks (from the Political Economy Context note)

- **H1** Education and the University squeezed. **H2** Wages contained, especially teachers'. **H3** Investment cut first. **H4** Monetary policy (Central Bank financing). **H5** IRA food subsidies cut.
- **Track A** the tax bargain. **Track B** the Honduras safety valve. **Track C** the protection bargain. **Track D** reforms announced and then diluted.

### After all years

A **synthesis note on the cost of the war** (1969 war with Honduras), built from War and Security Data/war_cost_1968_1980.csv. Later: Diario Oficial list; Spanish translations; retrofit of the 1964–67 notes.

## 4. Accounting conventions

- **Memoria concepts:** revenue = *rentas percibidas y devengadas*; **gastos**; **compromisos** (commitments pending); **créditos utilizados** = gastos + compromisos.
- **Current revenue excludes** bond proceeds and loan money booked as "otros ingresos para el financiamiento de programas de funcionamiento". These are financing, not revenue.
- **Commitment basis** = current revenue − credits used (amortization included).
- **Underlying result** = commitment basis, excluding (a) loan-financed capital commitments, (b) on-lending to public enterprises (CEL), (c) non-cash ANTEL payments. Report it also excluding amortization, and on a gastos basis.
- **Superávit Financiero** is shown, but check what it rests on: since 1974 it includes **authorized but unsold investment-plan bonds booked as accrued revenue receivable** (₡83M at end-1974; ₡144M at end-1975).
- **Ministry comparisons** use credits used (Cuadro 8). Cuadro 15 (economic classification) is on a **gastos** basis.
- **Pillars** (author's decision, November 2026): **Development** = Educación + Salud Pública + Obras Públicas. **Security** = Defensa + Seguridad Pública (the Defense-controlled security apparatus). **Agricultura y Ganadería** (agro-export support) and **Interior** are memo lines.
- **Unit lines come one year late.** From the 1974 volume on, Estado IV-2 in the Memoria for year Y reports the budget of year **Y−1**, liquidated at the end of Y. (MH 1975 gave the 1974 unit lines; MH 1976 should give 1975.) Earlier volumes (e.g., MH 1973) report the same year. Always check the year in the IV-2 header and the debt line.
- **Coffee prices:** World Bank Pink Sheet, "Graficas" sheet of *Commodities y recaudacion.xlsx* (US$/kg): 1971 0.992; 1972 1.110; 1973 1.373; 1974 1.452; 1975 1.442; 1976 3.147; 1977 5.169; 1978 3.590; 1979 3.826; 1980 3.466. The other columns are **sugar** and **cotton**, not oil.
- Currency: US$1.00 = ₡2.50.
- **Terminology (author's decision, 11 November 2026):** write **"corporate income tax"**, not "company income tax" or "company taxes" (Memoria: *impuesto sobre la renta de empresas / sociedades*). Personal income tax stays "personal income tax". Notes issued before this date use "company"; correct them when they are next revised.
- **Revenue classification across periods (author's decision, 11 November 2026):** the 1950s Memorias classify revenue by *Títulos* (patrimonial rights, official enterprises, taxes, etc.). For long series, map them into the modern categories (income, property, imports, exports, consumption, stamps and legal acts, non-tax) and explain the mapping in a footnote or methods note.

## 5. Verification rules

- **Every figure is read on the page image** and checked by arithmetic (sums of components must match printed totals exactly; use Python).
- **Remove a claim rather than carry it forward unverified** ("largest of the series", "first time", "since 19xx") unless checked against the earlier notes.
- Political facts: web search before stating them; cite the source in the event table.
- When a correction to an earlier note is found, fix the earlier note and its data file, and add a dated correction line.

## 6. Reading the scans (tools in this folder)

The Memorias from the 1970s are typewritten, poorly reproduced and turned sideways.

1. Rasterize: `pdftoppm -r 200 -gray -png "Memorias de Hacienda/MH 1976.pdf" p`.
2. Narrative pages (two-page spreads): `Tools/ocr_spreads.sh NN` (chooses 90° or 270° by Spanish stop-words, splits into halves). Then compile `h/NNL.txt`, `h/NNR.txt` and read.
3. Statement pages (one wide table): `Tools/ocr_tables.sh NN` (chooses orientation by number patterns). Grep `t/fNN.txt` for "ESTADO", "RAMO", unit names.
4. Small tables in the narrative (Cuadros 8, 11–15): re-render the page at 400 dpi (`pdftoppm -r 400 -f N -l N`), rotate by the saved angle, and crop/read the image directly.
5. Specific rows: `Tools/crop_row.sh NN 'regex' out.png 60 60` and read the image. Trust the tsv y-coordinate, not the image's displayed scale.
6. Run long OCR jobs with `nohup ... &` and poll for a done-file; use 2 parallel jobs.

## 7. Build and delivery

- Build: `pandoc X.md -o X.docx --reference-doc=Tools/ref_clean.docx`.
- Render check: `soffice --headless --convert-to pdf X.docx`, then `pdftoppm -r 60` one page and look at it.
- Copy .md and .docx to "Policy and Technical Notes/".
- **Data files: append only** (Python `open(f, 'a')`); never rewrite them through a text-mode read, because ministry_class has mixed CRLF/LF endings. Files:
  - Spending Data/ministry_class_1961_1979.csv (credits used by ramo; unit lines);
  - Spending Data/pillars_1961_1979.csv;
  - Loans and Investment/loan_financed_projects_1955_1979.csv;
  - War and Security Data/war_cost_1968_1980.csv;
  - Revenue Data 1952-1990/tax_burden_memoria_1967_YYYY.csv (rename with `git mv` each year).
- Commit and push with normal git (the session's GitHub connection). **Never write or use any personal access token,** including the one pasted in the project instructions. The remote may have new uploads from the author: `git fetch` and rebase before pushing. Commit messages end with the Co-Authored-By and Claude-Session lines given by the session.
- Write both notes to the claude.ai project (`claude/<name>.md`) and send the .docx files to the author.

## 8. Sources

- **Figures:** only the Memorias de Hacienda (repo folder "Memorias de Hacienda"; MH 1976, 1977, 1978, 1979 A/B remain).
- **Context only** (as with Turcios, Pleitez and Dada Hirezi): the Ellacuría editorials (1971 ANDES strike; 1976 ISTA law; 1976 agrarian reform comments; "A sus órdenes mi capital" 1976; December 1979 conjuncture), the World Bank Economic Memoranda (1974, 1983) and the May 1979 report, Griffith on colonels and industrial workers, and the IMF growth paper. The author regards **Ellacuría** as a balanced primary analytical voice, trying to find an equilibrium in a polarized society. Add each to the Bibliography note when first used.

## 9. Open questions to carry forward

| Question | Where to look |
|:--|:--|
| ₡30M General Fund loan to the Compañía de Café, S.A. (1975): purpose and terms | MH 1976; Diario Oficial; coffee company records |
| Coffee duty fell 13% in 1975 on a flat price | Export volumes (BCR); duty schedule |
| Dignity bonds held by the General Fund (₡16.5M at end-1974): acquisition and fate | MH 1976; Central Bank |
| Substitute teachers during strikes: "other personal services" in the teachers' line, 1971–76 (₡1.55M in 1974) | Estado IV-2 in MH 1971–1977 |
| 1975 unit lines | MH 1976, Estado IV-2 |
| Tax measures of late 1974 (income tax, sugar duty) | Diario Oficial |
| CPI and real wages (teachers) | BCR / DIGESTYC |
| University reopening date (1973); Cerrón Grande displacement | Secondary sources |
| Content of the Presidency and Defensa lump sums; the Presidency's community-development programme | Budget laws; IV-2 detail |

## 10. Key series so far (₡ millions)

| | 1971 | 1972 | 1973 | 1974 | 1975 |
|:--|--:|--:|--:|--:|--:|
| Current revenue | 299.8 | 326.4 | 402.7 | 484.3 | 576.8 |
| Credits used | 356.1 | 383.7 | 450.0 | 594.5 | 709.6 |
| Underlying result | −19.4 | −15.4 | −8.7 | −43.0 | −59.6 |
| — % of revenue | −6.5% | −4.7% | −2.2% | −8.9% | −10.3% |
| Superávit Financiero | 5.0 | 11.8 | 39.1 | 42.4 | 101.6 |
| Coffee duty | | | 74.0 | 90.0 | 78.0 |
| Income tax | 45.3 | 52.1 | 62.6 | 80.0 | 118.9 |
| Security share of credits used | | | 8.6% | 8.5% | 7.4% |
| Development share of credits used | | | 51.7% | 49.3% | 48.3% |
| Direct public debt | 170.2 | 205.7 | 258.3 | 323.2 | 397.3 |
| War bonds outstanding | 31.81 | 32.31 | 31.31 | 30.31 | 29.31 |
