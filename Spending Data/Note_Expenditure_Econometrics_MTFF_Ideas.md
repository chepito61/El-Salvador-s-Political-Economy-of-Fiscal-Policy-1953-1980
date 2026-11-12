# From the Revenue Series to Expenditure, Econometrics and a Retrospective Fiscal Framework: Ideas for Discussion

**Working note for discussion — 12 November 2026.** It answers the author's three ideas:

1. an expenditure workbook that mirrors the revenue series;
2. an econometric analysis of the impact of fiscal policy;
3. a medium-term fiscal framework (MTFF) to run fiscal stress tests for the period.

Nothing is built yet. The note proposes designs, flags what the data allow, and lists the decisions needed.

---

## 1. The expenditure workbook

### 1.1 What we already have

| Data | Years | Where |
|:--|:--|:--|
| Spending by ministry (*ramo*) and object class (wages, goods, transfers, equipment…) | 1952–56, 1960, 1961–79 (1957–59 not yet extracted) | Spending Data/ministry_class_*.csv |
| Credits used, gastos and commitments | 1961–79; 1950s on the old basis | Same files; yearly notes |
| Pillars (development, security, agriculture, interior) | 1952–79 | pillars_*.csv |
| Unit lines (teachers, University subsidy, Fuerza Armada, Presidency and Defence lump sums, IRA) | Most years from the 1960s | ministry_class (unit column); notes |
| Economic classification (Cuadro 15, gastos basis) | 1968–79 | Yearly notes |
| Loan-financed projects (lender, instrument, amounts) | 1955–79 | Loans and Investment/loan_financed_projects_1955_1979.csv |
| Cost of the war | 1968–80 | War and Security Data/war_cost_1968_1980.csv |
| Balances, financing and debt | Every year | Yearly notes (not yet in one table) |

The expenditure side is in better shape than the revenue side was: most figures are already extracted and verified. The work is mostly assembly and harmonization.

### 1.2 Proposed structure (mirroring the revenue workbook)

1. **README and Mapping:** the basis of each year and how the ministries and classes are harmonized. The *ramos* change name and scope: Cultura Popular becomes Educación; security bodies sit inside Defence in some years; the Presidency's units vary.
2. **Exp_Ministry:** credits used by ministry, 1952–79, with checks against the printed totals. One column per harmonized ministry, plus "other organs" (Legislative, Judicial, Corte de Cuentas, elections).
3. **Exp_Economic:** an economic classification built in the modern style, from the object classes for 1952–67 and Cuadro 15 for 1968–79:
   - wages and salaries;
   - goods and services;
   - current transfers (University, municipalities, autonomous institutions, pensions);
   - interest;
   - capital: real investment, capital transfers, and financial investment (on-lending, e.g. to CEL);
   - amortization (below the line).
4. **Exp_Function:** the pillars (development, security, agro-export support, administration), using the definitions already decided.
5. **Exp_Units:** the unit lines that test the hypotheses (H1 education and University; H2 teachers' wages; H5 IRA).
6. **Investment:** loan-financed and locally financed investment, with **execution rates** (budgeted against executed) and the lenders. This is where the financing and execution of infrastructure goes.
7. **Balances_Financing:** current revenue (linked to the revenue workbook) minus credits used, giving:
   - the commitment-basis result;
   - the underlying result (our convention);
   - financing: Central Bank, domestic bonds, external loans, use of deposits;
   - debt stock and interest.
8. **Periods and Indicators:** the same governments as the revenue series; shares; real values; per-person values when the deflator and population exist.
9. **Charts:**
   - development against security;
   - wages as a share of spending;
   - investment against the coffee duty;
   - a scatter of the change in spending against the change in revenue (procyclicality);
   - executed against budgeted investment.

**Decided (author, 12 November 2026): one fiscal workbook combining revenue, expenditure and interest on the debt.** Section 1.4 sets out how the three fit together.

### 1.3 Problems to solve

- **Basis.** Credits used (gastos + commitments) is our main basis; Cuadro 15 is on gastos. The workbook should carry both, and say which each chart uses.
- **Breaks in the classification.** The 1950s object classes differ from the 1960s ones and from Cuadro 15. Each mapping goes in a Mapping sheet, as for revenue.
- **1957–59.** Ministry detail still to be extracted from MH 1957–1959 (A–D).
- **Coverage.** Central government only. The autonomous institutions (CEL, ANTEL, ISSS, IRA, ISTA) enter only through transfers and on-lending. A consolidated public sector is not possible from the Memorias.

### 1.4 How revenue, expenditure and interest fit together

The core sheet is a **fiscal accounts table**, one row per year, 1952–79, in the format of a modern fiscal summary:

| Line | Source |
|:--|:--|
| Current revenue (tax and non-tax, by source) | Revenue_CM and Revenue_Detail (existing) |
| Primary spending: wages; goods and services; transfers; capital | Exp_Economic |
| **Interest on the public debt**, split domestic and external, and war bonds from 1968 | Public-debt service in each Memoria (Ramo de Hacienda, *Deuda Pública*) |
| **Primary balance** = current revenue − primary spending | Formula |
| **Overall balance** = primary balance − interest | Formula; checked against our commitment-basis result |
| Amortization, and gross financing needs = −overall balance + amortization | Debt service lines |
| Financing: external loans, domestic bonds, Central Bank, deposits | Yearly notes |
| Debt stock (domestic, external, war bonds) and the implicit interest rate (interest / previous year's debt) | Debt tables in the notes |

**Why interest deserves its own line:**

- **Primary against overall balance.** This separates what each government decided from what it inherited.
- **The cost of borrowing.** The implicit interest rate shows how this cost changed as lending shifted: from the IBRD, the IDB and AID, to commercial borrowing and domestic bonds.
- **The framework and the stress tests** in section 3 run on the debt dynamics: interest = rate × debt.
- **Interest against spending.** Interest can be set against education or security spending. That makes it a political-economy indicator in its own right: who was paid first in a bust.

**Data to assemble:** interest and amortization by year are in the debt-service lines of each Memoria and partly in the notes. They have not yet been put into one series, and the 1950s will need checking line by line.

---

## 2. Econometric analysis of fiscal policy

### 2.1 The constraint

We have **about 27 annual observations (1953–79)**, with structural breaks (1961–63 tax reforms, 1969 war, 1976–78 boom) and some measurement gaps. That rules out the standard large models (e.g., a structural VAR of the Blanchard–Perotti type) as a credible stand-alone exercise. It still allows well-chosen, transparent tests. **El Salvador has one advantage: an exogenous shock that is measured well**, the world coffee price. It works as a natural experiment.

### 2.2 Options, from the most robust to the most ambitious

1. **Cyclicality of fiscal policy (descriptive and simple regressions).**
   - Regress the change in real spending (total and by component) on the change in output, or directly on the coffee price or coffee duty.
   - Test **asymmetry**: do booms raise spending more than busts cut it? This is the "ratchet" and "when it rains, it pours" literature on Latin America (Gavin and Perotti, 1997; Kaminsky, Reinhart and Végh, 2004).
   - This answers a core question of the project: did the regime spend windfalls and cut in busts?
2. **Where the windfalls went.** Regress each component's change on the change in coffee revenue:
   - security; education; University; investment; wages; debt repayment.
   - The coefficients say what share of each extra colón of coffee revenue went to each use. This is the cleanest quantitative test of H1–H5.
3. **Event studies around the shocks:** the 1958–59 bust, the 1969 war, the 1970 windfall, the 1976–77 boom and the 1978 bust. Each component's path from two years before to two years after, against its trend. This is transparent and suited to policy readers.
4. **Revenue and spending dynamics.** Test whether revenue drives spending or the reverse ("tax-and-spend" against "spend-and-tax"), with an error-correction model. It checks whether the regime adjusted spending to revenue or ran deficits and financed them.
5. **Impact on growth (the fiscal multiplier).** Use local projections (Jordà, 2005), with the coffee price shock as an instrument for revenue and spending, and output as the outcome. This is the most ambitious option and the least precise with 27 observations.
   - **A way to gain power:** a **Central American panel** (Guatemala, Honduras, Nicaragua, Costa Rica; all coffee exporters, with similar regimes) from CEPAL and IMF data. It turns the national story into a comparative test and gives about five times the observations.
6. **The link to polarization.** Relate fiscal variables (education spending per student, teachers' real wages, security share) to measures of mobilization and repression: strikes, protest events, political killings. The best-known event data are in Almeida, *Waves of Protest: Popular Struggle in El Salvador, 1925–2005* (2008). This can show timing and association, **not causation**; the note would say so.

### 2.3 Data needed

| Series | Years needed | Source |
|:--|:--|:--|
| GDP (nominal and real) | 1950–79 | CEPAL, *Estudio Económico*; BCR; Penn World Table (to check coverage) |
| GDP deflator or CPI before 1974 | 1950–73 | BCR; ILO; CEPAL |
| Population | 1950–79 | Censuses 1950, 1961, 1971; CELADE |
| Coffee export volume and price received | 1950–79 | BCR; Compañía Salvadoreña de Café; ICO |
| Protest and repression events | 1960–79 | Almeida (2008); newspapers |
| Central American fiscal series (panel option) | 1960–79 | CEPAL; IMF *Government Finance Statistics* (from the 1970s) |

---

## 3. A retrospective medium-term fiscal framework and stress tests

### 3.1 The idea

Build a simple MTFF of the kind a finance ministry would use today, and run it **as if prepared at key dates** with the information then available (for example 1962, 1968, 1973, 1977). Then shock it. It answers: how exposed was the State, what would a prudent rule have required, and what did the regime actually do?

### 3.2 A model small enough to calibrate

- **Revenue:**
  - coffee duty = price × volume × the duty schedule;
  - other taxes = elasticity × GDP growth (elasticities estimated from the series);
  - non-tax revenue as a share.
- **Spending:**
  - wages and transfers with **inertia** (the rigid part);
  - investment as the **adjustment variable**, as observed;
  - interest = average rate × debt.
- **Financing:**
  - external loans (by lender, from the loan register);
  - domestic bonds;
  - Central Bank credit.
- **Outputs:**
  - balance;
  - debt path;
  - gross financing needs;
  - the adjustment required to keep debt or the deficit within a limit.

### 3.3 Stress tests (historical, so they are realistic)

| Shock | Calibration |
|:--|:--|
| Coffee price fall | −30% (1978), −13% (1971), and a bootstrapped distribution of 1950–79 price changes for fan charts |
| Coffee volume shock (frost, rust, labour) | One bad harvest |
| Loss of the Common Market | The 1969 war: exports and transit to Honduras |
| War or security spending shock | The 1968–70 war costs |
| Interest-rate and external-credit shock | 1979–80 |
| Combined | Coffee bust with an interest-rate shock (1978–80) |

### 3.4 Policy counterfactuals the framework allows

- **A coffee stabilization fund:** save the duty above a reference price in 1976–77, and spend it in 1978–79. How much of the 1978 deficit would it have covered?
- **A spending rule:** for example, current spending growing with trend revenue, not with the coffee duty.
- **Protecting development spending:** what the 1978 bust would have required if education and health had been protected.

These follow the logic of the modern fiscal-risk tools (IMF, *Analyzing and Managing Fiscal Risks: Best Practices*, 2016; the IMF debt sustainability framework). They are applied with modest, transparent calibration.

---

## 4. Proposed sequence

1. **Expenditure, interest and debt sheets** in the combined fiscal workbook (sections 1.2–1.4), around the fiscal accounts table. Everything else needs this.
2. **GDP, deflator and population, 1950–79.** These are also needed for the revenue charts (tax/GDP; real revenue).
3. **Descriptive cyclicality and event studies** (options 2.2-1 to 2.2-3). They give results quickly and suit the policy notes.
4. **The retrospective MTFF and stress tests** (section 3), built on the same workbook.
5. **The more ambitious econometrics** (options 2.2-4 to 2.2-6), and the Central American panel if the author wants the comparative angle.

## 5. Decisions for the author

1. ~~One fiscal workbook or two?~~ **Decided: one workbook combining revenue, expenditure and interest on the debt.**
2. ~~Basis?~~ **Decided: credits used as the main basis, with gastos alongside.**
3. ~~1957–59?~~ **Decided: extract.**
4. ~~Panel?~~ **Decided: national analysis.**
5. ~~MTFF dates~~ **Set aside.** This question was about which years to run the framework from, as if a ministry had prepared it then. It is not needed now. The present task is to build the data series and prepare the data the framework will need.


---

## 6. Work plan agreed (12 November 2026): the building blocks

**Scope.** Build the expenditure series to mirror the revenue series: verified against printed totals, harmonized across years, documented, with periods by government and charts. Combine them with revenue and interest in one fiscal workbook. At the same time, assemble the series a medium-term fiscal framework will need. No econometrics and no framework model yet.

### 6.1 Steps

1. **Inventory and gap table.** A year-by-sheet grid showing what we already hold (from the CSVs and notes) and what is missing. We already know:
   - ministry totals exist for every year;
   - ministry-by-class detail is missing for 1957–59;
   - 1955–56 is one combined volume (a 1955–56 fiscal period): this needs a rule (keep it as one observation, or split it).
2. **Extract 1957–59** (MH 1957 A–C, 1958 A–D, 1959 A–D): spending by ministry and object class, credits used and gastos. Check against the printed totals, as for the other years.
3. **Interest and debt service, 1952–79:** interest and amortization by creditor, from the public-debt lines of each Memoria. Debt stocks at year-end by creditor (domestic, external, war bonds).
4. **Economic classification, 1952–79:**
   - 1952–67: built from the object classes (wages, goods and services, transfers, capital), with interest taken from step 3;
   - 1968–79: Cuadro 15.
   - The overlap years are reconciled, as was done for revenue in 1962–63.
5. **Assemble the fiscal workbook:**
   - the fiscal accounts table (section 1.4);
   - expenditure by ministry, by economic class and by pillar;
   - unit lines;
   - interest and debt;
   - checks;
   - periods by government;
   - charts.
6. **A Word note** on the expenditure series: design, first findings and decisions, like the revenue note.

### 6.2 Data for the future framework: what each series needs and where it stands

| Series | Status | To do |
|:--|:--|:--|
| Revenue by source, 1952–79 | Done | Link |
| Coffee duty, and duty per unit of price | Done 1960–79 (1971–72 gap) | Fill 1971–72 |
| Coffee price | Pink Sheet 1960–79 | 1950–59 (ICO; FAO; BCR) |
| Coffee export volume and value | Not held | BCR, *Revista Mensual*; ICO |
| Spending by economic class | Partial | Steps 2 and 4 |
| Interest, amortization, debt stock by creditor | In the notes, scattered | Step 3 |
| Terms of loans (rate, maturity, grace) | Partly in the loan register | Add from the loan contracts cited in the Memorias |
| GDP, nominal and real | Memoria from 1968 | 1950–67 (CEPAL; BCR) |
| Price index | 1974–79 only | 1950–73 (BCR; CEPAL) |
| Population | Not held | Censuses 1950, 1961, 1971; CELADE |
| Exchange rate | Fixed at ₡2.50 per US$ | — |

### 6.3 Points of judgement to flag now

- **The 1950s basis.** In the 1950s Memorias the concept of commitments may differ from the 1960s *créditos utilizados*. Each year's basis will be recorded in a Basis column, as in the revenue workbook.
- **On-lending and loan money.** Spending financed by loans stays inside credits used (it is spending). Loan proceeds stay out of revenue (they are financing), consistent with the conventions.
- **Interest inside the ministries.** Interest is paid from the Ramo de Hacienda (Deuda Pública). It is taken out of primary spending and shown on its own line, so the totals are not double-counted.
- **Real terms.** Real values before 1974 wait for a deflator. Until then, the series are shown at current prices and as shares.

**Suggested first step on your go-ahead:** step 1 (the inventory and gap table). It is quick, and it will show exactly how much new extraction steps 2 and 3 require.

---

*References: Gavin, M. and R. Perotti (1997), "Fiscal Policy in Latin America", NBER Macroeconomics Annual; Kaminsky, G., C. Reinhart and C. Végh (2004), "When It Rains, It Pours: Procyclical Capital Flows and Macroeconomic Policies", NBER Macroeconomics Annual; Jordà, Ò. (2005), "Estimation and Inference of Impulse Responses by Local Projections", American Economic Review; Blanchard, O. and R. Perotti (2002), Quarterly Journal of Economics; Almeida, P. (2008), Waves of Protest: Popular Struggle in El Salvador, 1925–2005, University of Minnesota Press; IMF (2016), Analyzing and Managing Fiscal Risks: Best Practices.*
