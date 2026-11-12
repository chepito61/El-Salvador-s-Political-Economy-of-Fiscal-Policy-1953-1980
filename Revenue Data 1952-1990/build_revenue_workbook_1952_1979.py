import json
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment

OUT = "El_Salvador_Revenue_Memorias_1952_1979.xlsx"
csvd = {int(k): v for k, v in json.load(open("revenue_1956_1963_checked.json")).items()}

# ---------------- data (₡ millions) ----------------
# keys: inc, pers, corp, oth, prop, imp, exp, cof, cons, svc, emer, tax_p, cur_p, src, basis
D = {}
D[1952] = dict(tax_p=106.381, cur_p=129.690, basis="Old classification (Títulos); customs not split",
               src="MH 1952 (Note_MH_1952_1953): Título V Impuestos ₡106.38M incl. Renta de Aduanas ₡68.15M (imports + exports combined); rentas percibidas y devengadas ₡129.69M")
D[1953] = dict(inc=10.450, prop=1.181 + 3.851, imp=48.633, exp=35.029, cof=34.965, cons=20.273,
               svc=125.655 - (10.450 + 1.181 + 3.851 + 48.633 + 35.029 + 20.273), tax_p=125.655, cur_p=140.313,
               basis="Subclass codes 010–090 (MH 1954 comparatives)",
               src="MH 1954 comparative columns (Note_MH_1954); services/stamps/other derived as residual; MH 1953 itself reports rentas of ₡137.61M (see Reconciliation)")
D[1954] = dict(inc=10.634, prop=2.849 + 8.817, imp=53.510, exp=46.243, cof=46.162, cons=25.704,
               svc=3.899 + 0.995 + 1.816, tax_p=154.467, cur_p=167.426, basis="Subclass codes 010–090",
               src="MH 1954 (Note_MH_1954)")
D[1955] = dict(inc=10.988, prop=2.644 + 1.648, imp=55.443, exp=49.020, cof=48.988, cons=26.039,
               svc=4.072 + 1.109 + 1.740, tax_p=152.704, cur_p=164.793, basis="Subclass codes 010–090",
               src="MH 1955-1956 (Note_MH_1955)")
for y, v in csvd.items():
    D[y] = dict(inc=v["inc"], prop=v["prop"], imp=v["imp"], exp=v["exp"], cof=v["cof"], cons=v["cons"],
                svc=v["svc"], tax_p=v["tax"], cur_p=v["cur"], basis="Subclass codes 010–090",
                src=f"MH {y} item detail (revenue_1952_1964_long.csv), checked: subclasses sum to printed totals; totals agree with Note_MH_{y if y>1957 else '1957'}")
D[1956]["src"] = "MH 1955-1956 item detail; income tax proper ₡12.32M from Note_MH_1955_1956 (Vialidad = Renta y Vialidad − 12.32)"
D[1957]["src"] = "MH 1957 item detail; 'Otros Impuestos' corrected to ₡2,078,257.54 per Note_MH_1957"
D[1962]["src"] = "MH 1962 item detail; non-tax = rentas − taxes per Note_MH_1962 (item file ₡0.18M short)"
N = dict
D[1964] = N(inc=35.763, pers=17.758, corp=14.212, oth=3.793, prop=10.17, imp=63.93, exp=31.55, cof=30.46, cons=45.95, svc=4.65 + 7.07 + 0.52, tax_p=199.60, cur_p=210.46)
D[1965] = N(inc=34.662, pers=18.251, corp=11.492, oth=4.920, prop=9.74, imp=62.74, exp=36.79, cof=36.01, cons=54.49, svc=5.44 + 7.93 + 0.56, tax_p=212.35, cur_p=224.54)
D[1966] = N(inc=34.715, pers=19.925, corp=9.587, oth=5.203, prop=9.20, imp=62.68, exp=33.72, cof=32.56, cons=55.33, svc=5.63 + 8.84, tax_p=210.12, cur_p=223.52)
D[1967] = N(inc=38.667, pers=21.732, corp=11.261, oth=5.674, prop=11.49, imp=59.39, exp=31.06, cof=30.14, cons=56.38, svc=15.00, tax_p=211.98, cur_p=226.91)
D[1968] = N(inc=45.226, pers=28.328, corp=11.162, oth=5.736, prop=21.09, imp=52.91, exp=24.41, cof=23.66, cons=57.57, svc=15.81, tax_p=217.02, cur_p=230.98)
D[1969] = N(inc=38.063, pers=24.194, corp=11.060, oth=2.809, prop=24.82, imp=57.20, exp=27.40, cof=26.52, cons=60.66, svc=18.42, emer=9.30, tax_p=235.88, cur_p=251.95)
D[1970] = N(inc=39.661, pers=22.004, corp=11.600, oth=6.057, prop=17.82 + 7.08, imp=65.03, exp=48.47, cof=43.12, cons=64.63, emer=0.29, tax_p=264.59, cur_p=281.84)
D[1970]["svc"] = 264.59 - (39.661 + 24.90 + 65.03 + 48.47 + 64.63 + 0.29)
D[1971] = N(inc=45.32, prop=24.49, imp=68.33, exp=38.24, cons=69.35, svc=7.78 + 24.63 + 0.79, tax_p=278.93, cur_p=299.75)
D[1972] = N(inc=52.09, prop=23.50, imp=70.87, exp=47.50, cons=73.78, svc=8.29 + 28.04 + 0.83, tax_p=304.90, cur_p=326.38)
D[1973] = N(inc=62.626, pers=32.196, corp=20.515, oth=9.914, prop=30.81, imp=78.47, exp=76.12, cof=74.03, cons=83.31, svc=9.85 + 34.07 + 0.82, tax_p=376.07, cur_p=402.66)
D[1974] = N(inc=80.00, pers=36.28, corp=32.26, oth=11.46, prop=29.43, imp=93.11, exp=93.17, cof=89.95, cons=97.58, svc=11.64 + 47.56 + 0.78, tax_p=453.26, cur_p=484.35)
D[1975] = N(inc=118.92, pers=51.03, corp=47.28, oth=20.61, prop=35.28, imp=92.80, exp=111.80, cof=77.96, cons=115.22, svc=12.79 + 53.52 + 0.99, tax_p=541.32, cur_p=576.79)
D[1976] = N(inc=139.40, pers=64.85, corp=52.06, oth=22.49, prop=47.60, imp=115.26, exp=240.10, cof=227.17, cons=140.97, svc=16.03 + 67.19 + 1.01, tax_p=767.56, cur_p=809.98)
D[1977] = N(inc=164.15, pers=79.98, corp=55.95, oth=28.23, prop=73.30, imp=163.64, exp=460.05, cof=416.82, cons=167.14, svc=18.39 + 84.17 + 0.64, tax_p=1131.48, cur_p=1181.83)
D[1978] = N(inc=201.78, pers=90.45, corp=77.86, oth=33.47, prop=83.78, imp=157.34, exp=244.22, cof=126.89, cons=170.12, svc=20.91 + 93.31 + 0.62, tax_p=972.08, cur_p=1026.93)
D[1979] = N(inc=192.880, pers=97.825, corp=73.218, oth=21.836, prop=84.403, imp=178.279, exp=392.921, cof=379.549, cons=191.551, svc=20.887 + 100.756 + 0.600, tax_p=1162.277, cur_p=1215.390)
srcs = {1964: "MH 1964 Estado II-5 (Note_MH_1964); income tax split includes fines",
        1965: "MH 1965 Estado II-5 (Note_MH_1965)", 1966: "MH 1966 Estado II-5 (Note_MH_1966)",
        1967: "MH 1967 Estado II-5 (Note_MH_1967)", 1968: "MH 1968 Estado II-5 (Note_MH_1968)",
        1969: "MH 1969 Estado II-5 (Note_MH_1969); income tax split excludes fines; emergency tax D.L. 421",
        1970: "MH 1970 Memoria groupings (Note_MH_1970); tax total from MH 1972 Cuadro 12; services/stamps/other derived as residual",
        1971: "MH 1972 Cuadro 10, 1971 column (Note_MH_1972); coffee duty and income tax split not available",
        1972: "MH 1972 Cuadro 10 (Note_MH_1972); coffee duty and income tax split not available",
        1973: "MH 1973 Cuadro 10, Estado I-4 (Note_MH_1973)", 1974: "MH 1974 Cuadro 9, Estado I-4 (Note_MH_1974)",
        1975: "MH 1975 Estado I-4 (Note_MH_1975)", 1976: "MH 1976 Estado II-1 (Note_MH_1976)",
        1977: "MH 1977 Estado II-1 (Note_MH_1977)", 1978: "MH 1978 Estado II-1 (Note_MH_1978); excludes ₡60M ISSS bonds",
        1979: "MH 1979 Estado II-1 and I-4 (Note_MH_1979); excludes ₡2.1M foreign loans for operating programmes"}
for y in range(1964, 1980):
    D[y]["src"] = srcs[y]
    D[y]["basis"] = "Estado II-5 / II-1 economic classification" if y not in (1970,) else "Memoria groupings, mapped"

coffee_price = {1960: 0.9235, 1961: 0.8955, 1962: 0.8326, 1963: 0.8057, 1964: 1.0098, 1965: 1.0003, 1966: 0.9287,
                1967: 0.8642, 1968: 0.8670, 1969: 0.8767, 1970: 1.1466, 1971: 0.9919, 1972: 1.1096, 1973: 1.3735,
                1974: 1.4515, 1975: 1.4421, 1976: 3.1467, 1977: 5.1689, 1978: 3.5900, 1979: 3.8261}
gdp = {1968: 2291.911, 1969: 2381.819, 1970: 2564.570, 1971: 2697.564, 1972: 2890.835, 1973: 3197.668,
       1974: 3943.574, 1975: 4477.687, 1976: 5705.879, 1977: 7119.017, 1978: 7679.016, 1979: 8825.667}
gdp_src = {1968: "MH 1972", 1969: "MH 1972", 1970: "MH 1972", 1971: "MH 1973 (revised)", 1972: "MH 1973 (revised)",
           1973: "MH 1973", 1974: "MH 1978 (revised)", 1975: "MH 1979", 1976: "MH 1979 (revised)",
           1977: "MH 1979 (revised)", 1978: "MH 1979 (preliminary)", 1979: "MH 1979 (estimated)"}
cpi = {1974: 144.7, 1975: 172.2, 1976: 184.4, 1977: 206.1, 1978: 233.5}

laws = [
 (1950, "Coffee export tax made ad valorem and progressive; coffee producers exempt from income tax on coffee land", "D. 836 (13 Sep 1950, DO 200, t. 149)", "Pleitez; Context note"),
 (1951, "First modern income tax law (individuals 3–44%; companies 2.5–5%)", "D. 510 (10 Dec 1951, DO 232, t. 153)", "Pleitez; Context note"),
 (1952, "Sugar consumption tax ₡4/qq; sugar export tax removed", "D. 618 (12 Mar 1952)", "Pleitez"),
 (1952, "Ley de Fomento de Industrias de Transformación: import and income-tax exemptions", "D.L. 661 (22 May 1952, DO 102, t. 155)", "Pleitez"),
 (1957, "New tariff schedule: higher protective duties, lower duties on inputs (approx. 1957–59)", "Tarifa de Aforos (to identify)", "Turcios; Narrative Exercise"),
 (1959, "New consumption taxes on beer, soft drinks, matches and liquor (February 1959)", "To identify", "Turcios, Autoritarismo, p. 143"),
 (1961, "New Ley de Fomento Industrial", "D. 64 (18 Jan 1961)", "Pleitez"),
 (1961, "Shrimp export tax US$0.06/lb", "D. 154 (5 Jun 1961)", "Pleitez; Note_MH_1962"),
 (1961, "Income tax reform: top personal rate 76.5%; companies 2.5–15%", "D. 300 (18 Sep 1961)", "Pleitez"),
 (1961, "New income tax law: reinvested industrial profits exempt", "D. 442 (6 Dec 1961)", "Pleitez"),
 (1963, "New consumption taxes on petroleum products refined in the country and on flour", "To identify", "Note_MH_1963"),
 (1963, "New income tax law (in force 1964)", "D. 479 (19 Dec 1963)", "Pleitez; Note_MH_1964"),
 (1968, "Increase in Vialidad Serie 'A' (tax on capital)", "To identify (late 1967–early 1968)", "Note_MH_1968"),
 (1968, "CACM stabilization surcharge on imports (San José Protocol)", "To identify", "Note_MH_1969"),
 (1969, "Emergency tax after the war with Honduras", "D.L. 421 (DO 3 Jul 1969)", "Note_MH_1969"),
 (1970, "First Ley de Fomento de Exportaciones", "D.L. 148 (3 Nov 1970)", "Pleitez"),
 (1971, "Stamp tax yield doubles (+108%); legal basis to identify", "To identify", "Note_MH_1971"),
 (1974, "New export law: free zones, export credit guarantees", "D.L. 81 (5 Sep 1974)", "Pleitez"),
 (1974, "Progressive sugar export tax above US$10/qq", "D.L. 135 (19 Nov 1974)", "Pleitez; Note_MH_1975"),
 (1974, "Other tax measures of late 1974 (income tax)", "To identify", "Note_MH_1975"),
 (1976, "Tax amnesty cited by the World Bank", "D. 248 (date to identify)", "World Bank 1979; Note_MH_1976"),
 (1977, "'Other merchandise' export duty grows (₡41.9M; ₡116.6M in 1978); legal basis to identify", "To identify", "Note_MH_1977, 1978"),
 (1977, "Agricultural land tax (Impuesto Territorial Agropecuario)", "Decreed 22 Dec 1977 (number to identify)", "MH 1978, p. 36"),
 (1978, "Incentive measures for the taxation of enterprises (December 1978)", "To identify", "MH 1979, section 6.2"),
]

# ---------------- styles ----------------
F = "Arial"
fb = Font(name=F, bold=True); fn = Font(name=F); fin = Font(name=F, color="0000FF"); fgr = Font(name=F, color="008000")
fi = Font(name=F, italic=True, color="666666"); ft = Font(name=F, bold=True, size=13)
hdr_fill = PatternFill("solid", fgColor="DDE4EE"); thin = Side(style="thin", color="999999")
num = '#,##0.00;(#,##0.00);"-"'; pct = '0.0%;(0.0%);"-"'

wb = Workbook()

# README
ws = wb.active; ws.title = "README"
rows = [
 ("El Salvador: Central Government Revenue from the Memorias de Hacienda, 1952–1979", ft),
 ("Working dataset — November 2026. Unit: millions of colones (₡M), current prices. US$1.00 = ₡2.50.", fi),
 ("", fn),
 ("What this workbook contains", fb),
 ("Revenue_CM: revenue by source, one row per year, mapped into one modern classification. Totals and checks are formulas.", fn),
 ("Shares: each source as a share of current revenue (formulas).", fn),
 ("Indicators: coffee price and the coffee duty, direct taxes, corporate vs personal income tax, growth, tax/GDP, real revenue.", fn),
 ("Tax_Laws: tax laws and decrees named in the Memorias, the notes and the secondary sources — a checklist for the Diario Oficial.", fn),
 ("Mapping: how the 1950s subclass codes and the later classifications were mapped into the categories used here.", fn),
 ("Reconciliation: overlaps between classifications and differences with the co-authored dataset (All_revenue_1952_1990.csv).", fn),
 ("Charts: revenue composition, coffee duty vs coffee price, personal vs corporate income tax.", fn),
 ("", fn),
 ("Definitions", fb),
 ("Current revenue = rentas percibidas y devengadas (collected plus accrued). Excludes bond proceeds, loans and 'ingresos que no representan rentas' "
  "(e.g. the ₡60M Social Security bonds in 1978). From 1964 the Memoria's Estado II-5 includes small non-revenue receipts (item 900) in current revenue; see Reconciliation.", fn),
 ("Tax revenue includes the 1969–70 emergency tax (D.L. 421), shown in its own column.", fn),
 ("Income tax: personal, corporate (renta de empresas / sociedades) and other income. 1964–68 include fines; 1969 onward exclude fines.", fn),
 ("Property and wealth: Vialidad Serie 'A' (tax on capital), succession and gift taxes; from 1964 also property transfers and paving levies (Estado II-5 grouping). "
  "From 1977 also the agricultural land tax.", fn),
 ("Services, stamps and other taxes: especies fiscales (stamps, sealed paper), taxes on services and vehicles, levies on legal acts and business, other taxes.", fn),
 ("Non-tax revenue: sales, services of official enterprises, licences, fines, rents and interest, current transfers and other current revenue.", fn),
 ("", fn),
 ("Colour code", fb),
 ("Blue = figure entered from a source (each row's source is in column V of Revenue_CM). Black = formula. Green = link to another sheet.", fn),
 ("", fn),
 ("Main caveats", fb),
 ("1952 uses the old classification (Títulos); customs revenue is not split into imports and exports, so only totals are shown.", fn),
 ("The mapping of the 1953–1963 subclasses into the modern groups is close but not exact (see Mapping and Reconciliation: overlap years 1962–63).", fn),
 ("1970: categories mapped from the Memoria's own groupings; 'services, stamps and other' is the residual to the printed tax total.", fn),
 ("1971–72: the coffee duty and the personal/corporate split are not available (Estado II-5 missing from the scans).", fn),
 ("No annual coffee price before 1960 in the repository (World Bank Pink Sheet starts in 1960). GDP is the Memoria's own (latest vintage available).", fn),
 ("Figures are as verified in the yearly notes (Policy and Technical Notes), which give page references.", fn),
]
for i, (t, f) in enumerate(rows, 1):
    c = ws.cell(i, 1, t); c.font = f; c.alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 150

# Revenue_CM
ws = wb.create_sheet("Revenue_CM")
H = ["Year", "Income tax", "— personal", "— corporate", "— other income", "Property and wealth", "Import duties",
     "Export duties", "— coffee", "— other exports", "Consumption of goods (excises)", "Services, stamps and other taxes",
     "Emergency tax (1969–70)", "Tax revenue (sum)", "Tax revenue (printed)", "Check: sum − printed", "Non-tax revenue",
     "Current revenue (sum)", "Current revenue (printed)", "Check: sum − printed", "Basis", "Source"]
ws.cell(1, 1, "Central government revenue, ₡ millions (current prices) — Memorias de Hacienda, 1952–1979").font = ft
for j, h in enumerate(H, 1):
    c = ws.cell(3, j, h); c.font = fb; c.fill = hdr_fill; c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    c.border = Border(bottom=thin)
ws.row_dimensions[3].height = 45
first = 4
years = list(range(1952, 1980))
for i, y in enumerate(years):
    r = first + i; d = D[y]
    ws.cell(r, 1, str(y)).font = fb
    def put(col, key):
        if key in d and d[key] is not None:
            c = ws.cell(r, col, round(d[key], 4)); c.font = fin; c.number_format = num
    for col, key in [(2, "inc"), (3, "pers"), (4, "corp"), (5, "oth"), (6, "prop"), (7, "imp"), (8, "exp"), (9, "cof"),
                     (11, "cons"), (12, "svc"), (13, "emer"), (15, "tax_p"), (19, "cur_p")]:
        put(col, key)
    if y == 1952:
        ws.cell(r, 14, "n.a.").font = fi; ws.cell(r, 16, "n.a.").font = fi
        ws.cell(r, 17, f"=S{r}-O{r}").font = fn
        ws.cell(r, 18, "n.a.").font = fi; ws.cell(r, 20, "n.a.").font = fi
    else:
        if "cof" in d: ws.cell(r, 10, f"=H{r}-I{r}")
        ws.cell(r, 14, f"=B{r}+F{r}+G{r}+H{r}+K{r}+L{r}+M{r}")
        ws.cell(r, 16, f"=N{r}-O{r}")
        ws.cell(r, 17, round(d["cur_p"] - d["tax_p"], 4)).font = fin
        ws.cell(r, 18, f"=N{r}+Q{r}")
        ws.cell(r, 20, f"=R{r}-S{r}")
    for col in (10, 14, 16, 17, 18, 20):
        ws.cell(r, col).number_format = num
    ws.cell(r, 21, d.get("basis", "")).font = fi
    ws.cell(r, 22, d.get("src", "")).font = fi
last = first + len(years) - 1
ws.cell(last + 2, 1, "Notes").font = fb
notes = [
 "Non-tax revenue is entered as printed current revenue minus printed tax revenue, so the current-revenue check tests the tax categories only. Differences up to ±0.02 are rounding in the notes' ₡M figures.",
 "1952: Título V 'Impuestos' (₡106.38M) includes the Renta de Aduanas (₡68.15M, imports and exports together); non-tax shown is rentas minus Título V.",
 "1953: Memoria 1953 reports rentas of ₡137.61M (renta neta); the 1954 Memoria's comparative column, which carries the breakdown used here, gives ₡140.31M. See Reconciliation.",
 "1969–70: the emergency tax is part of tax revenue (₡9.30M and ₡0.29M). 1969 ordinary current revenue was ₡242.65M.",
 "1978: excludes ₡60.0M of bonds issued to the Social Security Institute. 1979: excludes ₡2.1M of foreign loans for operating programmes.",
]
for k, t in enumerate(notes):
    c = ws.cell(last + 3 + k, 1, t); c.font = fi
ws.freeze_panes = "B4"
ws.column_dimensions["A"].width = 8
for j in range(2, 21): ws.column_dimensions[L(j)].width = 12
ws.column_dimensions["U"].width = 34; ws.column_dimensions["V"].width = 90
REV = "Revenue_CM"

# Shares
ws2 = wb.create_sheet("Shares")
ws2.cell(1, 1, "Revenue by source, % of current revenue").font = ft
SH = [("Income tax", "B"), ("— personal", "C"), ("— corporate", "D"), ("Property and wealth", "F"), ("Import duties", "G"),
      ("Export duties", "H"), ("— coffee", "I"), ("Consumption of goods", "K"), ("Services, stamps and other", "L"),
      ("Emergency tax", "M"), ("Non-tax revenue", "Q"), ("Tax revenue", "N")]
ws2.cell(3, 1, "Year").font = fb; ws2.cell(3, 1).fill = hdr_fill
for j, (h, _) in enumerate(SH, 2):
    c = ws2.cell(3, j, h); c.font = fb; c.fill = hdr_fill; c.alignment = Alignment(wrap_text=True, horizontal="center")
ws2.row_dimensions[3].height = 32
for i, y in enumerate(years):
    r = 4 + i; sr = first + i
    ws2.cell(r, 1, str(y)).font = fb
    for j, (h, col) in enumerate(SH, 2):
        if y == 1952 and col not in ("Q",):
            ws2.cell(r, j, "n.a." if col != "N" else f"='{REV}'!O{sr}/'{REV}'!S{sr}").font = fi if col != "N" else fgr
            if col == "N": ws2.cell(r, j).number_format = pct
            continue
        f = f"=IF(ISNUMBER('{REV}'!{col}{sr}),'{REV}'!{col}{sr}/'{REV}'!S{sr},\"\")"
        c = ws2.cell(r, j, f); c.font = fgr; c.number_format = pct
for j in range(1, 14): ws2.column_dimensions[L(j)].width = 13
ws2.freeze_panes = "B4"

# Indicators
ws3 = wb.create_sheet("Indicators")
ws3.cell(1, 1, "Indicators").font = ft
IH = ["Year", "Coffee price (US$/kg)", "Coffee duty (₡M)", "Coffee duty / current revenue", "Coffee duty per US$ of price (₡M)",
      "Direct taxes / current revenue", "Corporate / personal income tax", "Corporate income tax / coffee duty",
      "Current revenue growth", "Tax revenue growth", "GDP (₡M, Memoria)", "GDP source", "Tax revenue / GDP",
      "CPI (BCR, 1972=100)", "Current revenue, 1978 prices (₡M)"]
for j, h in enumerate(IH, 1):
    c = ws3.cell(3, j, h); c.font = fb; c.fill = hdr_fill; c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
ws3.row_dimensions[3].height = 48
for i, y in enumerate(years):
    r = 4 + i; sr = first + i
    ws3.cell(r, 1, str(y)).font = fb
    if y in coffee_price:
        c = ws3.cell(r, 2, coffee_price[y]); c.font = fin; c.number_format = "0.000"
    c = ws3.cell(r, 3, f"=IF(ISNUMBER('{REV}'!I{sr}),'{REV}'!I{sr},\"\")"); c.font = fgr; c.number_format = num
    c = ws3.cell(r, 4, f"=IF(ISNUMBER(C{r}),C{r}/'{REV}'!S{sr},\"\")"); c.number_format = pct
    c = ws3.cell(r, 5, f"=IF(AND(ISNUMBER(B{r}),ISNUMBER(C{r})),C{r}/B{r},\"\")"); c.number_format = num
    if y != 1952:
        c = ws3.cell(r, 6, f"=('{REV}'!B{sr}+'{REV}'!F{sr})/'{REV}'!S{sr}"); c.number_format = pct
    c = ws3.cell(r, 7, f"=IF(AND(ISNUMBER('{REV}'!D{sr}),ISNUMBER('{REV}'!C{sr})),'{REV}'!D{sr}/'{REV}'!C{sr},\"\")"); c.number_format = "0.00"
    c = ws3.cell(r, 8, f"=IF(AND(ISNUMBER('{REV}'!D{sr}),ISNUMBER('{REV}'!I{sr})),'{REV}'!D{sr}/'{REV}'!I{sr},\"\")"); c.number_format = pct
    if i > 0:
        c = ws3.cell(r, 9, f"='{REV}'!S{sr}/'{REV}'!S{sr-1}-1"); c.number_format = pct
        if y != 1953:
            c = ws3.cell(r, 10, f"='{REV}'!O{sr}/'{REV}'!O{sr-1}-1"); c.number_format = pct
    if y in gdp:
        c = ws3.cell(r, 11, gdp[y]); c.font = fin; c.number_format = num
        ws3.cell(r, 12, gdp_src[y]).font = fi
        c = ws3.cell(r, 13, f"='{REV}'!O{sr}/K{r}"); c.number_format = pct
    if y in cpi:
        c = ws3.cell(r, 14, cpi[y]); c.font = fin; c.number_format = "0.0"
    if y == 1979:
        c = ws3.cell(r, 14, f"=N{r-1}*(1+0.1406)"); c.number_format = "0.0"
        ws3.cell(r, 14).comment = Comment("1979 = 1978 × (1 + 14.06%), World Bank WDI inflation (FRED FPCPITOTLZGSLV).", "Claude")
    if y in cpi or y == 1979:
        c = ws3.cell(r, 15, f"='{REV}'!S{sr}*$N${4+years.index(1978)}/N{r}"); c.number_format = num
ws3.cell(last + 2, 1, "Sources: coffee price, World Bank Pink Sheet ('Commodities y recaudacion.xlsx', sheet Graficas; annual average of monthly US$/kg). "
         "GDP: Memorias de Hacienda, tax-burden tables (tax_burden_memoria_1967_1979.csv), latest vintage for each year. "
         "CPI: Central Reserve Bank via World Bank Report 2287-ES (1979), Table 9.1; 1979 extrapolated with WDI inflation.").font = fi
ws3.cell(last + 3, 1, "Direct taxes = income tax + property and wealth. Corporate income tax = renta de empresas (sociedades). "
         "Coffee duty per US$ of price shows the progressivity of the duty and the effect of export volumes.").font = fi
for j in range(1, 16): ws3.column_dimensions[L(j)].width = 14
ws3.column_dimensions["L"].width = 20
ws3.freeze_panes = "B4"

# Tax_Laws
ws4 = wb.create_sheet("Tax_Laws")
ws4.cell(1, 1, "Tax laws and decrees affecting revenue — checklist for the Diario Oficial").font = ft
for j, h in enumerate(["Year", "Measure", "Decree / reference", "Where it is cited", "Diario Oficial checked?"], 1):
    c = ws4.cell(3, j, h); c.font = fb; c.fill = hdr_fill
for i, (y, m, dref, s) in enumerate(laws):
    r = 4 + i
    ws4.cell(r, 1, str(y)).font = fb; ws4.cell(r, 2, m).font = fn; ws4.cell(r, 3, dref).font = fin
    ws4.cell(r, 4, s).font = fi; c = ws4.cell(r, 5, "No"); c.fill = PatternFill("solid", fgColor="FFFF00")
    for j in (2, 3, 4): ws4.cell(r, j).alignment = Alignment(wrap_text=True, vertical="top")
ws4.cell(5 + len(laws), 1, "Yellow cells: mark 'Yes' and add the DO number, tomo and date once the decree has been read.").font = fi
ws4.column_dimensions["A"].width = 8; ws4.column_dimensions["B"].width = 70; ws4.column_dimensions["C"].width = 36
ws4.column_dimensions["D"].width = 30; ws4.column_dimensions["E"].width = 14

# Mapping
ws5 = wb.create_sheet("Mapping")
ws5.cell(1, 1, "Mapping into the categories used in this workbook").font = ft
MP = [("Category (this workbook)", "1953–1963: Memoria subclass codes", "1964–1969: Estado II-5", "1970–1971: Memoria groupings", "1972–1979: Cuadro 10 / Estado I-4 / II-1"),
      ("Income tax", "041 Impuesto sobre la Renta", "Impuesto sobre la Renta (personal, empresas, otras rentas)", "Income tax", "Income tax (personal, companies, other income)"),
      ("Property and wealth", "045 Vialidad Serie 'A' + 060 Sucesiones y Donaciones", "Patrimonio y propiedades (Vialidad, transfers, paving, successions, gifts)", "Vialidad + successions and gifts", "Property and wealth (incl. agricultural land tax from 1977)"),
      ("Import duties", "010 Derechos de Importación", "Import duties (incl. CACM stabilization tax from 1968)", "Import duties (incl. stabilization tax)", "Import duties"),
      ("Export duties", "020 Derechos de Exportación (021 coffee)", "Export duties (coffee, shrimp)", "Export duties (coffee)", "Export duties (coffee, sugar, shrimp, other merchandise)"),
      ("Consumption of goods", "030 Artículos Manufacturados y de Consumo", "Consumo de productos", "Manufactured goods and consumption", "Excises (consumption of goods)"),
      ("Services, stamps and other taxes", "050 Especies Fiscales e Impuestos por Servicios + 070 Matrículas de Vehículos + 090 Otros Impuestos (incl. alcabala)", "Use of services + productive and commercial activities (stamps) + other taxes", "Services, fees, stamps and other taxes (1970: residual)", "Use of services + levies on legal acts (stamps) and business + other taxes"),
      ("Emergency tax", "—", "Emergency tax (D.L. 421), 1969", "Residual emergency tax, 1970", "Included in other taxes (small)"),
      ("Non-tax revenue", "100 Ventas + 200 Servicios + 300 Derechos, licencias + 400 Multas + 700 Otras rentas", "Non-tax revenue + transfers and other current revenue", "Non-tax and transfers", "Non-tax revenue + transfers and other")]
for i, row in enumerate(MP):
    for j, t in enumerate(row, 1):
        c = ws5.cell(3 + i, j, t); c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = fb if i == 0 else fn
        if i == 0: c.fill = hdr_fill
ws5.cell(13, 1, "Main breaks: (1) in 1953–63 'property' excludes the alcabala (property-transfer tax), which sits in 'other taxes'; from 1964 it is in property. "
          "(2) In 1953–63 some consular and customs fees count as non-tax revenue. (3) 1952 uses Títulos (customs combined) and cannot be mapped. "
          "The overlap years 1962–63, published on both bases, show the size of the break (Reconciliation sheet).").font = fi
ws5.cell(13, 1).alignment = Alignment(wrap_text=True)
for j, w in enumerate([24, 42, 40, 32, 40], 1): ws5.column_dimensions[L(j)].width = w

# Reconciliation
ws6 = wb.create_sheet("Reconciliation")
ws6.cell(1, 1, "Reconciliation").font = ft
ws6.cell(3, 1, "A. Overlap 1962–63: old subclass basis (this workbook) vs Estado II-5 restatement in MH 1964 (₡M)").font = fb
hdr = ["Category", "1962 old", "1962 II-5", "Difference", "1963 old", "1963 II-5", "Difference"]
for j, h in enumerate(hdr, 1):
    c = ws6.cell(4, j, h); c.font = fb; c.fill = hdr_fill
ov = [("Income tax", "B", 23.54, 24.75), ("Property and wealth", "F", 7.47, 8.61), ("Import duties", "G", 60.02, 60.24),
      ("Export duties", "H", 21.09, 22.68), ("Consumption of goods", "K", 32.11, 38.89),
      ("Services, stamps and other taxes", "L", 3.49 + 5.25 + 0.34, 4.05 + 6.17 + 0.45), ("Tax revenue", "O", 153.31, 165.83),
      ("Current revenue", "S", 171.68, 181.13)]
r62 = first + years.index(1962); r63 = first + years.index(1963)
for i, (nm, col, v62, v63) in enumerate(ov):
    r = 5 + i
    ws6.cell(r, 1, nm).font = fn
    c = ws6.cell(r, 2, f"='{REV}'!{col}{r62}"); c.font = fgr; c.number_format = num
    c = ws6.cell(r, 3, round(v62, 2)); c.font = fin; c.number_format = num
    ws6.cell(r, 4, f"=C{r}-B{r}").number_format = num
    c = ws6.cell(r, 5, f"='{REV}'!{col}{r63}"); c.font = fgr; c.number_format = num
    c = ws6.cell(r, 6, round(v63, 2)); c.font = fin; c.number_format = num
    ws6.cell(r, 7, f"=F{r}-E{r}").number_format = num
ws6.cell(14, 1, "Source of II-5 figures: MH 1964, Estado II-5 (Note_MH_1964). The II-5 current revenue includes non-revenue receipts (item 900, ₡3.07M in 1962).").font = fi

ws6.cell(16, 1, "B. Current revenue: this workbook vs the co-authored dataset (All_revenue_1952_1990.csv, 'Total Ingresos Corrientes' or 'Total Ingresos Ordinarios') (₡M)").font = fb
for j, h in enumerate(["Year", "This workbook", "Co-authored dataset", "Difference", "Comment"], 1):
    c = ws6.cell(17, j, h); c.font = fb; c.fill = hdr_fill
import csv
lucas = {}
for x in csv.DictReader(open("All_revenue_1952_1990.csv", encoding="utf-8-sig")):
    y = int(x["year"]); dsc = x["description"]
    if x["value"] in ("NA", "", "0"): continue
    v = float(x["value"]) / 1e6
    if dsc.startswith("Total Ingresos Corrientes") or dsc == "TOTAL INGRESOS ORDINARIOS":
        lucas.setdefault(y, v)
cm = {1952: "Dataset gives Total Ingresos Ordinarios before the ₡2.0M customs-policy adjustment (Art. 99)",
      1953: "Both use the 1954 Memoria comparative total; MH 1953 itself reports ₡137.61M",
      1955: "No current-revenue total for 1955 in the dataset; some of its 1955 items are copied from other years (see _MIRROR_SOURCE.md)",
      1964: "No current-revenue total for 1964 matched in the dataset",
      1968: "Dataset ₡0.50M higher; this workbook follows MH 1968 Estado II-5",
      1969: "Both include the emergency tax",
      1975: "Dataset ₡1.15M higher; this workbook follows MH 1975 Estado I-4",
      1977: "Dataset ₡2.0M lower; this workbook follows MH 1977 Estado II-1"}
for i, y in enumerate(years):
    r = 18 + i
    ws6.cell(r, 1, str(y)).font = fb
    c = ws6.cell(r, 2, f"='{REV}'!S{first+i}"); c.font = fgr; c.number_format = num
    if y in lucas:
        c = ws6.cell(r, 3, round(lucas[y], 4)); c.font = fin; c.number_format = num
        ws6.cell(r, 4, f"=C{r}-B{r}").number_format = num
    else:
        ws6.cell(r, 3, "n.a.").font = fi
    ws6.cell(r, 5, cm.get(y, "")).font = fi
for j, w in enumerate([30, 14, 18, 12, 14, 14, 12], 1): ws6.column_dimensions[L(j)].width = w
ws6.column_dimensions["E"].width = 70

# Charts sheet + chart data
ws7 = wb.create_sheet("Charts")
ws7.cell(1, 1, "Charts (data below are linked to the other sheets)").font = ft
# data block for composition, 1953-1979
cats = [("Income tax", "B"), ("Property and wealth", "F"), ("Import duties", "G"), ("Export duties: coffee", None),
        ("Export duties: other", "J"), ("Consumption of goods", "K"), ("Services, stamps, other and emergency", None), ("Non-tax revenue", "Q")]
base = 70
ws7.cell(base, 1, "Year").font = fb
for j, (nm, _) in enumerate(cats, 2):
    ws7.cell(base, j, nm).font = fb
ys = [y for y in years if y >= 1953]
for i, y in enumerate(ys):
    r = base + 1 + i; sr = first + years.index(y)
    ws7.cell(r, 1, str(y))
    for j, (nm, col) in enumerate(cats, 2):
        if nm == "Export duties: coffee":
            f = f"=IF(ISNUMBER('{REV}'!I{sr}),'{REV}'!I{sr},'{REV}'!H{sr})"
        elif nm == "Export duties: other":
            f = f"=IF(ISNUMBER('{REV}'!I{sr}),'{REV}'!J{sr},0)"
        elif nm.startswith("Services"):
            f = f"='{REV}'!L{sr}+'{REV}'!M{sr}"
        else:
            f = f"='{REV}'!{col}{sr}"
        c = ws7.cell(r, j, f); c.font = fgr; c.number_format = num
endr = base + len(ys)
ch = BarChart(); ch.type = "col"; ch.grouping = "percentStacked"; ch.overlap = 100
ch.title = "Revenue composition, 1953–1979 (% of current revenue)"
ch.y_axis.title = "% of current revenue"; ch.y_axis.number_format = "0%"
ch.add_data(Reference(ws7, min_col=2, max_col=1 + len(cats), min_row=base, max_row=endr), titles_from_data=True)
ch.set_categories(Reference(ws7, min_col=1, min_row=base + 1, max_row=endr))
ch.height = 10; ch.width = 26; ch.legend.position = "b"
ws7.add_chart(ch, "A3")
# coffee chart data 1960-1979
cb = base; cc = 12
for j, h in enumerate(["Year", "Coffee duty (₡M)", "Coffee price (US$/kg)"], cc):
    ws7.cell(cb, j, h).font = fb
cys = list(range(1960, 1980))
for i, y in enumerate(cys):
    r = cb + 1 + i; ir = 4 + years.index(y)
    ws7.cell(r, cc, str(y))
    if "cof" in D[y]:
        c = ws7.cell(r, cc + 1, f"=Indicators!C{ir}"); c.font = fgr; c.number_format = num
    c = ws7.cell(r, cc + 2, f"=Indicators!B{ir}"); c.font = fgr; c.number_format = "0.000"
ce = cb + len(cys)
b1 = BarChart(); b1.title = "Coffee duty and the coffee price, 1960–1979"; b1.y_axis.title = "Coffee duty (₡M)"
b1.add_data(Reference(ws7, min_col=cc + 1, min_row=cb, max_row=ce), titles_from_data=True)
b1.set_categories(Reference(ws7, min_col=cc, min_row=cb + 1, max_row=ce))
l1 = LineChart(); l1.add_data(Reference(ws7, min_col=cc + 2, min_row=cb, max_row=ce), titles_from_data=True)
l1.y_axis.axId = 200; l1.y_axis.title = "Coffee price (US$/kg)"; l1.y_axis.crosses = "max"
l1.series[0].smooth = False
l1.y_axis.majorGridlines = None
b1 += l1; b1.height = 10; b1.width = 26; b1.legend.position = "b"
ws7.add_chart(b1, "A25")
# income tax chart data 1964-1979
ib = base; icn = 17
for j, h in enumerate(["Year", "Personal income tax", "Corporate income tax"], icn):
    ws7.cell(ib, j, h).font = fb
iys = list(range(1964, 1980))
for i, y in enumerate(iys):
    r = ib + 1 + i; sr = first + years.index(y)
    ws7.cell(r, icn, str(y))
    for k, col in enumerate(["C", "D"]):
        if ("pers" if col == "C" else "corp") in D[y]:
            c = ws7.cell(r, icn + 1 + k, f"='{REV}'!{col}{sr}"); c.font = fgr; c.number_format = num
ie = ib + len(iys)
l2 = LineChart(); l2.display_blanks = "gap"; l2.title = "Personal and corporate income tax, 1964–1979 (₡M; 1971–72 not available)"
l2.add_data(Reference(ws7, min_col=icn + 1, max_col=icn + 2, min_row=ib, max_row=ie), titles_from_data=True)
l2.set_categories(Reference(ws7, min_col=icn, min_row=ib + 1, max_row=ie)); l2.y_axis.title = "₡M"
l2.height = 10; l2.width = 18; l2.legend.position = "b"
for sr_ in l2.series: sr_.smooth = False
ws7.add_chart(l2, "P3")
# total revenue line
tb = base; tc = 21
for j, h in enumerate(["Year", "Current revenue (₡M)", "Tax revenue (₡M)"], tc):
    ws7.cell(tb, j, h).font = fb
for i, y in enumerate(years):
    r = tb + 1 + i; sr = first + i
    ws7.cell(r, tc, str(y))
    c = ws7.cell(r, tc + 1, f"='{REV}'!S{sr}"); c.font = fgr; c.number_format = num
    c = ws7.cell(r, tc + 2, f"='{REV}'!O{sr}"); c.font = fgr; c.number_format = num
te = tb + len(years)
l3 = LineChart(); l3.title = "Current and tax revenue, 1952–1979 (₡M, current prices)"
l3.add_data(Reference(ws7, min_col=tc + 1, max_col=tc + 2, min_row=tb, max_row=te), titles_from_data=True)
l3.set_categories(Reference(ws7, min_col=tc, min_row=tb + 1, max_row=te)); l3.y_axis.title = "₡M"
l3.height = 10; l3.width = 18; l3.legend.position = "b"
for sr_ in l3.series: sr_.smooth = False
ws7.add_chart(l3, "P25")
ws7.cell(base - 1, 1, "Chart data (linked). In the composition chart, 1971–72 export duties are shown as coffee because the coffee duty is not separately available for those years.").font = fb

for s in wb.worksheets:
    for row in s.iter_rows():
        for c in row:
            if c.value is not None and c.font.name != F:
                f0 = c.font
                c.font = Font(name=F, bold=f0.bold, italic=f0.italic, color=f0.color.rgb if f0.color is not None and isinstance(f0.color.rgb, str) else None, size=f0.size)
wb.save(OUT)
print("saved", OUT)
