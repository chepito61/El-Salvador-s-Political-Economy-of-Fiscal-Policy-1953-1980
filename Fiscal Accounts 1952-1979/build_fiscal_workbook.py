# Builds the combined fiscal workbook: the revenue workbook plus expenditure, interest, debt and analytics sheets.
# Run from this folder after the revenue workbook has been rebuilt and recalculated. Then recalculate this file.
import csv
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment

SRC_WB = "../Revenue Data 1952-1990/El_Salvador_Revenue_Memorias_1952_1979.xlsx"
OUT = "El_Salvador_Fiscal_Accounts_Memorias_1952_1979.xlsx"
YRS = list(range(1952, 1980))

# ---------------- inputs (₡ millions), from the yearly notes ----------------
CU = {1952: 131.24, 1953: 140.30, 1954: 155.71, 1955: 155.67, 1956: 166.44, 1957: 186.29, 1958: 184.67, 1959: 175.81, 1960: 174.26,
      1961: 182.48, 1962: 180.02, 1963: 186.25, 1964: 204.6, 1965: 239.0, 1966: 250.1, 1967: 236.64, 1968: 234.39, 1969: 291.89,
      1970: 303.3, 1971: 356.1, 1972: 383.71, 1973: 449.96, 1974: 594.50, 1975: 709.61, 1976: 1028.08, 1977: 1241.94, 1978: 1334.64, 1979: 1453.35}
CU_SRC = {**{y: "Note_MH (comparative table, 1952–56: Note_MH_1955 Table 9)" for y in range(1952, 1957)},
          1957: "Note_MH_1957", 1958: "Note_MH_1958", 1959: "Note_MH_1959", 1960: "Note_MH_1961 comparative", 1961: "Note_MH_1961", 1962: "Note_MH_1962",
          **{y: "Note_MH_1971 comparative, 1963–71" for y in range(1963, 1970)}, **{y: "Note_MH_1979 comparative, 1970–79" for y in range(1970, 1980)}}
GA = {1952: 112.07, 1953: 124.05, 1954: 142.43, 1955: 153.35, 1956: 152.97, 1963: 176.98, 1964: 192.33, 1965: 222.86, 1966: 232.11, 1969: 247.22,
      1970: 269.52, 1971: 294.55, 1972: 336.25, 1973: 401.62, 1974: 509.21, 1975: 599.57, 1976: 805.99, 1977: 915.31, 1978: 1095.03, 1979: 1169.92}
UND = {1963: -5.7, 1964: 5.9, 1965: -14.4, 1966: -26.6, 1967: -9.7, 1968: -3.4, 1969: -8.2, 1970: -5.6, 1971: -19.4, 1972: -15.4, 1973: -8.7,
       1974: -42.99, 1975: -59.60, 1976: -91.1, 1977: 98.5, 1978: -226.21, 1979: -112.56}
INT = {1953: 1.057, 1954: 0.949, 1955: 0.799, 1957: 1.658, 1958: 1.915, 1959: 1.957, 1960: 1.767, 1961: 1.623, 1962: 1.659, 1963: 1.795,
       1964: 2.375, 1965: 2.549, 1966: 2.307, 1967: 1.924, 1969: 3.06, 1970: 4.18, 1971: 5.32, 1972: 7.42, 1973: 7.77, 1974: 8.88,
       1975: 12.28, 1976: 14.44, 1977: 17.70, 1978: 22.27, 1979: 25.83}
INT_NOTE = {1953: "Interest paid on all public debt (MH 1954)", 1954: "Same", 1955: "Debt managed through the BCR only (MH 1955); understated",
            **{y: "Interest paid, debt-service statement (Estado 20 / 22 / VII-2)" for y in range(1957, 1968)},
            1969: "Economic classification (MH 1969 p. 42)", **{y: "Economic classification, credits used (MH 1972 Cuadro 16)" for y in (1970, 1971, 1972)},
            1973: "Interest and other debt charges, credits used (MH 1974 Cuadro 15)", **{y: "Economic classification, gastos (Cuadro 15)" for y in range(1974, 1980)}}
AMO = {1953: 5.131, 1957: 1.673, 1958: 1.743, 1959: 6.708, 1960: 5.837, 1961: 3.987, 1962: 5.644, 1963: 5.475, 1964: 6.225, 1965: 7.898,
       1966: 7.685, 1967: 2.984, 1970: 4.19, 1971: 4.92, 1972: 11.50, 1973: 13.23, 1974: 16.30, 1975: 20.10, 1976: 23.47, 1977: 29.32, 1978: 33.01, 1979: 28.84}
DSERV = {1952: 4.661, 1953: 4.80, 1954: 4.727, 1955: 10.242, 1956: 3.05, 1968: 5.26}  # debt service, not split
DEBT = {1952: 30.6, 1953: 28.4, 1954: 30.3, 1955: 23.9, 1956: 32.6, 1957: 42.57, 1958: 45.80, 1959: 39.61, 1960: 35.6, 1961: 36.1, 1962: 36.3,
        1963: 48.8, 1964: 58.9, 1965: 74.9, 1966: 82.6, 1967: 88.9, 1968: 99.7, 1969: 125.6, 1970: 137.6, 1971: 170.2, 1972: 205.7, 1973: 258.3,
        1974: 323.2, 1975: 397.3, 1976: 435.5, 1977: 466.2, 1978: 638.7, 1979: 857.7}
SF = {1952: 4.06, 1953: 5.47, 1954: 20.02, 1955: 35.08, 1956: 39.95, 1957: 45.72, 1958: 31.20, 1959: 18.03, 1960: 20.32, 1961: 0.50, 1962: -6.93,
      1963: 0.48, 1964: 13.9, 1965: 17.5, 1966: 0.0, 1967: 0.02, 1968: 0.1, 1969: -9.9, 1970: 0.4, 1971: 5.0, 1972: 11.8, 1973: 39.1, 1974: 42.43,
      1975: 101.60, 1976: 61.2, 1977: 75.5, 1978: 10.7, 1979: 97.0}

# pillars by ramo: choose one row per year (preferred basis)
PREF = {1963: "credits", 1966: "gastos", 1967: "credits", 1969: "credits", 1973: "credits"}
PIL = {}
for f in ["../Spending Data/pillars_1952_1960.csv", "../Spending Data/pillars_1961_1979.csv"]:
    for x in csv.DictReader(open(f, encoding="utf-8-sig", newline="")):
        b = x.get("basis", "") or "credits used"
        ys = [1955, 1956] if x["year"].startswith("1955-56") else [int(x["year"])]
        for y in ys:
            want = PREF.get(y)
            if y in PIL and want:
                if want == "credits" and not b.startswith("credits"): continue
                if want == "gastos" and not b.startswith("gastos"): continue
            elif y in PIL:
                continue
            PIL[y] = {k: float(x[k]) for k in ("tot", "Def", "Seg", "Cul", "Sal", "OP", "Agr", "Gob")}
            PIL[y]["basis"] = ("credits used" if b.startswith("credits") else "gastos" if b.startswith("gastos") else b.split(" (")[0]) + \
                              (" (1955–56 average)" if x["year"].startswith("1955-56") else "")
# comparability with the previous year (basis switches)
BREAK = {1955, 1956, 1957, 1964, 1967, 1970, 1973}

# economic classification
ECO_COLS = ["Wages and salaries", "Goods and services", "Interest", "Current transfers (incl. pensions)", "Real investment",
            "Financial investment (on-lending)", "Capital transfers", "Amortization", "Debt service, not split", "Commitments, not classified"]
ECO = {
 1952: dict(w=53.966, g=6.162 + 11.353, t=23.599 + 5.168, r=3.320 + 2.129, ds=4.661, tot=112.074, b="gastos, object classes (Estado 6)"),
 1953: dict(w=53.993 + 7.731, g=5.289 + 8.596, t=26.354 + 5.341, r=5.476 + 6.382, ds=4.800, tot=124.054, b="gastos, object classes"),
 1954: dict(w=59.744 + 10.234, g=5.806 + 9.317, t=21.698 + 7.909, r=7.140 + 15.658, ds=4.727, tot=142.427, b="gastos, object classes"),
 1955: dict(w=62.271 + 11.399, g=7.012 + 8.565, t=24.724 + 9.179, r=3.623 + 14.789, ds=10.242, tot=153.347, b="gastos, object classes"),
 1956: dict(w=63.910 + 10.457, g=6.800 + 6.557, t=30.404 + 10.488, r=2.473 + 16.884, ds=3.050, tot=152.970, b="gastos, object classes"),
 1970: dict(w=131.98, g=26.00, i=4.18, t=18.54 + 60.37, r=34.19, f=0.79, k=13.05, a=4.19, c=5.70, tot=301.19, b="credits used, MH 1972 Cuadro 16"),
 1971: dict(w=140.99, g=33.35, i=5.32, t=20.73 + 45.27, r=61.40, f=0.89, k=15.81, a=4.92, c=24.67, tot=355.84, b="credits used, MH 1972 Cuadro 16"),
 1972: dict(w=157.17, g=39.99, i=7.42, t=20.63 + 48.26, r=28.92, f=11.60, k=8.98, a=11.50, c=47.46, tot=383.71, b="credits used, MH 1972 Cuadro 16"),
 1974: dict(w=202.22, g=55.84, i=8.88, t=84.84 + 22.24, r=47.18, f=35.22, a=16.30, tot=509.21, b="gastos, Cuadro 15 (capital transfers and transfers abroad not shown: in 'other')"),
 1975: dict(w=228.75, g=68.72, i=12.28, t=135.48, r=61.55, f=29.89, k=42.79, a=20.10, tot=599.57, b="gastos, Cuadro 15"),
 1976: dict(w=298.74, g=89.44, i=14.44, t=144.76, r=102.35, f=16.49, k=116.30, a=23.47, tot=805.99, b="gastos, Cuadro 15"),
 1977: dict(w=348.98, g=115.45, i=17.70, t=174.51, r=120.39, f=8.68, k=100.28, a=29.32, tot=915.31, b="gastos, Cuadro 15"),
 1978: dict(w=411.22, g=130.82, i=22.27, t=255.35, r=117.18, f=2.19, k=123.00, a=33.01, tot=1095.03, b="gastos, Cuadro 15"),
 1979: dict(w=475.17, g=129.70, i=25.83, t=207.47, r=159.16, f=12.65, k=131.10, a=28.84, tot=1169.92, b="gastos, Cuadro 15"),
}
EKEY = ["w", "g", "i", "t", "r", "f", "k", "a", "ds", "c"]

PER = [("Osorio", 1953, 1956), ("Lemus", 1957, 1960), ("Junta/Directorio", 1961, 1961), ("Rivera", 1962, 1967),
       ("Sánchez Hernández", 1968, 1972), ("Molina", 1973, 1977), ("Romero", 1978, 1979)]

# ---------------- styles ----------------
F = "Arial"
fb = Font(name=F, bold=True); fn = Font(name=F); fin = Font(name=F, color="0000FF"); fgr = Font(name=F, color="008000")
fi = Font(name=F, italic=True, color="666666"); ft = Font(name=F, bold=True, size=13)
hdr = PatternFill("solid", fgColor="DDE4EE"); tot = PatternFill("solid", fgColor="F3F3F3"); brk = PatternFill("solid", fgColor="FFF2CC")
num = '#,##0.00;(#,##0.00);"-"'; pct = '0.0%;(0.0%);"-"'
def header(ws, row, labels, start=1, height=45):
    for j, h in enumerate(labels, start):
        c = ws.cell(row, j, h); c.font = fb; c.fill = hdr; c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.row_dimensions[row].height = height
def put(ws, r, c, v, font=fn, fmt=num):
    cell = ws.cell(r, c, v); cell.font = font
    if fmt: cell.number_format = fmt
    return cell

wb = load_workbook(SRC_WB)
REV = "Revenue_CM"; rr = lambda y: 4 + (y - 1952)

# ================= Fiscal_Accounts =================
fa = wb.create_sheet("Fiscal_Accounts", index=1)
fa.cell(1, 1, "Central government fiscal accounts, ₡ millions — Memorias de Hacienda, 1952–1979 (credits used basis)").font = ft
fa.cell(2, 1, "Blue = from the yearly notes; green = linked to the revenue sheets; black = formula. Underlying result = commitment-basis result "
              "excluding loan-financed capital commitments and other adjustments (Workflow conventions); before 1963 no adjustments are recorded.").font = fi
H = ["Year", "Current revenue", "Tax revenue", "Coffee duty", "Credits used", "Gastos", "Commitments (credits used − gastos)",
     "Commitment-basis result", "Adjustments (from notes)", "Underlying result", "Interest", "Primary underlying result",
     "Amortization", "Debt service, not split", "Gross financing need", "Direct public debt, end-year", "Superávit Financiero",
     "Underlying result / revenue", "Primary result / revenue", "Interest / revenue", "Debt / revenue", "Implicit interest rate",
     "Credits used growth", "Revenue growth", "Source of credits used", "Interest basis"]
header(fa, 4, H)
FR = lambda y: 5 + (y - 1952)
for y in YRS:
    r = FR(y); s = rr(y)
    fa.cell(r, 1, str(y)).font = fb
    put(fa, r, 2, f"='{REV}'!S{s}", fgr); put(fa, r, 3, f"='{REV}'!O{s}", fgr)
    put(fa, r, 4, f"=IF(ISNUMBER('{REV}'!I{s}),'{REV}'!I{s},\"\")", fgr)
    put(fa, r, 5, CU[y], fin)
    if y in GA: put(fa, r, 6, GA[y], fin); put(fa, r, 7, f"=E{r}-F{r}")
    put(fa, r, 8, f"=B{r}-E{r}")
    put(fa, r, 9, f"=J{r}-H{r}" if False else (None), fin)
    if y in UND:
        # adjustment = underlying (note) − commitment result (sheet); entered as a value so the underlying result is a formula
        fa.cell(r, 9).value = None
    put(fa, r, 10, f"=H{r}+I{r}", fb)
    if y in INT: put(fa, r, 11, INT[y], fin)
    put(fa, r, 12, f"=IF(ISNUMBER(K{r}),J{r}+K{r},\"\")")
    if y in AMO: put(fa, r, 13, AMO[y], fin)
    if y in DSERV: put(fa, r, 14, DSERV[y], fin)
    put(fa, r, 15, f"=IF(OR(ISNUMBER(M{r}),ISNUMBER(N{r})),-J{r}+IF(ISNUMBER(M{r}),M{r},N{r}),\"\")")
    put(fa, r, 16, DEBT[y], fin); put(fa, r, 17, SF[y], fin)
    put(fa, r, 18, f"=J{r}/B{r}", fmt=pct); put(fa, r, 19, f"=IF(ISNUMBER(L{r}),L{r}/B{r},\"\")", fmt=pct)
    put(fa, r, 20, f"=IF(ISNUMBER(K{r}),K{r}/B{r},\"\")", fmt=pct); put(fa, r, 21, f"=P{r}/B{r}", fmt=pct)
    if y > 1952:
        put(fa, r, 22, f"=IF(ISNUMBER(K{r}),K{r}/P{r-1},\"\")", fmt=pct)
        put(fa, r, 23, f"=E{r}/E{r-1}-1", fmt=pct); put(fa, r, 24, f"=B{r}/B{r-1}-1", fmt=pct)
    fa.cell(r, 25, CU_SRC[y]).font = fi; fa.cell(r, 26, INT_NOTE.get(y, "n.a." if y not in DSERV else "Only total debt service available")).font = fi
LASTF = FR(1979)
# fill adjustments after reading revenue values: use the values printed in the notes' revenue (= Revenue_CM S, current revenue)
wbv = load_workbook(SRC_WB, data_only=True); S = {y: wbv[REV][f"S{rr(y)}"].value for y in YRS}
for y, u in UND.items():
    adj = round(u - (S[y] - CU[y]), 2)
    if abs(adj) >= 0.1:
        c = put(fa, FR(y), 9, adj, fin)
        c.comment = Comment(f"Underlying result in the notes: {u:+.2f}. Adjustment = underlying − commitment-basis result.", "Claude")
for y, t in {1953: "Revenue_CM uses ₡140.31M for 1953 (revised); the 1953 note used ₡137.61M (result −2.68).",
             1955: "1955 and 1956 come from the 1955–56 volumes; ministry shares use the 1955–56 average.",
             1971: "From 1971 loan-financed capital commitments are booked with equal accrued revenue (D.L. 448); the underlying result removes them.",
             1977: "Underlying result positive: ₡105.4M of loan-financed programmes are excluded.",
             1978: "Corrected underlying result (−226.21); see dated correction in the 1978 notes."}.items():
    fa.cell(FR(y), 8).comment = Comment(t, "Claude")
n0 = LASTF + 2
for k, t in enumerate([
    "Interest: before 1968 from the debt-service statements (direct debt); from 1969 from the economic classification (1970–73 credits used; 1974–79 gastos). 1952, 1956 and 1968: only total debt service is available.",
    "Gross financing need = −underlying result + amortization (or total debt service where interest and amortization are not split).",
    "Debt: direct public debt at year-end, as in the notes (1950s: total public debt).",
    "Superávit Financiero: as reported; from 1974 it includes authorized but unsold bonds booked as revenue receivable.",
    "Credits used before 1960 are as printed in each year's Memoria; the 1958 Memoria restated earlier years (e.g. 1957: 183.0) — not used here."]):
    fa.cell(n0 + k, 1, t).font = fi
fa.freeze_panes = "B5"; fa.column_dimensions["A"].width = 7
for j in range(2, 25): fa.column_dimensions[L(j)].width = 11.5
fa.column_dimensions["Y"].width = 40; fa.column_dimensions["Z"].width = 50

# ================= Exp_Pillars =================
ep = wb.create_sheet("Exp_Ministry_Pillars", index=2)
ep.cell(1, 1, "Spending by ministry and pillar, ₡ millions, 1952–1979").font = ft
ep.cell(2, 1, "One basis per year (column M); yellow rows change basis against the previous year and are excluded from year-on-year analytics. "
              "Security = Defensa + Seguridad Pública (police inside Defensa in 1962–69). Development = Educación + Salud + Obras Públicas.").font = fi
H2 = ["Year", "Defensa", "Seguridad Pública", "Educación", "Salud Pública", "Obras Públicas", "Agricultura", "Interior (Gobernación)",
      "Other ministries and general obligations", "Total (this basis)", "Security", "Development", "Basis", "Comparable with previous year",
      "Security share", "Development share", "Education share", "Agriculture share", "Development / security"]
header(ep, 4, H2)
PR = lambda y: 5 + (y - 1952)
for y in YRS:
    r = PR(y); p = PIL[y]
    ep.cell(r, 1, str(y)).font = fb
    for j, k in enumerate(["Def", "Seg", "Cul", "Sal", "OP", "Agr", "Gob"], 2): put(ep, r, j, round(p[k], 3), fin)
    put(ep, r, 10, round(p["tot"], 3), fin); put(ep, r, 9, f"=J{r}-SUM(B{r}:H{r})")
    put(ep, r, 11, f"=B{r}+C{r}", fb); put(ep, r, 12, f"=D{r}+E{r}+F{r}", fb)
    ep.cell(r, 13, p["basis"]).font = fi
    ep.cell(r, 14, "No" if (y in BREAK or y == 1952) else "Yes").font = fn
    if y in BREAK:
        for j in range(1, 20): ep.cell(r, j).fill = brk
    for j, f in enumerate([f"=K{r}/J{r}", f"=L{r}/J{r}", f"=D{r}/J{r}", f"=G{r}/J{r}", f"=L{r}/K{r}"], 15):
        put(ep, r, j, f, fmt=pct if j < 19 else "0.00")
LASTP = PR(1979)
ep.freeze_panes = "B5"
for j in range(2, 20): ep.column_dimensions[L(j)].width = 11.5
ep.column_dimensions["M"].width = 26

# ================= Exp_Economic =================
ee = wb.create_sheet("Exp_Economic", index=3)
ee.cell(1, 1, "Spending by economic class, ₡ millions — years available (1952–56 object classes; 1970–72 Cuadro 16; 1974–79 Cuadro 15)").font = ft
ee.cell(2, 1, "Not yet available: 1957–69 and 1973 (to extract: MH 1957–59 object classes; MH 1960–69 economic classification). "
              "'Other / not classified' = total minus the classes, so each row ties to the printed total.").font = fi
H3 = ["Year"] + ECO_COLS + ["Other / not classified", "Total (as printed)", "Basis", "Wages share", "Interest share", "Real investment share", "Transfers share"]
header(ee, 4, H3)
er = 5; EROW = {}
for y in sorted(ECO):
    e = ECO[y]; EROW[y] = er
    ee.cell(er, 1, str(y)).font = fb
    for j, k in enumerate(EKEY, 2):
        if k in e: put(ee, er, j, round(e[k], 3), fin)
    put(ee, er, 12, f"=M{er}-SUM(B{er}:K{er})"); put(ee, er, 13, e["tot"], Font(name=F, bold=True, color="0000FF"))
    ee.cell(er, 14, e["b"]).font = fi
    for j, c in enumerate(["B", "D", "F", "E"], 15): put(ee, er, j, f"={c}{er}/M{er}", fmt=pct)
    er += 1
ee.freeze_panes = "B5"
for j in range(2, 20): ee.column_dimensions[L(j)].width = 12
ee.column_dimensions["N"].width = 44

# ================= Periods_Fiscal =================
pf = wb.create_sheet("Periods_Fiscal", index=4)
pf.cell(1, 1, "Fiscal indicators by government (weighted: period sums)").font = ft
pf.cell(2, 1, "Years assigned to the president inaugurated the previous July. Shares of ministries use the basis of each year (Exp_Ministry_Pillars). "
              "Interest uses only the years where it is available (hover for the years).").font = fi
header(pf, 4, ["Indicator"] + [f"{n}\n{a}–{b}" if a != b else f"{n}\n{a}" for n, a, b in PER], height=32)
def rngF(c, a, b): return f"Fiscal_Accounts!{c}{FR(a)}:{c}{FR(b)}"
def rngP(c, a, b): return f"Exp_Ministry_Pillars!{c}{PR(a)}:{c}{PR(b)}"
IND = [
 ("Revenue, average annual growth", lambda a, b: f"=(Fiscal_Accounts!B{FR(b)}/Fiscal_Accounts!B{FR(a-1)})^(1/{b-a+1})-1", pct),
 ("Credits used, average annual growth", lambda a, b: f"=(Fiscal_Accounts!E{FR(b)}/Fiscal_Accounts!E{FR(a-1)})^(1/{b-a+1})-1", pct),
 ("Commitment-basis result / revenue", lambda a, b: f"=SUM({rngF('H',a,b)})/SUM({rngF('B',a,b)})", pct),
 ("Underlying result / revenue", lambda a, b: f"=SUM({rngF('J',a,b)})/SUM({rngF('B',a,b)})", pct),
 ("Interest / revenue", None, pct),
 ("Debt / revenue, end of period", lambda a, b: f"=Fiscal_Accounts!P{FR(b)}/Fiscal_Accounts!B{FR(b)}", pct),
 ("Debt, change over the period (₡M)", lambda a, b: f"=Fiscal_Accounts!P{FR(b)}-Fiscal_Accounts!P{FR(a-1)}", num),
 ("Security share of spending", lambda a, b: f"=SUM({rngP('K',a,b)})/SUM({rngP('J',a,b)})", pct),
 ("Development share of spending", lambda a, b: f"=SUM({rngP('L',a,b)})/SUM({rngP('J',a,b)})", pct),
 ("Education share of spending", lambda a, b: f"=SUM({rngP('D',a,b)})/SUM({rngP('J',a,b)})", pct),
 ("Agriculture share of spending", lambda a, b: f"=SUM({rngP('G',a,b)})/SUM({rngP('J',a,b)})", pct),
 ("Coffee duty / revenue", None, pct),
]
for i, (lab, f, fmt) in enumerate(IND):
    r = 5 + i; pf.cell(r, 1, lab).font = fn
    for j, (n, a, b) in enumerate(PER, 2):
        if lab == "Interest / revenue":
            yy = [y for y in range(a, b + 1) if y in INT]
            if not yy: pf.cell(r, j, "n.a.").font = fi; continue
            c = put(pf, r, j, "=(" + "+".join(f"Fiscal_Accounts!K{FR(y)}" for y in yy) + ")/(" + "+".join(f"Fiscal_Accounts!B{FR(y)}" for y in yy) + ")", fmt=pct)
            if len(yy) < b - a + 1: c.comment = Comment("Years used: " + ", ".join(map(str, yy)), "Claude"); c.font = Font(name=F, italic=True)
        elif lab == "Coffee duty / revenue":
            yy = [y for y in range(a, b + 1) if y not in (1952, 1971, 1972)]
            c = put(pf, r, j, "=(" + "+".join(f"Fiscal_Accounts!D{FR(y)}" for y in yy) + ")/(" + "+".join(f"Fiscal_Accounts!B{FR(y)}" for y in yy) + ")", fmt=pct)
            if len(yy) < b - a + 1: c.comment = Comment("Years used: " + ", ".join(map(str, yy)), "Claude"); c.font = Font(name=F, italic=True)
        else:
            put(pf, r, j, f(a, b), fmt=fmt)
pf.column_dimensions["A"].width = 38
for j in range(2, 9): pf.column_dimensions[L(j)].width = 14

# ================= Analytics: booms and busts =================
an = wb.create_sheet("Analytics_Cycle", index=5)
an.cell(1, 1, "The coffee cycle and spending: year-on-year changes, ₡ millions").font = ft
an.cell(2, 1, "Boom = coffee duty up 10% or more; bust = down 10% or more (1971–72: export duties). Years where the ministry basis changes are excluded "
              "from the ministry columns (marked No). Thresholds are inputs (cells C3 and E3).").font = fi
an.cell(3, 1, "Boom threshold").font = fb; put(an, 3, 3, 0.10, fin, pct); an.cell(3, 4, "Bust threshold").font = fb; put(an, 3, 5, -0.10, fin, pct)
H4 = ["Year", "Coffee duty (or exports)", "Change in coffee duty", "Phase", "Δ Revenue", "Δ Credits used", "Δ Interest",
      "Comparable ministry basis", "Δ Security", "Δ Development", "Δ Education", "Δ Other spending", "Δ Underlying result"]
header(an, 5, H4)
AR = lambda y: 6 + (y - 1954)
AY = list(range(1954, 1980))
for y in AY:
    r = AR(y); f, p, fp = FR(y), PR(y), PR(y - 1)
    an.cell(r, 1, str(y)).font = fb
    s = rr(y)
    put(an, r, 2, f"=IF(ISNUMBER('{REV}'!I{s}),'{REV}'!I{s},'{REV}'!H{s})", fgr)
    if y in (1971, 1973):  # comparisons across the gap use export duties
        put(an, r, 3, f"='{REV}'!H{s}/'{REV}'!H{s-1}-1", fmt=pct)
        an.cell(r, 3).comment = Comment("Export duties used (coffee duty not available in 1971–72).", "Claude")
    else:
        put(an, r, 3, f"=B{r}/B{r-1}-1" if y > 1954 else f"=B{r}/'{REV}'!I{s-1}-1", fmt=pct)
    put(an, r, 4, f'=IF(C{r}>=$C$3,"Boom",IF(C{r}<=$E$3,"Bust","Neutral"))', fmt=None)
    put(an, r, 5, f"=Fiscal_Accounts!B{f}-Fiscal_Accounts!B{f-1}")
    put(an, r, 6, f"=Fiscal_Accounts!E{f}-Fiscal_Accounts!E{f-1}")
    if y in INT and (y - 1) in INT and y not in (1955, 1974): put(an, r, 7, f"=Fiscal_Accounts!K{f}-Fiscal_Accounts!K{f-1}")
    comp = "No" if y in BREAK else "Yes"
    an.cell(r, 8, comp).font = fn
    if comp == "Yes":
        put(an, r, 9, f"=Exp_Ministry_Pillars!K{p}-Exp_Ministry_Pillars!K{fp}")
        put(an, r, 10, f"=Exp_Ministry_Pillars!L{p}-Exp_Ministry_Pillars!L{fp}")
        put(an, r, 11, f"=Exp_Ministry_Pillars!D{p}-Exp_Ministry_Pillars!D{fp}")
        put(an, r, 12, f"=(Exp_Ministry_Pillars!J{p}-Exp_Ministry_Pillars!J{fp})-I{r}-J{r}")
    else:
        for j in range(1, 14): an.cell(r, j).fill = brk
    put(an, r, 13, f"=Fiscal_Accounts!J{f}-Fiscal_Accounts!J{f-1}")
LA = AR(1979); A0 = AR(1954)
# summary
sr = LA + 3
an.cell(sr, 1, "Summary by phase").font = ft
header(an, sr + 1, ["Phase", "Years", "Σ Δ revenue", "Σ Δ credits used", "Σ Δ credits used / Σ Δ revenue (negative: spending rose while revenue fell)",
                    "Years with comparable basis", "Σ Δ security", "Σ Δ development", "Σ Δ education", "Security per colón of total Δ spending",
                    "Development per colón", "Education per colón", "Average Δ underlying result"])
for i, ph in enumerate(["Boom", "Bust", "Neutral"]):
    r = sr + 2 + i; rngD = f"$D${A0}:$D${LA}"; rngH = f"$H${A0}:$H${LA}"
    an.cell(r, 1, ph).font = fb
    put(an, r, 2, f'=COUNTIF({rngD},"{ph}")', fmt="0")
    put(an, r, 3, f'=SUMIF({rngD},"{ph}",$E${A0}:$E${LA})'); put(an, r, 4, f'=SUMIF({rngD},"{ph}",$F${A0}:$F${LA})')
    put(an, r, 5, f"=IF(C{r}<>0,D{r}/C{r},\"\")", fmt="0.00")
    put(an, r, 6, f'=COUNTIFS({rngD},"{ph}",{rngH},"Yes")', fmt="0")
    for j, c in enumerate(["I", "J", "K"], 7):
        put(an, r, j, f'=SUMIFS(${c}${A0}:${c}${LA},{rngD},"{ph}",{rngH},"Yes")')
    tot_ = f'(SUMIFS($I${A0}:$I${LA},{rngD},"{ph}",{rngH},"Yes")+SUMIFS($J${A0}:$J${LA},{rngD},"{ph}",{rngH},"Yes")+SUMIFS($L${A0}:$L${LA},{rngD},"{ph}",{rngH},"Yes"))'
    for j, c in enumerate(["G", "H", "I"], 10):
        put(an, r, j, f"=IF({tot_}<>0,{c}{r}/{tot_},\"\")", fmt="0.00")
    put(an, r, 13, f'=IFERROR(AVERAGEIF({rngD},"{ph}",$M${A0}:$M${LA}),"")')
cr = sr + 6
an.cell(cr, 1, "Correlation of revenue growth and credits-used growth, 1954–79").font = fb
put(an, cr, 6, f"=CORREL(Fiscal_Accounts!X{FR(1954)}:X{FR(1979)},Fiscal_Accounts!W{FR(1954)}:W{FR(1979)})", fmt="0.00")
an.cell(cr + 1, 1, "Correlation of the change in the coffee duty and credits-used growth").font = fb
put(an, cr + 1, 6, f"=CORREL(C{A0}:C{LA},Fiscal_Accounts!W{FR(1954)}:W{FR(1979)})", fmt="0.00")
an.cell(cr + 3, 1, "Reading: in busts, a response per colón above 0 means spending kept growing while revenue fell (deficit financing); "
                   "the boom and bust rows compare where the extra or missing money went.").font = fi
an.column_dimensions["A"].width = 10
for j in range(2, 14): an.column_dimensions[L(j)].width = 13

# ================= Charts_Fiscal =================
cf = wb.create_sheet("Charts_Fiscal", index=6)
cf.cell(1, 1, "Fiscal charts (data linked from row 70)").font = ft
B = 70
for j, h in enumerate(["Year", "Underlying result / revenue", "Interest / revenue", "Debt / revenue", "Security share", "Development share",
                       "Education share", "Revenue", "Credits used", "Coffee duty"], 1):
    cf.cell(B, j, h).font = fb
for i, y in enumerate(YRS):
    r = B + 1 + i; f = FR(y); p = PR(y)
    cf.cell(r, 1, str(y))
    for j, ref in enumerate([f"=Fiscal_Accounts!R{f}", f"=IF(ISNUMBER(Fiscal_Accounts!T{f}),Fiscal_Accounts!T{f},NA())" if False else None,
                             f"=Fiscal_Accounts!U{f}", f"=Exp_Ministry_Pillars!O{p}", f"=Exp_Ministry_Pillars!P{p}", f"=Exp_Ministry_Pillars!Q{p}",
                             f"=Fiscal_Accounts!B{f}", f"=Fiscal_Accounts!E{f}", f"=IF(ISNUMBER(Fiscal_Accounts!D{f}),Fiscal_Accounts!D{f},\"\")"], 2):
        if ref is None:
            if y in INT: ref = f"=Fiscal_Accounts!T{f}"
            else: continue
        c = cf.cell(r, j, ref); c.number_format = pct if j <= 7 else num; c.font = fgr
E = B + len(YRS)
def line(title, cols, anchor, fmt="0%"):
    ch = LineChart(); ch.title = title; ch.display_blanks = "gap"
    ch.add_data(Reference(cf, min_col=cols[0], max_col=cols[-1], min_row=B, max_row=E), titles_from_data=True)
    ch.set_categories(Reference(cf, min_col=1, min_row=B + 1, max_row=E)); ch.y_axis.number_format = fmt
    for s_, colr in zip(ch.series, ["1F4E79", "8C510A", "2E7D32"]):
        s_.smooth = False; s_.graphicalProperties.line.solidFill = colr; s_.graphicalProperties.line.width = 22000
    ch.height = 9; ch.width = 18; ch.legend.position = "b"; cf.add_chart(ch, anchor)
bc = BarChart(); bc.title = "Underlying result as % of revenue, 1952–1979"
bc.add_data(Reference(cf, min_col=2, min_row=B, max_row=E), titles_from_data=True); bc.set_categories(Reference(cf, min_col=1, min_row=B + 1, max_row=E))
bc.y_axis.number_format = "0%"; bc.legend = None; bc.height = 9; bc.width = 18; bc.x_axis.tickLblPos = "low"; cf.add_chart(bc, "A3")
line("Interest and debt as % of revenue (interest gaps = not available)", [3, 4], "L3")
line("Security, development and education shares of spending", [5, 6, 7], "A22")
line("Revenue, credits used and the coffee duty (₡M)", [8, 9, 10], "L22", fmt="#,##0")

# ================= README additions =================
rd = wb["README"]
rd.insert_rows(3, 4)
rd.cell(3, 1, "COMBINED FISCAL WORKBOOK (12 November 2026): the revenue sheets plus Fiscal_Accounts (revenue, credits used, gastos, results, interest, "
              "amortization, debt), Exp_Ministry_Pillars, Exp_Economic, Periods_Fiscal, Analytics_Cycle and Charts_Fiscal.").font = Font(name=F, bold=True)
rd.cell(4, 1, "Expenditure figures come from the verified yearly notes and the Spending Data CSVs; each sheet states its basis. "
              "Built by build_fiscal_workbook.py from the revenue workbook.").font = Font(name=F)
rd.cell(5, 1, "Main gaps: economic classification 1957–69 and 1973; interest 1952, 1956, 1968; amortization 1952–56 and 1968–69; "
              "ministry-by-class detail 1957–59 (to extract).").font = Font(name=F)
for r in (3, 4, 5): rd.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
wb.save(OUT); print("saved", OUT)
