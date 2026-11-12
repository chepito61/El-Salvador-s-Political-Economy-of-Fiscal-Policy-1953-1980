# Run after build_revenue_workbook_1952_1979.py, then recalc, then build_revenue_charts_png.py, then recalc again.
# Adds Revenue_Detail, Periods and Charts_Detail sheets to the workbook built by build_rev.py
import json, sys
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment

SP = ""
OUT = "El_Salvador_Revenue_Memorias_1952_1979.xlsx"
exec(open("revenue_detail_1974_1979.py").read())  # yrs, d
d7479 = d
j5663 = json.load(open("revenue_detail_1956_1963.json"))
csvd = json.load(open("revenue_1956_1963_checked.json"))

# ---------- item data, ₡M ----------
# keys: vial, succ (successions+gifts), alca, pave, land, shrimp, sugx, othx, alc, cig, petr, sugc, bev, sel, cot, othc, timb, sell
X = {y: {} for y in range(1953, 1980)}
X[1953].update(vial=1.181, succ=3.851)
X[1954].update(vial=2.849, succ=8.817, alc=16.377, cig=6.449)
X[1955].update(vial=2.644, succ=1.648, alc=16.736, cig=6.521, sugc=1.854, bev=0.675)
for ys, v in j5663.items():
    y = int(ys)
    X[y].update(vial=csvd[ys]["vial"], succ=v["succ"], alca=v["alcabala"], shrimp=v["shrimp"] or None, othx=v["othx"] or None,
                alc=v["alcohol"], cig=v["cig"], sugc=v["sugar_c"], bev=v["bev"], cot=v["cotton"] or None,
                petr=v["petrol"] or None, othc=v["flour"] or None, timb=v["timbres"], sell=v["sellado"])
P6469 = {1964: (4.946, 0.168, 1.701, 0.788 + 2.567), 1965: (5.101, 0.238, 2.023, 0.316 + 2.058), 1966: (5.312, 0.209, 2.018, 0.563 + 1.099),
         1967: (5.697, 0.722, 2.054, 0.321 + 2.693), 1968: (15.408, 0.143, 2.089, 0.326 + 3.123), 1969: (17.509, 0.153, 2.416, 0.478 + 4.266)}
E6469 = {1964: (1.072, 0.021), 1965: (0.761, 0.017), 1966: (1.131, 0.034), 1967: (0.891, 0.025), 1968: (0.723, 0.032), 1969: (0.851, 0.028)}
C6469 = {1964: (22.770, 9.547, 3.973, 1.013, 6.493, 1.068 + 0.541 + 0.480 + 0.065),
         1965: (23.250, 10.463, 3.813, 1.110, 13.948, 1.323 + 0.057 + 0.491 + 0.038),
         1966: (23.701, 10.558, 4.101, 1.365, 13.856, 1.193 + 0.541 + 0.019),
         1967: (22.834, 10.937, 4.377, 1.314, 14.952, 1.423 + 0.525 + 0.013),
         1968: (22.649, 10.521, 4.435, 1.189, 16.853, 1.295 + 0.613 + 0.012),
         1969: (24.552, 11.088, 4.992, 1.309, 16.685, 1.426 + 0.603)}
S6469 = {1964: (4.202, 1.582), 1965: (4.649, 1.879), 1966: (4.815, 1.990), 1967: (5.245, 2.026), 1968: (5.653, 2.236), 1969: (5.900, 2.331)}
for y in range(1964, 1970):
    p, e, c, s = P6469[y], E6469[y], C6469[y], S6469[y]
    X[y].update(vial=p[0], pave=p[1], alca=p[2], succ=p[3], shrimp=e[0], othx=e[1],
                alc=c[0], cig=c[1], sugc=c[2], bev=c[3], petr=c[4], othc=c[5], timb=s[0], sell=s[1])
X[1970].update(vial=17.82, succ=7.08, shrimp=1.028, alc=26.53, timb=9.58)
X[1971].update(timb=20.00)
X[1972].update(timb=23.00)
X[1973].update(vial=18.333, pave=0.210, alca=4.613, succ=0.985 + 6.665, shrimp=0.974, othx=1.120, alc=34.552, cig=16.273,
               sugc=6.309, bev=2.105, petr=22.371, othc=0.709 + 0.850, timb=27.692, sell=3.671)
for i, y in enumerate(yrs):
    g = lambda k: d7479[k][i]
    X[y].update(vial=g("vialidad"), pave=g("paving"), land=g("landtax") or None, alca=g("alcabala"), succ=g("gifts") + g("succ"),
                shrimp=g("shrimp"), sugx=g("sugar_x") or None, othx=g("othmerch"), alc=g("alcohol"), cig=g("cig"), petr=g("petrol"),
                sugc=g("sugar_c"), bev=g("bev"), sel=g("selective") or None, othc=g("matches"), timb=g("timbres"), sell=g("sellado"))

SRC = {}
for y in (1953,): SRC[y] = "MH 1954 comparatives (Note_MH_1954): Vialidad and Sucesiones only"
SRC[1954] = "MH 1954 Estado 2 (Note_MH_1954): subclasses 031, 032, 045, 060"
SRC[1955] = "MH 1955-56 Estado 2-A (Note_MH_1955); consumption items from the item file (checked)"
for y in range(1956, 1964): SRC[y] = f"MH {y} item detail (codes 031–037, 045, 051–052, 060, 091), checked against subclass totals"
for y in range(1964, 1970): SRC[y] = f"MH {y}–{min(y+2,1970)} Estado II-5, cross-checked across volumes; items include fines; sub-lines checked to printed subtotals"
SRC[1970] = "MH 1970 Memoria groupings (Note_MH_1970); most items not available"
SRC[1971] = "MH 1972 Cuadro 10; only stamps available"; SRC[1972] = SRC[1971]
SRC[1973] = "MH 1973 Estado I-4, actual column (page image); sub-lines checked to printed subtotals"
for y in range(1974, 1980): SRC[y] = "MH 1978 Estado II-1 (1974) and MH 1979 Estado II-1 (1975–79)"

# ---------- styles ----------
F = "Arial"
fb = Font(name=F, bold=True); fn = Font(name=F); fin = Font(name=F, color="0000FF"); fgr = Font(name=F, color="008000")
fi = Font(name=F, italic=True, color="666666"); ft = Font(name=F, bold=True, size=13); fbi = Font(name=F, bold=True, italic=True)
hdr = PatternFill("solid", fgColor="DDE4EE"); grp = PatternFill("solid", fgColor="EEF1F5"); tot = PatternFill("solid", fgColor="F3F3F3")
thin = Side(style="thin", color="999999")
num = '#,##0.00;(#,##0.00);"-"'; pct = '0.0%;(0.0%);"-"'

wb = load_workbook(OUT)
REV = "Revenue_CM"
rr = lambda y: 4 + (y - 1952)  # row in Revenue_CM

# ---------- column layout ----------
# (key, label, kind) kind: item (entered), link:<col> (from Revenue_CM), res (residual), tot (group total)
G = [
 ("Income tax", [("pers", "Personal", "link:C"), ("corp", "Corporate", "link:D"), ("othi", "Other income", "link:E"),
                 ("resi", "Not itemized", "res"), ("TI", "Income tax total", "tot")]),
 ("Property and wealth", [("vial", "Net wealth tax (Vialidad Serie A)", "item"), ("succ", "Successions and gifts", "item"),
                 ("alca", "Property transfers (alcabala)", "item"), ("pave", "Paving levy", "item"),
                 ("land", "Agricultural land tax", "item"), ("resp", "Not itemized", "res"), ("TP", "Property total", "tot")]),
 ("Import duties", [("TM", "Import duties", "tot")]),
 ("Export duties", [("cof", "Coffee", "link:I"), ("sugx", "Sugar", "item"), ("shrimp", "Shrimp", "item"),
                 ("othx", "Other merchandise", "item"), ("resx", "Not itemized", "res"), ("TX", "Export total", "tot")]),
 ("Consumption of goods (excises)", [("alc", "Alcohol (liquor, beer)", "item"), ("cig", "Cigarettes", "item"),
                 ("petr", "Petroleum products", "item"), ("sugc", "Sugar", "item"), ("bev", "Soft drinks and beverages", "item"),
                 ("sel", "Selective consumption tax", "item"), ("cot", "Cotton (1956–58)", "item"),
                 ("othc", "Flour, matches, cosmetics, other", "item"), ("resc", "Not itemized", "res"), ("TC", "Consumption total", "tot")]),
 ("Stamps, services and other taxes", [("timb", "Stamp duties (timbres)", "item"), ("sell", "Sealed paper", "item"),
                 ("ress", "Services, other taxes and not itemized", "res"), ("TS", "Group total", "tot")]),
 ("Emergency tax", [("TE", "Emergency tax (1969–70)", "tot")]),
 ("Totals", [("TT", "Tax revenue (sum)", "sum"), ("TTp", "Tax revenue (Revenue_CM)", "link:O"), ("CK", "Check", "chk"),
             ("NT", "Non-tax revenue", "link:Q"), ("CR", "Current revenue", "link:S")]),
]
col = {}; c = 2
for gname, items in G:
    for k, lab, kind in items:
        col[k] = c; c += 1
col["src"] = c
CL = {k: L(v) for k, v in col.items()}

ws = wb.create_sheet("Revenue_Detail", index=2)
ws.cell(1, 1, "Detailed tax revenue, ₡ millions (current prices) — Memorias de Hacienda, 1953–1979").font = ft
ws.cell(2, 1, "Blue = entered from the Memoria; green = linked to Revenue_CM; black = formula. 'Not itemized' = group total minus the itemized lines "
              "(it holds items the source does not show separately in that year). Group totals tie to Revenue_CM.").font = fi
ws.cell(4, 1, "Year").font = fb; ws.cell(4, 1).fill = hdr
for gname, items in G:
    c0 = col[items[0][0]]; c1 = col[items[-1][0]]
    ws.cell(3, c0, gname).font = fb
    for cc in range(c0, c1 + 1):
        ws.cell(3, cc).fill = grp; ws.cell(3, cc).border = Border(bottom=thin)
    if c1 > c0: ws.merge_cells(start_row=3, start_column=c0, end_row=3, end_column=c1)
    ws.cell(3, c0).alignment = Alignment(horizontal="center")
    for k, lab, kind in items:
        h = ws.cell(4, col[k], lab); h.font = fb if kind in ("tot", "sum") else fn
        h.fill = hdr; h.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
ws.cell(4, col["src"], "Source of the itemized lines").font = fb; ws.cell(4, col["src"]).fill = hdr
ws.row_dimensions[4].height = 58
first = 5
YRS = list(range(1953, 1980))
R = lambda y: first + (y - 1953)
for y in YRS:
    r = R(y); s = rr(y); x = X[y]
    ws.cell(r, 1, str(y)).font = fb
    def cell(k): return f"{CL[k]}{r}"
    def put(k, v, font=fin):
        cc = ws.cell(r, col[k], v); cc.font = font; cc.number_format = num; return cc
    adj = (y <= 1963 and x.get("alca") is not None)  # 1953–63 alcabala sits in 'other taxes' in Revenue_CM
    for gname, items in G:
        its = [k for k, lab, kind in items if kind in ("item",) or kind.startswith("link:")]
        for k, lab, kind in items:
            if kind == "item":
                v = x.get(k)
                if v is not None: put(k, round(v, 4))
            elif kind.startswith("link:"):
                cl = kind[5:]
                put(k, f"=IF(ISNUMBER('{REV}'!{cl}{s}),'{REV}'!{cl}{s},\"\")", fgr)
            elif kind == "res":
                totk = items[-1][0]
                put(k, f"={cell(totk)}-SUM({CL[its[0]]}{r}:{CL[its[-1]]}{r})" if its else "", fn)
            elif kind == "tot":
                base = {"TI": "B", "TP": "F", "TM": "G", "TX": "H", "TC": "K", "TS": "L", "TE": "M"}[k]
                f = f"='{REV}'!{base}{s}"
                if k == "TP" and adj: f += f"+{cell('alca')}"
                if k == "TS" and adj: f += f"-{cell('alca')}"
                if k == "TE": f = f"=IF(ISNUMBER('{REV}'!M{s}),'{REV}'!M{s},0)"
                cc = put(k, f, fgr); cc.font = Font(name=F, bold=True, color="008000"); cc.fill = tot
            elif kind == "sum":
                cc = put(k, "=" + "+".join(cell(t) for t in ("TI", "TP", "TM", "TX", "TC", "TS", "TE")), fb); cc.fill = tot
            elif kind == "chk":
                put(k, f"={cell('TT')}-{cell('TTp')}", fn)
    ws.cell(r, col["src"], SRC.get(y, "")).font = fi
    if adj:
        ws.cell(r, col["alca"]).comment = Comment("1953–63: the Memoria classifies the alcabala under 'Otros Impuestos' (090). "
                                                  "Moved here to property; the 'Stamps, services and other' group is reduced by the same amount.", "Claude")
last = R(1979)
# comments on specific cells
def note(y, k, t): ws.cell(R(y), col[k]).comment = Comment(t, "Claude")
note(1970, "succ", "1970: successions and gifts reported together (₡7.08M) in the Memoria's grouping.")
note(1956, "vial", "1956: Vialidad = Renta y Vialidad (₡15.15M) − income tax proper (₡12.32M), per Note_MH_1955_1956.")
note(1974, "resx", "1974: Revenue_CM uses MH 1974 Cuadro 9 for coffee (₡89.95M); MH 1978 II-1 gives ₡89.93M. The difference shows here.")
note(1975, "sugx", "Progressive sugar export tax, D.L. 135 of 19 Nov 1974.")
note(1978, "othx", "'Other merchandise' export duty: ₡116.6M in 1978. Legal basis still to identify (Diario Oficial).")
note(1956, "cot", "Tax 'Sobre el Algodón' recorded under consumption taxes, 1956–58; no cotton export duty in the Memorias.")
note(1975, "sel", "Impuesto Selectivo al Consumo, shown separately from 1974 in Estado II-1.")
note(1964, "othc", "1964–65 include a tax on cosmetics (₡0.54M in 1964).")
note(1953, "resp", "1953–55: the alcabala is not available; it remains in 'Stamps, services and other'.")

nr = last + 2
notes = [
 "Coverage of the itemized lines: full in 1956–1969 and 1973–1979; partial in 1953–55 (property, and alcohol/cigarettes in 1954–55) and 1970; only stamp duties in 1971–72.",
 "Income tax split: 1964–79 except 1971–72 (statements missing from the scans). 1964–68 include fines; from 1969 fines are excluded.",
 "Successions and gifts: one line, because the 1950s Memorias report them together (060 Sucesiones y Donaciones); from 1964 the two taxes are added.",
 "Stamps: 1953–55 not separated from the other especies fiscales. 1970: sealed paper not available. 1971–72: stamp duties from Cuadro 10, sealed paper not available.",
 "Vialidad Serie A is a tax on net wealth (capital): Ley de Vialidad, D.L. 1308 of 22 Dec 1953; replaced by the Ley del Impuesto sobre el Patrimonio, D.L. 553 of 18 Dec 1986.",
]
for k, t in enumerate(notes):
    ws.cell(nr + k, 1, t).font = fi
ws.freeze_panes = "B5"
ws.column_dimensions["A"].width = 8
for k, cc in col.items():
    ws.column_dimensions[L(cc)].width = 11.5
ws.column_dimensions[CL["src"]].width = 80

# ---------- Periods ----------
PER = [("Osorio", 1953, 1956), ("Lemus", 1957, 1960), ("Junta/Directorio", 1961, 1961), ("Rivera", 1962, 1967),
       ("Sánchez Hernández", 1968, 1972), ("Molina", 1973, 1977), ("Romero", 1978, 1979)]
# composition lines: (label, list of detail keys summed, availability key or None)
COMP = [  # (label, item keys, parent group or None)
 ("Income tax", ["TI"], None),
 ("Net wealth tax (Vialidad A)", ["vial"], "TP"),
 ("Other property taxes", ["TP", "-vial"], "TP"),
 ("Import duties", ["TM"], None),
 ("Export duty: coffee", ["cof"], "TX"),
 ("Other export duties", ["TX", "-cof"], "TX"),
 ("Alcohol and cigarettes", ["alc", "cig"], "TC"),
 ("Other consumption taxes", ["TC", "-alc", "-cig"], "TC"),
 ("Stamps, services, other and emergency", ["TS", "TE"], None),
 ("Non-tax revenue", ["NT"], None),
]
def avail(keys, y):
    for k in keys:
        kk = k.lstrip("-")
        if kk in ("TI", "TP", "TM", "TX", "TC", "TS", "TE", "NT"): continue
        if kk == "cof":
            if y in (1971, 1972): return False
            continue
        if X[y].get(kk) is None: return False
    return True
def segs_of(yy):
    segs = []; s0 = yy[0]; p = yy[0]
    for y in yy[1:] + [None]:
        if y is None or y != p + 1:
            segs.append((s0, p)); s0 = y
        if y is not None: p = y
    return segs
wp = wb.create_sheet("Periods", index=3)
wp.cell(1, 1, "Revenue composition by government, % of current revenue (weighted: period sum of the source / period sum of current revenue)").font = ft
wp.cell(2, 1, "Years are assigned to the president inaugurated the previous July (1 July), e.g. 1962 to Rivera. Italic cells: the split inside a group is missing in some years; "
              "the group's share is for the whole period and is split using the years available (hover for the years).").font = fi
wp.cell(4, 1, "Source").font = fb; wp.cell(4, 1).fill = hdr
for j, (nm, a, b) in enumerate(PER, 2):
    h = wp.cell(4, j, f"{nm}\n{a}–{b}" if a != b else f"{nm}\n{a}"); h.font = fb; h.fill = hdr
    h.alignment = Alignment(wrap_text=True, horizontal="center")
wp.row_dimensions[4].height = 32
DS = "Revenue_Detail"
def rng(k, a, b): return f"'{DS}'!{CL[k]}{R(a)}:{CL[k]}{R(b)}"
def S(keys, segs): return "+".join(("-" if k.startswith("-") else "") + f"SUM({rng(k.lstrip('-'), x0, x1)})" for k in keys for x0, x1 in segs)
for i, (lab, keys, par) in enumerate(COMP):
    r = 5 + i
    wp.cell(r, 1, lab).font = fn
    for j, (nm, a, b) in enumerate(PER, 2):
        full = [(a, b)]
        yy = [y for y in range(a, b + 1) if avail(keys, y)]
        if not yy:
            wp.cell(r, j, "n.a.").font = fi; continue
        if len(yy) == b - a + 1:
            f = f"=({S(keys, full)})/({S(['CR'], full)})"
        else:
            sg = segs_of(yy)
            f = f"=({S(keys, sg)})/({S([par], sg)})*({S([par], full)})/({S(['CR'], full)})"
        cc = wp.cell(r, j, f); cc.number_format = pct
        if len(yy) < b - a + 1:
            cc.comment = Comment("Split estimated from the years available: " + ", ".join(map(str, yy)), "Claude"); cc.font = Font(name=F, italic=True)
rT = 5 + len(COMP)
wp.cell(rT, 1, "Total").font = fb
for j in range(2, 2 + len(PER)):
    cc = wp.cell(rT, j, f"=SUM({L(j)}5:{L(j)}{rT-1})"); cc.number_format = pct; cc.font = fi
# memo lines
memo_r = rT + 2
wp.cell(memo_r, 1, "Memo").font = fb
MEMO = [("Direct taxes (income + property)", "=(SUM({TI})+SUM({TP}))/SUM({CR})"),
        ("Trade taxes (imports + exports)", "=(SUM({TM})+SUM({TX}))/SUM({CR})"),
        ("Petroleum products", None), ("Selective consumption tax", None),
        ("Corporate income tax", None), ("Personal income tax", None),
        ("Current revenue, average annual growth", "growth"), ("Coffee price, average (US$/kg)", "price")]
for i, (lab, f) in enumerate(MEMO):
    r = memo_r + 1 + i
    wp.cell(r, 1, lab).font = fn
    for j, (nm, a, b) in enumerate(PER, 2):
        if f == "growth":
            cc = wp.cell(r, j, f"=('{REV}'!S{rr(b)}/'{REV}'!S{rr(a-1)})^(1/{b-a+1})-1"); cc.number_format = pct
        elif f == "price":
            if b < 1960: wp.cell(r, j, "n.a.").font = fi; continue
            a2 = max(a, 1960)
            cc = wp.cell(r, j, f"=AVERAGE(Indicators!B{4+(a2-1952)}:B{4+(b-1952)})"); cc.number_format = "0.00"
            if a2 != a: cc.comment = Comment(f"Pink Sheet starts in 1960: {a2}–{b} only.", "Claude")
        elif f:
            cc = wp.cell(r, j, f.format(TI=rng("TI", a, b), TP=rng("TP", a, b), TM=rng("TM", a, b), TX=rng("TX", a, b), CR=rng("CR", a, b)))
            cc.number_format = pct
        else:
            k = {"Petroleum products": "petr", "Selective consumption tax": "sel", "Corporate income tax": "corp", "Personal income tax": "pers"}[lab]
            yy = [y for y in range(a, b + 1) if (X[y].get(k) is not None if k in ("petr", "sel") else (y >= 1964 and y not in (1971, 1972)))]
            if not yy: wp.cell(r, j, "n.a.").font = fi; continue
            numr = "+".join(f"'{DS}'!{CL[k]}{R(y)}" for y in yy); den = "+".join(f"'{DS}'!{CL['CR']}{R(y)}" for y in yy)
            cc = wp.cell(r, j, f"=({numr})/({den})"); cc.number_format = pct
            if len(yy) < b - a + 1:
                cc.comment = Comment("Years used: " + ", ".join(map(str, yy)), "Claude"); cc.font = Font(name=F, italic=True)
mend = memo_r + len(MEMO)
wp.cell(mend + 2, 1, "Italic cells use only part of the period (hover for the years). Growth: compound annual growth of current revenue from the year before the period to its last year. "
                     "Osorio's growth starts from 1952 (old classification total).").font = fi
wp.column_dimensions["A"].width = 40
for j in range(2, 2 + len(PER)): wp.column_dimensions[L(j)].width = 14
wp.freeze_panes = "B5"

# chart: one 100% stacked bar per government
pc = BarChart(); pc.type = "col"; pc.grouping = "percentStacked"; pc.overlap = 100
pc.title = "Revenue composition by government (% of current revenue)"
pc.add_data(Reference(wp, min_col=1, max_col=1 + len(PER), min_row=5, max_row=rT - 1), from_rows=True, titles_from_data=True)
pc.set_categories(Reference(wp, min_col=2, max_col=1 + len(PER), min_row=4, max_row=4))
pc.y_axis.number_format = "0%"; pc.height = 11; pc.width = 24; pc.legend.position = "r"
wp.add_chart(pc, f"A{mend + 5}")

# ---------- Charts_Detail: the four recommended charts + detail charts ----------
wc = wb.create_sheet("Charts_Detail")
wc.cell(1, 1, "Charts from the detailed series (chart data from row 80, linked)").font = ft
B = 80; Y = list(range(1953, 1980))
# block 1: direct vs trade shares (cols A-D)
for j, h in enumerate(["Year", "Direct taxes", "Import duties", "Export duties"], 1): wc.cell(B, j, h).font = fb
for i, y in enumerate(Y):
    r = B + 1 + i; wc.cell(r, 1, str(y))
    for j, f in enumerate([f"=('{DS}'!{CL['TI']}{R(y)}+'{DS}'!{CL['TP']}{R(y)})/'{DS}'!{CL['CR']}{R(y)}",
                           f"='{DS}'!{CL['TM']}{R(y)}/'{DS}'!{CL['CR']}{R(y)}", f"='{DS}'!{CL['TX']}{R(y)}/'{DS}'!{CL['CR']}{R(y)}"], 2):
        cc = wc.cell(r, j, f); cc.number_format = pct; cc.font = fgr
E1 = B + len(Y)
c1 = LineChart(); c1.title = "Who paid: direct taxes vs border taxes, 1953–1979 (% of current revenue)"
c1.add_data(Reference(wc, min_col=2, max_col=4, min_row=B, max_row=E1), titles_from_data=True)
c1.set_categories(Reference(wc, min_col=1, min_row=B + 1, max_row=E1)); c1.y_axis.number_format = "0%"
for s_ in c1.series: s_.smooth = False
c1.height = 9; c1.width = 18; c1.legend.position = "b"; wc.add_chart(c1, "A3")
# block 2: corporate / coffee duty 1964-79 (cols F-G)
wc.cell(B, 6, "Year").font = fb; wc.cell(B, 7, "Corporate income tax / coffee duty").font = fb
Y2 = list(range(1964, 1980))
for i, y in enumerate(Y2):
    r = B + 1 + i; wc.cell(r, 6, str(y))
    if y not in (1971, 1972):
        cc = wc.cell(r, 7, f"='{DS}'!{CL['corp']}{R(y)}/'{DS}'!{CL['cof']}{R(y)}"); cc.number_format = pct; cc.font = fgr
E2 = B + len(Y2)
c2 = BarChart(); c2.title = "Corporate income tax as % of the coffee duty, 1964–1979 (1971–72 n.a.)"
c2.add_data(Reference(wc, min_col=7, min_row=B, max_row=E2), titles_from_data=True)
c2.set_categories(Reference(wc, min_col=6, min_row=B + 1, max_row=E2)); c2.y_axis.number_format = "0%"
c2.height = 9; c2.width = 18; c2.legend = None; wc.add_chart(c2, "L3")
# block 3: revenue growth vs coffee price change 1961-79 (cols I-K)
for j, h in enumerate(["Year", "Current revenue growth", "Coffee price change"], 9): wc.cell(B, j, h).font = fb
Y3 = list(range(1961, 1980))
for i, y in enumerate(Y3):
    r = B + 1 + i; wc.cell(r, 9, str(y)); ir = 4 + (y - 1952)
    cc = wc.cell(r, 10, f"=Indicators!I{ir}"); cc.number_format = pct; cc.font = fgr
    cc = wc.cell(r, 11, f"=Indicators!B{ir}/Indicators!B{ir-1}-1"); cc.number_format = pct; cc.font = fgr
E3 = B + len(Y3)
c3 = BarChart(); c3.title = "Revenue growth and the coffee price, 1961–1979 (annual % change)"
c3.add_data(Reference(wc, min_col=10, max_col=11, min_row=B, max_row=E3), titles_from_data=True)
c3.set_categories(Reference(wc, min_col=9, min_row=B + 1, max_row=E3)); c3.y_axis.number_format = "0%"
c3.height = 9; c3.width = 18; c3.legend.position = "b"; wc.add_chart(c3, "A22")
# block 4: boom vs bust change by source (cols M-O)
for j, h in enumerate(["Source", "1976→1977 (boom)", "1977→1978 (bust)", "1978→1979"], 13): wc.cell(B, j, h).font = fb
SRCS = [("Income tax", "TI"), ("Property", "TP"), ("Imports", "TM"), ("Coffee duty", "cof"), ("Other exports", "resx"),
        ("Consumption", "TC"), ("Stamps, services, other", "TS"), ("Non-tax", "NT")]
for i, (lab, k) in enumerate(SRCS):
    r = B + 1 + i; wc.cell(r, 13, lab)
    if k == "resx":
        f = lambda a, b: f"=('{DS}'!{CL['TX']}{R(b)}-'{DS}'!{CL['cof']}{R(b)})-('{DS}'!{CL['TX']}{R(a)}-'{DS}'!{CL['cof']}{R(a)})"
    else:
        f = lambda a, b, k=k: f"='{DS}'!{CL[k]}{R(b)}-'{DS}'!{CL[k]}{R(a)}"
    for j, (a, b) in enumerate([(1976, 1977), (1977, 1978), (1978, 1979)], 14):
        cc = wc.cell(r, j, f(a, b)); cc.number_format = num; cc.font = fgr
E4 = B + len(SRCS)
c4 = BarChart(); c4.type = "bar"; c4.title = "Change in revenue by source, boom and bust (₡M)"
c4.add_data(Reference(wc, min_col=14, max_col=16, min_row=B, max_row=E4), titles_from_data=True)
c4.set_categories(Reference(wc, min_col=13, min_row=B + 1, max_row=E4))
c4.x_axis.tickLblPos = "low"; c3.x_axis.tickLblPos = "low"; c4.height = 9; c4.width = 18; c4.legend.position = "b"; wc.add_chart(c4, "L22")
# block 5: consumption detail lines (cols R-V)
CK = [("Alcohol", "alc"), ("Cigarettes", "cig"), ("Petroleum products", "petr"), ("Selective consumption tax", "sel")]
wc.cell(B, 18, "Year").font = fb
for j, (lab, k) in enumerate(CK, 19): wc.cell(B, j, lab).font = fb
Y5 = [y for y in Y if y >= 1956]
for i, y in enumerate(Y5):
    r = B + 1 + i; wc.cell(r, 18, str(y))
    for j, (lab, k) in enumerate(CK, 19):
        if X[y].get(k) is not None:
            cc = wc.cell(r, j, f"='{DS}'!{CL[k]}{R(y)}"); cc.number_format = num; cc.font = fgr
E5 = B + len(Y5)
c5 = LineChart(); c5.title = "Main consumption taxes, 1956–1979 (₡M; gaps = not available)"; c5.display_blanks = "gap"
c5.add_data(Reference(wc, min_col=19, max_col=22, min_row=B, max_row=E5), titles_from_data=True)
c5.set_categories(Reference(wc, min_col=18, min_row=B + 1, max_row=E5))
for s_ in c5.series: s_.smooth = False
c5.height = 9; c5.width = 18; c5.legend.position = "b"; wc.add_chart(c5, "A41")
# block 6: property detail (cols X-AA)
PK = [("Net wealth tax (Vialidad A)", "vial"), ("Successions and gifts", "succ"), ("Property transfers", "alca"), ("Agricultural land tax", "land")]
wc.cell(B, 24, "Year").font = fb
for j, (lab, k) in enumerate(PK, 25): wc.cell(B, j, lab).font = fb
for i, y in enumerate(Y):
    r = B + 1 + i; wc.cell(r, 24, str(y))
    for j, (lab, k) in enumerate(PK, 25):
        if X[y].get(k) is not None:
            cc = wc.cell(r, j, f"='{DS}'!{CL[k]}{R(y)}"); cc.number_format = num; cc.font = fgr
E6 = B + len(Y)
c6 = LineChart(); c6.title = "Property and wealth taxes, 1953–1979 (₡M; gaps = not available)"; c6.display_blanks = "gap"
c6.add_data(Reference(wc, min_col=25, max_col=28, min_row=B, max_row=E6), titles_from_data=True)
c6.set_categories(Reference(wc, min_col=24, min_row=B + 1, max_row=E6))
for s_ in c6.series: s_.smooth = False
c6.height = 9; c6.width = 18; c6.legend.position = "b"; wc.add_chart(c6, "L41")
wc.cell(B - 1, 1, "Chart data (linked to Revenue_Detail and Indicators)").font = fb

# ---------- Tax_Laws additions ----------
tl = wb["Tax_Laws"]
lastr = max(r for r in range(4, tl.max_row + 1) if tl.cell(r, 2).value and tl.cell(r, 1).value)
add = [
 (1953, "Ley de Vialidad: Vialidad Serie 'A', a tax on net wealth (capital); replaced in 1986 by the Ley del Impuesto sobre el Patrimonio (D.L. 553)", "D.L. 1308 (22 Dec 1953, DO 233, t. 161)", "D.L. 553 of 1986, recitals"),
 (1956, "Tax 'Sobre el Algodón' recorded under consumption taxes (₡0.29M in 1956; ends 1958); no cotton export duty found", "To identify", "MH 1956–58 item detail"),
 (1964, "Tax on cosmetics (₡0.54M in 1964; almost nil from 1965)", "To identify", "MH 1964–65 Estado II-5"),
 (1977, "Net wealth tax (Vialidad A) yield doubles (₡25.9M in 1976 to ₡54.0M in 1977); rate change to identify", "To identify", "MH 1978–79 Estado II-1"),
 (1974, "Selective consumption tax (Impuesto Selectivo al Consumo), first shown in 1974", "To identify", "MH 1978–79 Estado II-1"),
]
note_row = lastr + 2
tl.cell(note_row - 1, 1).value = None
for i, (y, m, dref, s) in enumerate(add):
    r = lastr + 1 + i
    tl.cell(r, 1, str(y)).font = fb; tl.cell(r, 2, m).font = fn; tl.cell(r, 3, dref).font = fin; tl.cell(r, 4, s).font = fi
    cc = tl.cell(r, 5, "No"); cc.fill = PatternFill("solid", fgColor="FFFF00"); cc.font = fn
    for j in (2, 3, 4): tl.cell(r, j).alignment = Alignment(wrap_text=True, vertical="top")
# move footnote below
for r in range(lastr + 1 + len(add), tl.max_row + 1):
    v = tl.cell(r, 1).value
    if isinstance(v, str) and v.startswith("Yellow"): tl.cell(r, 1).value = None
tl.cell(lastr + len(add) + 2, 1, "Yellow cells: mark 'Yes' and add the DO number, tomo and date once the decree has been read.").font = fi

# ---------- README additions ----------
rd = wb["README"]
rd.insert_rows(12, 3)
rd.cell(12, 1, "Revenue_Detail: about 30 lines within the same groups (personal/corporate income tax; net wealth tax, successions and gifts, alcabala, paving, land tax; "
               "coffee, sugar, shrimp and other exports; alcohol, cigarettes, petroleum, sugar, beverages, selective consumption tax, cotton; stamps). "
               "'Not itemized' columns make every group tie to Revenue_CM.").font = Font(name=F)
rd.cell(13, 1, "Periods: composition by government (Osorio, Lemus, Junta/Directorio, Rivera, Sánchez Hernández, Molina, Romero), weighted averages, with a stacked-bar chart.").font = Font(name=F)
rd.cell(14, 1, "Charts_Detail: direct vs border taxes; corporate income tax vs coffee duty; revenue growth vs coffee price; boom and bust changes by source; "
               "consumption and property taxes. Charts_Shaded: annual composition with each government shaded (image).").font = Font(name=F)
for r in (12, 13, 14): rd.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
rd.cell(16 + 3 + 3, 1).value = rd.cell(16 + 3 + 3, 1).value  # no-op keep
wb.save(OUT)
print("ok", CL)
