# Investment and debt charts for the narrative briefing (draft, Nov 2026).
# Data: fiscal workbook (Memorias de Hacienda); guarantees from the yearly notes (Estado 21 / Cuadro 24);
# 1965-73 central-government saving and investment from World Bank (1974), Table 5.1.
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import os
B = os.path.dirname(os.path.abspath(__file__)); R = os.path.join(B, "..", "..")
OUT = os.path.join(B, "charts")
BLUE, ORANGE, AQUA, YELLOW, MAG, GREEN, VIOLET = "#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6"
INK, INK2, GRID = "#0b0b0b", "#52514e", "#e4e3df"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9, "axes.edgecolor": INK2, "axes.labelcolor": INK2,
                     "xtick.color": INK2, "ytick.color": INK2, "figure.facecolor": "white", "axes.facecolor": "white"})
def style(ax):
    ax.spines[["top", "right"]].set_visible(False); ax.grid(axis="y", color=GRID, lw=0.8); ax.set_axisbelow(True)

wb = load_workbook(os.path.join(R, "Fiscal Accounts 1952-1979/El_Salvador_Fiscal_Accounts_Memorias_1952_1979.xlsx"), data_only=True)
fa = {int(r[0]): r for r in wb["Fiscal_Accounts"].iter_rows(min_row=5, max_row=32, values_only=True) if r[0]}
ee = {int(r[0]): r for r in wb["Exp_Economic"].iter_rows(min_row=5, max_row=40, values_only=True) if r[0] and str(r[0]).isdigit()}
rev = {y: fa[y][1] for y in fa}; debt = {y: fa[y][15] for y in fa}

# State guarantees of autonomous-institution debt (₡M, end-year), from the yearly notes.
# 1964-75: narrower coverage (Banco Hipotecario mortgage bonds and IRA dropped from the statement);
# 1976-79: revised series of MH 1978 (Cuadro 24).
guar = {1959: 165.98, 1960: 182.48, 1961: 195.8, 1962: 200.6, 1963: 209.5,
        1964: 114.3, 1965: 130.3, 1966: 156.1, 1967: 172.2, 1968: 181.6, 1969: 190.5, 1970: 191.6,
        1971: 190.7, 1972: 197.67, 1973: 201.59, 1974: 225.12, 1975: 266.49,
        1976: 370.46, 1977: 461.95, 1978: 605.82, 1979: 688.21}

# Fig D: debt and guarantees
fig, ax = plt.subplots(figsize=(9.5, 5.0))
Y = sorted(debt)
ax.plot(Y, [debt[y] for y in Y], color=BLUE, lw=2, marker="o", ms=3.5, label="Direct public debt")
for seg, ls in (((1959, 1963), "-"), ((1964, 1975), "-"), ((1976, 1979), "-")):
    ys = [y for y in sorted(guar) if seg[0] <= y <= seg[1]]
    ax.plot(ys, [guar[y] for y in ys], color=ORANGE, lw=2, marker="o", ms=3.5, ls=ls,
            label="State-guaranteed debt of autonomous institutions" if seg[0] == 1959 else None)
ax.set_yscale("log"); ax.set_yticks([25, 50, 100, 200, 400, 800]); ax.set_yticklabels(["25", "50", "100", "200", "400", "800"])
style(ax); ax.set_xlim(1951.5, 1979.8); ax.set_xticks(range(1952, 1980, 2)); ax.set_ylabel("₡ millions, end of year (log scale)")
ax.annotate("1964: statement drops the\nBanco Hipotecario and IRA", (1964, guar[1964]), xytext=(1957.2, 70), fontsize=7.5, color=INK2,
            arrowprops=dict(arrowstyle="-", color=INK2, lw=0.8))
ax.annotate("1976: revised series", (1976, guar[1976]), xytext=(1969.6, 520), fontsize=7.5, color=INK2,
            arrowprops=dict(arrowstyle="-", color=INK2, lw=0.8))
ax.legend(frameon=False, loc="upper left", fontsize=8)
ax.set_title("Direct debt and State guarantees, 1952–1979", loc="left", fontsize=10, color=INK)
fig.text(0.01, 0.005, "Sources: Memorias de Hacienda (fiscal workbook; Estado 21 and MH 1978 Cuadro 24 via the yearly notes). Guarantee coverage changes in 1964 and 1976 (breaks in the line).", fontsize=7, color=INK2)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(os.path.join(OUT, "figD_debt_and_guarantees.png"), dpi=200); plt.close(fig)

# Fig E: current saving and capital spending, three bases
def cur(y):  # current spending: wages + goods and services + interest + current transfers
    r = ee[y]; return sum(x or 0 for x in (r[1], r[2], r[3], r[4]))
sav, cap, basis = {}, {}, {}
for y in (1952, 1953, 1954, 1955, 1956):
    sav[y] = rev[y] - cur(y); cap[y] = ee[y][5] or 0; basis[y] = "A"
wbt = {1965: (52.5, 53.3), 1966: (37.0, 51.2), 1967: (34.8, 45.3), 1968: (24.8, 21.4), 1969: (39.0, 39.2),
       1970: (19.2, 38.7), 1971: (43.3, 61.2), 1972: (46.3, 88.5), 1973: (80.8, 114.2)}
for y, (s, c) in wbt.items(): sav[y], cap[y], basis[y] = s, c, "B"
for y in (1974, 1975, 1976, 1977, 1978, 1979):
    r = ee[y]; sav[y] = rev[y] - cur(y); cap[y] = sum(x or 0 for x in (r[5], r[6], r[7])); basis[y] = "C"
fig, ax = plt.subplots(figsize=(10, 5.4))
ys = sorted(sav); w = 0.38
ax.bar([y - w / 2 for y in ys], [sav[y] for y in ys], width=w, color=BLUE, label="Current saving (revenue − current spending)")
ax.bar([y + w / 2 for y in ys], [cap[y] for y in ys], width=w, color=ORANGE, label="Capital spending (investment, on-lending, capital transfers)")
ax.set_yscale("symlog", linthresh=20); ax.set_yticks([0, 10, 20, 50, 100, 200, 500]); ax.set_yticklabels(["0", "10", "20", "50", "100", "200", "500"])
style(ax); ax.set_xlim(1951.3, 1979.7); ax.set_xticks(list(range(1952, 1957)) + list(range(1965, 1980))); ax.tick_params(axis="x", labelrotation=90, labelsize=8)
ax.axvspan(1956.6, 1964.4, color="#f3f2ee", zorder=0); ax.text(1960.5, 3, "1957–1964:\neconomic\nclassification\nnot yet\nextracted", ha="center", fontsize=7.5, color=INK2)
for (a, b, t) in ((1952, 1956, "Memorias, object\nclasses (payments)"), (1965, 1973, "World Bank (1974), Table 5.1"), (1974, 1979, "Memorias, Cuadro 15 (payments)")):
    ax.text((a + b) / 2, 700, t, ha="center", fontsize=7.5, color=INK2)
ax.set_ylim(0, 1300); ax.set_ylabel("₡ millions (scale linear to 20, logarithmic above)")
ax.legend(frameon=False, loc="upper left", fontsize=8, bbox_to_anchor=(0, 0.88))
ax.set_title("Central government current saving and capital spending, years available", loc="left", fontsize=10, color=INK)
fig.text(0.01, 0.005, "Bases differ by period and are not strictly comparable; 1974–79 exclude year-end commitments. Sources: fiscal workbook (Exp_Economic); World Bank, Economic Memorandum 1974.", fontsize=7, color=INK2)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(os.path.join(OUT, "figE_saving_and_capital.png"), dpi=200); plt.close(fig)

print("Year saving capital basis | debt debt/rev guar guar/rev")
for y in sorted(set(sav) | set(guar)):
    s = f"{sav[y]:7.1f} {cap[y]:7.1f} {basis[y]}" if y in sav else " " * 17
    g = f"{guar[y]:7.1f} {100*guar[y]/rev[y]:5.0f}%" if y in guar else ""
    print(y, s, f"| {debt[y]:6.1f} {100*debt[y]/rev[y]:4.0f}%", g)

# Fig F: how revenue, spending, investment and debt interact (schematic)
from matplotlib.patches import FancyBboxPatch
fig, ax = plt.subplots(figsize=(11, 6.6)); ax.set_xlim(-2.5, 103); ax.set_ylim(-17, 63); ax.axis("off")
BX = {"coffee": (6, 36, "Coffee price\nset abroad;\n₡2.50 per US$", YELLOW),
      "rev": (23, 36, "Revenue\ncoffee duty,\nimport duties", BLUE),
      "sav": (40, 36, "Current saving\nrevenue minus\ncurrent spending", BLUE),
      "cur": (40, 15, "Current spending\nteachers, health,\nsecurity, transfers", BLUE),
      "inv": (58, 36, "Central government\ninvestment and\ncounterpart funds", ORANGE),
      "rc": (58, 15, "Recurrent costs\nstaff for new\nschools, hospitals", ORANGE),
      "aut": (58, -4, "Autonomous\ninstitutions' invest-\nment: dams, port, water", ORANGE),
      "loan": (50, 55, "External loans\nIBRD, IDB, AID, BCIE", VIOLET),
      "bond": (68, 55, "Domestic bonds;\nCentral Bank credit", VIOLET),
      "debt": (76, 36, "Direct debt", VIOLET),
      "ds": (94, 36, "Debt service", VIOLET),
      "guar": (76, 15, "State guarantees\n(contingent debt)", VIOLET)}
W, H = 14.2, 9.4
for k, (x, y, t_, c) in BX.items():
    ax.add_patch(FancyBboxPatch((x - W / 2, y - H / 2), W, H, boxstyle="round,pad=0.25,rounding_size=1.0", fc="white", ec=c, lw=1.8))
    ax.text(x, y, t_, ha="center", va="center", fontsize=7.8, color=INK)
def edge(k, tx, ty):
    x, y = BX[k][0], BX[k][1]; dx, dy = tx - x, ty - y
    s = min((W / 2 + 0.9) / abs(dx) if dx else 1e9, (H / 2 + 0.9) / abs(dy) if dy else 1e9)
    return x + dx * s, y + dy * s
def arr(a, b, txt=None, dashed=False, off=(0, 0)):
    (x1, y1), (x2, y2) = BX[a][:2], BX[b][:2]
    p1, p2 = edge(a, x2, y2), edge(b, x1, y1)
    ax.annotate("", xy=p2, xytext=p1, arrowprops=dict(arrowstyle="-|>", color=INK2, lw=1.1, ls="--" if dashed else "-", shrinkA=0, shrinkB=0))
    if txt: ax.text((p1[0] + p2[0]) / 2 + off[0], (p1[1] + p2[1]) / 2 + off[1], txt, fontsize=7, color=INK2, ha="center", va="center",
                    bbox=dict(fc="white", ec="none", pad=0.5))
def path(pts, txt=None, tpos=None):
    xs, ys = zip(*pts); ax.plot(xs[:-1], ys[:-1], color=INK2, lw=1.1)
    ax.annotate("", xy=pts[-1], xytext=pts[-2], arrowprops=dict(arrowstyle="-|>", color=INK2, lw=1.1, shrinkA=0, shrinkB=0))
    if txt: ax.text(*tpos, txt, fontsize=7, color=INK2, ha="center", va="center", bbox=dict(fc="white", ec="none", pad=0.5))
arr("coffee", "rev"); arr("rev", "sav"); arr("rev", "cur"); arr("cur", "sav", "less", off=(2.6, 0))
arr("sav", "inv"); arr("loan", "inv"); arr("bond", "inv"); arr("bond", "debt", "loans and bonds\nadd to debt", off=(5.5, 1))
arr("inv", "rc", "builds", off=(3.2, 0)); arr("rc", "cur"); arr("debt", "ds")
xs, ys = [47.6, 49.0, 49.0], [31.0, 31.0, -4]
ax.plot(xs, ys, color=INK2, lw=1.1, ls="--"); ax.annotate("", xy=(58 - W / 2 - 0.9, -4), xytext=(49.0, -4), arrowprops=dict(arrowstyle="-|>", color=INK2, lw=1.1, shrinkA=0, shrinkB=0))
ax.text(42.5, 2.5, "surplus lent to\nown agencies", fontsize=7, color=INK2, ha="center", va="center")
arr("guar", "aut"); arr("guar", "debt", "if called", dashed=True, off=(3.6, 0))
path([(94, 36 - H / 2 - 0.9), (94, -13), (36, -13), (36, 15 - H / 2 - 0.9)], "debt service paid from the operating budget", (75, -13))
ax.set_title("How revenue, spending, investment and debt interact in the fiscal record", loc="left", fontsize=10, color=INK)
fig.text(0.01, 0.01, "Schematic drawn from the yearly notes and the note on loan-financed investment. Blue: operating budget; orange: investment; violet: financing and debt.", fontsize=7, color=INK2)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(os.path.join(OUT, "figF_interaction.png"), dpi=200); plt.close(fig)
