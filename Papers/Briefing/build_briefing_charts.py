# Charts for the narrative briefing (draft, Nov 2026). Data: fiscal workbook (Memorias de Hacienda) and World Bank Pink Sheet.
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import os
B = os.path.dirname(os.path.abspath(__file__)); R = os.path.join(B, "..", "..")
OUT = os.path.join(B, "charts")
BLUE, ORANGE, AQUA, YELLOW, MAG, GREEN, VIOLET, RED = "#2a78d6","#eb6834","#1baf7a","#eda100","#e87ba4","#008300","#6250d6","#e34948"
INK, INK2, GRID, SURF = "#0b0b0b", "#52514e", "#e4e3df", "#fcfcfb"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9, "axes.edgecolor": INK2, "axes.labelcolor": INK2,
                     "xtick.color": INK2, "ytick.color": INK2, "figure.facecolor": "white", "axes.facecolor": "white"})
def style(ax):
    ax.spines[["top", "right"]].set_visible(False); ax.grid(axis="y", color=GRID, lw=0.8); ax.set_axisbelow(True)
wb = load_workbook(os.path.join(R, "Fiscal Accounts 1952-1979/El_Salvador_Fiscal_Accounts_Memorias_1952_1979.xlsx"), data_only=True)
fa = wb["Fiscal_Accounts"]
rows = {int(r[0]): r for r in fa.iter_rows(min_row=5, max_row=32, values_only=True) if r[0]}
Y = sorted(rows)
rev = [rows[y][1] for y in Y]; cu = [rows[y][4] for y in Y]; duty = {y: rows[y][3] for y in Y if rows[y][3] is not None}
ac = wb["Analytics_Cycle"]
phase = {int(r[0]): r[3] for r in ac.iter_rows(min_row=6, max_row=31, values_only=True) if r[0]}
cw = load_workbook(os.path.join(R, "Revenue Data 1952-1990/Commodities y recaudacion.xlsx"), data_only=True)["Graficas"]
price = {int(r[0]): r[1] for r in cw.iter_rows(min_row=4, max_row=24, values_only=True) if r[0] and r[1] and int(r[0]) <= 1979}
PER = [("Osorio", 1953, 1956), ("Lemus", 1957, 1960), ("Junta", 1961, 1961), ("Rivera", 1962, 1967), ("Sánchez H.", 1968, 1972), ("Molina", 1973, 1977), ("Romero", 1978, 1979)]
def periods(ax, top):
    for i, (n, a, b) in enumerate(PER):
        if i % 2 == 0: ax.axvspan(a - 0.5, b + 0.5, color="#f3f2ee", zorder=0)
        ax.text((a + b) / 2, top, n, ha="center", va="bottom", fontsize=7.5, color=INK2)

# Fig A: coffee price and coffee duty (small multiples, one axis each)
fig, (a1, a2) = plt.subplots(2, 1, figsize=(9, 6.2), sharex=True)
py = sorted(price); a1.plot(py, [price[y] for y in py], color=BLUE, lw=2); style(a1)
a1.set_ylabel("US$ per kg"); a1.set_title("Coffee price (World Bank Pink Sheet), 1960–1979", loc="left", fontsize=10, color=INK)
a1.annotate("Brazilian frost:\n1976–77 boom", (1977, price[1977]), xytext=(1969.5, 4.3), fontsize=8, color=INK2, arrowprops=dict(arrowstyle="-", color=INK2, lw=0.8))
dy = sorted(duty); a2.bar(dy, [duty[y] for y in dy], color=ORANGE, width=0.7); style(a2)
a2.set_ylabel("₡ millions"); a2.set_title("Coffee export duty collected, 1953–1979", loc="left", fontsize=10, color=INK)
a2.set_xlim(1952.4, 1979.6); a2.set_xticks(range(1953, 1980, 2))
fig.text(0.01, 0.005, "Sources: World Bank Pink Sheet (\"Commodities y recaudacion.xlsx\"); Memorias de Hacienda (fiscal workbook). 1971–72 duty not separated in the series.", fontsize=7, color=INK2)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(os.path.join(OUT, "figA_coffee_price_and_duty.png"), dpi=200); plt.close(fig)

# Fig B: revenue and spending with bust years shaded
fig, ax = plt.subplots(figsize=(9.5, 5.2))
for y in Y:
    if phase.get(y) == "Bust": ax.axvspan(y - 0.5, y + 0.5, color="#fbe3da", zorder=0)
ax.plot(Y, rev, color=BLUE, lw=2, marker="o", ms=3.5, label="Current revenue")
ax.plot(Y, cu, color=ORANGE, lw=2, marker="o", ms=3.5, label="Spending (credits used)")
ax.set_yscale("log"); ax.set_yticks([100, 200, 400, 800, 1600]); ax.set_yticklabels(["100", "200", "400", "800", "1,600"])
style(ax); ax.set_ylabel("₡ millions (log scale)"); ax.set_xlim(1951.5, 1979.5); ax.set_xticks(range(1952, 1980, 2))
ax.text(1979.4, rev[-1], " Revenue", color=INK, fontsize=8, va="center"); ax.text(1979.4, cu[-1], " Spending", color=INK, fontsize=8, va="center")
ax.legend(frameon=False, loc="upper left", fontsize=8)
ax.set_title("Revenue and spending, 1952–1979: shaded years are coffee busts (coffee duty down 10% or more)", loc="left", fontsize=10, color=INK)
fig.text(0.01, 0.005, "Source: Memorias de Hacienda (fiscal workbook, Fiscal_Accounts and Analytics_Cycle). Log scale, so equal slopes mean equal growth rates.", fontsize=7, color=INK2)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(os.path.join(OUT, "figB_revenue_spending_cycle.png"), dpi=200); plt.close(fig)

# Fig C: timeline — coffee duty change and political events
chg = {int(r[0]): r[2] for r in ac.iter_rows(min_row=6, max_row=31, values_only=True) if r[0] and r[2] is not None}
fig, ax = plt.subplots(figsize=(10, 5.6))
cy = sorted(chg); vals = [max(min(chg[y] * 100, 120), -80) for y in cy]
ax.bar(cy, vals, color=[BLUE if v >= 0 else ORANGE for v in vals], width=0.7)
ax.axhline(0, color=INK2, lw=0.8); style(ax); ax.set_ylim(-95, 215); ax.set_xlim(1953.4, 1979.6); ax.set_xticks(range(1954, 1980, 2))
ax.set_ylabel("Change in coffee duty, % (capped at +120)")
for y in cy:
    if chg[y] * 100 > 120: ax.text(y, 112, f"+{chg[y]*100:.0f}%", ha="center", va="top", fontsize=6.8, color="white", rotation=90)
EV = [(1960, "Oct 1960\ncoup"), (1961, "Jan 1961\ncoup"), (1967, "ACERO\ngeneral strike"), (1968, "1st ANDES\nstrike"), (1969, "War with\nHonduras"),
      (1971, "53-day\nteachers' strike"), (1972, "Stolen election,\nUniversity occupied"), (1975, "30 July student\nmassacre"), (1976, "Agrarian project\nwithdrawn"),
      (1977, "Fraud;\nstate of siege"), (1979, "Oct 1979\ncoup")]
lv = [190, 155, 190, 155, 125, 190, 155, 190, 155, 125, 190]
for (y, t), h in zip(EV, lv):
    ax.plot([y, y], [-95, h - 4], color=INK2, lw=0.6, ls=":"); ax.text(y, h, t, ha="center", va="bottom", fontsize=6.8, color=INK)
ax.set_title("The coffee cycle and the political record, 1954–1979", loc="left", fontsize=10, color=INK)
fig.text(0.01, 0.005, "Sources: Memorias de Hacienda (Analytics_Cycle; 1971–72 export duties); event dates from the yearly notes and the Political Economy Context note.", fontsize=7, color=INK2)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(os.path.join(OUT, "figC_cycle_and_politics.png"), dpi=200); plt.close(fig)
print("ok", len(Y), sorted(price)[:2], {y: round(chg[y]*100) for y in cy})
