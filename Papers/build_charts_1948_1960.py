# Charts for the 1948-1960 note. Data: Memorias (via notes and workbooks), Dada Hirezi (1978) and Pleitez.
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9})
OUT = "charts_1948_1960/"
def shade(ax):
    ax.axvspan(1949.5, 1956.5, color="#f0f0f0", zorder=0)
    for x, t in [(1949.0, "CGR"), (1953, "Osorio"), (1958.5, "Lemus")]:
        ax.text(x, ax.get_ylim()[1] * 0.97, t, ha="center", va="top", fontsize=8.5, fontweight="bold")
# 1. revenue, credits used, coffee duty 1950-60
rev = {1950: 83.4, 1951: 117.7, 1952: 129.7, 1953: 137.6, 1954: 167.4, 1955: 164.8, 1956: 170.6, 1957: 189.5, 1958: 167.7, 1959: 156.7, 1960: 170.5}
cu = {1952: 131.2, 1953: 140.3, 1954: 155.7, 1955: 155.7, 1956: 166.4, 1957: 186.3, 1958: 184.7, 1959: 175.8, 1960: 174.3}
cof = {1953: 34.97, 1954: 46.16, 1955: 48.99, 1956: 43.47, 1957: 54.55, 1958: 35.43, 1959: 24.77, 1960: 26.28}
sf = {1952: 4.1, 1953: 5.5, 1954: 20.0, 1955: 35.1, 1956: 40.0, 1957: 45.7, 1958: 31.2, 1959: 18.0, 1960: 20.3}
fig, ax = plt.subplots(figsize=(10, 4.6))
ax.set_ylim(0, 210)
ax.axvspan(1949.6, 1956.5, color="#f0f0f0", zorder=0)
for x, t in [(1953.0, "Osorio (1950–56)"), (1958.5, "Lemus (1956–60)")]:
    ax.text(x, 205, t, ha="center", va="top", fontsize=8.5, fontweight="bold")
ax.plot(list(rev), list(rev.values()), color="#1f4e79", lw=2.2, marker="o", ms=3, label="Current revenue")
ax.plot(list(cu), list(cu.values()), color="#8c510a", lw=2, ls="--", label="Credits used (spending)")
ax.bar(list(cof), list(cof.values()), color="#d8b365", width=0.6, label="Coffee export duty", zorder=2)
ax.plot(list(sf), list(sf.values()), color="#2e7d32", lw=1.8, marker="s", ms=3, label="Superávit Financiero (end-year)")
ax.set_xticks(range(1950, 1961)); ax.set_ylabel("₡ millions"); ax.spines[["top", "right"]].set_visible(False)
ax.legend(frameon=False, loc="upper left", bbox_to_anchor=(0.0, 0.93), fontsize=8)
ax.set_title("Revenue, spending, the coffee duty and the surplus, 1950–1960 (₡ millions)", loc="left", fontsize=11)
fig.text(0.01, 0.01, "Sources: Memorias de Hacienda (ten-year table in MH 1960; Estado 1; yearly notes). Coffee duty from 1953.", fontsize=7.5, color="#555")
fig.tight_layout(); fig.savefig(OUT + "fig1_revenue_spending_1950_1960.png", dpi=150); plt.close(fig)
# 2. export structure
yrs = [1945, 1948, 1951, 1954, 1957, 1960]
cshare = [87.6, 86.8, 88.9, 87.6, 79.3, 65.7]
cval = [46.7, 90.4, 190.1, 230.0, 274.6, 191.7]
tot = [v / s * 100 for v, s in zip(cval, cshare)]
cot = {1945: 0.58, 1948: 4.87, 1951: 9.59, 1954: 16.14, 1957: 39.05}
cots = [cot[y] / t * 100 for y, t in zip(yrs[:5], tot[:5])]
fig, ax = plt.subplots(figsize=(10, 4.4))
ax.bar([y - 0.6 for y in yrs], cshare, width=1.1, color="#8c510a", label="Coffee")
ax.bar([y + 0.6 for y in yrs[:5]], cots, width=1.1, color="#5b9bd5", label="Cotton")
for y, v in zip(yrs, cshare): ax.text(y - 0.6, v + 1.5, f"{v:.0f}", ha="center", fontsize=8)
for y, v in zip(yrs[:5], cots): ax.text(y + 0.6, v + 1.5, f"{v:.0f}", ha="center", fontsize=8)
ax.set_xticks(yrs); ax.set_ylim(0, 100); ax.set_ylabel("% of total exports"); ax.spines[["top", "right"]].set_visible(False)
ax.legend(frameon=False, loc="upper right")
ax.set_title("Export concentration: coffee and cotton, % of total exports, 1945–1960", loc="left", fontsize=11)
fig.text(0.01, 0.01, "Coffee: Dada Hirezi (1978), Cuadro 2-1 (Anuarios Estadísticos). Cotton: Pleitez, Tabla 2.1 (CEPAL 1959), over total exports implied by Cuadro 2-1. Cotton 1960 n.a.", fontsize=7.2, color="#555")
fig.tight_layout(); fig.savefig(OUT + "fig2_export_concentration.png", dpi=150); plt.close(fig)
# 3. revenue composition 1953-60 (workbook)
wb = load_workbook("../Revenue Data 1952-1990/El_Salvador_Revenue_Memorias_1952_1979.xlsx", data_only=True); rc = wb["Revenue_CM"]
Y = list(range(1953, 1961)); g = lambda c, y: rc[f"{c}{4 + y - 1952}"].value or 0
S = [("Income tax", "B"), ("Property", "F"), ("Import duties", "G"), ("Export duties (coffee)", "H"), ("Consumption taxes", "K"),
     ("Stamps, services, other", "L"), ("Non-tax revenue", "Q")]
cols = ["#1f4e79", "#5b9bd5", "#7f7f7f", "#8c510a", "#2e7d32", "#9ccc65", "#c9c9c9"]
fig, ax = plt.subplots(figsize=(10, 4.8)); bottom = [0] * len(Y)
for (n, c), col in zip(S, cols):
    v = [g(c, y) / g("S", y) * 100 for y in Y]
    ax.bar(Y, v, bottom=bottom, color=col, label=n, width=0.7, edgecolor="white", linewidth=0.3)
    for x, b0, vv in zip(Y, bottom, v):
        if vv >= 6: ax.text(x, b0 + vv / 2, f"{vv:.0f}", ha="center", va="center", fontsize=7.5, color="white" if col in ("#1f4e79", "#7f7f7f", "#8c510a", "#2e7d32") else "black")
    bottom = [a + b for a, b in zip(bottom, v)]
ax.set_ylim(0, 100); ax.set_xticks(Y); ax.set_ylabel("% of current revenue"); ax.spines[["top", "right"]].set_visible(False)
ax.legend(frameon=False, loc="center left", bbox_to_anchor=(1.0, 0.5), fontsize=8)
ax.set_title("Composition of current revenue, 1953–1960 (%)", loc="left", fontsize=11)
fig.text(0.01, 0.01, "Source: Memorias de Hacienda; revenue workbook (Revenue_CM), modern classification.", fontsize=7.5, color="#555")
fig.tight_layout(); fig.savefig(OUT + "fig3_revenue_composition_1953_1960.png", dpi=150); plt.close(fig)
# 4. spending pillars
yy = ["1952", "1953", "1954", "1955–56", "1957", "1958", "1959", "1960"]
sec = [13.3, 14.5, 15.9, 14.1, 13.3, 13.5, 12.7, 12.6]; edu = [14.0, 13.2, 19.5, 14.5, 16.0, 15.5, 16.6, 18.6]
hea = [9.4, 9.5, 9.4, 9.5, 9.7, 10.1, 10.4, 11.0]; op = [15.2, 14.9, 16.5, 17.4, 17.7, 17.5, 14.7, 16.1]
fig, ax = plt.subplots(figsize=(10, 4.4)); x = range(len(yy)); w = 0.2
for k, (v, n, c) in enumerate([(sec, "Security (Defensa + Seguridad Pública)", "#8c510a"), (edu, "Education", "#5b9bd5"), (hea, "Health", "#9ccc65"), (op, "Public works", "#2e7d32")]):
    ax.bar([i + (k - 1.5) * w for i in x], v, width=w, color=c, label=n)
ax.set_xticks(list(x)); ax.set_xticklabels(yy); ax.set_ylabel("% of credits used"); ax.set_ylim(0, 24)
ax.spines[["top", "right"]].set_visible(False); ax.legend(frameon=False, ncol=2, loc="upper left", fontsize=8)
ax.set_title("Spending priorities: shares of credits used, 1952–1960 (%)", loc="left", fontsize=11)
fig.text(0.01, 0.01, "Source: Memorias de Hacienda, Estado N° 1 (Consolidated MISI Spending note). 1955–56: biennium.", fontsize=7.5, color="#555")
fig.tight_layout(); fig.savefig(OUT + "fig4_spending_shares_1952_1960.png", dpi=150); plt.close(fig)
print("ok", [round(t, 1) for t in tot], [round(c, 1) for c in cots])
