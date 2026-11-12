# Static charts for the Word note, from the recalculated fiscal workbook
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
wb = load_workbook("El_Salvador_Fiscal_Accounts_Memorias_1952_1979.xlsx", data_only=True)
fa, ep = wb["Fiscal_Accounts"], wb["Exp_Ministry_Pillars"]
Y = list(range(1952, 1980)); R = lambda y: 5 + y - 1952
g = lambda ws, c, y: ws[f"{c}{R(y)}"].value
PER = [("Osorio", 1953, 1956), ("Lemus", 1957, 1960), ("Junta", 1961, 1961), ("Rivera", 1962, 1967),
       ("Sánchez H.", 1968, 1972), ("Molina", 1973, 1977), ("Romero", 1978, 1979)]
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9})
def bands(ax, top):
    for i, (n, a, b) in enumerate(PER):
        if i % 2 == 0: ax.axvspan(a - 0.5, b + 0.5, color="#f0f0f0", zorder=0)
        ax.text((a + b) / 2, top, n, ha="center", va="bottom", fontsize=8.5, fontweight="bold")
# 1. underlying result and debt
fig, ax = plt.subplots(figsize=(10, 4.8))
u = [g(fa, "R", y) * 100 for y in Y]; d = [g(fa, "U", y) * 100 for y in Y]
bands(ax, 74)
ax.bar(Y, u, color=["#2e7d32" if v >= 0 else "#b23a3a" for v in u], width=0.7, zorder=2, label="Underlying result (% of revenue)")
ax.plot(Y, d, color="#1f4e79", lw=2, zorder=3, label="Direct public debt (% of revenue)")
ax.axhline(0, color="#333", lw=0.8); ax.set_ylim(-25, 80); ax.set_xlim(1951.4, 1979.6); ax.set_xticks(Y[::2])
ax.spines[["top", "right"]].set_visible(False); ax.legend(frameon=False, loc="upper left", bbox_to_anchor=(0, 0.93))
ax.set_title("Underlying budget result and public debt, 1952–1979 (% of current revenue)", loc="left", fontsize=11, pad=18)
fig.text(0.01, 0.01, "Source: Memorias de Hacienda via the yearly notes (Fiscal_Accounts sheet). Credits-used basis.", fontsize=7.5, color="#555")
fig.tight_layout(); fig.savefig("charts/underlying_result_and_debt.png", dpi=150); plt.close(fig)
# 2. shares
fig, ax = plt.subplots(figsize=(10, 4.6))
bands(ax, 53)
for c, lab, col in [("P", "Development (Education + Health + Public Works)", "#2e7d32"), ("Q", "Education", "#5b9bd5"), ("O", "Security (Defence + Public Security)", "#8c510a")]:
    ax.plot(Y, [g(ep, c, y) * 100 for y in Y], lw=2, color=col, label=lab)
ax.set_ylim(0, 58); ax.set_xlim(1951.4, 1979.6); ax.set_xticks(Y[::2]); ax.set_ylabel("% of spending")
ax.spines[["top", "right"]].set_visible(False); ax.legend(frameon=False, loc="center left", bbox_to_anchor=(0, 0.62), fontsize=8)
ax.set_title("Development, education and security shares of spending, 1952–1979", loc="left", fontsize=11, pad=18)
fig.text(0.01, 0.01, "Basis varies by year (credits used or gastos; Exp_Ministry_Pillars). 1955–56: average of the two years.", fontsize=7.5, color="#555")
fig.tight_layout(); fig.savefig("charts/spending_shares.png", dpi=150); plt.close(fig)
print("ok")
