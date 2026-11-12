# The three pillars of the budget, 1952-1979, with Pleitez's long-run benchmark for the agro-export State (1827-1948)
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
wb = load_workbook("../Fiscal Accounts 1952-1979/El_Salvador_Fiscal_Accounts_Memorias_1952_1979.xlsx", data_only=True); ep = wb["Exp_Ministry_Pillars"]
Y = list(range(1952, 1980)); g = lambda c, y: ep[f"{c}{5 + y - 1952}"].value * 100
dev = [g("P", y) for y in Y]; sec = [g("O", y) for y in Y]; agr = [g("R", y) for y in Y]; edu = [g("Q", y) for y in Y]
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9})
fig, ax = plt.subplots(figsize=(10.5, 5.8))
PER = [("Osorio", 1953, 1956), ("Lemus", 1957, 1960), ("Junta", 1961, 1961), ("Rivera", 1962, 1967), ("Sánchez H.", 1968, 1972), ("Molina", 1973, 1977), ("Romero", 1978, 1979)]
for i, (n, a, b) in enumerate(PER):
    if i % 2 == 0: ax.axvspan(a - 0.5, b + 0.5, color="#f0f0f0", zorder=0)
    ax.text((a + b) / 2, 57, n, ha="center", fontsize=8.5, fontweight="bold")
ax.plot(Y, dev, color="#2e7d32", lw=2.4, label="Development (Education + Health + Public Works)")
ax.plot(Y, edu, color="#5b9bd5", lw=1.6, ls="--", label="  of which Education")
ax.plot(Y, sec, color="#8c510a", lw=2.4, label="Security (Defensa + Seguridad Pública)")
ax.plot(Y, agr, color="#b8860b", lw=1.6, label="Agro-export support (Agricultura)")
for v, col, lab in [(31.1, "#8c510a", "Security, agro-export State 1827–1948: 31.1%"), (20.5, "#2e7d32", "Development, 1827–1948: 20.5%")]:
    ax.axhline(v, color=col, ls=":", lw=1.3); ax.text(1951.6, v + 0.6, lab, fontsize=7.5, color=col)
ax.set_xlim(1951.4, 1979.6); ax.set_ylim(0, 60); ax.set_xticks(Y[::2]); ax.set_ylabel("% of spending")
ax.spines[["top", "right"]].set_visible(False); ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.07), ncol=2, fontsize=8)
ax.set_title("The three pillars of the budget, 1952–1979 (% of spending)", loc="left", fontsize=11, pad=8)
fig.text(0.01, 0.01, "Sources: Memorias de Hacienda (Exp_Ministry_Pillars; basis varies by year, 1955–56 average). Benchmarks: Pleitez, Tabla 1.15 (averages 1827–1948).", fontsize=7.2, color="#555")
fig.tight_layout(); fig.savefig("charts_1948_1960/fig5_three_pillars_1952_1979.png", dpi=150); plt.close(fig)
for y, d, s, a, e in zip(Y, dev, sec, agr, edu): print(y, round(d, 1), round(s, 1), round(a, 1), round(e, 1), round(d / s, 1))
