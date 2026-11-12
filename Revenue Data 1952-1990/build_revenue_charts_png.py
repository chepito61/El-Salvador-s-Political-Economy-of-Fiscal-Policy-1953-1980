import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XImg
import os
OUT = "El_Salvador_Revenue_Memorias_1952_1979.xlsx"
CH = "charts/"
os.makedirs(CH, exist_ok=True)
wb = load_workbook(OUT, data_only=True); ws = wb["Revenue_Detail"]
def v(c, r):
    x = ws[f"{c}{r}"].value
    return x if isinstance(x, (int, float)) else 0.0
yrs = list(range(1953, 1980)); R = lambda y: 5 + y - 1953
series = [("Income tax", lambda r: v("F", r)),
          ("Property and wealth", lambda r: v("M", r)),
          ("Import duties", lambda r: v("N", r)),
          ("Export duty: coffee*", lambda r: v("O", r) if v("O", r) else v("T", r)),
          ("Other export duties", lambda r: v("T", r) - v("O", r) if v("O", r) else 0),
          ("Consumption taxes", lambda r: v("AD", r)),
          ("Stamps, services, other, emergency", lambda r: v("AH", r) + v("AI", r)),
          ("Non-tax revenue", lambda r: v("AM", r))]
cols = ["#1f4e79", "#5b9bd5", "#7f7f7f", "#8c510a", "#d8b365", "#2e7d32", "#9ccc65", "#c9c9c9"]
PER = [("Osorio", 1953, 1956), ("Lemus", 1957, 1960), ("Junta", 1961, 1961), ("Rivera", 1962, 1967),
       ("Sánchez H.", 1968, 1972), ("Molina", 1973, 1977), ("Romero", 1978, 1979)]
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9})
# 1. annual composition with government bands
fig, ax = plt.subplots(figsize=(12, 5.6))
for i, (a, b) in enumerate([(p[1], p[2]) for p in PER]):
    if i % 2 == 0: ax.axvspan(a - 0.5, b + 0.5, color="#f0f0f0", zorder=0)
bottom = [0] * len(yrs)
for (nm, f), c in zip(series, cols):
    vals = [f(R(y)) / v("AN", R(y)) * 100 for y in yrs]
    ax.bar(yrs, vals, bottom=bottom, color=c, width=0.75, label=nm, zorder=2, edgecolor="white", linewidth=0.3)
    bottom = [b + x for b, x in zip(bottom, vals)]
for nm, a, b in PER:
    ax.text((a + b) / 2, 102, nm, ha="center", va="bottom", fontsize=9, fontweight="bold")
    ax.axvline(a - 0.5, color="#555", lw=0.6, ls=":", zorder=3)
ax.set_ylim(0, 108); ax.set_xlim(1952.4, 1979.6); ax.set_ylabel("% of current revenue")
ax.set_xticks(yrs); ax.set_xticklabels([str(y) for y in yrs], rotation=90)
ax.spines[["top", "right"]].set_visible(False)
ax.legend(ncol=4, loc="upper center", bbox_to_anchor=(0.5, -0.13), frameon=False)
ax.set_title("El Salvador: composition of central government revenue by government, 1953–1979", loc="left", fontsize=11, pad=22)
fig.text(0.01, 0.01, "*1971–72: coffee duty not separately available; total export duties shown. Source: Memorias de Hacienda (workbook Revenue_Detail).", fontsize=7.5, color="#555")
fig.tight_layout(); fig.savefig(CH + "composition_by_government_annual.png", dpi=150); plt.close(fig)
# 2. period bars from Periods sheet
wp = wb["Periods"]
labs = [wp.cell(r, 1).value for r in range(5, 15)]
fig, ax = plt.subplots(figsize=(10, 5.4))
pc = ["#1f4e79", "#2f75b5", "#9dc3e6", "#7f7f7f", "#8c510a", "#d8b365", "#2e7d32", "#9ccc65", "#bfbfbf", "#e0e0e0"]
names = [wp.cell(4, j).value.replace("\n", "\n") for j in range(2, 9)]
bottom = [0] * 7
for i, lab in enumerate(labs):
    vals = []
    for j in range(2, 9):
        x = wp.cell(5 + i, j).value
        vals.append(x * 100 if isinstance(x, (int, float)) else 0)
    tot = None
    ax.bar(range(7), vals, bottom=bottom, color=pc[i], label=lab, edgecolor="white", linewidth=0.4)
    for k, (b0, x) in enumerate(zip(bottom, vals)):
        if x >= 4: ax.text(k, b0 + x / 2, f"{x:.0f}", ha="center", va="center", fontsize=7.5, color="white" if i in (0, 1, 3, 4, 6) else "black")
    bottom = [b + x for b, x in zip(bottom, vals)]
ax.set_xticks(range(7)); ax.set_xticklabels(names, fontsize=8.5); ax.set_ylabel("% of current revenue"); ax.set_ylim(0, 103)
ax.spines[["top", "right"]].set_visible(False)
ax.legend(loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False, fontsize=8)
ax.set_title("Revenue composition by government (weighted averages, % of current revenue)", loc="left", fontsize=11)
fig.text(0.01, 0.01, "Lines missing in some years are computed on the years available (Periods sheet). Source: Memorias de Hacienda.", fontsize=7.5, color="#555")
fig.tight_layout(); fig.savefig(CH + "composition_by_government_periods.png", dpi=150); plt.close(fig)
# 3. corporate vs coffee + direct vs trade (for the Word note)
fig, ax = plt.subplots(figsize=(10, 4.5))
d = [(v("F", R(y)) + v("M", R(y))) / v("AN", R(y)) * 100 for y in yrs]
t = [(v("N", R(y)) + v("T", R(y))) / v("AN", R(y)) * 100 for y in yrs]
for i, (a, b) in enumerate([(p[1], p[2]) for p in PER]):
    if i % 2 == 0: ax.axvspan(a - 0.5, b + 0.5, color="#f0f0f0", zorder=0)
ax.plot(yrs, d, color="#1f4e79", lw=2, label="Direct taxes (income + property)")
ax.plot(yrs, t, color="#8c510a", lw=2, label="Border taxes (imports + exports)")
for nm, a, b in PER: ax.text((a + b) / 2, 66, nm, ha="center", fontsize=8.5, fontweight="bold")
ax.set_ylim(0, 70); ax.set_ylabel("% of current revenue"); ax.set_xticks(yrs[::2])
ax.spines[["top", "right"]].set_visible(False); ax.legend(frameon=False, loc="lower left")
ax.set_title("Who paid: direct taxes and border taxes, 1953–1979", loc="left", fontsize=11)
fig.tight_layout(); fig.savefig(CH + "direct_vs_border_taxes.png", dpi=150); plt.close(fig)
wb.close()
# insert images
wb = load_workbook(OUT)
if "Charts_Shaded" in wb.sheetnames: del wb["Charts_Shaded"]
s = wb.create_sheet("Charts_Shaded")
from openpyxl.styles import Font
s.cell(1, 1, "Charts with each government shaded (static images, generated from Revenue_Detail and Periods; regenerate after changing data)").font = Font(name="Arial", bold=True, size=13)
for fn_, anchor, w in [("composition_by_government_annual.png", "A3", 1000), ("composition_by_government_periods.png", "A33", 860), ("direct_vs_border_taxes.png", "A62", 860)]:
    im = XImg(CH + fn_); r = w / im.width; im.width = w; im.height = int(im.height * r); s.add_image(im, anchor)
wb.save(OUT); print("ok")
