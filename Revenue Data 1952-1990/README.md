# EL-SALVADOR — Fiscal Revenue Data (1952–1990)

Historical government revenue data for El Salvador, digitised from Ministerio de Hacienda annual reports and compiled into a single long-format dataset spanning 1952–1990.

---

## Repository structure

```
EL-SALVADOR/
│
├── Data sources (raw)
│   ├── Data Exctractions - Revenue.xlsx      # Multi-sheet workbook; one sheet per fiscal year (1952–1964)
│   ├── El Salvador_1960-1990_28 feb.xlsx     # Single wide-format sheet "Ingresos" covering 1960–1990
│   ├── Commodities y recaudacion.xlsx        # Supplementary commodities / revenue reference
│   └── MH 1954.pdf                           # Ministerio de Hacienda annual report, 1954
│
├── R scripts
│   ├── extract_1952_1964.R                   # Step 1 — extract 1952-1964 from the multi-sheet workbook
│   ├── extract_1960_1990.R                   # Step 2 — extract 1960-1990 from the wide-format workbook
│   └── merge_all_revenue.R                   # Step 3 — merge, filter, and plot
│
├── Outputs
│   ├── revenue_1952_1964_long.csv            # Long-format output of Step 1
│   ├── revenue_1960_1990_long.csv            # Long-format output of Step 2
│   ├── All_revenue_1952_1990.csv             # Final merged dataset (1952–1990)
│   └── coffee_revenue_1952_1964.png          # Chart: coffee export tax revenue by year
│
└── README.md
```

---

## Data sources

| File | Years | Format | Notes |
|------|-------|--------|-------|
| `Data Exctractions - Revenue.xlsx` | 1952–1964 | One sheet per year | 1952 has a unique layout (data from row 8, value in col C); all other sheets share an Item / Description / Value structure |
| `El Salvador_1960-1990_28 feb.xlsx` | 1960–1990 | Single wide sheet ("Ingresos") | Years as columns (row 4), revenue categories as rows 8–60; values in millions of colones |

---

## Scripts

### 1. `extract_1952_1964.R`

Reads `Data Exctractions - Revenue.xlsx`, loops over each sheet (fiscal year), and normalises the data into a consistent long format.

- The **1952** sheet is handled separately (fixed layout: data starts at row 8, value in column C).
- All other sheets are detected dynamically by locating the `"Item"` header row.
- Keeps columns: item code (col A), description (col B), value (col D).
- Assigns item_code `1000` to grand-total rows.
- Removes duplicate descriptions within each sheet.
- Writes **`revenue_1952_1964_long.csv`** with a UTF-8 BOM for correct accent display in Excel.

Output schema: `year | item_code | description | value`

---

### 2. `extract_1960_1990.R`

Reads the `"Ingresos"` sheet from `El Salvador_1960-1990_28 feb.xlsx`.

- Extracts year labels from row 4 (columns 2 onward).
- Uses data rows 8–60 (last meaningful category: "Ingresos de Capital").
- Parses item codes embedded in description strings (3-digit numeric, Roman numeral, or 1–2 digit prefixes).
- Drops redundant section-header rows (`"1. Totales"`, `"2. Ingresos Tributarios"`, `"IV. Ingresos No Tributarios"`).
- Pivots from wide to long format and **multiplies values by 1,000,000** (source is in millions of colones).
- Writes **`revenue_1960_1990_long.csv`** with a UTF-8 BOM.

Output schema: `year | item_code | description | value`

---

### 3. `merge_all_revenue.R`

Merges the two cleaned datasets and produces a chart.

**Merge logic:**
- Loads both long-format CSVs.
- Keeps only years **1965–1990** from `revenue_1960_1990_long.csv` to avoid overlap with the 1952–1964 series.
- Harmonises `year` (to integer) and `item_code` (to character) before row-binding.
- Sorts by year → item_code → description.
- Writes **`All_revenue_1952_1990.csv`** with a UTF-8 BOM.

**Coffee chart:**
- Filters `item_code == "21"` + description containing "café" (item 21 = *Sobre el Café*, coffee export tax, recorded 1954–1963).
- Produces a line + point plot with revenue in millions of colones on the y-axis.
- Saves **`coffee_revenue_1952_1964.png`** at 300 dpi (9 × 5.5 in).

---

## Final dataset — `All_revenue_1952_1990.csv`

| Column | Type | Description |
|--------|------|-------------|
| `year` | integer | Fiscal year |
| `item_code` | character | Revenue classification code (blank where not present in source) |
| `description` | character | Revenue category label (in Spanish) |
| `value` | numeric | Revenue in colones (units, not millions) |

Year coverage: **1952–1990**
Source overlap: 1952–1964 comes exclusively from the multi-sheet workbook; 1965–1990 from the wide-format workbook.

---

## How to run

Run the scripts in order from the project root directory:

```r
source("extract_1952_1964.R")   # → revenue_1952_1964_long.csv
source("extract_1960_1990.R")   # → revenue_1960_1990_long.csv
source("merge_all_revenue.R")   # → All_revenue_1952_1990.csv + coffee_revenue_1952_1964.png
```

**Required R packages:** `readxl`, `dplyr`, `tidyr`, `ggplot2`, `scales`

---

## Notes

- All monetary values are in **Salvadoran colones**.
- The 1952–1964 source reports values in units; the 1960–1990 source reports in millions — the extraction scripts normalise both to units before saving.
- Spanish locale (`Sys.setlocale("LC_ALL", "Spanish")`) is set at the top of each script to handle accented characters correctly on Windows.
- All output CSVs include a UTF-8 BOM (`EF BB BF`) prepended so that Excel on Windows renders Spanish characters (á, é, ñ, etc.) without corruption.
