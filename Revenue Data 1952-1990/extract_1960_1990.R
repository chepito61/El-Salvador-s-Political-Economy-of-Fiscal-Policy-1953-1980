# =============================================================================
# extract_1960_1990.R
# -----------------------------------------------------------------------------
# Extracts and standardizes government revenue data for El Salvador (1960–1990)
# from a single Excel workbook that stores all years in one wide-format sheet
# ("Ingresos"), with years as columns and revenue categories as rows.
# The script pivots the data to long format and converts values from millions
# to units (multiplies by 1,000,000).
# Output: revenue_1960_1990_long.csv with columns
#         year | item_code | description | value
# =============================================================================

library(readxl)
library(dplyr)
library(tidyr)

# UTF-8 encoding so Spanish characters (accents, ñ, ©) are preserved
options(encoding = "UTF-8")
Sys.setlocale("LC_ALL", "Spanish")

# Path to the source workbook and target sheet
fp         <- "El Salvador_1960-1990_28 feb.xlsx"
sheet_name <- "Ingresos"

# ── Read raw (no headers) ─────────────────────────────────────────────────────
raw <- read_excel(
  fp, sheet = sheet_name, col_names = FALSE, .name_repair = "minimal"
)
cat("Raw dimensions:", nrow(raw), "rows x", ncol(raw), "cols\n")

# ── Extract year labels from row 4 (columns 2 onward) ────────────────────────
# Row 4 contains the fiscal years as numeric values; column 1 is the label column.
year_row  <- 4
year_vals <- as.numeric(unlist(raw[year_row, 2:ncol(raw)]))
years     <- as.integer(year_vals)
cat("Years detected:", min(years, na.rm = TRUE),
    "to", max(years, na.rm = TRUE), "\n")

# ── Data: rows 8–60 ───────────────────────────────────────────────────────────
# Row 60 is "Ingresos de Capital", the last meaningful revenue category.
data_raw <- raw[8:60, ]

# Assign column names: first column is "description", rest are year integers
colnames(data_raw) <- c("description", as.character(years))

# ── Clean description column ──────────────────────────────────────────────────
data_raw$description <- trimws(as.character(data_raw$description))

# Remove rows with empty or NA descriptions
data_raw <- data_raw[
  !is.na(data_raw$description) &
    data_raw$description != "" &
    data_raw$description != "NA",
]

cat("Data rows after cleaning:", nrow(data_raw), "\n")

# ── Drop redundant section-header rows ───────────────────────────────────────
# These rows are subtotals/aggregates already covered by their child rows.
drop_descs <- c(
  "1. Totales",
  "2. Ingresos Tributarios",
  "IV. Ingresos No Tributarios"
)
data_raw <- data_raw[!data_raw$description %in% drop_descs, ]
cat("Data rows after dropping headers:", nrow(data_raw), "\n")

# ── Extract item_code from the description string ─────────────────────────────
# Patterns handled:
#   "040---Description"  or  "020--Description"  → item_code = "040" / "020"
#   "IV. Description"                             → item_code = "IV"
#   "1. Description"  or  "1-Description"         → item_code = "1"
extract_code <- function(x) {
  # 3-digit numeric code followed by two or more dashes
  m <- regmatches(x, regexpr("^\\d{3}(?=--)", x, perl = TRUE))
  if (length(m) > 0) return(m)
  # Roman numeral followed by period + space
  m <- regmatches(x, regexpr("^[IVX]+(?=\\.\\s)", x, perl = TRUE))
  if (length(m) > 0) return(m)
  # Single or double digit followed by period or dash
  m <- regmatches(x, regexpr("^\\d{1,2}(?=[.\\-])", x, perl = TRUE))
  if (length(m) > 0) return(m)
  ""
}

# Remove the code prefix from the description to leave a clean label
strip_code <- function(x) {
  x <- sub("^\\d{3}-{2,}\\s*", "", x)   # "040---" prefix
  x <- sub("^[IVX]+\\.\\s+", "", x)      # "IV. " prefix
  x <- sub("^\\d{1,2}[.\\-]\\s*", "", x) # "1. " or "1-" prefix
  trimws(x)
}

data_raw$item_code   <- sapply(data_raw$description, extract_code, USE.NAMES = FALSE)
data_raw$description <- sapply(data_raw$description, strip_code,   USE.NAMES = FALSE)

# ── Coerce year columns to character before pivoting ─────────────────────────
# Ensures consistent type across columns regardless of how Excel stored the values
year_cols <- setdiff(names(data_raw), "description")
data_raw[year_cols] <- lapply(data_raw[year_cols], as.character)

# ── Pivot wide → long ─────────────────────────────────────────────────────────
long_df <- data_raw |>
  pivot_longer(
    cols      = -c(description, item_code),
    names_to  = "year",
    values_to = "value"
  ) |>
  mutate(
    year  = as.integer(year),
    value = suppressWarnings(as.numeric(value)) * 1e6  # millions → units
  ) |>
  filter(!is.na(value)) |>           # remove year-category combinations with no data
  select(year, item_code, description, value) |>
  arrange(year, description)

cat("Long format rows:", nrow(long_df), "\n")
cat("Years in output:",
    paste(sort(unique(long_df$year)), collapse = ", "), "\n")

# ── Save output ───────────────────────────────────────────────────────────────
out_file <- "revenue_1960_1990_long.csv"
write.csv(long_df, out_file, row.names = FALSE, fileEncoding = "UTF-8")

# Prepend UTF-8 BOM so Excel opens the file with correct accent rendering
raw_bytes <- readBin(out_file, "raw", file.info(out_file)$size)
con <- file(out_file, "wb")
writeBin(as.raw(c(0xEF, 0xBB, 0xBF)), con)
writeBin(raw_bytes, con)
close(con)

cat("Saved to:", out_file, "\n")
