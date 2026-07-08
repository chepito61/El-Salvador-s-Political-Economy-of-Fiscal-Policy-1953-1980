# =============================================================================
# merge_all_revenue.R
# -----------------------------------------------------------------------------
# Merges the two cleaned revenue datasets into a single long-format file
# covering El Salvador fiscal years 1952–1990:
#   - revenue_1952_1964_long.csv  (from extract_1952_1964.R)
#   - revenue_1960_1990_long.csv  (from extract_1960_1990.R), years 1965–1990 only
#
# Output 1: All_revenue_1952_1990.csv  — full merged dataset
# Output 2: coffee_revenue_1952_1964.png — line graph of coffee tax (item_code 21)
# =============================================================================

library(dplyr)
library(ggplot2)

# UTF-8 encoding for Spanish characters
options(encoding = "UTF-8")

# ── Load source files ─────────────────────────────────────────────────────────

df_1952_1964 <- read.csv(
  "revenue_1952_1964_long.csv",
  stringsAsFactors = FALSE,
  fileEncoding     = "UTF-8-BOM"   # handles the BOM written by the extract scripts
)

df_1960_1990 <- read.csv(
  "revenue_1960_1990_long.csv",
  stringsAsFactors = FALSE,
  fileEncoding     = "UTF-8-BOM"
)

cat("1952-1964 rows loaded:", nrow(df_1952_1964), "\n")
cat("1960-1990 rows loaded:", nrow(df_1960_1990), "\n")

# ── Keep only 1965–1990 from the second dataset ───────────────────────────────
# The overlap period (1960-1964) is already covered by the first dataset.
df_1965_1990 <- df_1960_1990 |>
  filter(year >= 1965)

cat("1965-1990 rows after trimming:", nrow(df_1965_1990), "\n")

# ── Harmonise column types before binding ─────────────────────────────────────
# In revenue_1952_1964_long.csv the year column is stored as character ("1952");
# in revenue_1960_1990_long.csv it is numeric. Convert both to integer.
df_1952_1964$year      <- as.integer(df_1952_1964$year)
df_1965_1990$year      <- as.integer(df_1965_1990$year)

# item_code is character in both files; ensure consistency
df_1952_1964$item_code <- as.character(df_1952_1964$item_code)
df_1965_1990$item_code <- as.character(df_1965_1990$item_code)

# ── Row-bind the two datasets ─────────────────────────────────────────────────
all_revenue <- bind_rows(df_1952_1964, df_1965_1990) |>
  arrange(year, item_code, description)

cat("Total rows in merged dataset:", nrow(all_revenue), "\n")
cat("Year range:", min(all_revenue$year), "–", max(all_revenue$year), "\n")

# ── Save merged dataset ───────────────────────────────────────────────────────
out_file <- "All_revenue_1952_1990.csv"
write.csv(all_revenue, out_file, row.names = FALSE, fileEncoding = "UTF-8")

# Prepend UTF-8 BOM so Excel opens accents correctly
raw_bytes <- readBin(out_file, "raw", file.info(out_file)$size)
con <- file(out_file, "wb")
writeBin(as.raw(c(0xEF, 0xBB, 0xBF)), con)
writeBin(raw_bytes, con)
close(con)

cat("Merged dataset saved to:", out_file, "\n\n")

# ── Coffee tax: filter item_code == "21" ─────────────────────────────────────
# In the 1952-1964 data, item_code 21 = "Sobre el Café" (export tax on coffee).
# This code does not appear in the 1965-1990 aggregate source, so the series
# covers 1954–1963 (the years where the code was recorded).
coffee <- all_revenue |>
  filter(item_code == "21",
         grepl("caf", description, ignore.case = TRUE))

cat("Coffee (item_code 21) rows:", nrow(coffee), "\n")
cat("Years with coffee data:", paste(sort(unique(coffee$year)), collapse = ", "), "\n")

# ── Plot coffee tax revenue over time ─────────────────────────────────────────
p <- ggplot(coffee, aes(x = year, y = value / 1e6)) +         # convert to millions
  geom_line(colour = "#6F4E37", linewidth = 1.1) +
  geom_point(colour = "#6F4E37", size = 3) +
  scale_x_continuous(breaks = sort(unique(coffee$year))) +
  scale_y_continuous(labels = scales::comma_format(suffix = "M")) +
  labs(
    title    = "El Salvador: Coffee Export Tax Revenue (1952–1964)",
    subtitle = "Item code 21 — Sobre el Café",
    x        = "Fiscal Year",
    y        = "Revenue (millions of colones)",
    caption  = "Source: Ministerio de Hacienda, El Salvador"
  ) +
  theme_minimal(base_size = 13) +
  theme(
    plot.title    = element_text(face = "bold"),
    axis.text.x   = element_text(angle = 45, hjust = 1),
    panel.grid.minor = element_blank()
  )

# Save graph
ggsave(
  filename = "coffee_revenue_1952_1964.png",
  plot     = p,
  width    = 9,
  height   = 5.5,
  dpi      = 300
)

cat("Graph saved to: coffee_revenue_1952_1964.png\n")
