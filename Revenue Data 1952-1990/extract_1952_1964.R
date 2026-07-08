# =============================================================================
# extract_revenue_1952_1964.R
# -----------------------------------------------------------------------------
# Extracts and standardizes government revenue data for El Salvador (1952–1964)
# from a multi-sheet Excel file.  Each sheet corresponds to one fiscal year.
# The 1952 sheet has a unique layout and is handled separately; all other
# sheets share a common "Item / Description / Value" structure.
# Output: a single long-format CSV (revenue_long.csv) with columns
#         year | item_code | description | value
# =============================================================================

library(readxl)
library(dplyr)

# UTF-8 encoding so Spanish characters (accents, ñ, ©) are preserved
options(encoding = "UTF-8")
Sys.setlocale("LC_ALL", "Spanish")

# Path to the source workbook
fp <- "Data Exctractions - Revenue.xlsx"

# List all sheets (each sheet = one fiscal year)
sheets <- excel_sheets(fp)
cat("Found", length(sheets), "sheets:", paste(sheets, collapse = ", "), "\n\n")

# Container to accumulate one data frame per year
all_years <- list()

for (sheet in sheets) {

  # ── Special handling for 1952 ───────────────────────────────────────────────
  # The 1952 sheet differs from all others: data begins at row 8 and
  # the value column is C (index 3) instead of D (index 4).
  if (sheet == "1952") {
    cat("Sheet: 1952 -> Using fixed layout (data from row 8, cols A & C)\n")

    raw <- read_excel(fp, sheet = sheet, col_names = FALSE, .name_repair = "minimal")

    # Data starts at row 8; keep columns A (description) and C (value)
    data_rows <- raw[8:nrow(raw), ]

    df <- data.frame(
      description = as.character(data_rows[[1]]),
      value       = suppressWarnings(as.numeric(data_rows[[3]])),
      stringsAsFactors = FALSE
    )

    # Extract item code: the leading "TITULOS I … VII" portion of column A
    df$item_code <- trimws(
      ifelse(
        grepl("TITULOS\\s+[IVX]+", df$description, ignore.case = TRUE),
        regmatches(df$description,
                   regexpr("TITULOS\\s+[IVX]+", df$description, ignore.case = TRUE)),
        ""
      )
    )

    df$description <- trimws(df$description)

    # Remove blank/NA description rows
    df <- df[!is.na(df$description) & df$description != "", ]

    # Stop at "TOTAL GENERAL PERCIBIDO..." and assign it item_code 1000
    stop_idx <- which(grepl("^TOTAL GENERAL PERCIBIDO", df$description, ignore.case = TRUE))[1]
    if (!is.na(stop_idx)) {
      df <- df[1:stop_idx, ]
      df$item_code[stop_idx] <- "1000"
    }

    # Remove duplicate descriptions (keep first occurrence)
    df <- df[!duplicated(df$description), ]
    df$year <- sheet

    all_years[[sheet]] <- df[, c("year", "item_code", "description", "value")]
    next
  }
  # ───────────────────────────────────────────────────────────────────────────

  # Read the full sheet without headers to inspect its structure
  raw <- read_excel(fp, sheet = sheet, col_names = FALSE, .name_repair = "minimal")

  # Find the header row: the first row where column A starts with "Item"
  col_a <- as.character(raw[[1]])
  header_row <- which(grepl("^Item", col_a, ignore.case = TRUE))[1]

  if (is.na(header_row)) {
    cat("WARNING: Could not find 'Item...' header row in sheet:", sheet, "\n")
    next
  }

  cat("Sheet:", sheet, "-> Header row found at row", header_row, "\n")

  # Re-read using the detected header row so column names are assigned correctly
  df <- read_excel(fp, sheet = sheet, skip = header_row - 1, col_names = TRUE, .name_repair = "minimal")

  # Keep only the three relevant columns: A (item code), B (description), D (value)
  if (ncol(df) < 4) {
    cat("WARNING: Sheet", sheet, "has fewer than 4 columns after header. Skipping.\n")
    next
  }

  df <- df[, c(1, 2, 4)]
  colnames(df) <- c("item_code", "description", "value")

  # Drop rows where both item_code and description are missing
  df <- df[!is.na(df$item_code) | !is.na(df$description), ]

  # Drop rows with no meaningful label
  df <- df[!is.na(df$description), ]

  # Ensure value column is numeric
  df$value <- suppressWarnings(as.numeric(df$value))

  # Trim leading/trailing whitespace
  df$description <- trimws(df$description)
  df$item_code   <- trimws(as.character(df$item_code))

  # Remove duplicate descriptions within the same sheet (keep first occurrence)
  df <- df[!duplicated(df$description), ]

  # Assign item_code 1000 to the grand-total row where item_code is blank/NA
  total_idx <- grepl("^TOTAL GENERAL DE INGRESOS", df$description, ignore.case = TRUE)
  df$item_code[total_idx & (is.na(df$item_code) | df$item_code == "NA")] <- "1000"

  # Tag with the fiscal year (taken from the sheet name)
  df$year <- sheet

  all_years[[sheet]] <- df[, c("year", "item_code", "description", "value")]
}

cat("\n")

# Combine all yearly data frames into a single long-format table
long_df <- bind_rows(all_years)

# ── Save long-format output ───────────────────────────────────────────────────
# Columns: year | item_code | description | value
out_long <- "revenue_1952_1964_long.csv"
write.csv(long_df, out_long, row.names = FALSE, fileEncoding = "UTF-8")

# Prepend UTF-8 BOM so Excel opens the file with correct accent rendering
raw <- readBin(out_long, "raw", file.info(out_long)$size)
con <- file(out_long, "wb")
writeBin(as.raw(c(0xEF, 0xBB, 0xBF)), con)
writeBin(raw, con)
close(con)

cat("Long format saved to:", out_long, "\n")
