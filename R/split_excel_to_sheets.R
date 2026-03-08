
# ─────────────────────────────────────────────────────────────────────────────
# split_excel_to_sheets()
#
Reads an Excel file, splits it into `n` equal chunks, writes each chunk
# to its own sheet inside a single workbook, and returns the workbook object.
#
# Dependencies: openxlsx, data.table
# Install once:  install.packages(c("openxlsx", "data.table"))
# ─────────────────────────────────────────────────────────────────────────────

library(openxlsx)
library(data.table)          # Fast in-memory operations (~10–50× base R)

split_excel_to_sheets <- function(
    file_path,               # Path to the source .xlsx file
    n,                       # Number of splits (sheets)
    sheet         = 1,       # Source sheet: index or name
    output_path   = NULL,    # Optional: save to disk; NULL = in-memory only
    sheet_prefix  = "Part",  # Sheet name prefix  →  "Part_1", "Part_2", …
    header_style  = TRUE,    # Apply a styled header row to each sheet
    col_widths    = "auto"   # "auto" | numeric vector | NULL
) {

  # ── 0. Validate inputs ─────────────────────────────────────────────────────
  stopifnot(
    is.character(file_path), file.exists(file_path),
    is.numeric(n), n >= 1, n == as.integer(n)
  )
  n <- as.integer(n)

  # ── 1. Read source data (fast path) ────────────────────────────────────────
  # openxlsx::read.xlsx is faster when detectDates = FALSE and we skip type
  # inference; convert afterwards only if needed.
  message(sprintf("[1/4] Reading '%s' …", basename(file_path)))
  t0 <- proc.time()

  dt <- setDT(
    read.xlsx(
      file_path,
      sheet       = sheet,
      detectDates = FALSE,    # Skip slow date-detection pass
      colNames    = TRUE
    )
  )                           # setDT() converts in-place — zero copy cost

  n_rows <- nrow(dt)
  message(sprintf("      %s rows × %s cols  (%.2f s)",
                  format(n_rows, big.mark = ","),
                  ncol(dt),
                  (proc.time() - t0)[["elapsed"]]))

  if (n > n_rows) {
    warning(sprintf(
      "n (%d) > nrow (%d). Reducing n to nrow.", n, n_rows
    ))
    n <- n_rows
  }

  # ── 2. Compute chunk boundaries (vectorised, no loop) ──────────────────────
  # cut() assigns each row to a chunk label in one vectorised call.
  message("[2/4] Computing chunk boundaries …")

  row_idx    <- seq_len(n_rows)
  chunk_ids  <- as.integer(cut(row_idx, breaks = n, labels = FALSE))
  # data.table split: produces a named list of data.tables — very fast
  chunks     <- split(dt, chunk_ids)

  # ── 3. Build workbook ──────────────────────────────────────────────────────
  message("[3/4] Writing sheets …")
  t1 <- proc.time()

  wb <- createWorkbook()

  # Pre-build header style once (reused across all sheets)
  hs <- if (isTRUE(header_style)) {
    createStyle(
      fontName    = "Arial",
      fontSize    = 10,
      fontColour  = "#FFFFFF",
      fgFill      = "#2F5496",
      halign      = "LEFT",
      textDecoration = "bold",
      border      = "Bottom",
      borderColour = "#FFFFFF"
    )
  } else NULL

  body_style <- createStyle(
    fontName = "Arial",
    fontSize = 10
  )

  for (i in seq_len(n)) {
    sheet_name <- sprintf("%s_%d", sheet_prefix, i)
    chunk      <- chunks[[as.character(i)]]

    addWorksheet(wb, sheetName = sheet_name)

    writeData(
      wb,
      sheet       = sheet_name,
      x           = chunk,
      startRow    = 1,
      startCol    = 1,
      headerStyle = hs,       # applies style to header row automatically
      borders     = "rows",
      borderStyle = "thin",
      borderColour = "#D9D9D9"
    )

    # Apply body style to data rows
    addStyle(
      wb,
      sheet = sheet_name,
      style = body_style,
      rows  = seq(2, nrow(chunk) + 1),
      cols  = seq_len(ncol(chunk)),
      gridExpand = TRUE       # avoids an inner R loop
    )

    # Column widths
    if (!is.null(col_widths)) {
      setColWidths(
        wb,
        sheet  = sheet_name,
        cols   = seq_len(ncol(chunk)),
        widths = col_widths   # "auto" triggers openxlsx's own sizing
      )
    }

    # Freeze the header row for readability
    freezePane(wb, sheet = sheet_name, firstRow = TRUE)

    message(sprintf("      Sheet %-15s → %s rows",
                    paste0('"', sheet_name, '"'),
                    format(nrow(chunk), big.mark = ",")))
  }

  message(sprintf("      Done  (%.2f s)", (proc.time() - t1)[["elapsed"]]))

  # ── 4. Optionally save to disk ─────────────────────────────────────────────
  if (!is.null(output_path)) {
    message(sprintf("[4/4] Saving to '%s' …", output_path))
    t2 <- proc.time()
    saveWorkbook(wb, file = output_path, overwrite = TRUE)
    message(sprintf("      Saved  (%.2f s)", (proc.time() - t2)[["elapsed"]]))
  } else {
    message("[4/4] Skipping disk save (output_path = NULL)")
  }

  invisible(wb)   # return workbook object for further manipulation
}


# ─────────────────────────────────────────────────────────────────────────────
# USAGE EXAMPLES
# ─────────────────────────────────────────────────────────────────────────────

# -- Basic: split into 5 sheets, save to disk ---------------------------------
# wb <- split_excel_to_sheets(
#   file_path   = "data/sales_2024.xlsx",
#   n           = 5,
#   output_path = "data/sales_2024_split.xlsx"
# )

# -- Advanced: second sheet of source, custom prefix, no auto col widths ------
# wb <- split_excel_to_sheets(
#   file_path    = "data/large_export.xlsx",
#   n            = 10,
#   sheet        = "RawData",
#   output_path  = "data/large_export_chunked.xlsx",
#   sheet_prefix = "Chunk",
#   col_widths   = NULL          # skip width calc for maximum speed
# )

# -- In-memory only (no disk write) — pipe into further openxlsx calls --------
# wb <- split_excel_to_sheets("input.xlsx", n = 3)
# addWorksheet(wb, "Summary")
# saveWorkbook(wb, "final.xlsx", overwrite = TRUE)
