# ─────────────────────────────────────────────────────────────────────────────
# test_split_excel.R
# testthat test suite for split_excel_to_sheets()
#
# Run all tests:
#   testthat::test_file("test_split_excel.R")
#
# Dependencies: testthat, openxlsx, data.table
# ─────────────────────────────────────────────────────────────────────────────

library(testthat)
library(openxlsx)
library(data.table)

source("split_excel_to_sheets.R")   # adjust path if needed

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS — build reusable temp fixtures
# ─────────────────────────────────────────────────────────────────────────────

#' Write a data.frame to a temp .xlsx and return its path
make_fixture <- function(df, sheet_name = "Sheet1") {
  path <- tempfile(fileext = ".xlsx")
  wb   <- createWorkbook()
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet = sheet_name, x = df)
  saveWorkbook(wb, path, overwrite = TRUE)
  path
}

# Standard 100-row fixture
df_100 <- data.frame(
  id     = 1:100,
  name   = paste0("item_", 1:100),
  value  = round(runif(100, 1, 1000), 2),
  active = sample(c(TRUE, FALSE), 100, replace = TRUE)
)

# Minimal 3-row fixture (edge-case sizing)
df_3 <- data.frame(x = 1:3, y = c("a", "b", "c"))

# Single-row fixture
df_1 <- data.frame(col1 = 42L, col2 = "only row")

# Named-sheet fixture (two sheets)
make_two_sheet_fixture <- function() {
  path <- tempfile(fileext = ".xlsx")
  wb   <- createWorkbook()
  addWorksheet(wb, "Primary")
  writeData(wb, sheet = "Primary", x = df_100)
  addWorksheet(wb, "Secondary")
  writeData(wb, sheet = "Secondary", x = df_3)
  saveWorkbook(wb, path, overwrite = TRUE)
  path
}

# ─────────────────────────────────────────────────────────────────────────────
# 1. INPUT VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

test_that("errors on non-existent file", {
  expect_error(
    split_excel_to_sheets("/no/such/file.xlsx", n = 2),
    regexp = NULL   # any error is acceptable
  )
})

test_that("errors when file_path is not a character", {
  expect_error(split_excel_to_sheets(123, n = 2))
})

test_that("errors when n is not numeric", {
  f <- make_fixture(df_100)
  expect_error(split_excel_to_sheets(f, n = "five"))
})

test_that("errors when n is less than 1", {
  f <- make_fixture(df_100)
  expect_error(split_excel_to_sheets(f, n = 0))
})

test_that("errors when n is a non-integer numeric", {
  f <- make_fixture(df_100)
  expect_error(split_excel_to_sheets(f, n = 2.5))
})

test_that("errors when n is NA", {
  f <- make_fixture(df_100)
  expect_error(split_excel_to_sheets(f, n = NA_integer_))
})

# ─────────────────────────────────────────────────────────────────────────────
# 2. RETURN VALUE
# ─────────────────────────────────────────────────────────────────────────────

test_that("returns a Workbook object invisibly", {
  f  <- make_fixture(df_100)
  wb <- split_excel_to_sheets(f, n = 4)
  expect_s4_class(wb, "Workbook")
})

test_that("return value is usable — can add a sheet after the call", {
  f  <- make_fixture(df_100)
  wb <- split_excel_to_sheets(f, n = 2)
  expect_silent(addWorksheet(wb, "Extra"))
  expect_true("Extra" %in% names(wb))
})

# ─────────────────────────────────────────────────────────────────────────────
# 3. SHEET COUNT
# ─────────────────────────────────────────────────────────────────────────────

test_that("workbook contains exactly n sheets for n = 1", {
  f  <- make_fixture(df_100)
  wb <- split_excel_to_sheets(f, n = 1)
  expect_equal(length(names(wb)), 1L)
})

test_that("workbook contains exactly n sheets for n = 5", {
  f  <- make_fixture(df_100)
  wb <- split_excel_to_sheets(f, n = 5)
  expect_equal(length(names(wb)), 5L)
})

test_that("workbook contains exactly n sheets for n = 10", {
  f  <- make_fixture(df_100)
  wb <- split_excel_to_sheets(f, n = 10)
  expect_equal(length(names(wb)), 10L)
})

test_that("workbook contains exactly n sheets when n equals nrow", {
  f  <- make_fixture(df_3)
  wb <- split_excel_to_sheets(f, n = 3)
  expect_equal(length(names(wb)), 3L)
})

# ─────────────────────────────────────────────────────────────────────────────
# 4. SHEET NAMING
# ─────────────────────────────────────────────────────────────────────────────

test_that("default sheet names follow 'Part_N' pattern", {
  f  <- make_fixture(df_100)
  wb <- split_excel_to_sheets(f, n = 3)
  expect_equal(names(wb), c("Part_1", "Part_2", "Part_3"))
})

test_that("custom sheet_prefix is applied correctly", {
  f  <- make_fixture(df_100)
  wb <- split_excel_to_sheets(f, n = 3, sheet_prefix = "Chunk")
  expect_equal(names(wb), c("Chunk_1", "Chunk_2", "Chunk_3"))
})

test_that("sheet names are unique across all n sheets", {
  f      <- make_fixture(df_100)
  wb     <- split_excel_to_sheets(f, n = 7)
  sheets <- names(wb)
  expect_equal(length(sheets), length(unique(sheets)))
})

# ─────────────────────────────────────────────────────────────────────────────
# 5. ROW DISTRIBUTION
# ─────────────────────────────────────────────────────────────────────────────

#' Helper: read all sheets from a saved workbook into a list of data.frames
read_all_sheets <- function(path) {
  lapply(getSheetNames(path), function(s) read.xlsx(path, sheet = s))
}

test_that("total rows across all sheets equals source row count", {
  f    <- make_fixture(df_100)
  out  <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 5, output_path = out)
  sheets    <- read_all_sheets(out)
  total_rows <- sum(sapply(sheets, nrow))
  expect_equal(total_rows, nrow(df_100))
})

test_that("no row is duplicated across sheets", {
  f   <- make_fixture(df_100)
  out <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 4, output_path = out)
  sheets   <- read_all_sheets(out)
  all_ids  <- unlist(lapply(sheets, `[[`, "id"))
  expect_equal(length(all_ids), length(unique(all_ids)))
})

test_that("chunks are roughly equal in size (max diff ≤ 1 row)", {
  f   <- make_fixture(df_100)   # 100 rows, n = 7 → uneven split
  out <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 7, output_path = out)
  sheets     <- read_all_sheets(out)
  chunk_rows <- sapply(sheets, nrow)
  expect_lte(max(chunk_rows) - min(chunk_rows), 1L)
})

test_that("n = 1 returns entire dataset in one sheet", {
  f   <- make_fixture(df_100)
  out <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 1, output_path = out)
  result <- read.xlsx(out, sheet = 1)
  expect_equal(nrow(result), nrow(df_100))
})

test_that("row order is preserved across chunks", {
  f   <- make_fixture(df_100)
  out <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 5, output_path = out)
  sheets    <- read_all_sheets(out)
  all_ids   <- unlist(lapply(sheets, `[[`, "id"))
  expect_equal(all_ids, df_100$id)
})

# ─────────────────────────────────────────────────────────────────────────────
# 6. COLUMN INTEGRITY
# ─────────────────────────────────────────────────────────────────────────────

test_that("all sheets contain the same column names as the source", {
  f      <- make_fixture(df_100)
  out    <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 4, output_path = out)
  sheets <- read_all_sheets(out)
  for (s in sheets) {
    expect_equal(sort(colnames(s)), sort(colnames(df_100)))
  }
})

test_that("column count per sheet matches source", {
  f      <- make_fixture(df_100)
  out    <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 3, output_path = out)
  sheets <- read_all_sheets(out)
  for (s in sheets) {
    expect_equal(ncol(s), ncol(df_100))
  }
})

# ─────────────────────────────────────────────────────────────────────────────
# 7. EDGE CASES
# ─────────────────────────────────────────────────────────────────────────────

test_that("n > nrow triggers a warning and reduces n to nrow", {
  f <- make_fixture(df_3)   # 3 rows
  expect_warning(
    split_excel_to_sheets(f, n = 10),
    regexp = "Reducing n to nrow"
  )
})

test_that("after n > nrow warning, sheet count equals nrow not original n", {
  f  <- make_fixture(df_3)
  wb <- suppressWarnings(split_excel_to_sheets(f, n = 10))
  expect_equal(length(names(wb)), nrow(df_3))
})

test_that("single-row source with n = 1 works without error", {
  f  <- make_fixture(df_1)
  wb <- split_excel_to_sheets(f, n = 1)
  expect_s4_class(wb, "Workbook")
  expect_equal(length(names(wb)), 1L)
})

test_that("handles a source file with a single column", {
  f  <- make_fixture(data.frame(only = 1:20))
  wb <- split_excel_to_sheets(f, n = 4)
  expect_equal(length(names(wb)), 4L)
})

test_that("handles a source file with many columns (wide table)", {
  wide <- as.data.frame(matrix(1:500, nrow = 10, ncol = 50))
  f    <- make_fixture(wide)
  wb   <- split_excel_to_sheets(f, n = 2)
  expect_equal(length(names(wb)), 2L)
})

# ─────────────────────────────────────────────────────────────────────────────
# 8. SOURCE SHEET SELECTION
# ─────────────────────────────────────────────────────────────────────────────

test_that("reads from a named source sheet correctly", {
  f   <- make_two_sheet_fixture()
  out <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 2, sheet = "Primary", output_path = out)
  sheets    <- read_all_sheets(out)
  total_rows <- sum(sapply(sheets, nrow))
  expect_equal(total_rows, nrow(df_100))
})

test_that("reads from the secondary sheet when specified by name", {
  f   <- make_two_sheet_fixture()
  out <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 2, sheet = "Secondary", output_path = out)
  sheets    <- read_all_sheets(out)
  total_rows <- sum(sapply(sheets, nrow))
  expect_equal(total_rows, nrow(df_3))
})

test_that("reads from a sheet specified by index (index = 2)", {
  f   <- make_two_sheet_fixture()
  out <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 1, sheet = 2, output_path = out)
  result <- read.xlsx(out, sheet = 1)
  expect_equal(nrow(result), nrow(df_3))
})

# ─────────────────────────────────────────────────────────────────────────────
# 9. DISK OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

test_that("no file is written when output_path = NULL", {
  f   <- make_fixture(df_100)
  out <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 3, output_path = NULL)
  expect_false(file.exists(out))
})

test_that("file is written when output_path is provided", {
  f   <- make_fixture(df_100)
  out <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 3, output_path = out)
  expect_true(file.exists(out))
})

test_that("saved file is a valid readable xlsx", {
  f   <- make_fixture(df_100)
  out <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 3, output_path = out)
  expect_no_error(getSheetNames(out))
})

test_that("saved file has correct number of sheets", {
  f   <- make_fixture(df_100)
  out <- tempfile(fileext = ".xlsx")
  split_excel_to_sheets(f, n = 6, output_path = out)
  expect_equal(length(getSheetNames(out)), 6L)
})

# ─────────────────────────────────────────────────────────────────────────────
# 10. STYLE & FORMATTING OPTIONS
# ─────────────────────────────────────────────────────────────────────────────

test_that("header_style = FALSE runs without error", {
  f  <- make_fixture(df_100)
  expect_no_error(split_excel_to_sheets(f, n = 3, header_style = FALSE))
})

test_that("col_widths = NULL runs without error", {
  f  <- make_fixture(df_100)
  expect_no_error(split_excel_to_sheets(f, n = 3, col_widths = NULL))
})

test_that("col_widths as numeric vector runs without error", {
  f  <- make_fixture(df_100)
  expect_no_error(
    split_excel_to_sheets(f, n = 2, col_widths = rep(15, ncol(df_100)))
  )
})

# ─────────────────────────────────────────────────────────────────────────────
# 11. PERFORMANCE SMOKE TEST
# ─────────────────────────────────────────────────────────────────────────────

test_that("processes 10,000-row file into 10 sheets in under 30 seconds", {
  big_df <- data.frame(
    id    = 1:10000,
    grp   = sample(letters, 10000, replace = TRUE),
    val   = runif(10000),
    score = rnorm(10000)
  )
  f     <- make_fixture(big_df)
  t0    <- proc.time()
  wb    <- split_excel_to_sheets(f, n = 10)
  elapsed <- (proc.time() - t0)[["elapsed"]]
  expect_lte(elapsed, 30)
  expect_equal(length(names(wb)), 10L)
})
