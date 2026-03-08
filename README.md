A detailed description of the `split_excel_to_sheets()` function:

---

### `split_excel_to_sheets()`

**Purpose**
Reads a source Excel file, partitions its data into `n` equal-sized chunks, and writes each chunk as a separate styled sheet within a single returned workbook object.

---

### Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `file_path` | character | required | Path to the source `.xlsx` file to read |
| `n` | integer | required | Number of splits — determines how many sheets are created |
| `sheet` | integer or character | `1` | Source sheet to read from — by index or name |
| `output_path` | character or NULL | `NULL` | If provided, saves the workbook to disk at this path; if `NULL`, returns in-memory only |
| `sheet_prefix` | character | `"Part"` | Prefix for generated sheet names — e.g. `"Part"` → `Part_1`, `Part_2`, … |
| `header_style` | logical | `TRUE` | Applies a dark blue styled header row to each sheet when `TRUE` |
| `col_widths` | character or numeric | `"auto"` | `"auto"` for openxlsx auto-sizing, a numeric vector for fixed widths, or `NULL` to skip entirely |

---

### What it does — step by step

**Step 1 — Validate**
Checks that the file exists, and that `n` is a positive whole number. If `n` exceeds the number of rows, it downgrades `n` with a warning rather than erroring out.

**Step 2 — Read**
Reads the source sheet using `openxlsx::read.xlsx()` with `detectDates = FALSE` for speed, then immediately converts the result to a `data.table` in-place via `setDT()`.

**Step 3 — Chunk**
Uses `cut()` to assign each row to a chunk group in a single vectorised operation, then `data.table::split()` to partition the table into a named list — no loops, no repeated subsetting.

**Step 4 — Write sheets**
Iterates over each chunk, creates a named sheet, writes the data with `writeData()`, applies a pre-built header style and body style, sets column widths, and freezes the top row. The header `createStyle()` object is built once and reused across all sheets.

**Step 5 — Save (optional)**
If `output_path` is provided, calls `saveWorkbook()`. Otherwise, the step is skipped and the workbook lives in memory only.

---

### Return value
Returns the `openxlsx` workbook object **invisibly**, meaning it can be captured and further manipulated — additional sheets, summaries, or charts — before a final `saveWorkbook()` call.

---

### Timing
Each of the 4 phases is timed and printed to the console via `message()`, giving the caller visibility into where time is being spent across large files.
