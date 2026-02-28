# go-xlsb

A Go port of [pyxlsb](https://github.com/willtrnr/pyxlsb) by [William Turner](https://github.com/willtrnr). All credit for the original design, format research, and reference implementation goes to him.

Reads Microsoft Excel Binary Workbook (`.xlsb`) files. Pure Go, no CGO.

## Installation

```sh
go get github.com/TsubasaBE/go-xlsb
```

Requires Go 1.22 or later (uses range-over-func iterators).

> This library is intentionally read-only. Writing `.xlsb` files is out of scope and will never be implemented.

## Status

This package is incomplete. The core reading path works — cell values, number formatting, dates, merged cells, hyperlinks — but large parts of the XLSB spec are not yet parsed.

### Implemented

Cell values: blank, number, boolean, string (shared string table), error, and formula results for all of the above. Rich text strings are read as plain text; the individual formatting runs are discarded.

Worksheet metadata: sheet list with visibility levels, used-range dimension, column definitions (width and style), merged cell ranges, and hyperlinks. Hyperlinks are stored as a `[row, col] -> rId` map; there is currently no public method to resolve an `rId` to its URL.

Number formatting via `wb.FormatCell`: integer and decimal rendering, thousands separator, percent, literal prefix/suffix, multi-section formats, date and datetime formats (built-in and custom), elapsed time (`[h]:mm:ss`), AM/PM, day-of-week and month names, and both the 1900 and 1904 date systems.

### Not implemented

Cell styling: font (name, size, weight, style, underline, color), fill (background color and pattern), borders, and alignment (horizontal, vertical, wrap, indent, rotation). The XF records are parsed only far enough to extract the number format ID; everything else is skipped.

Named cell styles, conditional formatting, and differential formatting (Dxfs) are not parsed.

Worksheet features not yet read: row height, hidden rows, default row and column sizes, sheet view properties (freeze panes, zoom, active cell), page setup (margins, print options, headers and footers), tables, AutoFilter, and comments.

Chart sheets open without error but always return zero rows. No chart data is exposed.

Workbook features not yet read: defined names, external references, data connections, OLE objects, and drawings.

Password-protected files are not supported.

Number format gaps: accounting alignment (column-fill `*` and alignment `_` tokens produce a single character rather than actual column padding).

## Usage

```go
import "github.com/TsubasaBE/go-xlsb"

wb, err := xlsb.Open("Book1.xlsb")
if err != nil {
    log.Fatal(err)
}
defer wb.Close()

fmt.Println(wb.Sheets()) // ["Sheet1", "Sheet2"]

sheet, err := wb.Sheet(1) // 1-based index
if err != nil {
    log.Fatal(err)
}

for row := range sheet.Rows(false) {
    for _, cell := range row {
        fmt.Printf("(%d,%d) = %v\n", cell.R, cell.C, cell.V)
    }
}
```

Open from an `io.ReaderAt` (e.g. an in-memory buffer):

```go
data, _ := os.ReadFile("Book1.xlsb")
wb, err := xlsb.OpenReader(bytes.NewReader(data), int64(len(data)))
```

## API

### `xlsb` package

| Symbol | Description |
|---|---|
| `Version string` | Current library version |
| `Open(name string) (*workbook.Workbook, error)` | Open a `.xlsb` file by path |
| `OpenReader(r io.ReaderAt, size int64) (*workbook.Workbook, error)` | Open from any `io.ReaderAt` |
| `ConvertDate(date float64) (time.Time, error)` | Convert an Excel date serial to `time.Time` (1900 system) |
| `ConvertDateEx(date float64, date1904 bool) (time.Time, error)` | Convert a date serial respecting the workbook's date system |
| `IsDateFormat(id int, formatStr string) bool` | Report whether a number-format ID represents a date/datetime format |

### `workbook.Workbook`

| Field / Method | Description |
|---|---|
| `Date1904 bool` | True when the workbook uses the 1904 date system |
| `Styles styles.StyleTable` | Full XF style table parsed from `xl/styles.bin` |
| `Sheets() []string` | Ordered list of all sheet names (visible and hidden) |
| `Sheet(idx int) (*worksheet.Worksheet, error)` | 1-based index lookup |
| `SheetByName(name string) (*worksheet.Worksheet, error)` | Case-insensitive name lookup |
| `SheetVisible(name string) bool` | Report whether a named sheet is visible |
| `SheetVisibility(name string) int` | Return visibility level: `SheetVisible` (0), `SheetHidden` (1), `SheetVeryHidden` (2), or -1 if not found |
| `FormatCell(v any, styleIdx int) string` | Render a raw cell value to its Excel display string |
| `Close() error` | Release the underlying file handle |

#### Sheet visibility constants

```go
workbook.SheetVisible     = 0 // tab is visible
workbook.SheetHidden      = 1 // hidden; user can unhide via Excel UI
workbook.SheetVeryHidden  = 2 // hidden; only accessible via VBA / programmatic access
```

### `worksheet.Worksheet`

| Field / Method | Description |
|---|---|
| `Name string` | Sheet display name |
| `Dimension *Dimension` | Used cell range (`nil` if not present in the file) |
| `Cols []Col` | Column definitions |
| `Hyperlinks map[[2]int]string` | `[row, col]` to relationship ID |
| `MergeCells []MergeArea` | All merged cell ranges in the sheet |
| `Rows(sparse bool) func(yield func([]Cell) bool)` | Range-over-func row iterator |
| `FormatCell(cell Cell) string` | Render a cell to its Excel display string (delegates to `wb.FormatCell`) |

`Rows(false)` emits empty rows between data rows, matching pyxlsb's default behaviour. Pass `true` to skip empty rows.

### `worksheet.Cell`

```go
type Cell struct {
    R     int // 0-based row index
    C     int // 0-based column index
    V     any // nil | string | float64 | bool
    Style int // 0-based XF index into wb.Styles
}
```

### `worksheet.Dimension`

```go
type Dimension struct {
    R int // first row index (0-based)
    C int // first column index (0-based)
    H int // height (number of rows)
    W int // width (number of columns)
}
```

### `worksheet.Col`

```go
type Col struct {
    C1    int
    C2    int
    Width float64
    Style int
}
```

### `worksheet.MergeArea`

```go
type MergeArea struct {
    R, C, H, W int // top-left anchor (0-based row/col), height, width
}
```

### `styles.StyleTable`

`wb.Styles` is a `styles.StyleTable` (a `[]styles.XFStyle` slice indexed by XF index).

| Method | Description |
|---|---|
| `IsDate(s int) bool` | Report whether XF index `s` maps to a date/datetime format |
| `FmtStr(s int) string` | Return the raw custom format string for XF index `s` |

```go
type XFStyle struct {
    NumFmtID  int    // numFmtId from the BrtXF record (0–163 built-in, ≥164 custom)
    FormatStr string // custom format string; empty for built-in IDs
}
```

`styles.BuiltInNumFmt` is a `map[int]string` of canonical format strings for built-in IDs (0–58) as defined by ECMA-376 §18.8.30.

## Cell formatting

`Rows` always returns raw values (`nil`, `string`, `float64`, or `bool`). To obtain the display string that Excel would show — respecting number formats, date formats, elapsed time, literal prefixes, decimal precision, and so on — call `wb.FormatCell`:

```go
for row := range sheet.Rows(false) {
    for _, cell := range row {
        raw       := cell.V
        formatted := wb.FormatCell(cell.V, cell.Style)
        fmt.Printf("raw=%v  formatted=%s\n", raw, formatted)
    }
}
```

`sheet.FormatCell(cell)` is a convenience wrapper that accepts a `Cell` directly and delegates to `wb.FormatCell`.

## Dates

Excel stores dates as floating-point serial numbers (days since 1900-01-00, fractional part is time of day).

`wb.FormatCell` handles date rendering automatically when the cell's number format is a date or datetime format. For direct access to the underlying `time.Time` value, use `wb.Styles.IsDate` to detect date cells and `ConvertDateEx` to convert them:

```go
for row := range sheet.Rows(false) {
    for _, cell := range row {
        if f, ok := cell.V.(float64); ok && wb.Styles.IsDate(cell.Style) {
            t, err := xlsb.ConvertDateEx(f, wb.Date1904)
            if err == nil {
                fmt.Println(t)
            }
        }
    }
}
```

`ConvertDate` is a convenience wrapper for the common 1900 date system (`wb.Date1904 == false`). `ConvertDateEx` additionally handles the 1904 date system used by some workbooks (notably those originally created on macOS).

Both functions handle the Lotus 1-2-3 leap-year bug that Excel carries forward in the 1900 system: serial 60 is a phantom date (1900-02-29 never existed) but Excel counts it anyway, so serial 61 onwards is off by one without compensation. The library compensates automatically.

## Credits

This library is a port of [pyxlsb](https://github.com/willtrnr/pyxlsb), written by [William Turner](https://github.com/willtrnr). The BIFF12 parsing logic, shared string table handling, date conversion, and overall design are all derived from his work.

## License

This library is licensed under the [GNU Lesser General Public License v3 or later (LGPL-3.0-or-later)](LICENSE), matching the license of [pyxlsb](https://github.com/willtrnr/pyxlsb) from which it is derived. The full LGPL-3.0 text is in [LICENSE](LICENSE) and the underlying GPL-3.0 text is in [COPYING](COPYING).
