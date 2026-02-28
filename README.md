# go-xlsb

A Go port of [pyxlsb](https://github.com/willtrnr/pyxlsb) by [William Turner](https://github.com/willtrnr). All credit for the original design, format research, and reference implementation goes to him.

Reads Microsoft Excel Binary Workbook (`.xlsb`) files. Pure Go, no CGO, no external dependencies.

## Installation

```sh
go get github.com/TsubasaBE/go-xlsb
```

Requires Go 1.22 or later (uses range-over-func iterators).

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
| `Sheets() []string` | Ordered list of all sheet names (visible and hidden) |
| `Sheet(idx int) (*worksheet.Worksheet, error)` | 1-based index lookup |
| `SheetByName(name string) (*worksheet.Worksheet, error)` | Case-insensitive name lookup |
| `SheetVisible(name string) bool` | Report whether a named sheet is visible |
| `SheetVisibility(name string) int` | Return visibility level: `SheetVisible` (0), `SheetHidden` (1), `SheetVeryHidden` (2), or -1 if not found |
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
| `IsDateCell(style int) bool` | Report whether a cell's XF style index maps to a date/datetime format |

`Rows(false)` emits empty rows between data rows, matching pyxlsb's default behaviour. Pass `true` to skip empty rows.

### `worksheet.Cell`

```go
type Cell struct {
    R     int // 0-based row index
    C     int // 0-based column index
    V     any // nil | string | float64 | bool
    Style int // 0-based XF index; use Worksheet.IsDateCell(cell.Style) to detect dates
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

### Dates

Excel stores dates as floating-point serial numbers (days since 1900-01-00, fractional part is time of day).

Use `ConvertDateEx` together with `wb.Date1904` and `sheet.IsDateCell` to correctly identify and convert date cells:

```go
for row := range sheet.Rows(false) {
    for _, cell := range row {
        if f, ok := cell.V.(float64); ok && sheet.IsDateCell(cell.Style) {
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
