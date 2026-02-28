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
| `ConvertDate(date float64) (time.Time, error)` | Convert an Excel date serial to `time.Time` |

### `workbook.Workbook`

| Method | Description |
|---|---|
| `Sheets() []string` | Ordered list of sheet names |
| `Sheet(idx int) (*worksheet.Worksheet, error)` | 1-based index lookup |
| `SheetByName(name string) (*worksheet.Worksheet, error)` | Case-insensitive name lookup |
| `Close() error` | Release the underlying file handle |

### `worksheet.Worksheet`

| Field / Method | Description |
|---|---|
| `Name string` | Sheet display name |
| `Dimension *Dimension` | Used cell range (`nil` if not present in the file) |
| `Cols []Col` | Column definitions |
| `Hyperlinks map[[2]int]string` | `[row, col]` to relationship ID |
| `MergeCells []MergeArea` | All merged cell ranges in the sheet |
| `Rows(sparse bool) func(yield func([]Cell) bool)` | Range-over-func row iterator |

`Rows(false)` emits empty rows between data rows, matching pyxlsb's default behaviour. Pass `true` to skip empty rows.

### `worksheet.Cell`

```go
type Cell struct {
    R int // 0-based row index
    C int // 0-based column index
    V any // nil | string | float64 | bool
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

Excel stores dates as floating-point serial numbers (days since 1900-01-00, fractional part is time of day). Use `ConvertDate` to get a `time.Time`:

```go
for row := range sheet.Rows(false) {
    for _, cell := range row {
        if f, ok := cell.V.(float64); ok {
            t, err := xlsb.ConvertDate(f)
            if err == nil {
                fmt.Println(t)
            }
        }
    }
}
```

`ConvertDate` handles the Lotus 1-2-3 leap-year bug that Excel carries forward: serial 60 is a phantom date (1900-02-29 never existed) but Excel counts it anyway, so serial 61 onwards is off by one without compensation. The library skips serial 60 and maps it to 1900-03-01, keeping all subsequent serials correct.

## Credits

This library is a port of [pyxlsb](https://github.com/willtrnr/pyxlsb), written by [William Turner](https://github.com/willtrnr). The BIFF12 parsing logic, shared string table handling, date conversion, and overall design are all derived from his work.

## License

This library is licensed under the [GNU Lesser General Public License v3 or later (LGPL-3.0-or-later)](LICENSE), matching the license of [pyxlsb](https://github.com/willtrnr/pyxlsb) from which it is derived. The full LGPL-3.0 text is in [LICENSE](LICENSE) and the underlying GPL-3.0 text is in [COPYING](COPYING).
