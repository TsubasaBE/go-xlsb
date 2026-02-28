// Package xlsb provides a pure-Go reader for Microsoft Excel Binary Workbook
// (.xlsb) files.  No cgo is required.
//
// # Quick start
//
//	wb, err := xlsb.Open("Book1.xlsb")
//	if err != nil { ... }
//	defer wb.Close()
//
//	fmt.Println(wb.Sheets()) // ["Sheet1", "Sheet2"]
//
//	sheet, err := wb.Sheet(1)
//	if err != nil { ... }
//
//	for row := range sheet.Rows(false) {
//	    for _, cell := range row {
//	        fmt.Printf("(%d,%d) = %v\n", cell.R, cell.C, cell.V)
//	    }
//	}
//
// # Dates
//
// Excel stores dates as floating-point serial numbers.  Use [ConvertDate] to
// turn such a value into a [time.Time]:
//
//	if f, ok := cell.V.(float64); ok {
//	    t, err := xlsb.ConvertDate(f)
//	}
package xlsb

import (
	"fmt"
	"io"
	"math"
	"time"

	"github.com/TsubasaBE/go-xlsb/workbook"
)

// Open opens the named .xlsb file.  The caller must call Close on the returned
// Workbook when done.
func Open(name string) (*workbook.Workbook, error) {
	return workbook.Open(name)
}

// OpenReader reads an .xlsb workbook from an arbitrary [io.ReaderAt].
// size must equal the total byte length of the data.
func OpenReader(r io.ReaderAt, size int64) (*workbook.Workbook, error) {
	return workbook.OpenReader(r, size)
}

// ConvertDate converts an Excel date serial number to a [time.Time] value.
//
// Excel (and the BIFF12 format) represents dates as the number of days since
// 1900-01-00, with the fractional part representing the time of day.  Lotus
// 1-2-3 incorrectly treated 1900 as a leap year, so Excel perpetuates the bug:
// serial 60 is treated as 1900-02-29 (which never existed).  This function
// handles the three resulting branches exactly as pyxlsb does:
//
//   - serial == 0  → midnight on 1900-01-01
//   - serial >= 61 → subtract one day to compensate for the phantom leap day
//   - 1 ≤ serial ≤ 60 → no compensation (serial 60 yields 1900-03-01)
func ConvertDate(date float64) (time.Time, error) {
	if math.IsNaN(date) || math.IsInf(date, 0) {
		return time.Time{}, fmt.Errorf("xlsb: ConvertDate: invalid value %v", date)
	}
	if date < 0 {
		return time.Time{}, fmt.Errorf("xlsb: ConvertDate: negative serial %v not supported", date)
	}
	// Excel dates only reach serial 2,958,465 (year 9999).  Values above this
	// would overflow time.Duration arithmetic (int64 nanoseconds).
	const maxSerial = 2_958_466
	if date > maxSerial {
		return time.Time{}, fmt.Errorf("xlsb: ConvertDate: serial %v exceeds maximum supported value %d", date, maxSerial)
	}

	base := time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)
	intPart := int(date)
	// fractional seconds (sub-day time component).
	// Clamp to [0, 86399] to guard against floating-point rounding artefacts
	// near whole numbers (e.g. 0.9999999... rounding up to 86400, which would
	// add a phantom extra day on top of intPart).
	fracSec := int64(math.Round((date - math.Trunc(date)) * 24 * 60 * 60))
	if fracSec < 0 {
		fracSec = 0
	} else if fracSec > 86399 {
		fracSec = 86399
	}

	var t time.Time
	switch {
	case intPart == 0:
		// Serial 0 → 1900-01-01 plus fractional seconds
		t = time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC).Add(time.Duration(fracSec) * time.Second)
	case intPart >= 61:
		// Subtract 1 day to skip the phantom 1900-02-29
		t = base.Add(time.Duration(intPart-1)*24*time.Hour + time.Duration(fracSec)*time.Second)
	default:
		// Serials 1–60: no correction
		t = base.Add(time.Duration(intPart)*24*time.Hour + time.Duration(fracSec)*time.Second)
	}
	return t, nil
}
