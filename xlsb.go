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
// # Cell formatting
//
// [worksheet.Worksheet.Rows] always returns raw values (nil, string, float64, or bool).  To obtain
// the display string that Excel would show — respecting number formats, date
// formats, custom formats, and so on — call [workbook.Workbook.FormatCell]:
//
//	for row := range sheet.Rows(false) {
//	    for _, cell := range row {
//	        raw       := cell.V
//	        formatted := wb.FormatCell(cell.V, cell.Style)
//	        _, _ = raw, formatted
//	    }
//	}
//
// [worksheet.Worksheet.FormatCell] is a convenience wrapper on the sheet that
// accepts a [worksheet.Cell] directly.
//
// # Dates
//
// Excel stores dates as floating-point serial numbers.  [FormatCell] handles
// date rendering automatically when the cell's number format is a date or
// datetime format.  For direct access to the underlying [time.Time] value use
// [ConvertDateEx], passing wb.Date1904 so the correct date system is used:
//
//	if f, ok := cell.V.(float64); ok && wb.Styles.IsDate(cell.Style) {
//	    t, err := xlsb.ConvertDateEx(f, wb.Date1904)
//	}
//
// [ConvertDate] is a convenience wrapper for the common 1900 date system
// (Date1904 == false).
//
// # Format detection
//
// [IsDateFormat] checks whether a number-format ID (and optional custom format
// string) represents a date or datetime format.  It is a lower-level helper for
// callers that inspect format metadata without going through [workbook.Workbook.FormatCell].
package xlsb

import (
	"fmt"
	"io"
	"math"
	"time"

	"github.com/TsubasaBE/go-xlsb/internal/dateformat"
	"github.com/TsubasaBE/go-xlsb/workbook"
)

// Version is the current version of the go-xlsb library.
const Version = "1.1.0"

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
//
// The fractional-day component is converted to whole seconds using the same
// rounding algorithm as excelize (roundEpsilon + half-second rounding) so
// that this function produces identical results to [workbook.Workbook.FormatCell]
// for date/time cells.
func ConvertDate(date float64) (time.Time, error) {
	if math.IsNaN(date) || math.IsInf(date, 0) {
		return time.Time{}, fmt.Errorf("xlsb: ConvertDate: invalid value %v", date)
	}
	if date < 0 {
		return time.Time{}, fmt.Errorf("xlsb: ConvertDate: negative serial %v not supported", date)
	}
	// Excel dates only reach serial 2,958,465 (year 9999-12-31).  The constant
	// below is the exclusive upper bound (one above the last valid serial) used
	// in the comparison.  Values above maxSerial-1 would overflow time.Duration
	// arithmetic (int64 nanoseconds).
	const maxSerial = 2_958_466
	if date > maxSerial {
		return time.Time{}, fmt.Errorf("xlsb: ConvertDate: serial %v exceeds maximum supported value %d", date, maxSerial)
	}

	fracSec, dayRollover := serialToFracSec(date)

	base := time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)
	intPart := int(date) + dayRollover
	var t time.Time
	switch {
	case intPart == 0:
		t = time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC).Add(time.Duration(fracSec) * time.Second)
	case intPart >= 61:
		t = base.Add(time.Duration(intPart-1)*24*time.Hour + time.Duration(fracSec)*time.Second)
	default:
		t = base.Add(time.Duration(intPart)*24*time.Hour + time.Duration(fracSec)*time.Second)
	}
	return t, nil
}

// ConvertDateEx converts an Excel date serial number to a [time.Time] value,
// respecting the workbook's date system.
//
// Pass wb.Date1904 as the date1904 argument. When date1904 is false the
// function is identical to [ConvertDate] (1900 date system). When date1904 is
// true the workbook uses the 1904 date system:
//   - Serial 0 corresponds to 1904-01-01.
//   - Serials increase by one day per unit, with no phantom leap-day
//     correction (the Lotus 1-2-3 bug does not apply to the 1904 system).
func ConvertDateEx(date float64, date1904 bool) (time.Time, error) {
	if !date1904 {
		return ConvertDate(date)
	}
	if math.IsNaN(date) || math.IsInf(date, 0) {
		return time.Time{}, fmt.Errorf("xlsb: ConvertDateEx: invalid value %v", date)
	}
	if date < 0 {
		return time.Time{}, fmt.Errorf("xlsb: ConvertDateEx: negative serial %v not supported", date)
	}
	// In the 1904 system the maximum representable date is the same calendar
	// day as in the 1900 system.  Serial 0 = 1904-01-01, so the 1904 serials
	// are offset by 1462 days from the 1900 serials (4 years including one
	// leap year, 1904 itself).  The maximum 1900 serial is 2,958,465
	// (9999-12-31); subtracting 1462 gives the 1904-system maximum.
	const maxSerial = 2_958_466 - 1462
	if date > maxSerial {
		return time.Time{}, fmt.Errorf("xlsb: ConvertDateEx: serial %v exceeds maximum supported value %d", date, maxSerial)
	}

	fracSec, dayRollover := serialToFracSec(date)

	// Base: 1904-01-01. Serial 0 = 1904-01-01, serial 1 = 1904-01-02, etc.
	// No phantom leap-day correction is needed for the 1904 date system.
	base := time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)
	intPart := int(date) + dayRollover
	t := base.Add(time.Duration(intPart)*24*time.Hour + time.Duration(fracSec)*time.Second)
	return t, nil
}

// serialToFracSec converts the fractional-day part of an Excel serial to a
// whole-second count within the day (0–86399), plus a day-rollover flag (0 or 1).
//
// The algorithm mirrors excelize's timeFromExcelTime and numfmt.convertSerial:
//   - add roundEpsilon (1e-9) to the fractional day to avoid floating-point drift
//   - convert to nanoseconds and round to the nearest second (half-second rule)
//   - when rounding pushes the result to exactly 86400 s (midnight), roll over
//     to the next day rather than clamping
//
// Both ConvertDate and ConvertDateEx use this helper so they remain in sync
// with numfmt.convertSerial without a circular import.
func serialToFracSec(serial float64) (fracSec int64, dayRollover int) {
	const roundEpsilon = 1e-9
	fracDay := (serial - math.Trunc(serial)) + roundEpsilon
	const nanosInADay = float64(24 * 60 * 60 * 1e9)
	durNanos := time.Duration(fracDay * nanosInADay)
	ns := int(durNanos % time.Second)
	secs := int64(durNanos / time.Second)
	if ns > 500_000_000 {
		secs++
	}
	if secs < 0 {
		secs = 0
	}
	// When rounding pushes secs to 86400 (midnight), roll over to the next day.
	rollover := int(secs / 86400)
	secs = secs % 86400
	return secs, rollover
}

// IsDateFormat reports whether a number-format ID (and optional custom format
// string) represents a date or datetime format.
//
// id is the numFmtId stored in the XF record.  For built-in formats (id < 164)
// formatStr is ignored; for custom formats (id >= 164) formatStr must be the
// format string read from the BrtFmt record in xl/styles.bin.
//
// Built-in date/time IDs follow ECMA-376 §18.8.30.  This function recognises
// the following as date or datetime formats:
//
//	14–17, 22, 27–36, 45–47, 50–58
//
// Note: built-in time-only IDs 18–21 (h:mm AM/PM, h:mm:ss AM/PM, h:mm,
// h:mm:ss) are intentionally excluded; those formats carry no calendar date
// component and converting them to [time.Time] is usually not meaningful.
// Use the internal isDateFormatID copies (in workbook/ and styles/) when
// rendering number-formatted output that includes time-only formats.
//
// For custom formats the function scans the unquoted portion of formatStr for
// any of the characters d, D, m, M, y, Y, h, H.  Sections enclosed in double
// quotes or square brackets are skipped.
func IsDateFormat(id int, formatStr string) bool {
	// Built-in date/time numFmtIds.
	switch {
	case id >= 14 && id <= 17:
		return true
	case id == 22:
		return true
	case id >= 27 && id <= 36:
		return true
	case id >= 45 && id <= 47:
		return true
	case id >= 50 && id <= 58:
		return true
	}
	if id < 164 {
		return false // other built-in IDs are not dates
	}
	// Custom format: scan unquoted characters for date/time tokens.
	return dateformat.ScanFormatStr(formatStr)
}
