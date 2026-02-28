// Package styles holds the resolved number-format metadata parsed from
// xl/styles.bin.  It is a deliberately small, import-cycle-free package so
// that both workbook/ and worksheet/ can depend on it without introducing
// circular imports.
package styles

// XFStyle holds the resolved formatting information for one XF (cell-format)
// index as read from the CellXfs table in xl/styles.bin.
type XFStyle struct {
	// NumFmtID is the numFmtId stored in the BrtXF record.  Values 0–163 are
	// built-in Excel formats; values ≥ 164 are custom formats defined by a
	// BrtFmt record in the same file.
	NumFmtID int
	// FormatStr is the raw format string from the corresponding BrtFmt record.
	// It is empty for built-in IDs that have no custom override.
	FormatStr string
}

// StyleTable maps XF index → XFStyle.  The slice index is the 0-based XF
// index as stored in cell records (Cell.Style).
type StyleTable []XFStyle

// IsDate reports whether the XF at index s represents a date or datetime
// number format.  It returns false when s is out of range or when styles
// information is unavailable (nil / empty table).
//
// This method uses the broader internal date-detection logic that includes
// time-only built-in IDs 18–21 (h:mm AM/PM etc.).  The public
// [xlsb.IsDateFormat] function deliberately excludes those IDs because they
// carry no calendar-date component.  Use IsDate when you need to decide
// whether FormatCell will render a cell as a date or time string.
//
// The logic is intentionally duplicated from workbook.isDateFormatID and
// xlsb.IsDateFormat to avoid import cycles.  All internal copies must stay in
// sync.
func (st StyleTable) IsDate(s int) bool {
	if s < 0 || s >= len(st) {
		return false
	}
	return isDateFormatID(st[s].NumFmtID, st[s].FormatStr)
}

// FmtStr returns the raw format string for style index s, or an empty
// string when s is out of range.
func (st StyleTable) FmtStr(s int) string {
	if s < 0 || s >= len(st) {
		return ""
	}
	return st[s].FormatStr
}

// BuiltInNumFmt maps built-in numFmtId values to their canonical format
// strings as defined by ECMA-376 §18.8.30.  IDs 27–36 and 50–58 are
// locale-specific (CJK/Thai) in the spec; the entries here are neutral
// Western fallbacks used when no custom BrtFmt record overrides the ID in
// the file.  This ensures the serial is always rendered as a human-readable
// date rather than a raw number.
var BuiltInNumFmt = map[int]string{
	0:  "General",
	1:  "0",
	2:  "0.00",
	3:  "#,##0",
	4:  "#,##0.00",
	5:  `($#,##0_);($#,##0)`,
	6:  `($#,##0_);[Red]($#,##0)`,
	7:  `($#,##0.00_);($#,##0.00)`,
	8:  `($#,##0.00_);[Red]($#,##0.00)`,
	9:  "0%",
	10: "0.00%",
	11: "0.00E+00",
	12: "# ?/?",
	13: "# ??/??",
	14: "mm-dd-yy",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm AM/PM",
	19: "h:mm:ss AM/PM",
	20: "hh:mm",
	21: "hh:mm:ss",
	22: "m/d/yy hh:mm",
	// IDs 27–36: locale-specific CJK date formats.  Real files embed the
	// locale string via BrtFmt records (which override this table).  These
	// neutral Western fallbacks are used when no BrtFmt override is present.
	27: "MM-DD-YYYY",
	28: "D-MMM-YY",
	29: "D-MMM-YY",
	30: "M/D/YY",
	31: "YYYY-M-D",
	32: "H:MM",
	33: "H:MM:SS",
	34: "H:MM AM/PM",
	35: "H:MM:SS AM/PM",
	36: "MM-DD-YYYY",
	37: `(#,##0_);(#,##0)`,
	38: `(#,##0_);[Red](#,##0)`,
	39: `(#,##0.00_);(#,##0.00)`,
	40: `(#,##0.00_);[Red](#,##0.00)`,
	41: `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)`,
	42: `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)`,
	43: `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`,
	44: `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)`,
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mm:ss.0",
	48: "##0.0E+0",
	49: "@",
	// IDs 50–58: locale-specific CJK date formats (variant set).  Same
	// fallback strategy as IDs 27–36 above.
	50: "MM-DD-YYYY",
	51: "D-MMM-YY",
	52: "H:MM AM/PM",
	53: "H:MM:SS AM/PM",
	54: "D-MMM-YY",
	55: "H:MM AM/PM",
	56: "H:MM:SS AM/PM",
	57: "MM-DD-YYYY",
	58: "D-MMM-YY",
}

// ── date-format detection (copy of workbook.isDateFormatID / xlsb.IsDateFormat) ─

// isDateFormatID reports whether the given numFmtId (and optional custom
// format string) represents a date or datetime format.  This is a shared
// internal copy used by styles and the rendering engine; all internal copies
// (styles, workbook, numfmt) must stay in sync.
//
// Unlike the public xlsb.IsDateFormat, this function treats time-only built-in
// IDs 18–21 as date/time formats so that cells with time-only number formats
// are handled correctly in [StyleTable.IsDate] and during rendering.
func isDateFormatID(id int, formatStr string) bool {
	switch {
	case id >= 14 && id <= 22:
		// IDs 14-17: date formats (m/d/yy, d-mmm-yy, d-mmm, mmm-yy)
		// IDs 18-21: time formats (h:mm AM/PM, h:mm:ss AM/PM, h:mm, h:mm:ss)
		// ID 22: datetime format (m/d/yy h:mm)
		return true
	case id >= 27 && id <= 36:
		return true
	case id >= 45 && id <= 47:
		return true
	case id >= 50 && id <= 58:
		return true
	}
	if id < 164 {
		return false
	}
	// Custom format: scan the unquoted portion for date/time token characters.
	inDoubleQuote := false
	inBracket := false
	var prev rune
	for _, ch := range formatStr {
		switch {
		case inDoubleQuote:
			if ch == '"' {
				inDoubleQuote = false
			}
		case inBracket:
			if ch == ']' {
				inBracket = false
			}
		case ch == '"':
			inDoubleQuote = true
		case ch == '[':
			inBracket = true
		case ch == 'd' || ch == 'D' ||
			ch == 'm' || ch == 'M' ||
			ch == 'y' || ch == 'Y' ||
			ch == 'h' || ch == 'H' ||
			ch == 's' || ch == 'S':
			return true
		case ch == 'e' || ch == 'E':
			if prev != '0' && prev != '#' && prev != '?' && prev != '.' {
				return true
			}
		}
		if !inDoubleQuote && !inBracket {
			prev = ch
		}
	}
	return false
}
