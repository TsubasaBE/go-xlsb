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
// The logic is intentionally duplicated from workbook.isDateFormatID and
// xlsb.IsDateFormat to avoid import cycles.  All three copies must stay in
// sync.
func (st StyleTable) IsDate(s int) bool {
	if s < 0 || s >= len(st) {
		return false
	}
	return isDateFormatID(st[s].NumFmtID, st[s].FormatStr)
}

// FormatStr returns the raw format string for style index s, or an empty
// string when s is out of range.
func (st StyleTable) FmtStr(s int) string {
	if s < 0 || s >= len(st) {
		return ""
	}
	return st[s].FormatStr
}

// BuiltInNumFmt maps built-in numFmtId values (0–49) to their canonical
// format strings as defined by ECMA-376 §18.8.30.  IDs not present in this
// map are built-in IDs whose format string is locale-dependent or otherwise
// not representable as a static string.
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
	14: "MM-DD-YY",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm AM/PM",
	19: "h:mm:ss AM/PM",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yy h:mm",
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
}

// ── date-format detection (copy of workbook.isDateFormatID / xlsb.IsDateFormat) ─

// isDateFormatID reports whether the given numFmtId (and optional custom
// format string) represents a date or datetime format.  This is the third
// copy of this logic in the codebase; all copies must stay in sync.  The
// duplication is intentional to avoid circular imports between workbook/,
// worksheet/, and styles/.
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
			ch == 'h' || ch == 'H':
			return true
		}
	}
	return false
}
