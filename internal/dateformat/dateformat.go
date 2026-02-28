// Package dateformat provides shared date-format detection helpers used by
// multiple packages within go-xlsb.
//
// It exists solely to eliminate duplicated code; it has no public-API
// contract of its own.  All callers are within the same module.
package dateformat

// IsBuiltInDateID reports whether id is a built-in Excel numFmtId that
// represents a date, datetime, or time format.
//
// The recognised IDs follow ECMA-376 §18.8.30:
//
//	14–22   date and time formats (IDs 18–21 are time-only)
//	27–36   locale-specific CJK date formats
//	45–47   elapsed-time / seconds formats
//	50–58   locale-specific CJK date formats (variant set)
//
// Unlike the public [xlsb.IsDateFormat], this function intentionally includes
// the time-only IDs 18–21 so that the rendering engine and [styles.StyleTable]
// treat them as date/time values requiring serial-number conversion.
func IsBuiltInDateID(id int) bool {
	switch {
	case id >= 14 && id <= 22:
		// IDs 14-17: date formats (m/d/yy, d-mmm-yy, d-mmm, mmm-yy)
		// IDs 18-21: time formats (h:mm AM/PM, h:mm:ss AM/PM, h:mm, h:mm:ss)
		// ID 22:     datetime format (m/d/yy h:mm)
		return true
	case id >= 27 && id <= 36:
		return true
	case id >= 45 && id <= 47:
		return true
	case id >= 50 && id <= 58:
		return true
	}
	return false
}

// ScanFormatStr scans the unquoted portion of a custom number-format string
// for date/time token characters and returns true if any are found.
//
// The following characters are treated as date/time tokens when they appear
// outside double-quoted literals and outside square-bracket sections:
//
//   - d, D — day
//   - m, M — month
//   - y, Y — year
//   - h, H — hour
//   - s, S — second
//   - e, E — Japanese era (only when NOT preceded by a digit placeholder
//     0, #, ?, or .)
func ScanFormatStr(formatStr string) bool {
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
			// E/e is a scientific-notation exponent marker when preceded by a
			// digit placeholder (0, #, ?) — in that context it is NOT a date
			// token.  Only treat it as the Japanese era date token otherwise.
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
