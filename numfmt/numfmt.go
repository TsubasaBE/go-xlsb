// Package numfmt renders Excel cell values to their display string using
// a number format string.  It is the rendering engine behind
// [workbook.Workbook.FormatCell] and [worksheet.Worksheet.FormatCell].
//
// The public entry point is [FormatValue].  All format-string parsing is
// delegated to [github.com/xuri/nfp]; this package only implements the
// rendering logic on top of the resulting token stream.
package numfmt

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"
	"unicode"

	"github.com/xuri/nfp"

	"github.com/TsubasaBE/go-xlsb/styles"
)

// FormatValue renders a raw cell value v using the given number format.
//
//   - numFmtID is the numFmtId from the XF record (0 = General).
//   - fmtStr is the custom format string from the BrtFmt record; pass ""
//     for built-in IDs that have no custom override.
//   - date1904 should match [workbook.Workbook.Date1904].
//
// The dynamic type of v must be one of: nil, string, bool, float64.
// Any other type falls back to [fmt.Sprint].
func FormatValue(v any, numFmtID int, fmtStr string, date1904 bool) string {
	// Resolve the effective format string.
	effective := resolveFormat(numFmtID, fmtStr)

	// Type-specific short-circuits.
	switch val := v.(type) {
	case nil:
		return ""
	case string:
		// Text format (@) or General — return as-is.
		return val
	case bool:
		if val {
			return "TRUE"
		}
		return "FALSE"
	case float64:
		return formatFloat(val, numFmtID, effective, date1904)
	default:
		return fmt.Sprint(v)
	}
}

// ── format-string resolution ──────────────────────────────────────────────────

// resolveFormat returns the effective format string: the custom fmtStr when
// non-empty, the built-in string for numFmtID when known, or "General".
func resolveFormat(numFmtID int, fmtStr string) string {
	if fmtStr != "" {
		return fmtStr
	}
	if s, ok := styles.BuiltInNumFmt[numFmtID]; ok {
		return s
	}
	return "General"
}

// ── float64 dispatch ──────────────────────────────────────────────────────────

func formatFloat(val float64, numFmtID int, effective string, date1904 bool) string {
	if effective == "General" || numFmtID == 0 && effective == "General" {
		return renderGeneral(val)
	}

	// Parse the format string into sections.
	ps := nfp.NumberFormatParser()
	sections := ps.Parse(effective)
	if len(sections) == 0 {
		return renderGeneral(val)
	}

	// Determine which section applies.
	sec := selectSection(sections, val)

	// Date / elapsed path.
	if isDateFormat(numFmtID, effective) {
		return renderDateTime(val, sec, date1904)
	}

	// Number path.
	return renderNumber(val, sec, sections)
}

// selectSection picks the correct section based on the value's sign.
//
//	1 section  → applies to all values
//	2 sections → [0]=positive+zero  [1]=negative
//	3 sections → [0]=positive  [1]=negative  [2]=zero
//	4 sections → [0]=positive  [1]=negative  [2]=zero  [3]=text
func selectSection(sections []nfp.Section, val float64) nfp.Section {
	switch {
	case len(sections) == 1:
		return sections[0]
	case len(sections) == 2:
		if val < 0 {
			return sections[1]
		}
		return sections[0]
	default: // 3 or 4
		switch {
		case val > 0:
			return sections[0]
		case val < 0:
			return sections[1]
		default: // zero
			return sections[2]
		}
	}
}

// ── General rendering ─────────────────────────────────────────────────────────

// renderGeneral formats a float64 in Excel's "General" style:
//   - integer values are rendered without a decimal point
//   - fractional values use Go's shortest-representation float
func renderGeneral(val float64) string {
	if math.IsNaN(val) || math.IsInf(val, 0) {
		return strconv.FormatFloat(val, 'G', -1, 64)
	}
	if val == math.Trunc(val) && math.Abs(val) < 1e15 {
		return strconv.FormatInt(int64(val), 10)
	}
	return strconv.FormatFloat(val, 'G', -1, 64)
}

// ── date-format detection (mirrors styles.isDateFormatID) ────────────────────

// isDateFormat reports whether the format is a date/datetime format.
// Mirrors the logic in styles.isDateFormatID and xlsb.IsDateFormat.
func isDateFormat(id int, fmtStr string) bool {
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
	if id < 164 && id != 0 {
		return false
	}
	// For custom formats (id >= 164) or id==0 with a non-"General" fmtStr,
	// scan unquoted content for date token characters.
	inDoubleQuote := false
	inBracket := false
	for _, ch := range fmtStr {
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

// ── date/time renderer ────────────────────────────────────────────────────────

// renderDateTime renders a date/time serial number using the tokens in sec.
// serial is the raw Excel serial (fractional days since the epoch).
func renderDateTime(serial float64, sec nfp.Section, date1904 bool) string {
	// Convert to time.Time for calendar tokens.
	t, err := convertSerial(serial, date1904)
	if err != nil {
		// Fallback: render the raw number.
		return renderGeneral(serial)
	}

	// Pre-scan to determine if any AM/PM token is present — affects hour rendering.
	hasAmPm := false
	for _, tok := range sec.Items {
		if tok.TType == nfp.TokenTypeDateTimes {
			upper := strings.ToUpper(tok.TValue)
			if upper == "AM/PM" || upper == "A/P" {
				hasAmPm = true
				break
			}
		}
	}

	var sb strings.Builder
	lastWasHour := false

	for _, tok := range sec.Items {
		switch tok.TType {

		case nfp.TokenTypeDateTimes:
			upper := strings.ToUpper(tok.TValue)
			s := renderDateToken(upper, t, serial, hasAmPm, lastWasHour)
			sb.WriteString(s)
			// Track whether this token was an hour (H / HH) for M/MM disambiguation.
			lastWasHour = upper == "H" || upper == "HH"

		case nfp.TokenTypeElapsedDateTimes:
			// Elapsed tokens operate on the raw serial (fractional days).
			upper := strings.ToUpper(tok.TValue)
			s := renderElapsed(upper, serial)
			sb.WriteString(s)
			// An elapsed hour token ([h] or [hh]) acts like a regular hour
			// token for M/MM disambiguation: the next M/MM should be minutes.
			lastWasHour = upper == "H" || upper == "HH"

		case nfp.TokenTypeLiteral:
			// Quoted text or escape sequences — emit the value verbatim.
			// Do NOT reset lastWasHour: a literal separator (e.g. ":") between
			// an hour token and a following M/MM must not break the
			// minute-vs-month disambiguation.
			sb.WriteString(tok.TValue)

		default:
			// Ignore colour codes, conditions, alignment, etc.
			lastWasHour = false
		}
	}

	// Guard: if no token produced any output (e.g. the format string contained
	// only unrecognised or purely decorative tokens), fall back to the raw
	// serial so the numeric value is never silently dropped.
	if sb.Len() == 0 {
		return renderGeneral(serial)
	}
	return sb.String()
}

// renderDateToken renders a single date/time token value (already upper-cased).
func renderDateToken(upper string, t time.Time, serial float64, hasAmPm bool, lastWasHour bool) string {
	switch upper {
	// ── year ────────────────────────────────────────────────────────────────
	case "YYYY":
		return fmt.Sprintf("%04d", t.Year())
	case "YY":
		return fmt.Sprintf("%02d", t.Year()%100)

	// ── month / minute (disambiguated by lastWasHour) ────────────────────
	case "MMMM":
		return t.Month().String() // "January" … "December"
	case "MMM":
		return t.Month().String()[:3] // "Jan" … "Dec"
	case "MM":
		if lastWasHour {
			return fmt.Sprintf("%02d", t.Minute())
		}
		return fmt.Sprintf("%02d", int(t.Month()))
	case "M":
		if lastWasHour {
			return strconv.Itoa(t.Minute())
		}
		return strconv.Itoa(int(t.Month()))

	// ── day ─────────────────────────────────────────────────────────────────
	case "DDDD":
		return t.Weekday().String() // "Sunday" … "Saturday"
	case "DDD":
		return t.Weekday().String()[:3] // "Sun" … "Sat"
	case "DD":
		return fmt.Sprintf("%02d", t.Day())
	case "D":
		return strconv.Itoa(t.Day())

	// ── hour ─────────────────────────────────────────────────────────────────
	case "HH":
		h := t.Hour()
		if hasAmPm {
			h = h % 12
			if h == 0 {
				h = 12
			}
		}
		return fmt.Sprintf("%02d", h)
	case "H":
		h := t.Hour()
		if hasAmPm {
			h = h % 12
			if h == 0 {
				h = 12
			}
		}
		return strconv.Itoa(h)

	// ── second ───────────────────────────────────────────────────────────────
	case "SS":
		return fmt.Sprintf("%02d", t.Second())
	case "S":
		return strconv.Itoa(t.Second())

	// ── AM/PM ────────────────────────────────────────────────────────────────
	case "AM/PM":
		if t.Hour() < 12 {
			return "AM"
		}
		return "PM"
	case "A/P":
		if t.Hour() < 12 {
			return "A"
		}
		return "P"
	}
	return ""
}

// renderElapsed renders an elapsed-time token (h, hh, mm, ss — as emitted by
// the nfp parser with brackets stripped) using the raw serial (fractional days).
func renderElapsed(upper string, serial float64) string {
	switch upper {
	case "H", "HH":
		return strconv.Itoa(int(serial * 24))
	case "MM":
		return fmt.Sprintf("%02d", int(serial*24*60)%60)
	case "M":
		return strconv.Itoa(int(serial*24*60) % 60)
	case "SS":
		return fmt.Sprintf("%02d", int(serial*24*3600)%60)
	case "S":
		return strconv.Itoa(int(serial*24*3600) % 60)
	}
	return ""
}

// convertSerial converts an Excel serial to time.Time, handling both date
// systems.  It mirrors xlsb.ConvertDateEx without importing the root package
// to keep the import graph simple.
func convertSerial(serial float64, date1904 bool) (time.Time, error) {
	if math.IsNaN(serial) || math.IsInf(serial, 0) || serial < 0 {
		return time.Time{}, fmt.Errorf("numfmt: invalid serial %v", serial)
	}
	fracSec := int64(math.Round((serial - math.Trunc(serial)) * 86400))
	if fracSec < 0 {
		fracSec = 0
	} else if fracSec > 86399 {
		fracSec = 86399
	}
	if date1904 {
		base := time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)
		intPart := int(serial)
		return base.Add(time.Duration(intPart)*24*time.Hour + time.Duration(fracSec)*time.Second), nil
	}
	base := time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)
	intPart := int(serial)
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

// ── number renderer ───────────────────────────────────────────────────────────

// renderNumber renders a numeric (non-date) float64 value using the token
// section sec.  sections is the full parsed set (needed to check whether the
// negative section has its own sign tokens).
func renderNumber(val float64, sec nfp.Section, sections []nfp.Section) string {
	// ── pass 1: collect format metadata ──────────────────────────────────────
	type meta struct {
		hasPercent      bool
		hasThousands    bool
		decZeros        int // count of '0' placeholders after decimal point
		decHashes       int // count of '#' placeholders after decimal point
		intZeros        int // count of '0' placeholders before decimal point
		hasDecimal      bool
		hasExplicitSign bool // literal '+' or '-' in the section
	}
	var m meta
	afterDecimal := false
	for _, tok := range sec.Items {
		switch tok.TType {
		case nfp.TokenTypePercent:
			m.hasPercent = true
		case nfp.TokenTypeThousandsSeparator:
			m.hasThousands = true
		case nfp.TokenTypeDecimalPoint:
			m.hasDecimal = true
			afterDecimal = true
		case nfp.TokenTypeZeroPlaceHolder:
			if afterDecimal {
				m.decZeros += len(tok.TValue)
			} else {
				m.intZeros += len(tok.TValue)
			}
		case nfp.TokenTypeHashPlaceHolder:
			if afterDecimal {
				m.decHashes += len(tok.TValue)
			}
		case nfp.TokenTypeLiteral:
			tv := tok.TValue
			if tv == "+" || tv == "-" {
				m.hasExplicitSign = true
			}
		}
	}
	totalDecPlaces := m.decZeros + m.decHashes

	// ── apply scaling ─────────────────────────────────────────────────────────
	absVal := math.Abs(val)
	if m.hasPercent {
		absVal *= 100
	}

	// ── format the absolute value ─────────────────────────────────────────────
	var intStr, fracStr string
	if m.hasDecimal {
		// Format with the required number of decimal places.
		formatted := strconv.FormatFloat(absVal, 'f', totalDecPlaces, 64)
		dotIdx := strings.IndexByte(formatted, '.')
		if dotIdx >= 0 {
			intStr = formatted[:dotIdx]
			fracStr = formatted[dotIdx+1:]
		} else {
			intStr = formatted
			fracStr = strings.Repeat("0", totalDecPlaces)
		}
		// Trim trailing zeros beyond what '0' placeholders require (# placeholders).
		if m.decHashes > 0 && len(fracStr) > m.decZeros {
			trimTo := len(fracStr)
			for trimTo > m.decZeros && trimTo > 0 && fracStr[trimTo-1] == '0' {
				trimTo--
			}
			fracStr = fracStr[:trimTo]
		}
	} else {
		intStr = strconv.FormatFloat(absVal, 'f', 0, 64)
	}

	// ── apply integer zero-padding ────────────────────────────────────────────
	for len(intStr) < m.intZeros {
		intStr = "0" + intStr
	}

	// ── apply thousands separator ─────────────────────────────────────────────
	if m.hasThousands && len(intStr) > 3 {
		intStr = insertThousandsSep(intStr)
	}

	// ── determine sign ────────────────────────────────────────────────────────
	// When the negative section is selected (val<0) and it has no explicit sign
	// tokens, we must not prepend a minus (the section itself encodes the sign
	// visually, e.g. via parentheses).
	needsMinus := false
	if val < 0 && !m.hasExplicitSign {
		// Check whether we are in the negative section (index 1 when len>=2).
		// If the section has a Literal that looks like a sign wrapper we skip.
		if len(sections) < 2 {
			// Only one section: we must prepend the minus.
			needsMinus = true
		}
		// Two+ sections: negative section handles its own sign display.
	}

	// ── reassemble by walking tokens ──────────────────────────────────────────
	var sb strings.Builder
	if needsMinus {
		sb.WriteByte('-')
	}

	intConsumed := false
	fracConsumed := false
	afterDecimal = false

	for _, tok := range sec.Items {
		switch tok.TType {
		case nfp.TokenTypeLiteral:
			sb.WriteString(tok.TValue)

		case nfp.TokenTypeDecimalPoint:
			if len(fracStr) > 0 {
				sb.WriteByte('.')
			}
			afterDecimal = true

		case nfp.TokenTypeZeroPlaceHolder, nfp.TokenTypeHashPlaceHolder:
			if afterDecimal {
				if !fracConsumed {
					sb.WriteString(fracStr)
					fracConsumed = true
				}
			} else {
				if !intConsumed {
					sb.WriteString(intStr)
					intConsumed = true
				}
			}

		case nfp.TokenTypePercent:
			sb.WriteByte('%')

		case nfp.TokenTypeThousandsSeparator:
			// Already applied to intStr; don't emit the raw comma token.

		case nfp.TokenTypeColor, nfp.TokenTypeCondition,
			nfp.TokenTypeCurrencyLanguage, nfp.TokenTypeAlignment:
			// Ignore formatting-only tokens.
		}
	}

	// If the format had no placeholder tokens at all, just emit the integer.
	if !intConsumed && !afterDecimal {
		sb.WriteString(intStr)
	}

	// Guard: if nothing was written (e.g. a format string whose only tokens
	// are colours, conditions, or other non-output tokens), fall back to the
	// raw value so the numeric value is never silently dropped.
	if sb.Len() == 0 {
		return renderGeneral(val)
	}

	return sb.String()
}

// insertThousandsSep inserts commas every three digits from the right in an
// integer string (digits only, no sign).
func insertThousandsSep(s string) string {
	n := len(s)
	if n <= 3 {
		return s
	}
	var b strings.Builder
	b.Grow(n + n/3)
	rem := n % 3
	if rem == 0 {
		rem = 3
	}
	b.WriteString(s[:rem])
	for i := rem; i < n; i += 3 {
		b.WriteByte(',')
		b.WriteString(s[i : i+3])
	}
	return b.String()
}

// isDigit reports whether b is an ASCII digit.
func isDigit(b byte) bool { return b >= '0' && b <= '9' }

// suppress unused-import warning for unicode (used in future extension points).
var _ = unicode.IsDigit
