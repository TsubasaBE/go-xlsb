// Package numfmt renders Excel cell values to their display string using
// a number format string.  It is the rendering engine behind
// [workbook.Workbook.FormatCell] and [worksheet.Worksheet.FormatCell].
//
// The public entry point is [FormatValue].  All format-string parsing is
// delegated to the github.com/xuri/nfp package; this package only implements
// the rendering logic on top of the resulting token stream.
package numfmt

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/nfp"

	"github.com/TsubasaBE/go-xlsb/internal/dateformat"
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
		return formatString(val, effective)
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

// formatString renders a string cell value using the text section (section 4)
// of the format string.  If the format has a fourth section, any "@"
// TextPlaceHolder token is substituted with the cell value and surrounding
// literals are emitted; otherwise the string is returned as-is.
func formatString(val string, effective string) string {
	if effective == "General" || effective == "" || effective == "@" {
		return val
	}
	ps := nfp.NumberFormatParser()
	sections := ps.Parse(effective)
	// The text section is section index 3 (0-based).  With fewer sections the
	// format does not constrain text values.
	if len(sections) < 4 {
		return val
	}
	sec := sections[3]
	var sb strings.Builder
	for _, tok := range sec.Items {
		switch tok.TType {
		case nfp.TokenTypeTextPlaceHolder:
			sb.WriteString(val)
		case nfp.TokenTypeLiteral:
			sb.WriteString(tok.TValue)
		case nfp.TokenTypeAlignment:
			// _x → pad with one space (alignment hint, not measurable in plain text).
			sb.WriteByte(' ')
		case nfp.TokenTypeRepeatsChar:
			// *x → emit one instance of x (column-fill is a display concept only).
			if tok.TValue != "" {
				sb.WriteString(tok.TValue)
			}
		case nfp.TokenTypeCurrencyLanguage:
			sb.WriteString(currencySymbol(tok))
		case nfp.TokenTypeColor, nfp.TokenTypeCondition:
			// ignore decorative tokens
		}
	}
	if sb.Len() == 0 {
		return val
	}
	return sb.String()
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

// ── currency symbol extraction ────────────────────────────────────────────────

// currencySymbol extracts the currency-symbol sub-part from a
// TokenTypeCurrencyLanguage token.  The nfp parser stores the symbol in
// Parts[0] when the bracket contains a "$symbol" prefix before the "-locale"
// suffix.  If no symbol part is found the raw TValue (which includes the
// surrounding brackets) is returned as a fallback.
//
// Brackets that do not begin with "[$" are not currency brackets — they may be
// color brackets like "[Color 3]" or "[Red]" that nfp also classifies as
// CurrencyLanguage tokens.  Those return "" so they are treated as decorative.
func currencySymbol(tok nfp.Token) string {
	// Only real currency brackets start with "[$".  Color brackets like
	// "[Color 3]" or "[Red]" look similar but carry no printable symbol.
	raw := tok.TValue
	if len(raw) < 3 || raw[0] != '[' || raw[1] != '$' {
		return ""
	}
	for _, p := range tok.Parts {
		if p.Token.TType == nfp.TokenSubTypeCurrencyString {
			return p.Token.TValue
		}
	}
	// Fallback: strip the outer brackets and any "-locale" suffix.
	if len(raw) >= 2 && raw[0] == '[' && raw[len(raw)-1] == ']' {
		inner := raw[1 : len(raw)-1]
		// Strip the leading '$' that the bracket format uses.
		if len(inner) > 0 && inner[0] == '$' {
			inner = inner[1:]
		}
		// Drop everything from '-' onward (locale code).
		if idx := strings.IndexByte(inner, '-'); idx >= 0 {
			inner = inner[:idx]
		}
		return inner
	}
	return raw
}

// ── float64 dispatch ──────────────────────────────────────────────────────────

func formatFloat(val float64, numFmtID int, effective string, date1904 bool) string {
	if effective == "General" {
		return renderGeneral(val)
	}

	// Format ID 49 (or any format whose only section is "@") applied to a
	// numeric cell: Excel renders the number using General formatting, i.e.
	// the same as if no format were applied.  The "@" token is a text
	// placeholder and has no meaningful rendering for numeric values.
	if effective == "@" {
		return renderGeneral(val)
	}

	// Parse the format string into sections.
	ps := nfp.NumberFormatParser()
	sections := ps.Parse(effective)
	if len(sections) == 0 {
		return renderGeneral(val)
	}

	// Determine which section applies.
	// useAbs is true when the selected section was chosen by a conditional
	// bracket (e.g. [>=1000]) — in that case Excel renders math.Abs(val)
	// because the condition itself encodes the sign semantics.
	sec, useAbs := selectSection(sections, val)

	renderVal := val
	if useAbs {
		renderVal = math.Abs(val)
	}

	// Date / elapsed path.
	if isDateFormat(numFmtID, effective) {
		return renderDateTime(renderVal, sec, date1904)
	}

	// Number path.
	return renderNumber(renderVal, sec, sections)
}

// extractCondition returns the operator and numeric operand from the first
// TokenTypeCondition token found in sec, if any.  The operator is one of the
// six Excel condition operators: "<", "<=", ">", ">=", "<>", "=".
func extractCondition(sec nfp.Section) (op string, operand float64, ok bool) {
	for _, tok := range sec.Items {
		if tok.TType != nfp.TokenTypeCondition {
			continue
		}
		if len(tok.Parts) < 2 {
			continue
		}
		op = tok.Parts[0].Token.TValue
		operandStr := tok.Parts[1].Token.TValue
		var err error
		operand, err = strconv.ParseFloat(operandStr, 64)
		if err != nil {
			continue
		}
		return op, operand, true
	}
	return "", 0, false
}

// evalCondition returns true when val satisfies the condition op operand.
func evalCondition(val float64, op string, operand float64) bool {
	switch op {
	case "<":
		return val < operand
	case "<=":
		return val <= operand
	case ">":
		return val > operand
	case ">=":
		return val >= operand
	case "<>":
		return val != operand
	case "=":
		return val == operand
	}
	return false
}

// selectSection picks the correct section for val and reports whether the
// caller should use math.Abs(val) when rendering.
//
// Conditional formats (sections whose first token is a [condition] bracket)
// are evaluated first: sections are tested in order and the first match wins.
// When a conditional section is selected, useAbs is true — the condition
// itself encodes the sign semantics, so the renderer receives the magnitude.
//
// When no section carries a condition token the classic sign-based rules apply:
//
//	1 section  → applies to all values                          (useAbs=false)
//	2 sections → [0]=positive+zero  [1]=negative                (useAbs=false)
//	3 sections → [0]=positive  [1]=negative  [2]=zero           (useAbs=false)
//	4 sections → [0]=positive  [1]=negative  [2]=zero  [3]=text (useAbs=false)
//
// If a conditional format is present but no section matches val, the function
// falls through to a final unconditional section if one exists, or returns the
// last section as a safe fallback.
func selectSection(sections []nfp.Section, val float64) (sec nfp.Section, useAbs bool) {
	// First pass: collect which sections carry a condition token.
	hasAnyCondition := false
	for i := range sections {
		if op, operand, ok := extractCondition(sections[i]); ok {
			hasAnyCondition = true
			if evalCondition(val, op, operand) {
				return sections[i], true
			}
			_ = operand // already used above
			_ = op
		}
	}

	if hasAnyCondition {
		// No condition matched.  Excel's behaviour: use the last section that
		// has no condition token as a fallback, otherwise use the last section.
		for i := len(sections) - 1; i >= 0; i-- {
			if _, _, ok := extractCondition(sections[i]); !ok {
				return sections[i], false
			}
		}
		// All sections have conditions and none matched — last section is safest.
		return sections[len(sections)-1], true
	}

	// Classic sign-based selection (no conditions present).
	switch {
	case len(sections) == 1:
		return sections[0], false
	case len(sections) == 2:
		if val < 0 {
			return sections[1], false
		}
		return sections[0], false
	default: // 3 or 4
		switch {
		case val > 0:
			return sections[0], false
		case val < 0:
			return sections[1], false
		default: // zero
			return sections[2], false
		}
	}
}

// ── General rendering ─────────────────────────────────────────────────────────

// renderGeneral formats a float64 in Excel's "General" style:
//   - integer values (within safe int64 range) are rendered without a decimal point
//   - values whose absolute value is >= 1e11 use E+ scientific notation with up to
//     5 significant digits (matching Excel's General threshold behaviour)
//   - all other fractional values use 10 significant digits (G10), matching
//     excelize's RawCellValue=false behaviour.
func renderGeneral(val float64) string {
	if math.IsNaN(val) || math.IsInf(val, 0) {
		return strconv.FormatFloat(val, 'G', -1, 64)
	}
	abs := math.Abs(val)
	// Excel switches to E+ scientific notation for magnitudes >= 1e11.
	// This applies to both integers and fractional values.
	if abs >= 1e11 {
		s := strconv.FormatFloat(val, 'E', 5, 64)
		// Normalise Go's "E+06" → "E+6" (Excel drops leading zero in exponent
		// when it is a single digit, matching excelize output).
		if idx := strings.IndexByte(s, 'E'); idx >= 0 {
			sign := s[idx+1]
			expPart := s[idx+2:]
			// Trim leading zeros from exponent, keeping at least one digit.
			for len(expPart) > 1 && expPart[0] == '0' {
				expPart = expPart[1:]
			}
			s = s[:idx+1] + string(sign) + expPart
		}
		// Trim trailing zeros from the significand (e.g. "1.00000E+11" → "1E+11").
		if idx := strings.IndexByte(s, 'E'); idx >= 0 {
			mantissa := s[:idx]
			rest := s[idx:]
			if dotIdx := strings.IndexByte(mantissa, '.'); dotIdx >= 0 {
				trimmed := strings.TrimRight(mantissa, "0")
				trimmed = strings.TrimRight(trimmed, ".")
				s = trimmed + rest
			}
		}
		return s
	}
	// Use integer formatting when val is a whole number and small enough that
	// float64 can represent it exactly.  float64 has 53-bit mantissa precision,
	// so integers above 2^53 ≈ 9.0e15 may not round-trip through int64(val)
	// accurately.  1e15 is a conservative threshold safely below that limit.
	if val == math.Trunc(val) && abs < 1e15 {
		return strconv.FormatInt(int64(val), 10)
	}
	// Excel General format uses 10 significant digits (G10), not full float64
	// precision. This matches excelize's GetCellValue with RawCellValue=false.
	return strconv.FormatFloat(val, 'G', 10, 64)
}

// ── date-format detection (mirrors styles.isDateFormatID) ────────────────────

// isDateFormat reports whether the format is a date/datetime format.
// Mirrors the logic in styles.isDateFormatID and workbook.isDateFormatID exactly:
// built-in IDs 14–22, 27–36, 45–47, 50–58 are date/time; all other built-in IDs
// (id < 164) are not; custom formats (id >= 164) are scanned for date tokens.
func isDateFormat(id int, fmtStr string) bool {
	if dateformat.IsBuiltInDateID(id) {
		return true
	}
	if id < 164 {
		return false
	}
	return dateformat.ScanFormatStr(fmtStr)
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

	// Pre-scan 1: determine if any AM/PM token is present — affects hour rendering.
	hasAmPm := false
	for _, tok := range sec.Items {
		if tok.TType == nfp.TokenTypeDateTimes {
			upper := strings.ToUpper(tok.TValue)
			if upper == "AM/PM" || upper == "A/P" || tok.TValue == "上午/下午" {
				hasAmPm = true
				break
			}
		}
	}

	// Pre-scan 2: build isMinuteIndex — maps item index → true when an M/MM
	// token should be rendered as minutes rather than months.
	//
	// Rules (matches Excel behaviour):
	//   (a) M/MM is a minute if the previous DateTimes/ElapsedDateTimes token
	//       was H/HH (hour), OR
	//   (b) M/MM is a minute if the next DateTimes token is S/SS (seconds).
	//
	// The existing single-pass `lastWasHour` flag handles rule (a) but not (b).
	// Rule (b) covers formats like "m:ss" where no hour token precedes the M.
	isMinuteIndex := make(map[int]bool)
	{
		// Collect indices of only DateTimes/ElapsedDateTimes tokens.
		type dtEntry struct {
			idx   int    // index into sec.Items
			upper string // uppercased TValue
		}
		var dtTokens []dtEntry
		for i, tok := range sec.Items {
			if tok.TType == nfp.TokenTypeDateTimes || tok.TType == nfp.TokenTypeElapsedDateTimes {
				dtTokens = append(dtTokens, dtEntry{i, strings.ToUpper(tok.TValue)})
			}
		}
		for j, entry := range dtTokens {
			if entry.upper != "M" && entry.upper != "MM" {
				continue
			}
			// Rule (a): previous dt token was an hour.
			if j > 0 {
				prev := dtTokens[j-1].upper
				if prev == "H" || prev == "HH" {
					isMinuteIndex[entry.idx] = true
					continue
				}
			}
			// Rule (b): next dt token is seconds.
			if j < len(dtTokens)-1 {
				next := dtTokens[j+1].upper
				if next == "S" || next == "SS" {
					isMinuteIndex[entry.idx] = true
				}
			}
		}
	}

	var sb strings.Builder
	lastWasHour := false
	lastWasSecs := false // true after an SS or S DateTimes token (for milliseconds)

	for i, tok := range sec.Items {
		switch tok.TType {

		case nfp.TokenTypeDateTimes:
			upper := strings.ToUpper(tok.TValue)
			// For M/MM, override lastWasHour with the pre-computed minute flag.
			lWH := lastWasHour
			if upper == "M" || upper == "MM" {
				lWH = isMinuteIndex[i]
			}
			s := renderDateToken(upper, t, serial, hasAmPm, lWH, date1904)
			sb.WriteString(s)
			// Track whether this token was an hour (H / HH) for M/MM disambiguation.
			lastWasHour = upper == "H" || upper == "HH"
			lastWasSecs = upper == "SS" || upper == "S"

		case nfp.TokenTypeElapsedDateTimes:
			// Elapsed tokens operate on the raw serial (fractional days).
			upper := strings.ToUpper(tok.TValue)
			s := renderElapsed(upper, serial)
			sb.WriteString(s)
			// An elapsed hour token ([h] or [hh]) acts like a regular hour
			// token for M/MM disambiguation: the next M/MM should be minutes.
			lastWasHour = upper == "H" || upper == "HH"
			lastWasSecs = upper == "SS" || upper == "S"

		case nfp.TokenTypeDecimalPoint:
			// When a DecimalPoint immediately follows a seconds token, it
			// introduces sub-second (millisecond) digits — do NOT emit the
			// decimal point yet; the ZeroPlaceHolder that follows will handle
			// both the dot and the digits.
			if !lastWasSecs {
				sb.WriteByte('.')
			}
			// Do not reset lastWasSecs here — the next token may be the
			// sub-second ZeroPlaceHolder that needs it.
			lastWasHour = false

		case nfp.TokenTypeZeroPlaceHolder:
			if lastWasSecs {
				// Sub-second digits: nfp emits ZeroPlaceHolder "0"/"00"/"000"
				// immediately after the DecimalPoint that followed a seconds token.
				digits := len(tok.TValue)
				// Fractional seconds = (serial mod 1) * 86400 mod 60, then
				// take only the sub-second fraction.
				fracSec := serial - math.Trunc(serial) // fractional day
				ms := fracSec * 86400                  // fractional seconds within the day
				ms = ms - math.Trunc(ms)               // keep only sub-second part
				msInt := int(math.Round(ms * math.Pow10(digits)))
				// Clamp to avoid overflow (e.g. rounding 0.9999 up to 1000).
				max := int(math.Pow10(digits))
				if msInt >= max {
					msInt = max - 1
				}
				sb.WriteByte('.')
				sb.WriteString(fmt.Sprintf("%0*d", digits, msInt))
				lastWasSecs = false
			} else {
				// Regular zero placeholder in a non-seconds context — emit verbatim.
				sb.WriteString(tok.TValue)
				lastWasSecs = false
			}
			lastWasHour = false

		case nfp.TokenTypeLiteral:
			// Quoted text or escape sequences — emit the value verbatim.
			// Do NOT reset lastWasHour: a literal separator (e.g. ":") between
			// an hour token and a following M/MM must not break the
			// minute-vs-month disambiguation.
			sb.WriteString(tok.TValue)
			lastWasSecs = false

		case nfp.TokenTypeAlignment:
			// _x → one space.
			sb.WriteByte(' ')
			lastWasHour = false
			lastWasSecs = false

		case nfp.TokenTypeRepeatsChar:
			// *x → one instance of x.
			if tok.TValue != "" {
				sb.WriteString(tok.TValue)
			}
			lastWasHour = false
			lastWasSecs = false

		case nfp.TokenTypeCurrencyLanguage:
			sb.WriteString(currencySymbol(tok))
			lastWasHour = false
			lastWasSecs = false

		default:
			// Ignore colour codes, conditions, etc.
			lastWasHour = false
			lastWasSecs = false
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

// excelDay returns the calendar day number for a serial.  date1904 selects
// the 1904 date system.
//
// For 1900-system (date1904=false):
//   - serial < 1        -> 0   (pure time value, no date component)
//   - 1 <= serial < 60  -> t.Day() (convertSerial already maps these correctly:
//     serial 1 = Jan 1 1900, serial 2 = Jan 2, etc.)
//   - 60 <= serial < 61 -> 29 (fake Feb-29-1900; the Lotus 1-2-3 phantom day)
//   - serial >= 61      -> t.Day()
//
// For 1904-system (date1904=true):
//   - serial < 1 -> 0
//   - otherwise  -> t.Day()
//
// NOTE: A previous version applied serial+1 for the 1-59 range under the
// incorrect assumption that convertSerial had an off-by-one there.
// convertSerial is correct for that range; the +1 was itself the bug, causing
// every date in Jan 1900 to render one day too late.
func excelDay(t time.Time, serial float64, date1904 bool) int {
	if serial < 1 {
		return 0
	}
	if serial >= 60 && serial < 61 {
		return 29 // fake Feb-29-1900 (Lotus 1-2-3 phantom day)
	}
	return t.Day()
}

// renderDateToken renders a single date/time token value (already upper-cased).
func renderDateToken(upper string, t time.Time, serial float64, hasAmPm bool, lastWasHour bool, date1904 bool) string {
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
		return fmt.Sprintf("%02d", excelDay(t, serial, date1904))
	case "D":
		return strconv.Itoa(excelDay(t, serial, date1904))

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
		// Single 'h' token: no zero-padding (Excel h vs hh distinction).
		return strconv.Itoa(h)

	// ── second ───────────────────────────────────────────────────────────────
	// IMPORTANT: seconds are computed by flooring the fractional-day part of
	// the serial, NOT from t.Second().  t uses excelize-compatible half-second
	// rounding (so minutes/hours are correct), but that rounding would shift
	// 62.75s → 63s, making the displayed whole-second wrong when sub-second
	// decimal digits are also rendered.  Flooring the serial gives the correct
	// whole-second base (62) while the decimal part shows the .75 fraction.
	case "SS":
		flooredSec := int(math.Floor((serial-math.Trunc(serial))*86400)) % 60
		return fmt.Sprintf("%02d", flooredSec)
	case "S":
		flooredSec := int(math.Floor((serial-math.Trunc(serial))*86400)) % 60
		return strconv.Itoa(flooredSec)

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
	case "上午/下午":
		// Chinese AM/PM: 上午 = morning (AM), 下午 = afternoon (PM).
		if t.Hour() < 12 {
			return "上午"
		}
		return "下午"

	// ── single-letter month initial ──────────────────────────────────────────
	// Excel "MMMMM" renders J F M A M J J A S O N D (first letter of month).
	case "MMMMM":
		return string([]rune(t.Month().String())[:1])

	// ── Era tokens (Japanese/CJK) ────────────────────────────────────────────
	// E / EE are Japanese imperial era tokens.  For non-CJK (western) locales
	// excelize falls back to the Gregorian year as a plain integer.
	case "E", "EE":
		return fmt.Sprintf("%d", t.Year())

	// G / GG / GGG are era-name tokens (e.g. "R" / "Rei" / "Reiwa" in Japanese).
	// For non-CJK locales excelize produces no output — silently ignored.
	case "G", "GG", "GGG", "GGGG", "GGGGG":
		return ""

	// R / RR are era-abbreviation tokens used in some Japanese locale formats.
	// For non-CJK locales produce no output (matches excelize western behaviour).
	case "R", "RR":
		return ""

	// ── Buddhist Era / Gregorian calendar mode indicators ────────────────────
	// B1 = Buddhist Era, B2 = Gregorian. Excel uses these as calendar-system
	// switches; we silently ignore them (matches excelize behaviour).
	case "B1", "B2":
		return ""
	}
	return ""
}

// renderElapsed renders an elapsed-time token (h, hh, mm, ss — as emitted by
// the nfp parser with brackets stripped) using the raw serial (fractional days).
//
// Elapsed tokens (produced by nfp from [h], [mm], [ss] in the format string)
// represent total elapsed duration, not a clock component:
//   - [h] / [hh] → total elapsed hours (no modulo)
//   - [mm]       → total elapsed minutes (no modulo)
//   - [ss]       → total elapsed seconds (no modulo)
//
// roundEpsilon is added before truncation to prevent floating-point drift
// from causing values that are exactly on a whole-unit boundary (e.g. exactly
// 48 hours) to truncate one unit too low due to representational error.
//
// int64 is used throughout to avoid overflow on 32-bit platforms when the
// serial represents a large elapsed duration (e.g. many thousands of hours).
func renderElapsed(upper string, serial float64) string {
	switch upper {
	case "H", "HH":
		return strconv.FormatInt(int64(serial*24+roundEpsilon), 10)
	case "MM":
		return strconv.FormatInt(int64(serial*24*60+roundEpsilon), 10)
	case "SS":
		return strconv.FormatInt(int64(serial*24*3600+roundEpsilon), 10)
	}
	return ""
}

// roundEpsilon is added to the fractional-day part before rounding to seconds.
// This matches excelize's timeFromExcelTime behaviour which adds 1e-9 to avoid
// floating-point drift causing floor/truncate to land one second below a whole
// second boundary (e.g. 0.4999999999 → 0.5000000000 → rounds up correctly).
// When rounding pushes the result to 86400 seconds (exactly midnight), the
// overflow is rolled into the next calendar day (intPart++) rather than being
// clamped to 23:59:59 — matching excelize's time.Round-based implementation.
const roundEpsilon = 1e-9

// convertSerial converts an Excel serial to time.Time, handling both date
// systems.  It mirrors xlsb.ConvertDateEx without importing the root package
// to keep the import graph simple.
//
// The fractional-seconds component is rounded (not floored) to match
// excelize's timeFromExcelTime: excelize adds roundEpsilon to the float part
// then rounds to the nearest second when nanoseconds > 500 ms.  Using floor
// causes off-by-one minute errors near minute boundaries.  Sub-second
// rendering in renderDateTime reads the remainder directly from the serial,
// not from the time.Time, so this rounding does not affect millisecond output.
func convertSerial(serial float64, date1904 bool) (time.Time, error) {
	if math.IsNaN(serial) || math.IsInf(serial, 0) || serial < 0 {
		return time.Time{}, fmt.Errorf("numfmt: invalid serial %v", serial)
	}
	// Add roundEpsilon to the fractional day to avoid floating-point drift
	// (mirrors excelize's `floatPart := excelTime - float64(wholeDaysPart) + roundEpsilon`).
	fracDay := (serial - math.Trunc(serial)) + roundEpsilon
	// Convert fractional day to nanoseconds, then round to the nearest second
	// (mirrors excelize's half-second rounding logic).
	const nanosInADay = float64(24 * 60 * 60 * 1e9)
	durNanos := time.Duration(fracDay * nanosInADay)
	// Round to nearest second: if nanosecond remainder > 500ms, round up.
	ns := int(durNanos % time.Second)
	fracSec := int64(durNanos / time.Second)
	if ns > 500_000_000 {
		fracSec++
	}
	if fracSec < 0 {
		fracSec = 0
	}
	// When rounding pushes fracSec to 86400 (midnight), roll over to the next
	// day rather than clamping to 23:59:59.  This matches excelize's behaviour:
	// it uses time.Round which naturally carries the overflow into the date part.
	dayRollover := int(fracSec / 86400)
	fracSec = fracSec % 86400
	if date1904 {
		base := time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)
		intPart := int(serial) + dayRollover
		return base.Add(time.Duration(intPart)*24*time.Hour + time.Duration(fracSec)*time.Second), nil
	}
	base := time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)
	intPart := int(serial) + dayRollover
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
	// ── empty section (e.g. ";;;" format) → suppress the value entirely ─────
	// An empty section (no items) means the format intentionally hides the value.
	// This is used by formats like ";;;" which suppress all output for every sign.
	if len(sec.Items) == 0 {
		return ""
	}

	// ── quick dispatch for specialised sub-formats ────────────────────────────
	for _, tok := range sec.Items {
		switch tok.TType {
		case nfp.TokenTypeExponential:
			return renderScientific(val, sec, sections)
		case nfp.TokenTypeFraction:
			return renderFraction(val, sec, sections)
		}
	}

	// ── pass 1: collect format metadata ──────────────────────────────────────
	type meta struct {
		hasPercent          bool
		hasThousands        bool
		trailingCommas      int // number of trailing commas after last digit placeholder (scaling)
		decZeros            int // count of '0' placeholders after decimal point
		decHashes           int // count of '#' placeholders after decimal point
		decQuestions        int // count of '?' placeholders after decimal point
		intZeros            int // count of '0' placeholders before decimal point
		intHashes           int // count of '#' placeholders before decimal point
		intQuestions        int // count of '?' placeholders before decimal point
		hasDecimal          bool
		hasExplicitSign     bool // literal '+' or '-' in the section
		hasDigitPlaceholder bool // any digit placeholder (0, #, ?) in the section
		hasOutputToken      bool // any token that can produce visible output (not purely decorative)
	}
	var m meta
	afterDecimal := false
	// Track trailing commas for scaling: commas after the last digit placeholder.
	// In nfp, trailing commas appear as ThousandsSeparator or Literal "," tokens
	// after all digit placeholders.
	lastDigitIdx := -1
	for i, tok := range sec.Items {
		switch tok.TType {
		case nfp.TokenTypeZeroPlaceHolder, nfp.TokenTypeHashPlaceHolder,
			nfp.TokenTypeDigitalPlaceHolder:
			lastDigitIdx = i
		}
	}
	// Count consecutive comma tokens (ThousandsSeparator or Literal ",") after
	// the last digit placeholder.  These are scaling commas, not separators.
	if lastDigitIdx >= 0 {
		for i := lastDigitIdx + 1; i < len(sec.Items); i++ {
			tok := sec.Items[i]
			if tok.TType == nfp.TokenTypeThousandsSeparator ||
				(tok.TType == nfp.TokenTypeLiteral && tok.TValue == ",") {
				m.trailingCommas++
			} else {
				break
			}
		}
	}
	for _, tok := range sec.Items {
		switch tok.TType {
		case nfp.TokenTypeColor, nfp.TokenTypeCondition:
			// Purely decorative — do not set hasOutputToken.
		default:
			m.hasOutputToken = true
		}
		switch tok.TType {
		case nfp.TokenTypePercent:
			m.hasPercent = true
		case nfp.TokenTypeThousandsSeparator:
			// Grouping vs. scaling is determined after the loop (line 798);
			// nothing to do here.
		case nfp.TokenTypeDecimalPoint:
			m.hasDecimal = true
			// nfp quirk: when '?' or '#' immediately precede the '.' in the
			// format string (e.g. "?.??"), nfp lumps them into the DecimalPoint
			// token's TValue (e.g. TValue="?.").  Count those leading chars as
			// integer-side digital placeholders.
			for _, ch := range tok.TValue {
				if ch == '?' {
					m.intQuestions++
				} else if ch == '#' {
					// treated as hash before decimal (suppresses leading zero)
				} else if ch == '.' {
					break
				}
			}
			afterDecimal = true
		case nfp.TokenTypeZeroPlaceHolder:
			m.hasDigitPlaceholder = true
			if afterDecimal {
				m.decZeros += len(tok.TValue)
			} else {
				m.intZeros += len(tok.TValue)
			}
		case nfp.TokenTypeHashPlaceHolder:
			m.hasDigitPlaceholder = true
			if afterDecimal {
				m.decHashes += len(tok.TValue)
			} else {
				m.intHashes += len(tok.TValue)
			}
		case nfp.TokenTypeDigitalPlaceHolder:
			m.hasDigitPlaceholder = true
			// '?' — space-padded digit placeholder.
			if afterDecimal {
				m.decQuestions += len(tok.TValue)
			} else {
				m.intQuestions += len(tok.TValue)
			}
		case nfp.TokenTypeLiteral:
			tv := tok.TValue
			if tv == "+" || tv == "-" {
				m.hasExplicitSign = true
			}
		}
	}
	// If all ThousandsSeparator tokens are trailing (scaling), clear hasThousands.
	// Count non-trailing ThousandsSeparator tokens.
	nonTrailingThousands := 0
	for i, tok := range sec.Items {
		if tok.TType == nfp.TokenTypeThousandsSeparator && i <= lastDigitIdx {
			nonTrailingThousands++
		}
	}
	m.hasThousands = nonTrailingThousands > 0

	totalDecPlaces := m.decZeros + m.decHashes + m.decQuestions

	// ── apply scaling ─────────────────────────────────────────────────────────
	absVal := math.Abs(val)
	if m.hasPercent {
		absVal *= 100
	}
	// Thousands-scaling: each trailing comma divides by 1,000.
	if m.trailingCommas > 0 {
		scale := math.Pow(1000, float64(m.trailingCommas))
		absVal /= scale
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
		// Trim trailing digits beyond what '0' placeholders require (# and ?
		// placeholders allow trimming).  For '?', trim but note that the token
		// walk will pad with spaces if needed.
		trimmable := m.decHashes + m.decQuestions
		if trimmable > 0 && len(fracStr) > m.decZeros {
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
	// '?' pads with leading spaces instead of zeros.
	for len(intStr) < m.intZeros+m.intQuestions {
		intStr = " " + intStr
	}
	// '#' placeholders suppress a leading zero: if the integer part is "0"
	// and the format has ONLY '#' integer placeholders (no '0'), suppress it.
	// This matches Excel's behaviour: "#" → "", "#.##" with 0.5 → ".5".
	// Formats with no integer placeholders at all (e.g. ".00") must NOT
	// suppress: their intStr will be prepended later in the afterDecimal path.
	hashSuppressed := false
	if m.intZeros == 0 && m.intHashes > 0 && intStr == "0" {
		intStr = ""
		hashSuppressed = true
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
		} else {
			// Two+ sections: check whether the negative section (sec) contains
			// a visual sign indicator — parentheses, minus, or plus literals.
			// A colour modifier alone (e.g. "0;[Red]0") does NOT suppress the
			// minus sign; Excel still renders "-5" in red for that format.
			// (E.g. "0;0" has no wrapper → "-5" not "5".)
			hasSignWrapper := false
			for _, tok := range sec.Items {
				if tok.TType == nfp.TokenTypeLiteral {
					if tok.TValue == "(" || tok.TValue == ")" ||
						tok.TValue == "-" || tok.TValue == "+" {
						hasSignWrapper = true
						break
					}
				}
			}
			if !hasSignWrapper {
				needsMinus = true
			}
		}
	}

	// ── reassemble by walking tokens ──────────────────────────────────────────
	var sb strings.Builder
	if needsMinus {
		sb.WriteByte('-')
	}

	intConsumed := false
	fracConsumed := false
	afterDecimal = false

	// Build a set of token indices that are trailing scaling commas so the
	// token walk can skip them.
	trailingCommaIdxs := make(map[int]bool)
	if m.trailingCommas > 0 {
		count := 0
		for i := lastDigitIdx + 1; i < len(sec.Items) && count < m.trailingCommas; i++ {
			tok := sec.Items[i]
			if tok.TType == nfp.TokenTypeThousandsSeparator ||
				(tok.TType == nfp.TokenTypeLiteral && tok.TValue == ",") {
				trailingCommaIdxs[i] = true
				count++
			} else {
				break
			}
		}
	}

	for i, tok := range sec.Items {
		if trailingCommaIdxs[i] {
			continue // scaling comma — already applied to absVal, do not emit
		}
		switch tok.TType {
		case nfp.TokenTypeLiteral:
			sb.WriteString(tok.TValue)

		case nfp.TokenTypeDecimalPoint:
			// Handle nfp quirk: when '?' or '#' immediately precede '.' in the
			// format string, nfp includes them in the DecimalPoint token's TValue
			// (e.g. "?.??" → TValue="?.").  Emit the integer part first, then '.'.
			if !intConsumed && len(tok.TValue) > 1 {
				sb.WriteString(intStr)
				intConsumed = true
			}
			if len(fracStr) > 0 {
				sb.WriteByte('.')
			}
			afterDecimal = true

		case nfp.TokenTypeZeroPlaceHolder, nfp.TokenTypeHashPlaceHolder,
			nfp.TokenTypeDigitalPlaceHolder:
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

		case nfp.TokenTypeAlignment:
			// _x → one space.
			sb.WriteByte(' ')

		case nfp.TokenTypeRepeatsChar:
			// *x → one instance of x.
			if tok.TValue != "" {
				sb.WriteString(tok.TValue)
			}

		case nfp.TokenTypeCurrencyLanguage:
			sb.WriteString(currencySymbol(tok))

		case nfp.TokenTypeColor, nfp.TokenTypeCondition:
			// Ignore formatting-only tokens.
		}
	}

	// If the format had placeholder tokens but none consumed the integer part,
	// emit it now.  This covers edge cases like ".00" (integer precedes the
	// decimal but has no explicit integer-side placeholder token).
	// Do NOT emit for literal-only sections (hasDigitPlaceholder==false): a
	// section such as `[<0]"neg "` is intentionally literal-only and should
	// not have any numeric digits appended.
	if !intConsumed && !afterDecimal && m.hasDigitPlaceholder {
		sb.WriteString(intStr)
	} else if !intConsumed && afterDecimal {
		// ".00"-style: no integer placeholder, so intStr was never emitted.
		// Prepend it to whatever has been written so far.  When needsMinus is
		// true sb already starts with '-'; we must insert intStr after the
		// sign so the result is "-5.50" not "5-.50".
		current := sb.String()
		sb.Reset()
		if needsMinus && len(current) > 0 && current[0] == '-' {
			sb.WriteByte('-')
			sb.WriteString(intStr)
			sb.WriteString(current[1:])
		} else {
			sb.WriteString(intStr)
			sb.WriteString(current)
		}
	}

	// Guard: if nothing was written, decide whether to fall back to renderGeneral.
	//   • Section has only decorative tokens (Color/Condition, no output tokens
	//     at all, e.g. "[Red]"): fall back so the numeric value is never dropped.
	//   • Section has output tokens but none are digit placeholders (literal-only,
	//     e.g. `[<0]"neg "` or `[=0]"zero"`): empty string is correct — return as-is.
	//   • Section has digit placeholders and hash suppression produced "": return "".
	//   • Section has digit placeholders but rendered nothing for another reason
	//     (shouldn't normally happen): fall back to renderGeneral as a safety net.
	if sb.Len() == 0 {
		if !m.hasOutputToken {
			// Colour-only or condition-only section — fall back to raw value.
			return renderGeneral(val)
		}
		if m.hasDigitPlaceholder && !hashSuppressed {
			// Digit placeholders exist but produced no output unexpectedly.
			return renderGeneral(val)
		}
		// Literal-only section or intentional hash suppression: return "".
	}

	return sb.String()
}

// ── scientific notation renderer ──────────────────────────────────────────────

// renderScientific renders val using an E+/E- scientific notation format.
//
// The format token sequence is:
//
//	[int-placeholders] [DecimalPoint] [frac-placeholders] Exponential [exp-placeholders]
//
// The number of integer placeholders determines the normalisation base: e.g.
// "##0.0E+0" has 3 integer placeholders so the mantissa is normalised to the
// nearest power of 3 (engineering notation), while "0.00E+00" normalises to
// exactly one significant integer digit (standard scientific notation).
func renderScientific(val float64, sec nfp.Section, sections []nfp.Section) string {
	// ── pass 1: collect metadata ──────────────────────────────────────────────
	type sciMeta struct {
		intPlaces       int    // total int placeholders (0+#)
		intZeros        int    // '0' int placeholders (for zero-padding)
		fracZeros       int    // '0' frac placeholders
		fracHashes      int    // '#' frac placeholders
		expZeros        int    // '0' exp placeholders
		expSign         string // "E+" or "E-"
		hasExplicitSign bool
	}
	var m sciMeta
	phase := 0 // 0=int, 1=frac, 2=exp
	for _, tok := range sec.Items {
		switch tok.TType {
		case nfp.TokenTypeDecimalPoint:
			phase = 1
		case nfp.TokenTypeExponential:
			m.expSign = tok.TValue // "E+" or "E-"
			phase = 2
		case nfp.TokenTypeZeroPlaceHolder:
			switch phase {
			case 0:
				m.intPlaces += len(tok.TValue)
				m.intZeros += len(tok.TValue)
			case 1:
				m.fracZeros += len(tok.TValue)
			case 2:
				m.expZeros += len(tok.TValue)
			}
		case nfp.TokenTypeHashPlaceHolder:
			switch phase {
			case 0:
				m.intPlaces += len(tok.TValue)
			case 1:
				m.fracHashes += len(tok.TValue)
			}
		case nfp.TokenTypeLiteral:
			if tok.TValue == "+" || tok.TValue == "-" {
				m.hasExplicitSign = true
			}
		}
	}
	if m.intPlaces == 0 {
		m.intPlaces = 1
	}
	totalFracPlaces := m.fracZeros + m.fracHashes

	// ── compute exponent and mantissa ─────────────────────────────────────────
	absVal := math.Abs(val)
	var exp int
	if absVal == 0 {
		exp = 0
	} else {
		// raw base-10 exponent
		rawExp := math.Floor(math.Log10(absVal))
		// Normalise so that the integer part has intPlaces significant digits.
		// For intPlaces=1: standard form (e.g. 1.23E+04).
		// For intPlaces=3: engineering form, exponent multiple of 3.
		if m.intPlaces == 1 {
			exp = int(rawExp)
		} else {
			// Round down to the nearest multiple of intPlaces.
			exp = int(rawExp) - (int(rawExp)%m.intPlaces+m.intPlaces)%m.intPlaces
		}
	}
	mantissa := absVal / math.Pow(10, float64(exp))

	// ── format mantissa ───────────────────────────────────────────────────────
	var mantissaStr string
	if totalFracPlaces > 0 {
		mantissaStr = strconv.FormatFloat(mantissa, 'f', totalFracPlaces, 64)
		// Trim trailing zeros in the '#' portion.
		if m.fracHashes > 0 {
			dotIdx := strings.IndexByte(mantissaStr, '.')
			if dotIdx >= 0 {
				intPart := mantissaStr[:dotIdx]
				fracPart := mantissaStr[dotIdx+1:]
				trimTo := len(fracPart)
				for trimTo > m.fracZeros && trimTo > 0 && fracPart[trimTo-1] == '0' {
					trimTo--
				}
				if trimTo == 0 {
					mantissaStr = intPart
				} else {
					mantissaStr = intPart + "." + fracPart[:trimTo]
				}
			}
		}
	} else {
		mantissaStr = strconv.FormatFloat(mantissa, 'f', 0, 64)
	}
	// Zero-pad integer part to intZeros width.
	dotIdx := strings.IndexByte(mantissaStr, '.')
	var mInt, mFrac string
	if dotIdx >= 0 {
		mInt, mFrac = mantissaStr[:dotIdx], mantissaStr[dotIdx+1:]
	} else {
		mInt = mantissaStr
	}
	for len(mInt) < m.intZeros {
		mInt = "0" + mInt
	}
	if mFrac != "" {
		mantissaStr = mInt + "." + mFrac
	} else {
		mantissaStr = mInt
	}

	// ── format exponent ───────────────────────────────────────────────────────
	expAbs := exp
	if expAbs < 0 {
		expAbs = -expAbs
	}
	expStr := strconv.Itoa(expAbs)
	for len(expStr) < m.expZeros {
		expStr = "0" + expStr
	}
	var expSign string
	if exp < 0 {
		expSign = "-"
	} else if m.expSign == "E+" {
		expSign = "+"
	}

	// ── sign ──────────────────────────────────────────────────────────────────
	needsMinus := val < 0 && !m.hasExplicitSign && len(sections) < 2

	// ── walk tokens to assemble output ────────────────────────────────────────
	var sb strings.Builder
	if needsMinus {
		sb.WriteByte('-')
	}
	phase = 0
	mantissaEmitted := false
	expEmitted := false
	for _, tok := range sec.Items {
		switch tok.TType {
		case nfp.TokenTypeDecimalPoint:
			phase = 1
		case nfp.TokenTypeExponential:
			// Emit the full exponent: "E+" or "E-" sign + digits.
			sb.WriteString(tok.TValue[:1]) // "E"
			sb.WriteString(expSign)
			sb.WriteString(expStr)
			expEmitted = true
			phase = 2
		case nfp.TokenTypeZeroPlaceHolder, nfp.TokenTypeHashPlaceHolder:
			switch phase {
			case 0, 1:
				if !mantissaEmitted {
					sb.WriteString(mantissaStr)
					mantissaEmitted = true
				}
			case 2:
				// Exponent digits already emitted with the Exponential token.
				if !expEmitted {
					sb.WriteString(expStr)
					expEmitted = true
				}
			}
		case nfp.TokenTypeLiteral:
			sb.WriteString(tok.TValue)
		case nfp.TokenTypeAlignment:
			sb.WriteByte(' ')
		case nfp.TokenTypeRepeatsChar:
			if tok.TValue != "" {
				sb.WriteString(tok.TValue)
			}
		case nfp.TokenTypeCurrencyLanguage:
			sb.WriteString(currencySymbol(tok))
		case nfp.TokenTypeColor, nfp.TokenTypeCondition:
			// ignore
		}
	}
	if sb.Len() == 0 {
		return renderGeneral(val)
	}
	return sb.String()
}

// ── fraction renderer ─────────────────────────────────────────────────────────

// renderFraction renders val as a mixed-number fraction using the format tokens.
//
// Format examples: "# ?/?" (max denom 9), "# ??/??" (max denom 99),
// "??/??" (no integer part, improper fraction).
//
// The number of '?' characters in the denominator placeholder sets the maximum
// denominator: 1 → 9, 2 → 99, 3 → 999, etc.
func renderFraction(val float64, sec nfp.Section, sections []nfp.Section) string {
	absVal := math.Abs(val)

	// ── scan tokens to find numerator/denominator placeholder widths and
	//    whether an integer-part placeholder exists ─────────────────────────────
	hasIntPart := false
	numWidth := 0
	denWidth := 0
	fixedDen := 0 // >0 when a TokenTypeDenominator (literal integer) is present
	phase := 0    // 0=pre-fraction, 1=numerator seen (after Fraction token = denominator)
	for _, tok := range sec.Items {
		switch tok.TType {
		case nfp.TokenTypeHashPlaceHolder:
			if phase == 0 {
				hasIntPart = true
			}
		case nfp.TokenTypeDigitalPlaceHolder:
			if phase == 0 {
				numWidth = len(tok.TValue)
			} else {
				denWidth = len(tok.TValue)
			}
		case nfp.TokenTypeDenominator:
			// Fixed denominator: TValue is a literal integer string (e.g. "4").
			if v, err := strconv.Atoi(tok.TValue); err == nil && v > 0 {
				fixedDen = v
				denWidth = len(tok.TValue)
			}
			phase = 1 // denominator tokens only appear after the slash
		case nfp.TokenTypeFraction:
			phase = 1
		}
	}
	if denWidth == 0 {
		denWidth = 1
	}
	maxDen := 1
	for range denWidth {
		maxDen *= 10
	}
	maxDen-- // e.g. denWidth=1 → maxDen=9; denWidth=2 → maxDen=99

	// ── separate integer and fractional parts ─────────────────────────────────
	var intPart int64
	var frac float64
	if hasIntPart {
		intPart = int64(math.Trunc(absVal))
		frac = absVal - float64(intPart)
	} else {
		intPart = 0
		frac = absVal
	}

	// ── rational approximation (Stern–Brocot / mediants) ─────────────────────
	var num, den int
	if fixedDen > 0 {
		// Fixed-denominator format (e.g. "# ?/4"): compute nearest numerator.
		num = int(math.Round(frac * float64(fixedDen)))
		den = fixedDen
	} else if hasIntPart {
		// Mixed-number format (e.g. "# ??/??"):  only the fractional part is
		// approximated; the integer part is rendered separately.
		num, den = bestFraction(frac, maxDen)
	} else {
		// Improper-fraction format (e.g. "??/??"):  the whole value (including
		// any integer part) is approximated as a single p/q with q ≤ maxDen.
		num, den = bestImproperFraction(frac, maxDen)
	}

	// Carry: for mixed-number formats only, if the fractional approximation
	// rounds to 1 (num == den), fold it back into the integer part.
	// This is not needed for improper fractions because bestImproperFraction
	// already approximates the whole value as a single p/q.
	if hasIntPart && den > 0 && num >= den {
		intPart++
		num = 0
		den = 1
	}

	// ── sign ──────────────────────────────────────────────────────────────────
	needsMinus := val < 0 && len(sections) < 2

	// ── walk tokens to assemble output ────────────────────────────────────────
	var sb strings.Builder
	if needsMinus {
		sb.WriteByte('-')
	}

	intEmitted := false
	numEmitted := false
	denEmitted := false
	phase = 0
	// When the fractional part is zero (exact integer), suppress the fraction
	// display entirely and replace numerator, slash, and denominator with spaces
	// to preserve column alignment.
	zeroFrac := num == 0
	var numStr, denStr string
	if zeroFrac {
		// Space-pad to width of numerator, slash, denominator.
		numStr = strings.Repeat(" ", numWidth)
		if fixedDen > 0 {
			// Fixed denominator is always shown (e.g. "# ?/4" shows "4" even for integers).
			denStr = strconv.Itoa(fixedDen)
		} else {
			denStr = strings.Repeat(" ", denWidth)
		}
	} else {
		numStr = strconv.FormatInt(int64(num), 10)
		denStr = strconv.FormatInt(int64(den), 10)
		// Space-pad numerator (right-aligned) and denominator (left-aligned).
		for len(numStr) < numWidth {
			numStr = " " + numStr
		}
		for len(denStr) < denWidth {
			denStr = denStr + " "
		}
	}
	intStr := strconv.FormatInt(intPart, 10)

	for _, tok := range sec.Items {
		switch tok.TType {
		case nfp.TokenTypeHashPlaceHolder:
			if phase == 0 && !intEmitted {
				// Integer part — suppress if zero and value has no integer.
				if intPart != 0 || !hasIntPart {
					sb.WriteString(intStr)
				}
				intEmitted = true
			}
		case nfp.TokenTypeDigitalPlaceHolder:
			if phase == 0 && !numEmitted {
				sb.WriteString(numStr)
				numEmitted = true
			} else if phase == 1 && !denEmitted {
				sb.WriteString(denStr)
				denEmitted = true
			}
		case nfp.TokenTypeDenominator:
			// Fixed denominator literal — emit denStr (either the fixed number or spaces).
			if !denEmitted {
				sb.WriteString(denStr)
				denEmitted = true
			}
		case nfp.TokenTypeFraction:
			if zeroFrac && fixedDen == 0 {
				sb.WriteByte(' ') // replace '/' with space for alignment (variable-denominator only)
			} else {
				sb.WriteByte('/')
			}
			phase = 1
		case nfp.TokenTypeLiteral:
			// Suppress the space separator between integer and numerator when
			// the integer part is zero and the integer placeholder is "#"
			// (which suppresses leading zero).
			if tok.TValue == " " && hasIntPart && intPart == 0 && !intEmitted {
				sb.WriteByte(' ') // keep the space for alignment
			} else {
				sb.WriteString(tok.TValue)
			}
		case nfp.TokenTypeAlignment:
			sb.WriteByte(' ')
		case nfp.TokenTypeRepeatsChar:
			if tok.TValue != "" {
				sb.WriteString(tok.TValue)
			}
		case nfp.TokenTypeCurrencyLanguage:
			sb.WriteString(currencySymbol(tok))
		case nfp.TokenTypeColor, nfp.TokenTypeCondition:
			// ignore
		}
	}
	if sb.Len() == 0 {
		return renderGeneral(val)
	}
	return sb.String()
}

// bestImproperFraction finds the best rational approximation p/q of x ≥ 0
// such that q ≤ maxDen and |x - p/q| is minimised.  Unlike bestFraction, it
// handles x ≥ 1 (improper fractions) by decomposing x = floor(x) + frac and
// delegating the fractional part to bestFraction, then reconstructing the
// improper numerator as floor(x)*q + frac_p.
func bestImproperFraction(x float64, maxDen int) (num, den int) {
	if x <= 0 {
		return 0, 1
	}
	intPart := int(math.Trunc(x))
	frac := x - float64(intPart)
	// Use bestFraction for the fractional part.
	fracNum, fracDen := bestFraction(frac, maxDen)
	// Reconstruct improper numerator.
	return intPart*fracDen + fracNum, fracDen
}

// bestFraction finds the best rational approximation p/q of x such that
// q ≤ maxDen and |x - p/q| is minimised, using the Stern–Brocot tree.
func bestFraction(x float64, maxDen int) (num, den int) {
	if x <= 0 {
		return 0, 1
	}
	if x >= 1 {
		return 1, 1
	}
	// Stern–Brocot search.
	loN, loD := 0, 1
	hiN, hiD := 1, 1
	for {
		medN := loN + hiN
		medD := loD + hiD
		if medD > maxDen {
			break
		}
		med := float64(medN) / float64(medD)
		if math.Abs(x-med) < 1e-10 {
			return medN, medD
		}
		if x < med {
			hiN, hiD = medN, medD
		} else {
			loN, loD = medN, medD
		}
	}
	// Pick the closer of lo and hi boundaries.
	loErr := math.Abs(x - float64(loN)/float64(loD))
	hiErr := math.Abs(x - float64(hiN)/float64(hiD))
	if loErr <= hiErr {
		return loN, loD
	}
	return hiN, hiD
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
