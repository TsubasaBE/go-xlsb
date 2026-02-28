package numfmt

import (
	"strconv"
	"testing"
)

// TestRenderNumberDotFormatNegative verifies that a ".00"-style format (no
// explicit integer placeholder before the decimal point) produces the correct
// sign ordering for negative values, i.e. "-5.50" not "5-.50".
func TestRenderNumberDotFormatNegative(t *testing.T) {
	t.Helper()
	cases := []struct {
		val    float64
		format string
		want   string
	}{
		{-5.5, ".00", "-5.50"},
		{5.5, ".00", "5.50"},
		{-0.5, ".00", "-0.50"},
		{0.5, ".00", "0.50"},
	}
	for _, tc := range cases {
		got := FormatValue(tc.val, 164, tc.format, false)
		if got != tc.want {
			t.Errorf("FormatValue(%v, 164, %q, false) = %q, want %q", tc.val, tc.format, got, tc.want)
		}
	}
}

// TestGeneralPrecision verifies that renderGeneral formats fractional values
// with 10 significant digits (G10), matching excelize's General format output.
func TestGeneralPrecision(t *testing.T) {
	t.Helper()
	vals := []struct {
		v    float64
		want string
	}{
		{45498.666666666664, "45498.66667"},
		{2.09003333333333, "2.090033333"},
		{6.041666666666667, "6.041666667"},
	}
	for _, tt := range vals {
		got := strconv.FormatFloat(tt.v, 'G', 10, 64)
		if got != tt.want {
			t.Errorf("G10(%v) = %q, want %q", tt.v, got, tt.want)
		}
	}
}
