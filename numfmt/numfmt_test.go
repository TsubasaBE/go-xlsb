package numfmt

import (
	"strconv"
	"testing"
)

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
