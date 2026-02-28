package numfmt

import (
	"fmt"
	"strconv"
	"testing"
)

func TestGeneralPrecision(t *testing.T) {
	vals := []struct {
		v    float64
		want string
	}{
		{45498.666666666664, "45498.66667"},
		{2.09003333333333, "2.090033333"},
		{6.041666666666667, "6.041666667"},
	}
	for _, tt := range vals {
		// Try G10
		g10 := strconv.FormatFloat(tt.v, 'G', 10, 64)
		fmt.Printf("val=%-25v  G10=%s  want=%s  match=%v\n", tt.v, g10, tt.want, g10 == tt.want)
	}
}
