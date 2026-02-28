//go:build ignore

package main

import (
	"fmt"
	"os"
	"path/filepath"
	"github.com/TsubasaBE/go-xlsb/workbook"
)

func main() {
	files, _ := filepath.Glob("xls/*.xlsb")
	for _, f := range files {
		wb, err := workbook.Open(f)
		if err != nil {
			fmt.Printf("ERROR opening %s: %v\n", filepath.Base(f), err)
			continue
		}
		sheets := wb.Sheets()
		fmt.Printf("\n=== %s ===\n", filepath.Base(f))
		fmt.Printf("  Sheets (%d): %v\n", len(sheets), sheets)
		for i, name := range sheets {
			sheet, err := wb.Sheet(i + 1)
			if err != nil {
				fmt.Printf("  Sheet(%d) %q: ERROR %v\n", i+1, name, err)
				continue
			}
			rowCount := 0
			cellCount := 0
			var firstRow []interface{}
			for row := range sheet.Rows(true) {
				if rowCount == 0 {
					for _, c := range row {
						firstRow = append(firstRow, c.V)
					}
				}
				rowCount++
				cellCount += len(row)
			}
			dim := sheet.Dimension
			if dim != nil {
				fmt.Printf("  [%d] %q dim=(%d,%d %dx%d) rows=%d cells=%d first=%v\n",
					i+1, name, dim.R, dim.C, dim.H, dim.W, rowCount, cellCount, firstRow)
			} else {
				fmt.Printf("  [%d] %q dim=nil rows=%d cells=%d first=%v\n",
					i+1, name, rowCount, cellCount, firstRow)
			}
		}
		wb.Close()
		_ = os.Stderr
	}
}
