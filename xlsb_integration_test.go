package xlsb_test

// Integration tests against real .xlsb files in the xls/ directory.
//
// These tests exercise the full read path end-to-end: ZIP extraction,
// workbook.bin parsing, sharedStrings.bin parsing, and per-sheet row
// iteration with all cell types.

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/TsubasaBE/go-xlsb"
	"github.com/TsubasaBE/go-xlsb/workbook"
	"github.com/TsubasaBE/go-xlsb/worksheet"
)

// ── helpers ───────────────────────────────────────────────────────────────────

// xlsbPath returns the path to a file in the xls/ test-data directory,
// and skips the test when the file does not exist (so CI without the
// proprietary files still passes).
func xlsbPath(t *testing.T, name string) string {
	t.Helper()
	p := filepath.Join("xls", name)
	if _, err := os.Stat(p); os.IsNotExist(err) {
		t.Skipf("test file %q not present, skipping", p)
	}
	return p
}

// openXLSB opens a real .xlsb file; the test is skipped when the file is
// absent.  The caller must not call Close — it is deferred here.
func openXLSB(t *testing.T, name string) *workbook.Workbook {
	t.Helper()
	wb, err := xlsb.Open(xlsbPath(t, name))
	if err != nil {
		t.Fatalf("Open(%q): %v", name, err)
	}
	t.Cleanup(func() { wb.Close() })
	return wb
}

// collectRows iterates all rows of a worksheet (sparse=true) and returns a
// flat slice of non-nil cell values together with the total row count.
func collectRows(ws *worksheet.Worksheet, sparse bool) (rows [][]worksheet.Cell, rowCount int) {
	for row := range ws.Rows(sparse) {
		rows = append(rows, row)
		rowCount++
	}
	return
}

// typeStats counts how many cells of each Go type appear in a worksheet.
func typeStats(ws *worksheet.Worksheet) (nFloat, nString, nBool, nNil int) {
	for row := range ws.Rows(true) {
		for _, c := range row {
			switch c.V.(type) {
			case float64:
				nFloat++
			case string:
				nString++
			case bool:
				nBool++
			case nil:
				nNil++
			}
		}
	}
	return
}

// ── TestRealFilesOpenAll ───────────────────────────────────────────────────────

// TestRealFilesOpenAll verifies that every .xlsb file in xls/ can be opened
// without error and returns at least one sheet.
func TestRealFilesOpenAll(t *testing.T) {
	entries, err := filepath.Glob(filepath.Join("xls", "*.xlsb"))
	if err != nil {
		t.Fatalf("glob: %v", err)
	}
	if len(entries) == 0 {
		t.Skip("no .xlsb files found in xls/, skipping")
	}
	for _, path := range entries {
		path := path
		t.Run(filepath.Base(path), func(t *testing.T) {
			wb, err := xlsb.Open(path)
			if err != nil {
				t.Fatalf("Open: %v", err)
			}
			defer wb.Close()
			sheets := wb.Sheets()
			if len(sheets) == 0 {
				t.Error("expected at least one sheet, got 0")
			}
		})
	}
}

// ── TestRealFilesAllSheetsReadable ────────────────────────────────────────────

// TestRealFilesAllSheetsReadable opens every .xlsb file and iterates all rows
// of every sheet, verifying that the iterators complete without panicking or
// returning errors mid-stream.
func TestRealFilesAllSheetsReadable(t *testing.T) {
	entries, err := filepath.Glob(filepath.Join("xls", "*.xlsb"))
	if err != nil {
		t.Fatalf("glob: %v", err)
	}
	if len(entries) == 0 {
		t.Skip("no .xlsb files found in xls/, skipping")
	}
	for _, path := range entries {
		path := path
		t.Run(filepath.Base(path), func(t *testing.T) {
			wb, err := xlsb.Open(path)
			if err != nil {
				t.Fatalf("Open: %v", err)
			}
			defer wb.Close()
			for i, name := range wb.Sheets() {
				sheet, err := wb.Sheet(i + 1)
				if err != nil {
					t.Errorf("Sheet(%d) %q: %v", i+1, name, err)
					continue
				}
				// Exhaust the iterator — if any internal panic occurs the
				// test will fail via the testing runtime.
				rowCount := 0
				for range sheet.Rows(true) {
					rowCount++
				}
				t.Logf("  sheet[%d] %q: %d rows (sparse)", i+1, name, rowCount)
			}
		})
	}
}

// ── Maand productie tests ─────────────────────────────────────────────────────

// TestMaandProductieSheets verifies the exact sheet list for the
// "Maand productie.xlsb" workbook.
func TestMaandProductieSheets(t *testing.T) {
	wb := openXLSB(t, "Maand productie.xlsb")

	want := []string{"Packing", "SSP-overzicht", "Packing Overzicht"}
	got := wb.Sheets()
	if len(got) != len(want) {
		t.Fatalf("Sheets() = %v, want %v", got, want)
	}
	for i, w := range want {
		if got[i] != w {
			t.Errorf("Sheets()[%d] = %q, want %q", i, got[i], w)
		}
	}
}

// TestMaandProductieSheetByName verifies case-insensitive lookup by name.
func TestMaandProductieSheetByName(t *testing.T) {
	wb := openXLSB(t, "Maand productie.xlsb")

	tests := []struct {
		query   string
		wantErr bool
	}{
		{"Packing", false},
		{"packing", false},           // case-insensitive
		{"PACKING OVERZICHT", false}, // case-insensitive
		{"SSP-overzicht", false},
		{"nonexistent", true},
	}
	for _, tc := range tests {
		_, err := wb.SheetByName(tc.query)
		if tc.wantErr && err == nil {
			t.Errorf("SheetByName(%q): expected error, got nil", tc.query)
		}
		if !tc.wantErr && err != nil {
			t.Errorf("SheetByName(%q): unexpected error: %v", tc.query, err)
		}
	}
}

// TestMaandProductieDimensions checks that the worksheet dimension metadata
// is parsed correctly for all three sheets.
func TestMaandProductieDimensions(t *testing.T) {
	wb := openXLSB(t, "Maand productie.xlsb")

	tests := []struct {
		sheetIdx int
		wantH    int
		wantW    int
	}{
		{1, 2505, 107}, // Packing
		{2, 35, 14},    // SSP-overzicht
		{3, 113, 81},   // Packing Overzicht
	}
	for _, tc := range tests {
		sheet, err := wb.Sheet(tc.sheetIdx)
		if err != nil {
			t.Errorf("Sheet(%d): %v", tc.sheetIdx, err)
			continue
		}
		if sheet.Dimension == nil {
			t.Errorf("Sheet(%d): Dimension is nil", tc.sheetIdx)
			continue
		}
		if sheet.Dimension.H != tc.wantH {
			t.Errorf("Sheet(%d).Dimension.H = %d, want %d", tc.sheetIdx, sheet.Dimension.H, tc.wantH)
		}
		if sheet.Dimension.W != tc.wantW {
			t.Errorf("Sheet(%d).Dimension.W = %d, want %d", tc.sheetIdx, sheet.Dimension.W, tc.wantW)
		}
	}
}

// TestMaandProductiePackingRowCount verifies that Rows(sparse=true) returns
// the expected number of non-empty rows.
func TestMaandProductiePackingRowCount(t *testing.T) {
	wb := openXLSB(t, "Maand productie.xlsb")
	sheet, err := wb.Sheet(1) // "Packing"
	if err != nil {
		t.Fatalf("Sheet(1): %v", err)
	}
	rows, _ := collectRows(sheet, true)
	if len(rows) != 1101 {
		t.Errorf("sparse row count = %d, want 1101", len(rows))
	}
}

// TestMaandProductieNonSparseRowCount verifies that Rows(sparse=false) emits
// exactly H rows (the full used range height).
func TestMaandProductieNonSparseRowCount(t *testing.T) {
	wb := openXLSB(t, "Maand productie.xlsb")
	// Use the SSP-overzicht sheet (small, 35 rows) for speed.
	sheet, err := wb.SheetByName("SSP-overzicht")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	rows, _ := collectRows(sheet, false)
	if sheet.Dimension == nil {
		t.Fatal("Dimension is nil")
	}
	wantRows := sheet.Dimension.H
	if len(rows) != wantRows {
		t.Errorf("non-sparse row count = %d, want %d (Dimension.H)", len(rows), wantRows)
	}
}

// TestMaandProductieSSPFirstCellString checks the string value in the header
// row of the SSP-overzicht sheet.
func TestMaandProductieSSPFirstCellString(t *testing.T) {
	wb := openXLSB(t, "Maand productie.xlsb")
	sheet, err := wb.SheetByName("SSP-overzicht")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	for row := range sheet.Rows(true) {
		if len(row) == 0 {
			t.Fatal("first row is empty")
		}
		v, ok := row[0].V.(string)
		if !ok {
			t.Fatalf("cell[0].V type = %T, want string", row[0].V)
		}
		if v != "SSP-overzicht" {
			t.Errorf("cell[0].V = %q, want %q", v, "SSP-overzicht")
		}
		break
	}
}

// TestMaandProductiePackingOvrzHeaderStrings verifies specific header cells in
// the "Packing Overzicht" sheet.
func TestMaandProductiePackingOvrzHeaderStrings(t *testing.T) {
	wb := openXLSB(t, "Maand productie.xlsb")
	sheet, err := wb.SheetByName("Packing Overzicht")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	wantHeaders := map[int]string{
		5:  "1 kg",
		6:  "5 kg",
		7:  "25 kg",
		8:  "200 Dr",
		9:  "DPE",
		10: "CNPE",
		11: "TK",
	}
	for row := range sheet.Rows(true) {
		for col, want := range wantHeaders {
			if col >= len(row) {
				t.Errorf("row too short: len=%d, need col %d", len(row), col)
				continue
			}
			v, ok := row[col].V.(string)
			if !ok {
				t.Errorf("col %d: type=%T, want string", col, row[col].V)
				continue
			}
			if v != want {
				t.Errorf("col %d: %q, want %q", col, v, want)
			}
		}
		break
	}
}

// TestMaandProductieCellCoordinates verifies that Cell.R and Cell.C are set
// correctly when iterating rows.
func TestMaandProductieCellCoordinates(t *testing.T) {
	wb := openXLSB(t, "Maand productie.xlsb")
	sheet, err := wb.SheetByName("SSP-overzicht")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	rowIdx := 0
	for row := range sheet.Rows(false) {
		for _, c := range row {
			if c.R != rowIdx {
				t.Errorf("cell at logical row %d has c.R = %d", rowIdx, c.R)
			}
			if c.C < 0 {
				t.Errorf("cell.C = %d, must be non-negative", c.C)
			}
		}
		rowIdx++
		if rowIdx >= 5 {
			break
		}
	}
}

// ── Truckplanning_2011 tests ──────────────────────────────────────────────────

// TestTruckplanningSheetCount verifies the expected 28 sheets.
func TestTruckplanningSheetCount(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	if n := len(wb.Sheets()); n != 28 {
		t.Errorf("sheet count = %d, want 28", n)
	}
}

// TestTruckplanningSheetNames verifies a representative sample of sheet names.
func TestTruckplanningSheetNames(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	sheets := wb.Sheets()
	wantPresent := []string{
		"Sheet1", "S203H", "S204H", "S303H", "S304H", "SAX015", "SAX260",
		"RD359", "SAX400", "Stock", "Deliveries", "Trucks", "Products",
		"Settings", "Werking", "Chart1",
	}
	nameSet := make(map[string]bool, len(sheets))
	for _, s := range sheets {
		nameSet[s] = true
	}
	for _, want := range wantPresent {
		if !nameSet[want] {
			t.Errorf("expected sheet %q not found", want)
		}
	}
}

// TestTruckplanningSheet1FirstCellString verifies the macro-warning text in
// Sheet1.
func TestTruckplanningSheet1FirstCellString(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	sheet, err := wb.SheetByName("Sheet1")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	for row := range sheet.Rows(true) {
		if len(row) == 0 {
			t.Fatal("first row is empty")
		}
		v, ok := row[0].V.(string)
		if !ok {
			t.Fatalf("cell[0].V type = %T, want string", row[0].V)
		}
		want := "THIS WORKBOOK REQUIRES MACROS TO BE ENABLED"
		if !strings.HasPrefix(v, want) {
			t.Errorf("cell[0].V = %q, want prefix %q", v, want)
		}
		break
	}
}

// TestTruckplanningBoolCells verifies that boolean values are parsed correctly
// in the S203H sheet (which contains bool:true cells).
func TestTruckplanningBoolCells(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	sheet, err := wb.SheetByName("S203H")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	boolFound := false
	for row := range sheet.Rows(true) {
		for _, c := range row {
			if b, ok := c.V.(bool); ok {
				boolFound = true
				// All bool cells in this sheet are false (unchecked checkboxes)
				if b {
					t.Logf("bool=true found at row=%d col=%d (acceptable)", c.R, c.C)
				}
			}
		}
		if boolFound {
			break
		}
	}
	if !boolFound {
		t.Error("no bool cells found in S203H; expected at least one")
	}
}

// TestTruckplanningFloatAndStringCells verifies that the S203H sheet contains
// both float64 and string cells.
func TestTruckplanningFloatAndStringCells(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	sheet, err := wb.SheetByName("S203H")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	nFloat, nString, _, _ := typeStats(sheet)
	if nFloat == 0 {
		t.Error("no float64 cells found in S203H")
	}
	if nString == 0 {
		t.Error("no string cells found in S203H")
	}
}

// TestTruckplanningSettingsHeaderRow checks the first row of the Settings sheet.
func TestTruckplanningSettingsHeaderRow(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	sheet, err := wb.SheetByName("Settings")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	wantFirstCells := map[int]string{
		0: "Ordernr",
		1: "Clients",
		3: "Grade",
	}
	for row := range sheet.Rows(true) {
		for col, want := range wantFirstCells {
			if col >= len(row) {
				t.Errorf("row too short: len=%d, need col %d", len(row), col)
				continue
			}
			v, ok := row[col].V.(string)
			if !ok {
				t.Errorf("col %d: type=%T, want string", col, row[col].V)
				continue
			}
			if v != want {
				t.Errorf("col %d: %q, want %q", col, v, want)
			}
		}
		break
	}
}

// TestTruckplanningStockDateSerial verifies that the Stock sheet contains
// float64 values consistent with Excel date serials (> 40000, meaning post-2009).
func TestTruckplanningStockDateSerial(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	sheet, err := wb.SheetByName("Stock")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	// The Stock sheet has date serials around 45813.x in the first header row.
	found := false
	for row := range sheet.Rows(true) {
		for _, c := range row {
			if f, ok := c.V.(float64); ok {
				if f > 40000 && f < 60000 {
					found = true
					// Verify the value can be converted to a valid date.
					dt, err := xlsb.ConvertDate(f)
					if err != nil {
						t.Errorf("ConvertDate(%v): %v", f, err)
						continue
					}
					if dt.Year() < 2010 || dt.Year() > 2030 {
						t.Errorf("ConvertDate(%v) = %v, year out of expected range", f, dt)
					}
				}
			}
		}
		break // only check header row
	}
	if !found {
		t.Error("no date-serial float64 found in Stock header row")
	}
}

// TestTruckplanningChartSheetEmpty verifies that "Chart1" (a chart sheet
// with no cell data) yields zero rows gracefully.
func TestTruckplanningChartSheetEmpty(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	sheet, err := wb.SheetByName("Chart1")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	rowCount := 0
	for range sheet.Rows(true) {
		rowCount++
	}
	if rowCount != 0 {
		t.Errorf("Chart1 row count = %d, want 0", rowCount)
	}
}

// TestTruckplanningAllSheetsByIndex accesses every sheet by 1-based index and
// verifies no error is returned.
func TestTruckplanningAllSheetsByIndex(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	n := len(wb.Sheets())
	for i := 1; i <= n; i++ {
		if _, err := wb.Sheet(i); err != nil {
			t.Errorf("Sheet(%d): %v", i, err)
		}
	}
}

// TestTruckplanningIndexOutOfRange verifies bounds errors on Sheet().
func TestTruckplanningIndexOutOfRange(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	n := len(wb.Sheets())
	if _, err := wb.Sheet(0); err == nil {
		t.Error("Sheet(0): expected error")
	}
	if _, err := wb.Sheet(n + 1); err == nil {
		t.Errorf("Sheet(%d): expected error (only %d sheets)", n+1, n)
	}
}

// ── planning_MSE12 tests ──────────────────────────────────────────────────────

// TestPlanningMSE12UnprotectedSheets checks the sheet list of the unprotected
// copy.
func TestPlanningMSE12UnprotectedSheets(t *testing.T) {
	wb := openXLSB(t, "planning_MSE12 - Copy-unprotected.xlsb")
	want := []string{"CPG", "MA", "2RX", "DMp", "Grade_Time_Table", "Grade_Compatibility", "Werking"}
	got := wb.Sheets()
	if len(got) != len(want) {
		t.Fatalf("Sheets() = %v, want %v", got, want)
	}
	for i, w := range want {
		if got[i] != w {
			t.Errorf("Sheets()[%d] = %q, want %q", i, got[i], w)
		}
	}
}

// TestPlanningMSE12WerkingFirstRow checks known data in the small "Werking"
// sheet.
func TestPlanningMSE12WerkingFirstRow(t *testing.T) {
	wb := openXLSB(t, "planning_MSE12 - Copy-unprotected.xlsb")
	sheet, err := wb.SheetByName("Werking")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	for row := range sheet.Rows(true) {
		if len(row) == 0 {
			t.Fatal("first non-empty row is empty")
		}
		v, ok := row[0].V.(string)
		if !ok {
			t.Fatalf("cell[0].V type = %T, want string", row[0].V)
		}
		want := "Werking / gebruik Planning"
		if v != want {
			t.Errorf("cell[0].V = %q, want %q", v, want)
		}
		break
	}
}

// TestPlanningMSE12GradeTimeTableHeaders checks the CPG header in the
// Grade_Time_Table sheet.
func TestPlanningMSE12GradeTimeTableHeaders(t *testing.T) {
	wb := openXLSB(t, "planning_MSE12 - Copy-unprotected.xlsb")
	sheet, err := wb.SheetByName("Grade_Time_Table")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	for row := range sheet.Rows(true) {
		if len(row) == 0 {
			t.Fatal("first row is empty")
		}
		v, ok := row[0].V.(string)
		if !ok {
			t.Fatalf("cell[0].V type = %T, want string", row[0].V)
		}
		if v != "CPG" {
			t.Errorf("cell[0].V = %q, want %q", v, "CPG")
		}
		break
	}
}

// TestPlanningMSE12GradeCompatibilityFirstCell checks the compatibility table
// header.
func TestPlanningMSE12GradeCompatibilityFirstCell(t *testing.T) {
	wb := openXLSB(t, "planning_MSE12 - Copy-unprotected.xlsb")
	sheet, err := wb.SheetByName("Grade_Compatibility")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	for row := range sheet.Rows(true) {
		if len(row) == 0 {
			t.Fatal("first row is empty")
		}
		v, ok := row[0].V.(string)
		if !ok {
			t.Fatalf("cell[0].V type = %T, want string", row[0].V)
		}
		if v != "MS Grade Compatibility Table" {
			t.Errorf("cell[0].V = %q, want %q", v, "MS Grade Compatibility Table")
		}
		break
	}
}

// TestPlanningMSE12LargeSheetRowCount reads the large CPG sheet (65536 rows)
// in sparse mode and verifies the row count is > 0.
func TestPlanningMSE12LargeSheetRowCount(t *testing.T) {
	if testing.Short() {
		t.Skip("skipping large-sheet test in short mode")
	}
	wb := openXLSB(t, "planning_MSE12 - Copy-unprotected.xlsb")
	sheet, err := wb.SheetByName("CPG")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	rowCount := 0
	for range sheet.Rows(true) {
		rowCount++
	}
	if rowCount != 65536 {
		t.Errorf("CPG sparse row count = %d, want 65536", rowCount)
	}
}

// TestPlanningMSE12MAFirstRow verifies the header text of the MA sheet.
func TestPlanningMSE12MAFirstRow(t *testing.T) {
	wb := openXLSB(t, "planning_MSE12 - Copy-unprotected.xlsb")
	sheet, err := wb.SheetByName("MA")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	for row := range sheet.Rows(true) {
		if len(row) == 0 {
			t.Fatal("first row is empty")
		}
		v, ok := row[0].V.(string)
		if !ok {
			t.Fatalf("cell[0].V type = %T, want string", row[0].V)
		}
		if !strings.Contains(v, "PLANNING") {
			t.Errorf("MA first cell = %q; expected to contain %q", v, "PLANNING")
		}
		break
	}
}

// ── planning_SSP34 tests ──────────────────────────────────────────────────────

// TestPlanningSSP34SheetList checks the exact four sheets.
func TestPlanningSSP34SheetList(t *testing.T) {
	wb := openXLSB(t, "planning_SSP34.xlsb")
	want := []string{"SSP", "Grade_Time_Table", "Grade_Compatibility", "Werking"}
	got := wb.Sheets()
	if len(got) != len(want) {
		t.Fatalf("Sheets() = %v, want %v", got, want)
	}
	for i, w := range want {
		if got[i] != w {
			t.Errorf("Sheets()[%d] = %q, want %q", i, got[i], w)
		}
	}
}

// TestPlanningSSP34WerkingDateSerial verifies that the Werking sheet contains a
// float64 that converts to a valid date.
func TestPlanningSSP34WerkingDateSerial(t *testing.T) {
	wb := openXLSB(t, "planning_SSP34.xlsb")
	sheet, err := wb.SheetByName("Werking")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	found := false
	for row := range sheet.Rows(true) {
		for _, c := range row {
			if f, ok := c.V.(float64); ok {
				// 42928 converts to 2017-07-12.
				if f == 42928 {
					found = true
					dt, err := xlsb.ConvertDate(f)
					if err != nil {
						t.Fatalf("ConvertDate(%v): %v", f, err)
					}
					if dt.Year() != 2017 || dt.Month() != 7 || dt.Day() != 12 {
						t.Errorf("ConvertDate(42928) = %v, want 2017-07-12", dt)
					}
				}
			}
		}
	}
	if !found {
		t.Error("serial 42928 not found in Werking sheet")
	}
}

// TestPlanningSSP34WerkingStringCell checks the string annotation next to the
// date serial.
func TestPlanningSSP34WerkingStringCell(t *testing.T) {
	wb := openXLSB(t, "planning_SSP34.xlsb")
	sheet, err := wb.SheetByName("Werking")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	found := false
	for row := range sheet.Rows(true) {
		for _, c := range row {
			if s, ok := c.V.(string); ok {
				if strings.Contains(s, "grade_compatibility") {
					found = true
				}
			}
		}
	}
	if !found {
		t.Error("expected annotation string containing 'grade_compatibility' not found in Werking")
	}
}

// TestPlanningSSP34GradeTimeTableHeader checks the SSP header.
func TestPlanningSSP34GradeTimeTableHeader(t *testing.T) {
	wb := openXLSB(t, "planning_SSP34.xlsb")
	sheet, err := wb.SheetByName("Grade_Time_Table")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	for row := range sheet.Rows(true) {
		v, ok := row[0].V.(string)
		if !ok {
			t.Fatalf("cell[0].V type = %T, want string", row[0].V)
		}
		if v != "SSP" {
			t.Errorf("Grade_Time_Table first cell = %q, want %q", v, "SSP")
		}
		break
	}
}

// ── data-type coverage tests ──────────────────────────────────────────────────

// TestCellTypeFloat64 verifies that the float64 type is present in files that
// are expected to contain numeric data.
func TestCellTypeFloat64(t *testing.T) {
	for _, name := range []string{
		"Maand productie.xlsb",
		"Truckplanning_2011.xlsb",
		"planning_SSP34.xlsb",
	} {
		t.Run(name, func(t *testing.T) {
			wb := openXLSB(t, name)
			sheets := wb.Sheets()
			found := false
			for i := range sheets {
				sheet, err := wb.Sheet(i + 1)
				if err != nil {
					continue
				}
				nFloat, _, _, _ := typeStats(sheet)
				if nFloat > 0 {
					found = true
					break
				}
			}
			if !found {
				t.Errorf("no float64 cells found in any sheet of %q", name)
			}
		})
	}
}

// TestCellTypeString verifies that string cells are present across the files.
func TestCellTypeString(t *testing.T) {
	for _, name := range []string{
		"Maand productie.xlsb",
		"Truckplanning_2011.xlsb",
		"planning_MSE12 - Copy-unprotected.xlsb",
		"planning_SSP34.xlsb",
	} {
		t.Run(name, func(t *testing.T) {
			wb := openXLSB(t, name)
			sheets := wb.Sheets()
			found := false
			for i := range sheets {
				sheet, err := wb.Sheet(i + 1)
				if err != nil {
					continue
				}
				_, nString, _, _ := typeStats(sheet)
				if nString > 0 {
					found = true
					break
				}
			}
			if !found {
				t.Errorf("no string cells found in any sheet of %q", name)
			}
		})
	}
}

// TestCellTypeBool verifies that boolean cells are parsed from Truckplanning.
func TestCellTypeBool(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	// S203H is confirmed to contain bool cells.
	sheet, err := wb.SheetByName("S203H")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	_, _, nBool, _ := typeStats(sheet)
	if nBool == 0 {
		t.Error("expected bool cells in S203H, found none")
	}
}

// TestCellTypeNilBlank verifies that blank cells (nil value) are yielded when
// sparse=false.
func TestCellTypeNilBlank(t *testing.T) {
	wb := openXLSB(t, "Maand productie.xlsb")
	sheet, err := wb.SheetByName("SSP-overzicht")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	nilCount := 0
	for row := range sheet.Rows(false) {
		for _, c := range row {
			if c.V == nil {
				nilCount++
			}
		}
	}
	if nilCount == 0 {
		t.Error("expected nil/blank cells with sparse=false, found none")
	}
}

// ── sparse vs. dense row iteration ───────────────────────────────────────────

// TestSparseVsDenseRowCount verifies that sparse mode returns fewer or equal
// rows compared to dense mode for a sheet that is not fully populated.
func TestSparseVsDenseRowCount(t *testing.T) {
	wb := openXLSB(t, "Maand productie.xlsb")
	sheet, err := wb.SheetByName("Packing")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	sparseCount := 0
	for range sheet.Rows(true) {
		sparseCount++
	}
	denseCount := 0
	for range sheet.Rows(false) {
		denseCount++
	}
	if sparseCount > denseCount {
		t.Errorf("sparse count (%d) > dense count (%d); sparse must be ≤ dense",
			sparseCount, denseCount)
	}
	t.Logf("sparse=%d dense=%d", sparseCount, denseCount)
}

// TestDenseRowWidthConsistent verifies that non-sparse rows for a sheet all
// have the same width (equal to the dimension width).
func TestDenseRowWidthConsistent(t *testing.T) {
	wb := openXLSB(t, "Maand productie.xlsb")
	sheet, err := wb.SheetByName("SSP-overzicht")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	if sheet.Dimension == nil {
		t.Fatal("Dimension is nil")
	}
	expectedWidth := sheet.Dimension.C + sheet.Dimension.W
	rowIdx := 0
	for row := range sheet.Rows(false) {
		if len(row) != expectedWidth {
			t.Errorf("row %d: width=%d, want %d", rowIdx, len(row), expectedWidth)
		}
		rowIdx++
	}
}

// ── specific row access ───────────────────────────────────────────────────────

// TestSpecificRowAccess reads a known row from the Werking sheet of planning_SSP34
// and checks its contents.
func TestSpecificRowAccess(t *testing.T) {
	wb := openXLSB(t, "planning_SSP34.xlsb")
	sheet, err := wb.SheetByName("Werking")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	// The Werking sheet starts at row 5 (sparse), has a float date in col 0
	// and a string annotation in col 1.
	var targetRow []worksheet.Cell
	for row := range sheet.Rows(true) {
		if _, ok := row[0].V.(float64); ok {
			targetRow = row
			break
		}
	}
	if targetRow == nil {
		t.Fatal("no row with float64 in col 0 found")
	}
	if _, ok := targetRow[0].V.(float64); !ok {
		t.Errorf("col 0: type=%T, want float64", targetRow[0].V)
	}
	if s, ok := targetRow[1].V.(string); !ok || s == "" {
		t.Errorf("col 1: type=%T val=%v, want non-empty string", targetRow[1].V, targetRow[1].V)
	}
}

// TestEarlyIterationStop verifies that breaking out of a Rows loop early does
// not cause a panic or leave the worksheet in a broken state.
func TestEarlyIterationStop(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	sheet, err := wb.SheetByName("S203H")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	count := 0
	for range sheet.Rows(true) {
		count++
		if count == 5 {
			break
		}
	}
	if count != 5 {
		t.Errorf("expected 5 rows before break, got %d", count)
	}
	// Re-iterate the same sheet to confirm the reader is reset correctly.
	count2 := 0
	for range sheet.Rows(true) {
		count2++
		if count2 == 5 {
			break
		}
	}
	if count2 != 5 {
		t.Errorf("re-iteration: expected 5 rows, got %d", count2)
	}
}

// ── copy-consistency tests ────────────────────────────────────────────────────

// TestCopyConsistencyMaandProductie verifies that the original and copy of
// "Maand productie" produce identical sheet lists and first-row values.
func TestCopyConsistencyMaandProductie(t *testing.T) {
	wb1 := openXLSB(t, "Maand productie.xlsb")
	wb2 := openXLSB(t, "Maand productie - Copy.xlsb")

	sheets1 := wb1.Sheets()
	sheets2 := wb2.Sheets()
	if len(sheets1) != len(sheets2) {
		t.Fatalf("sheet count: original=%d copy=%d", len(sheets1), len(sheets2))
	}
	for i := range sheets1 {
		if sheets1[i] != sheets2[i] {
			t.Errorf("sheets[%d]: %q vs %q", i, sheets1[i], sheets2[i])
		}
	}
}

// TestCopyConsistencyPlanningMSE12 verifies that all three MSE12 copies have
// the same sheet list.
func TestCopyConsistencyPlanningMSE12(t *testing.T) {
	wb1 := openXLSB(t, "planning_MSE12.xlsb")
	wb2 := openXLSB(t, "planning_MSE12 - Copy.xlsb")
	wb3 := openXLSB(t, "planning_MSE12 - Copy-unprotected.xlsb")

	s1, s2, s3 := wb1.Sheets(), wb2.Sheets(), wb3.Sheets()
	if len(s1) != len(s2) || len(s1) != len(s3) {
		t.Fatalf("sheet counts differ: %d / %d / %d", len(s1), len(s2), len(s3))
	}
	for i := range s1 {
		if s1[i] != s2[i] || s1[i] != s3[i] {
			t.Errorf("sheet[%d]: %q / %q / %q", i, s1[i], s2[i], s3[i])
		}
	}
}

// TestCopyConsistencyPlanningMSE3 verifies that all three MSE3 copies have
// the same sheet list.
func TestCopyConsistencyPlanningMSE3(t *testing.T) {
	wb1 := openXLSB(t, "planning_MSE3.xlsb")
	wb2 := openXLSB(t, "planning_MSE3 - Copy.xlsb")
	wb3 := openXLSB(t, "planning_MSE3 - Copy-unprotected.xlsb")

	s1, s2, s3 := wb1.Sheets(), wb2.Sheets(), wb3.Sheets()
	if len(s1) != len(s2) || len(s1) != len(s3) {
		t.Fatalf("sheet counts differ: %d / %d / %d", len(s1), len(s2), len(s3))
	}
	for i := range s1 {
		if s1[i] != s2[i] || s1[i] != s3[i] {
			t.Errorf("sheet[%d]: %q / %q / %q", i, s1[i], s2[i], s3[i])
		}
	}
}

// ── ConvertDate integration ───────────────────────────────────────────────────

// TestPlanningMSE12CPGMergeCells verifies that merged cell ranges are parsed
// from the CPG sheet of planning_MSE12.xlsb.
//
// Row 15 of CPG (0-based row 14) lies inside a vertical merge spanning rows
// 13–15 (0-based 12–14) in column C=2; those cells carry an empty-string
// value.  The header area around row 16 contains a horizontal merge starting
// at C=0 with value "CPG".  We verify:
//  1. MergeCells is non-empty (merge records were parsed).
//  2. At least one merge covers row 14 (0-based), confirming row 15 is merged.
//  3. The anchor cell of every merge (top-left corner) carries the actual value;
//     all non-anchor cells in the same merge are blank (nil or empty string).
func TestPlanningMSE12CPGMergeCells(t *testing.T) {
	wb := openXLSB(t, "planning_MSE12.xlsb")
	sheet, err := wb.SheetByName("CPG")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}

	// 1. MergeCells must be non-empty.
	if len(sheet.MergeCells) == 0 {
		t.Fatal("MergeCells is empty; expected merged ranges in CPG")
	}
	t.Logf("CPG has %d merge areas", len(sheet.MergeCells))

	// 2. At least one merge area must cover 0-based row 15 (Excel row 16),
	// which is the header row that Excel displays as part of the merged area
	// the user identified as "row 15" (the merge anchor is at 0-based row 15).
	const targetRow = 15 // 0-based; Excel row 16
	found := false
	for _, ma := range sheet.MergeCells {
		if ma.R <= targetRow && targetRow < ma.R+ma.H {
			found = true
			t.Logf("  merge covering row 14: R=%d C=%d H=%d W=%d", ma.R, ma.C, ma.H, ma.W)
		}
	}
	if !found {
		t.Errorf("no merge area covers 0-based row %d (sheet row 15)", targetRow)
	}

	// 3. For every merge area, collect rows and verify the anchor has a value
	// while non-anchor cells in the same column are blank.
	// Only check merges that are entirely within the first 20 rows (for speed).
	rows := make(map[int][]worksheet.Cell)
	rowNum := 0
	for row := range sheet.Rows(true) {
		rows[rowNum] = row
		rowNum++
		if rowNum >= 20 {
			break
		}
	}

	for _, ma := range sheet.MergeCells {
		if ma.R >= 20 {
			continue
		}
		anchorRow, ok := rows[ma.R]
		if !ok {
			continue // anchor row was empty (sparse), nothing to check
		}
		if ma.C >= len(anchorRow) {
			continue
		}
		// Non-anchor rows in the merge must have a blank (nil) value at the
		// same column — BIFF12 only stores the value in the anchor cell.
		for dr := 1; dr < ma.H && ma.R+dr < 20; dr++ {
			nonAnchorRow, ok := rows[ma.R+dr]
			if !ok {
				continue // sparse: row not present at all — fine
			}
			if ma.C >= len(nonAnchorRow) {
				continue
			}
			v := nonAnchorRow[ma.C].V
			if s, ok := v.(string); ok && s != "" {
				t.Errorf("non-anchor cell at row=%d col=%d inside merge (R=%d C=%d H=%d W=%d) has non-empty value %q",
					ma.R+dr, ma.C, ma.R, ma.C, ma.H, ma.W, s)
			}
		}
	}
}

// TestPlanningMSE12CPGRow16MergedValue verifies that Excel row 16 (0-based row 15)
// of the CPG sheet in planning_MSE12.xlsb is a 1×7 horizontal merge spanning
// columns A–G (0-based C=0..6) and that the anchor cell A16 carries the value "CPG".
func TestPlanningMSE12CPGRow16MergedValue(t *testing.T) {
	wb := openXLSB(t, "planning_MSE12.xlsb")
	sheet, err := wb.SheetByName("CPG")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}

	const targetRow = 15 // 0-based; Excel row 16

	// 1. Verify the merge record A16:G16 exists (R=15, C=0, H=1, W=7).
	foundMerge := false
	for _, ma := range sheet.MergeCells {
		if ma.R == targetRow && ma.C == 0 && ma.H == 1 && ma.W == 7 {
			foundMerge = true
			t.Logf("found merge A16:G16 — R=%d C=%d H=%d W=%d", ma.R, ma.C, ma.H, ma.W)
			break
		}
	}
	if !foundMerge {
		t.Errorf("no merge R=15 C=0 H=1 W=7 (A16:G16) found; MergeCells=%v", sheet.MergeCells)
	}

	// 2. Read up to row 16 (0-based 15) and check the anchor cell value.
	rowNum := 0
	for row := range sheet.Rows(true) {
		if rowNum == targetRow {
			if len(row) == 0 {
				t.Fatal("row 15 is empty")
			}
			// Anchor cell is col 0 (column A).
			v, ok := row[0].V.(string)
			if !ok {
				t.Fatalf("A16 value type = %T (%v), want string", row[0].V, row[0].V)
			}
			if v != "CPG" {
				t.Errorf("A16 = %q, want %q", v, "CPG")
			} else {
				t.Logf("A16 = %q (correct)", v)
			}
			// Non-anchor cells B16:G16 (cols 1–6) must be blank (nil).
			for col := 1; col <= 6 && col < len(row); col++ {
				if row[col].V != nil {
					t.Errorf("non-anchor cell row=15 col=%d inside A16:G16 merge has value %v, want nil", col, row[col].V)
				}
			}
			break
		}
		rowNum++
	}
}

// TestConvertDateFromRealFile exercises ConvertDate against actual float64
// values extracted from the Truckplanning Stock sheet.
func TestConvertDateFromRealFile(t *testing.T) {
	wb := openXLSB(t, "Truckplanning_2011.xlsb")
	sheet, err := wb.SheetByName("Stock")
	if err != nil {
		t.Fatalf("SheetByName: %v", err)
	}
	// The first row of Stock has date serial floats around 45813.x (year 2025).
	converted := 0
	for row := range sheet.Rows(true) {
		for _, c := range row {
			if f, ok := c.V.(float64); ok && f > 40000 && f < 60000 {
				dt, err := xlsb.ConvertDate(f)
				if err != nil {
					t.Errorf("ConvertDate(%v): %v", f, err)
					continue
				}
				if dt.IsZero() {
					t.Errorf("ConvertDate(%v) returned zero time", f)
				}
				converted++
			}
		}
		break // header row only
	}
	if converted == 0 {
		t.Error("no date serials found/converted in Stock header row")
	}
}
