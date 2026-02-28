// Package workbook opens and parses an .xlsb workbook file (a ZIP archive).
package workbook

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"encoding/xml"
	"fmt"
	"io"
	"strings"

	"github.com/TsubasaBE/go-xlsb/biff12"
	"github.com/TsubasaBE/go-xlsb/numfmt"
	"github.com/TsubasaBE/go-xlsb/record"
	"github.com/TsubasaBE/go-xlsb/stringtable"
	"github.com/TsubasaBE/go-xlsb/styles"
	"github.com/TsubasaBE/go-xlsb/worksheet"
)

// Sheet visibility levels, as stored in the hsState field of a BrtBundleSh
// record (MS-XLSB §2.4.720). Use these constants with SheetVisibility.
const (
	// SheetVisible indicates the sheet tab is visible (hsState == 0).
	SheetVisible = 0
	// SheetHidden indicates the sheet is hidden but can be unhidden by the
	// user via Excel's "Unhide" dialog (hsState == 1).
	SheetHidden = 1
	// SheetVeryHidden indicates the sheet is hidden and cannot be unhidden
	// through the Excel UI — only via VBA or programmatic access (hsState == 2).
	SheetVeryHidden = 2
)

// sheetEntry holds the display name and the zip-internal path target for one
// worksheet.
type sheetEntry struct {
	name       string
	target     string // e.g. "worksheets/sheet1.bin"
	visibility int    // SheetVisible, SheetHidden, or SheetVeryHidden
}

// Workbook represents an open .xlsb workbook.
type Workbook struct {
	zr          *zip.ReadCloser // non-nil when opened by file name
	zf          *zip.Reader     // always non-nil
	sheets      []sheetEntry
	stringTable *stringtable.StringTable
	// Styles is the full XF style table parsed from xl/styles.bin.  It is
	// exported so that callers who need low-level access to format metadata
	// can inspect it directly; normal callers should use FormatCell.
	Styles styles.StyleTable
	// Date1904 is true when the workbook uses the 1904 date system (base
	// date 1904-01-01, serial 0 = 1904-01-01). Most workbooks use the
	// default 1900 system (Date1904 == false). Pass this value to
	// ConvertDateEx when converting numeric cell values to time.Time.
	Date1904 bool
}

// Open opens the named .xlsb file and parses its workbook metadata.
// The caller must call Close on the returned Workbook when done to release the
// underlying file handle.
func Open(name string) (*Workbook, error) {
	rc, err := zip.OpenReader(name)
	if err != nil {
		return nil, fmt.Errorf("workbook: open %q: %w", name, err)
	}
	wb := &Workbook{zr: rc, zf: &rc.Reader}
	if err := wb.parse(); err != nil {
		_ = rc.Close()
		return nil, err
	}
	return wb, nil
}

// OpenReader parses an .xlsb workbook from an in-memory ReaderAt.
// size must be the total byte size of the ZIP data.
func OpenReader(r io.ReaderAt, size int64) (*Workbook, error) {
	zf, err := zip.NewReader(r, size)
	if err != nil {
		return nil, fmt.Errorf("workbook: open reader: %w", err)
	}
	wb := &Workbook{zf: zf}
	if err := wb.parse(); err != nil {
		return nil, err
	}
	return wb, nil
}

// Sheets returns the display names of all worksheets in order.
func (wb *Workbook) Sheets() []string {
	names := make([]string, len(wb.sheets))
	for i, s := range wb.sheets {
		names[i] = s.name
	}
	return names
}

// Sheet returns the worksheet at the given 1-based index.
// Index 1 refers to the first sheet. An out-of-range index returns a non-nil
// error describing the valid range.
func (wb *Workbook) Sheet(idx int) (*worksheet.Worksheet, error) {
	if idx < 1 || idx > len(wb.sheets) {
		return nil, fmt.Errorf("workbook: sheet index %d out of range [1, %d]", idx, len(wb.sheets))
	}
	return wb.openSheet(wb.sheets[idx-1])
}

// SheetByName returns the worksheet with the given name (case-insensitive).
// It returns a non-nil error if no sheet with that name exists.
func (wb *Workbook) SheetByName(name string) (*worksheet.Worksheet, error) {
	lower := strings.ToLower(name)
	for _, s := range wb.sheets {
		if strings.ToLower(s.name) == lower {
			return wb.openSheet(s)
		}
	}
	return nil, fmt.Errorf("workbook: sheet %q not found", name)
}

// SheetVisible reports whether the named sheet is visible (case-insensitive).
// It returns false for hidden sheets, very-hidden sheets, and unknown names.
// To distinguish hidden from very-hidden, use SheetVisibility.
func (wb *Workbook) SheetVisible(name string) bool {
	return wb.SheetVisibility(name) == SheetVisible
}

// SheetVisibility returns the visibility level of the named sheet
// (case-insensitive): SheetVisible (0), SheetHidden (1), or SheetVeryHidden (2).
// It returns -1 if no sheet with that name exists.
func (wb *Workbook) SheetVisibility(name string) int {
	lower := strings.ToLower(name)
	for _, s := range wb.sheets {
		if strings.ToLower(s.name) == lower {
			return s.visibility
		}
	}
	return -1
}

// FormatCell renders the cell value v using the XF style at index styleIdx.
// Pass cell.V as v and cell.Style as styleIdx.
//
// The returned string is the same display string that Excel would show in the
// cell.  Use this alongside Rows() to get both the raw value (cell.V) and
// the formatted display string:
//
//	for row := range sheet.Rows(false) {
//	    for _, cell := range row {
//	        raw       := cell.V
//	        formatted := wb.FormatCell(cell.V, cell.Style)
//	        _ = raw
//	        _ = formatted
//	    }
//	}
//
// When styleIdx is out of range (e.g. because styles.bin was absent), the
// function falls back to fmt.Sprint(v).
func (wb *Workbook) FormatCell(v any, styleIdx int) string {
	if styleIdx < 0 || styleIdx >= len(wb.Styles) {
		if v == nil {
			return ""
		}
		return fmt.Sprint(v)
	}
	s := wb.Styles[styleIdx]
	return numfmt.FormatValue(v, s.NumFmtID, s.FormatStr, wb.Date1904)
}

// Close releases the underlying ZIP file handle.
// It is a no-op when the workbook was opened via OpenReader (no file handle to
// release), and always returns nil in that case.
func (wb *Workbook) Close() error {
	if wb.zr != nil {
		return wb.zr.Close()
	}
	return nil
}

// ── internal ─────────────────────────────────────────────────────────────────

// parse reads workbook.bin, sharedStrings.bin (if present), and styles.bin.
func (wb *Workbook) parse() error {
	if err := wb.parseWorkbook(); err != nil {
		return err
	}
	if err := wb.parseSharedStrings(); err != nil {
		return err
	}
	if err := wb.parseStyles(); err != nil {
		return err
	}
	return nil
}

// parseWorkbook reads xl/_rels/workbook.bin.rels (XML) and xl/workbook.bin
// to build the sheet list.
func (wb *Workbook) parseWorkbook() error {
	// Step 1: load relationship ID → target map from the .rels XML.
	rels, err := wb.readRels("xl/_rels/workbook.bin.rels")
	if err != nil {
		return fmt.Errorf("workbook: parse rels: %w", err)
	}

	// Step 2: read workbook.bin record stream.
	data, err := wb.readZipEntry("xl/workbook.bin")
	if err != nil {
		return fmt.Errorf("workbook: read workbook.bin: %w", err)
	}

	rdr := record.NewReader(bytes.NewReader(data))
	for {
		recID, recData, err := rdr.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			return fmt.Errorf("workbook: %w", err)
		}

		switch recID {
		case biff12.WorkbookPr:
			// BrtWbProp payload (MS-XLSB §2.4.822): first uint32 is a flags field.
			// Bit 3 (0x08) is f1904DateSystem — set when the workbook uses the
			// 1904 date system (base date 1904-01-01, serial 0 = 1904-01-01).
			if len(recData) >= 4 {
				flags := binary.LittleEndian.Uint32(recData[:4])
				wb.Date1904 = (flags & 0x08) != 0
			}
		case biff12.Sheet:
			entry, err := parseSheetRecord(recData, rels)
			if err != nil {
				return fmt.Errorf("workbook: parse SHEET record: %w", err)
			}
			wb.sheets = append(wb.sheets, entry)
		case biff12.SheetsEnd:
			return nil
		}
	}
	return nil
}

// parseSharedStrings reads xl/sharedStrings.bin if it exists.
func (wb *Workbook) parseSharedStrings() error {
	data, err := wb.readZipEntry("xl/sharedStrings.bin")
	if err != nil {
		// File is optional — no shared strings in this workbook.
		return nil
	}
	st, err := stringtable.New(bytes.NewReader(data))
	if err != nil {
		return fmt.Errorf("workbook: shared strings: %w", err)
	}
	wb.stringTable = st
	return nil
}

// parseStyles reads xl/styles.bin and builds the StyleTable.
// Failures are silently ignored so that workbooks without styles.bin
// (or with malformed styles) still open correctly — FormatCell will fall
// back to fmt.Sprint for all cells.
func (wb *Workbook) parseStyles() error {
	data, err := wb.readZipEntry("xl/styles.bin")
	if err != nil {
		return nil // optional
	}
	st, err := parseStyleTable(data)
	if err != nil {
		return nil // degrade gracefully
	}
	wb.Styles = st
	return nil
}

// parseStyleTable parses the BIFF12 styles stream and returns a StyleTable
// mapping each XF index to its resolved XFStyle.
//
// BrtFmt record layout (MS-XLSB §2.4.697):
//
//	numFmtId  uint16
//	stFmtCode ReadString (4-byte char-count + UTF-16LE)
//
// BrtXF record layout (MS-XLSB §2.4.674) — we only read the first two fields:
//
//	ixfe      uint16   (parent XF index; ignored)
//	numFmtId  uint16
//	...       (remaining fields ignored)
func parseStyleTable(data []byte) (styles.StyleTable, error) {
	// fmts maps numFmtId → format string for custom formats (id >= 164).
	fmts := make(map[int]string)
	var table styles.StyleTable

	rdr := record.NewReader(bytes.NewReader(data))
	inCellXfs := false

	for {
		recID, recData, err := rdr.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("workbook: styles: %w", err)
		}

		switch recID {
		case biff12.NumFmt:
			// BrtFmt: numFmtId(uint16) + format string
			if len(recData) < 2 {
				continue
			}
			fmtID := int(binary.LittleEndian.Uint16(recData[:2]))
			rr := record.NewRecordReader(recData[2:])
			fmtStr, _ := rr.ReadString() // ignore error — use empty string
			fmts[fmtID] = fmtStr

		case biff12.CellXfs:
			inCellXfs = true

		case biff12.CellXfsEnd:
			inCellXfs = false

		case biff12.Xf:
			if !inCellXfs {
				continue // skip style-XF entries in CellStyleXfs
			}
			// BrtXF: ixfe(uint16) + numFmtId(uint16) + ...
			if len(recData) < 4 {
				table = append(table, styles.XFStyle{})
				continue
			}
			// ixfe is at bytes 0–1; numFmtId is at bytes 2–3.
			numFmtID := int(binary.LittleEndian.Uint16(recData[2:4]))
			fmtStr := fmts[numFmtID] // empty string for built-in IDs
			table = append(table, styles.XFStyle{
				NumFmtID:  numFmtID,
				FormatStr: fmtStr,
			})
		}
	}
	return table, nil
}

// isDateFormatID is the internal counterpart of xlsb.IsDateFormat.
// It is kept here (rather than delegating to styles.isDateFormatID) so that
// workbook remains self-contained when the styles package is not imported by
// callers.  All three copies must stay in sync.
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
			ch == 'h' || ch == 'H' ||
			ch == 's' || ch == 'S':
			return true
		}
	}
	return false
}

// openSheet reads the binary data for the given sheet entry and returns a
// ready-to-use Worksheet.
func (wb *Workbook) openSheet(entry sheetEntry) (*worksheet.Worksheet, error) {
	// Resolve "worksheets/sheet1.bin" → "xl/worksheets/sheet1.bin".
	// Absolute targets (starting with "/") are used as-is after stripping the
	// leading slash; relative targets are prefixed with "xl/".
	target := strings.TrimPrefix(entry.target, "/")
	var zipPath string
	if strings.HasPrefix(target, "xl/") {
		zipPath = target
	} else {
		zipPath = "xl/" + target
	}

	data, err := wb.readZipEntry(zipPath)
	if err != nil {
		return nil, fmt.Errorf("workbook: open sheet %q: %w", entry.name, err)
	}

	// Attempt to load the sheet .rels file (optional; needed for hyperlinks).
	lastSlash := strings.LastIndex(zipPath, "/")
	relsPath := zipPath[:lastSlash+1] + "_rels/" + zipPath[lastSlash+1:] + ".rels"
	relsData, _ := wb.readZipEntry(relsPath) // ignore error — it's optional

	return worksheet.New(entry.name, data, relsData, wb.stringTable, wb.Styles, wb.FormatCell)
}

// readZipEntry reads the full contents of a named entry from the ZIP archive.
func (wb *Workbook) readZipEntry(name string) ([]byte, error) {
	for _, f := range wb.zf.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			data, readErr := io.ReadAll(rc)
			closeErr := rc.Close()
			if readErr != nil {
				return nil, readErr
			}
			// Propagate decompressor checksum / close errors even when the read
			// appeared to succeed (e.g. truncated gzip stream).
			if closeErr != nil {
				return nil, closeErr
			}
			return data, nil
		}
	}
	return nil, fmt.Errorf("%q not found in archive", name)
}

// readRels parses a .rels XML file and returns a map of Id → Target.
func (wb *Workbook) readRels(name string) (map[string]string, error) {
	data, err := wb.readZipEntry(name)
	if err != nil {
		return nil, err
	}
	return parseRelsXML(data)
}

// ── XML relationship parsing ──────────────────────────────────────────────────

type xmlRelationships struct {
	Relationships []xmlRelationship `xml:"Relationship"`
}

type xmlRelationship struct {
	ID     string `xml:"Id,attr"`
	Target string `xml:"Target,attr"`
}

func parseRelsXML(data []byte) (map[string]string, error) {
	var rels xmlRelationships
	if err := xml.Unmarshal(data, &rels); err != nil {
		return nil, fmt.Errorf("parse rels XML: %w", err)
	}
	m := make(map[string]string, len(rels.Relationships))
	for _, r := range rels.Relationships {
		m[r.ID] = r.Target
	}
	return m, nil
}

// ── SHEET record parsing ───────────────────────────────────────────────────────

// parseSheetRecord decodes a SHEET record payload.
//
// BrtBundleSh layout (MS-XLSB §2.4.720):
//
//	hsState = read_uint32() & 0x03   # low 2 bits: 0=visible, 1=hidden, 2=veryHidden
//	sheetId = read_uint32()
//	relId   = read_string()
//	name    = read_string()
func parseSheetRecord(data []byte, rels map[string]string) (sheetEntry, error) {
	rr := record.NewRecordReader(data)

	flags, err := rr.ReadUint32()
	if err != nil {
		return sheetEntry{}, fmt.Errorf("read state flags: %w", err)
	}
	visibility := int(flags & 0x03)

	if _, err := rr.ReadUint32(); err != nil { // sheetId — not used by us
		return sheetEntry{}, fmt.Errorf("read sheetId: %w", err)
	}
	relID, err := rr.ReadString()
	if err != nil {
		return sheetEntry{}, fmt.Errorf("read relId: %w", err)
	}
	name, err := rr.ReadString()
	if err != nil {
		return sheetEntry{}, fmt.Errorf("read sheet name: %w", err)
	}

	target, ok := rels[relID]
	if !ok {
		return sheetEntry{}, fmt.Errorf("no relationship found for rId %q", relID)
	}
	return sheetEntry{name: name, target: target, visibility: visibility}, nil
}
