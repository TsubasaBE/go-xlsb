// Package worksheet parses a single .xlsb worksheet binary part and provides
// row/cell iteration.
package worksheet

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"

	"github.com/TsubasaBE/go-xlsb/biff12"
	"github.com/TsubasaBE/go-xlsb/record"
	"github.com/TsubasaBE/go-xlsb/stringtable"
)

// Dimension describes the used range of a worksheet.
type Dimension struct {
	// R is the first row index (0-based).
	R int
	// C is the first column index (0-based).
	C int
	// H is the height (number of rows).
	H int
	// W is the width (number of columns).
	W int
}

// Col describes a column definition record.
// C1 and C2 are the 0-based first and last column indices of the range this
// definition applies to (inclusive). Width is the column width in character
// units (equivalent to the "characters" unit displayed in Excel's column-width
// dialog). Style is the 0-based index into the workbook's cell-format (XF)
// table that applies to blank cells in the range.
type Col struct {
	// C1 is the 0-based index of the first column in the range.
	C1 int
	// C2 is the 0-based index of the last column in the range (inclusive).
	C2 int
	// Width is the column width in character units.
	Width float64
	// Style is the 0-based cell-format (XF) index applied to blank cells.
	Style int
}

// MergeArea describes a merged cell range.
// R and C are the 0-based row and column of the top-left anchor cell.
// H is the height (number of rows) and W is the width (number of columns)
// spanned by the merge.
type MergeArea struct {
	// R is the 0-based row index of the top-left anchor cell.
	R int
	// C is the 0-based column index of the top-left anchor cell.
	C int
	// H is the number of rows spanned by the merge (always >= 1).
	H int
	// W is the number of columns spanned by the merge (always >= 1).
	W int
}

// Cell is a single worksheet cell.
type Cell struct {
	// R is the 0-based row index of the cell.
	R int
	// C is the 0-based column index of the cell.
	C int
	// V holds the typed cell value. The dynamic type is one of:
	//   - nil          — blank / empty cell
	//   - string       — text, formula-string result, or Excel error string (e.g. "#DIV/0!")
	//   - float64      — numeric value, date serial, or formula-float result
	//   - bool         — boolean value
	V any
	// Style is the 0-based index into the workbook's cell-format (XF) table.
	// It is 0 for cells whose record carried no explicit style or for empty
	// padding cells emitted in dense (sparse=false) mode.
	// Use Worksheet.IsDateCell(cell.Style) to test whether the format is a date.
	Style int
}

// Worksheet holds parsed metadata and provides row iteration for one sheet.
type Worksheet struct {
	// Name is the display name of the worksheet as it appears on the sheet tab.
	Name string
	// Dimension describes the used range of the worksheet. It is nil if no
	// DIMENSION record was found in the binary stream.
	Dimension *Dimension
	// Cols contains the column-definition entries parsed from COL records.
	// The slice may be empty if the sheet defines no explicit column widths.
	Cols []Col
	// Hyperlinks maps each hyperlink cell coordinate [row, col] (both 0-based)
	// to its relationship ID, which can be resolved via the workbook's .rels
	// file to obtain the target URL.
	Hyperlinks map[[2]int]string
	// MergeCells contains all merged-cell ranges defined in the sheet.
	MergeCells []MergeArea

	data         []byte                   // full binary payload
	dataOffset   int64                    // byte offset of SHEETDATA record payload
	hasSheetData bool                     // true once SHEETDATA record was found
	stringTable  *stringtable.StringTable // may be nil
	rels         map[string]string        // relationship ID → URL (may be nil)
	dateXFs      map[int]bool             // XF indices that represent date/time formats; may be nil
}

// New parses the pre-loaded binary data and optional rels XML for a worksheet.
// stringTable may be nil if the workbook has no shared strings.
// dateXFs maps XF indices to true when the corresponding number format is a
// date/time format; it may be nil when styles information is unavailable.
func New(name string, data []byte, relsData []byte, st *stringtable.StringTable, dateXFs map[int]bool) (*Worksheet, error) {
	ws := &Worksheet{
		Name:        name,
		Hyperlinks:  make(map[[2]int]string),
		data:        data,
		stringTable: st,
		dateXFs:     dateXFs,
	}
	if len(relsData) > 0 {
		rels, err := parseRelsXML(relsData)
		if err == nil {
			ws.rels = rels
		}
	}
	if err := ws.parse(); err != nil {
		return nil, err
	}
	return ws, nil
}

// IsDateCell reports whether the given XF style index maps to a date or
// datetime number format according to the workbook's styles.
// It returns false when styles information is unavailable or the index is
// not in the date-format set.
func (ws *Worksheet) IsDateCell(style int) bool {
	return ws.dateXFs[style] // nil map returns false safely
}

// Rows iterates over the worksheet rows in order, calling yield for each one.
//
// When sparse is false (the default) empty rows between used rows are emitted
// as slices of nil-valued Cells (matching pyxlsb behaviour).  When sparse is
// true only rows that contain at least one record are yielded.
//
// Merged-cell regions are reflected faithfully from the underlying binary
// storage: only the anchor cell (top-left of the region) carries a value;
// all satellite cells — whether in the same row (horizontal merge) or a
// subsequent row (vertical merge) — return nil, matching the behaviour of
// excelize's GetRows.
//
// Rows uses Go 1.22+ range-over-func semantics.
func (ws *Worksheet) Rows(sparse bool) func(yield func([]Cell) bool) {
	return func(yield func([]Cell) bool) {
		// If the pre-scan never found a SHEETDATA record, the file is malformed.
		// Yield nothing rather than seeking to byte 0 and parsing garbage.
		if !ws.hasSheetData {
			return
		}

		rdr := record.NewReader(bytes.NewReader(ws.data))
		if _, err := rdr.Seek(ws.dataOffset, io.SeekStart); err != nil {
			return
		}

		dim := ws.effectiveDim()
		rowNum := -1
		var row []Cell

		for {
			recID, recData, err := rdr.Next()
			if err != nil {
				break
			}

			switch {
			case recID == biff12.Row:
				r, err := parseRowRecord(recData)
				if err != nil {
					continue
				}
				// Skip duplicate ROW records (same index as the current row).
				// pyxlsb guards: `if item[1].r != row_num`.  Without this a
				// duplicate record would flush the partially-built row early and
				// reset it to an empty slice, losing any cells already parsed.
				if r == rowNum {
					continue
				}
				if row != nil {
					if !yield(row) {
						return
					}
				}
				if !sparse {
					for rowNum < r-1 {
						rowNum++
						if !yield(makeEmptyRow(rowNum, dim)) {
							return
						}
					}
				}
				rowNum = r
				row = makeEmptyRow(rowNum, dim)

			case recID >= biff12.Blank && recID <= biff12.FormulaBoolErr:
				if row == nil {
					continue
				}
				c, err := parseCellRecord(recData, recID, ws.stringTable)
				if err != nil {
					continue
				}
				if c.C >= 0 && c.C < len(row) {
					row[c.C] = Cell{R: rowNum, C: c.C, V: c.V, Style: c.Style}
				}

			case recID == biff12.SheetDataEnd:
				if row != nil {
					yield(row) // caller may have stopped; return either way
				}
				return
			}
		}
		if row != nil {
			yield(row) // caller may have stopped; nothing to do after this
		}
	}
}

// ── internal helpers ──────────────────────────────────────────────────────────

// effectiveDim returns the dimension, or a zero-based 1×1 fallback so we
// never have a nil pointer when building rows.
func (ws *Worksheet) effectiveDim() *Dimension {
	if ws.Dimension != nil {
		return ws.Dimension
	}
	return &Dimension{R: 0, C: 0, H: 1, W: 1}
}

// makeEmptyRow returns a row of nil-valued Cells spanning [dim.C, dim.C+dim.W).
func makeEmptyRow(rowNum int, dim *Dimension) []Cell {
	w := dim.C + dim.W
	if w <= 0 {
		w = 1
	}
	cells := make([]Cell, w)
	for i := range cells {
		cells[i] = Cell{R: rowNum, C: i}
	}
	return cells
}

// parse does the pre-scan pass: reads Dimension, Col defs, Hyperlinks and
// records the byte offset of the SHEETDATA payload start.
func (ws *Worksheet) parse() error {
	rdr := record.NewReader(bytes.NewReader(ws.data))
	for {
		recID, recData, err := rdr.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			return fmt.Errorf("worksheet parse: %w", err)
		}

		switch recID {
		case biff12.Dimension:
			dim, err := parseDimensionRecord(recData)
			if err == nil {
				ws.Dimension = &dim
			}

		case biff12.Col:
			col, err := parseColRecord(recData)
			if err == nil {
				ws.Cols = append(ws.Cols, col)
			}

		case biff12.SheetData:
			// Record the position immediately after the SheetData marker so
			// Rows() can seek back here quickly.
			off, err := rdr.Tell()
			if err != nil {
				return fmt.Errorf("worksheet: tell after SHEETDATA: %w", err)
			}
			ws.dataOffset = off
			ws.hasSheetData = true

			// Skip over the row/cell records inside SheetData so we can
			// continue pre-scanning for MergeCell, Hyperlink, etc. which
			// appear after SheetDataEnd in the stream.
			for {
				id, _, err := rdr.Next()
				if err != nil {
					return nil // truncated stream — treat as end
				}
				if id == biff12.SheetDataEnd {
					break
				}
			}

		case biff12.MergeCell:
			ma, err := parseMergeCellRecord(recData)
			if err == nil {
				ws.MergeCells = append(ws.MergeCells, ma)
			}

		case biff12.Hyperlink:
			if ws.rels == nil {
				continue
			}
			hl, err := parseHyperlinkRecord(recData)
			if err != nil {
				continue
			}
			for dr := range hl.H {
				for dc := range hl.W {
					ws.Hyperlinks[[2]int{hl.R + dr, hl.C + dc}] = hl.RID
				}
			}
		}
	}
	return nil
}

// ── record decoders ───────────────────────────────────────────────────────────

// parseDimensionRecord decodes a DIMENSION record.
//
//	r1 = read_int()   # first row
//	r2 = read_int()   # last row
//	c1 = read_int()   # first col
//	c2 = read_int()   # last col
func parseDimensionRecord(data []byte) (Dimension, error) {
	rr := record.NewRecordReader(data)
	r1, err := rr.ReadUint32()
	if err != nil {
		return Dimension{}, err
	}
	r2, err := rr.ReadUint32()
	if err != nil {
		return Dimension{}, err
	}
	c1, err := rr.ReadUint32()
	if err != nil {
		return Dimension{}, err
	}
	c2, err := rr.ReadUint32()
	if err != nil {
		return Dimension{}, err
	}
	// Validate: last must be >= first to avoid uint32 wrap-around producing
	// enormous H/W values that would cause gigabyte allocations.
	if r2 < r1 {
		return Dimension{}, fmt.Errorf("dimension: r2 (%d) < r1 (%d)", r2, r1)
	}
	if c2 < c1 {
		return Dimension{}, fmt.Errorf("dimension: c2 (%d) < c1 (%d)", c2, c1)
	}
	// Cap to Excel maxima to prevent OOM via makeEmptyRow.
	// Max row index: 1,048,575 (0xFFFFF); max col index: 16,383 (0x3FFF).
	const maxRow = 0xFFFFF
	const maxCol = 0x3FFF
	if r2 > maxRow {
		return Dimension{}, fmt.Errorf("dimension: r2 (%d) exceeds Excel maximum row index %d", r2, maxRow)
	}
	if c2 > maxCol {
		return Dimension{}, fmt.Errorf("dimension: c2 (%d) exceeds Excel maximum column index %d", c2, maxCol)
	}
	return Dimension{
		R: int(r1),
		C: int(c1),
		H: int(r2-r1) + 1,
		W: int(c2-c1) + 1,
	}, nil
}

// parseColRecord decodes a COL record.
//
//	c1    = read_int()
//	c2    = read_int()
//	width = read_int() / 256
//	style = read_int()
func parseColRecord(data []byte) (Col, error) {
	rr := record.NewRecordReader(data)
	c1, err := rr.ReadUint32()
	if err != nil {
		return Col{}, err
	}
	c2, err := rr.ReadUint32()
	if err != nil {
		return Col{}, err
	}
	widthRaw, err := rr.ReadUint32()
	if err != nil {
		return Col{}, err
	}
	style, err := rr.ReadUint32()
	if err != nil {
		return Col{}, err
	}
	// Guard: style is an unvalidated uint32; cap to MaxInt32 so that
	// int(style) produces the same value on both 32-bit and 64-bit platforms.
	const maxStyleIndex = 0x7FFFFFFF
	if style > maxStyleIndex {
		return Col{}, fmt.Errorf("worksheet: column style index %d exceeds limit", style)
	}
	return Col{
		C1:    int(c1),
		C2:    int(c2),
		Width: float64(widthRaw) / 256,
		Style: int(style),
	}, nil
}

// parseRowRecord decodes a ROW record and returns the row index (0-based).
//
//	r = read_int()
func parseRowRecord(data []byte) (int, error) {
	rr := record.NewRecordReader(data)
	r, err := rr.ReadUint32()
	if err != nil {
		return 0, err
	}
	// Excel's maximum row index is 1,048,575 (0xFFFFF).  A corrupt record with
	// a larger value would cause the sparse=false fill loop in Rows() to iterate
	// billions of times and exhaust memory.
	const maxRowIndex = 0xFFFFF
	if r > maxRowIndex {
		return 0, fmt.Errorf("worksheet: row index %d exceeds Excel maximum %d", r, maxRowIndex)
	}
	return int(r), nil
}

// errStrings maps BIFF12 BErr byte codes (MS-XLSB §2.5.97.2) to the
// corresponding Excel error string displayed in a cell.
var errStrings = map[byte]string{
	0x00: "#NULL!",
	0x07: "#DIV/0!",
	0x0F: "#VALUE!",
	0x17: "#REF!",
	0x1D: "#NAME?",
	0x24: "#NUM!",
	0x2A: "#N/A",
	0x2B: "#GETTING_DATA",
}

// errString returns the Excel error string for the given BErr byte code, or a
// hex fallback (e.g. "0xff") for unknown codes.
func errString(b byte) string {
	if s, ok := errStrings[b]; ok {
		return s
	}
	return fmt.Sprintf("0x%02x", b)
}

// internalCell is used only during parsing.
type internalCell struct {
	C     int
	V     any
	Style int
}

// parseCellRecord decodes a cell record (BLANK, NUM, BOOLERR, BOOL, FLOAT,
// STRING, FORMULA_STRING, FORMULA_FLOAT, FORMULA_BOOL, FORMULA_BOOLERR).
//
// Python layout:
//
//	col   = read_int()
//	style = read_int()
//	... type-specific value ...
func parseCellRecord(data []byte, recID int, st *stringtable.StringTable) (internalCell, error) {
	rr := record.NewRecordReader(data)
	col, err := rr.ReadUint32()
	if err != nil {
		return internalCell{}, err
	}
	styleRaw, err := rr.ReadUint32()
	if err != nil {
		return internalCell{C: int(col)}, nil
	}
	// Guard: cap to MaxInt32 so int(styleRaw) is identical on 32- and 64-bit.
	const maxStyleIndex = 0x7FFFFFFF
	style := int(styleRaw)
	if styleRaw > maxStyleIndex {
		style = 0
	}

	var v any
	switch recID {
	case biff12.Num:
		f, err := rr.ReadFloat()
		if err != nil {
			break
		}
		v = f
	case biff12.BoolErr:
		b, err := rr.ReadUint8()
		if err != nil {
			break
		}
		v = errString(b)
	case biff12.Bool:
		b, err := rr.ReadUint8()
		if err != nil {
			break
		}
		v = b != 0
	case biff12.Float:
		f, err := rr.ReadDouble()
		if err != nil {
			break
		}
		v = f
	case biff12.String:
		idx, err := rr.ReadUint32()
		if err != nil {
			break
		}
		// Use uint32 comparison to stay safe on 32-bit platforms where
		// int(uint32(0xFFFFFFFF)) wraps to -1, making -1 < st.Len() always
		// true and causing st.Get(-1) to panic.
		if st != nil && idx < uint32(st.Len()) {
			v = st.Get(int(idx))
		} else {
			v = fmt.Sprintf("<%d>", idx) // fallback if no string table
		}
	case biff12.FormulaString:
		s, err := rr.ReadString()
		if err != nil {
			break
		}
		v = s
	case biff12.FormulaFloat:
		f, err := rr.ReadDouble()
		if err != nil {
			break
		}
		v = f
	case biff12.FormulaBool:
		b, err := rr.ReadUint8()
		if err != nil {
			break
		}
		v = b != 0
	case biff12.FormulaBoolErr:
		b, err := rr.ReadUint8()
		if err != nil {
			break
		}
		v = errString(b)
		// biff12.Blank: v remains nil
	}

	return internalCell{C: int(col), V: v, Style: style}, nil
}

// parseMergeCellRecord decodes a MERGE_CELL record.
//
// Layout is identical to DIMENSION:
//
//	r1 = read_uint32()  // first row (0-based)
//	r2 = read_uint32()  // last  row (0-based, inclusive)
//	c1 = read_uint32()  // first col (0-based)
//	c2 = read_uint32()  // last  col (0-based, inclusive)
func parseMergeCellRecord(data []byte) (MergeArea, error) {
	rr := record.NewRecordReader(data)
	r1, err := rr.ReadUint32()
	if err != nil {
		return MergeArea{}, err
	}
	r2, err := rr.ReadUint32()
	if err != nil {
		return MergeArea{}, err
	}
	c1, err := rr.ReadUint32()
	if err != nil {
		return MergeArea{}, err
	}
	c2, err := rr.ReadUint32()
	if err != nil {
		return MergeArea{}, err
	}
	if r2 < r1 {
		return MergeArea{}, fmt.Errorf("mergecell: r2 (%d) < r1 (%d)", r2, r1)
	}
	if c2 < c1 {
		return MergeArea{}, fmt.Errorf("mergecell: c2 (%d) < c1 (%d)", c2, c1)
	}
	return MergeArea{
		R: int(r1),
		C: int(c1),
		H: int(r2-r1) + 1,
		W: int(c2-c1) + 1,
	}, nil
}

// hyperlinkRecord is a temporary struct for HYPERLINK record parsing.
type hyperlinkRecord struct {
	R, C, H, W int
	RID        string
}

// parseHyperlinkRecord decodes a HYPERLINK record.
//
//	r1  = read_int()
//	r2  = read_int()
//	c1  = read_int()
//	c2  = read_int()
//	rId = read_string()
func parseHyperlinkRecord(data []byte) (hyperlinkRecord, error) {
	rr := record.NewRecordReader(data)
	r1, err := rr.ReadUint32()
	if err != nil {
		return hyperlinkRecord{}, err
	}
	r2, err := rr.ReadUint32()
	if err != nil {
		return hyperlinkRecord{}, err
	}
	c1, err := rr.ReadUint32()
	if err != nil {
		return hyperlinkRecord{}, err
	}
	c2, err := rr.ReadUint32()
	if err != nil {
		return hyperlinkRecord{}, err
	}
	rID, err := rr.ReadString()
	if err != nil {
		return hyperlinkRecord{}, err
	}
	// Validate: last must be >= first to avoid uint32 wrap-around producing
	// enormous H/W values that would cause billions of iterations in the
	// hyperlink-population loop.
	if r2 < r1 {
		return hyperlinkRecord{}, fmt.Errorf("hyperlink: r2 (%d) < r1 (%d)", r2, r1)
	}
	if c2 < c1 {
		return hyperlinkRecord{}, fmt.Errorf("hyperlink: c2 (%d) < c1 (%d)", c2, c1)
	}
	// Cap to Excel maxima.
	const maxRow = 0xFFFFF
	const maxCol = 0x3FFF
	if r2 > maxRow {
		return hyperlinkRecord{}, fmt.Errorf("hyperlink: r2 (%d) exceeds Excel maximum row index %d", r2, maxRow)
	}
	if c2 > maxCol {
		return hyperlinkRecord{}, fmt.Errorf("hyperlink: c2 (%d) exceeds Excel maximum column index %d", c2, maxCol)
	}
	return hyperlinkRecord{
		R:   int(r1),
		C:   int(c1),
		H:   int(r2-r1) + 1,
		W:   int(c2-c1) + 1,
		RID: rID,
	}, nil
}

// ── XML relationship parsing (duplicated from workbook to avoid circular dep) ─

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
