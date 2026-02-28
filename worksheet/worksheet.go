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
type Col struct {
	C1    int
	C2    int
	Width float64
	Style int
}

// MergeArea describes a merged cell range.
// R and C are the 0-based row and column of the top-left anchor cell.
// H is the height (number of rows) and W is the width (number of columns)
// spanned by the merge.
type MergeArea struct {
	R, C, H, W int
}

// Cell is a single worksheet cell.
// V holds the typed value:
//   - nil          — blank / empty
//   - string       — text or formula-string result
//   - float64      — numeric, date-serial, formula-float result
//   - bool         — boolean
//   - string       — error value formatted as hex (e.g. "0x2a")
type Cell struct {
	R int
	C int
	V any
}

// Worksheet holds parsed metadata and provides row iteration for one sheet.
type Worksheet struct {
	Name       string
	Dimension  *Dimension
	Cols       []Col
	Hyperlinks map[[2]int]string // [row,col] → relationship ID
	MergeCells []MergeArea       // all merged ranges in this sheet

	data         []byte                   // full binary payload
	dataOffset   int64                    // byte offset of SHEETDATA record payload
	hasSheetData bool                     // true once SHEETDATA record was found
	stringTable  *stringtable.StringTable // may be nil
	rels         map[string]string        // relationship ID → URL (may be nil)
}

// New parses the pre-loaded binary data and optional rels XML for a worksheet.
// stringTable may be nil if the workbook has no shared strings.
func New(name string, data []byte, relsData []byte, st *stringtable.StringTable) (*Worksheet, error) {
	ws := &Worksheet{
		Name:        name,
		Hyperlinks:  make(map[[2]int]string),
		data:        data,
		stringTable: st,
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

// Rows iterates over the worksheet rows in order, calling yield for each one.
//
// When sparse is false (the default) empty rows between used rows are emitted
// as slices of nil-valued Cells (matching pyxlsb behaviour).  When sparse is
// true only rows that contain at least one record are yielded.
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
						empty := makeEmptyRow(rowNum, dim)
						if !yield(empty) {
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
					row[c.C] = Cell{R: rowNum, C: c.C, V: c.V}
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

// internalCell is used only during parsing.
type internalCell struct {
	C int
	V any
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
	if _, err := rr.ReadUint32(); err != nil { // style — not exposed at this level
		return internalCell{C: int(col)}, nil
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
		v = fmt.Sprintf("0x%02x", b)
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
		v = fmt.Sprintf("0x%02x", b)
		// biff12.Blank: v remains nil
	}

	return internalCell{C: int(col), V: v}, nil
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
