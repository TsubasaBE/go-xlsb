package xlsb_test

// Unit tests for the go-xlsb library.
//
// The tests are intentionally self-contained: they build all binary fixtures
// in memory so no external .xlsb file is required.

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"math"
	"testing"
	"time"

	"github.com/TsubasaBE/go-xlsb"
	"github.com/TsubasaBE/go-xlsb/numfmt"
	"github.com/TsubasaBE/go-xlsb/record"
	"github.com/TsubasaBE/go-xlsb/stringtable"
	"github.com/TsubasaBE/go-xlsb/workbook"
	"github.com/TsubasaBE/go-xlsb/worksheet"
)

// ── ConvertDate ───────────────────────────────────────────────────────────────

func TestConvertDate(t *testing.T) {
	tests := []struct {
		name    string
		input   float64
		want    time.Time
		wantErr bool
	}{
		{
			name:  "serial 0 gives 1900-01-01",
			input: 0,
			want:  time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC),
		},
		{
			name:  "serial 0 with time component",
			input: 0.5,
			want:  time.Date(1900, 1, 1, 12, 0, 0, 0, time.UTC),
		},
		{
			name:  "serial 1 gives 1900-01-01 (base+1 day)",
			input: 1,
			want:  time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC),
		},
		{
			name:  "serial 60 gives 1900-03-01 (phantom leap day)",
			input: 60,
			want:  time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC),
		},
		{
			name:  "serial 61 compensates for Lotus leap-year bug",
			input: 61,
			want:  time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC),
		},
		{
			// pyxlsb: convert_date(41235.45578) == datetime(2012, 11, 22, 10, 56, 19)
			name:  "pyxlsb example: 41235.45578",
			input: 41235.45578,
			want:  time.Date(2012, 11, 22, 10, 56, 19, 0, time.UTC),
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got, err := xlsb.ConvertDate(tc.input)
			if tc.wantErr {
				if err == nil {
					t.Fatalf("expected error, got nil")
				}
				return
			}
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if !got.Equal(tc.want) {
				t.Errorf("ConvertDate(%v) = %v, want %v", tc.input, got, tc.want)
			}
		})
	}
}

// ── RecordReader ──────────────────────────────────────────────────────────────

func TestRecordReaderReadUint8(t *testing.T) {
	rr := record.NewRecordReader([]byte{0xAB})
	v, err := rr.ReadUint8()
	if err != nil {
		t.Fatal(err)
	}
	if v != 0xAB {
		t.Errorf("got 0x%02X, want 0xAB", v)
	}
}

func TestRecordReaderReadUint32(t *testing.T) {
	buf := []byte{0x01, 0x00, 0x00, 0x00} // little-endian 1
	rr := record.NewRecordReader(buf)
	v, err := rr.ReadUint32()
	if err != nil {
		t.Fatal(err)
	}
	if v != 1 {
		t.Errorf("got %d, want 1", v)
	}
}

func TestRecordReaderReadDouble(t *testing.T) {
	// IEEE-754 little-endian encoding of 42.0
	var buf [8]byte
	binary.LittleEndian.PutUint64(buf[:], 0x4045000000000000)
	rr := record.NewRecordReader(buf[:])
	v, err := rr.ReadDouble()
	if err != nil {
		t.Fatal(err)
	}
	if v != 42.0 {
		t.Errorf("got %v, want 42.0", v)
	}
}

func TestRecordReaderReadString(t *testing.T) {
	// "Hi" in UTF-16LE = 0x48 0x00 0x69 0x00; length = 2
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, uint32(2))
	buf.Write([]byte{0x48, 0x00, 0x69, 0x00})
	rr := record.NewRecordReader(buf.Bytes())
	s, err := rr.ReadString()
	if err != nil {
		t.Fatal(err)
	}
	if s != "Hi" {
		t.Errorf("got %q, want %q", s, "Hi")
	}
}

func TestRecordReaderSkip(t *testing.T) {
	rr := record.NewRecordReader([]byte{0x00, 0x00, 0xFF})
	if err := rr.Skip(2); err != nil {
		t.Fatal(err)
	}
	v, err := rr.ReadUint8()
	if err != nil {
		t.Fatal(err)
	}
	if v != 0xFF {
		t.Errorf("got 0x%02X, want 0xFF", v)
	}
}

func TestRecordReaderUnderflow(t *testing.T) {
	rr := record.NewRecordReader([]byte{0x01})
	if _, err := rr.ReadUint32(); err == nil {
		t.Error("expected error reading uint32 from 1-byte buffer")
	}
}

// ── record.Reader (BIFF12 stream) ─────────────────────────────────────────────

// encodeRecord writes a minimal BIFF12 record (single-byte ID + single-byte
// length) with the given payload into buf.
func encodeRecord(recID int, payload []byte) []byte {
	var buf bytes.Buffer
	// ID: if < 0x80, single byte
	buf.WriteByte(byte(recID))
	// Length: simple single-byte (payload length must be < 128 for these tests)
	buf.WriteByte(byte(len(payload)))
	buf.Write(payload)
	return buf.Bytes()
}

func TestReaderNextSingleRecord(t *testing.T) {
	payload := []byte{0xDE, 0xAD}
	raw := encodeRecord(0x07, payload) // record type 7
	rdr := record.NewReader(bytes.NewReader(raw))
	id, data, err := rdr.Next()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if id != 0x07 {
		t.Errorf("id = 0x%02X, want 0x07", id)
	}
	if !bytes.Equal(data, payload) {
		t.Errorf("data = %v, want %v", data, payload)
	}
}

func TestReaderNextEOF(t *testing.T) {
	rdr := record.NewReader(bytes.NewReader([]byte{}))
	_, _, err := rdr.Next()
	if err == nil {
		t.Fatal("expected EOF, got nil")
	}
}

func TestReaderMultipleRecords(t *testing.T) {
	var buf bytes.Buffer
	buf.Write(encodeRecord(0x01, []byte{0xAA}))
	buf.Write(encodeRecord(0x02, []byte{0xBB}))

	rdr := record.NewReader(bytes.NewReader(buf.Bytes()))

	id1, d1, err := rdr.Next()
	if err != nil || id1 != 0x01 || d1[0] != 0xAA {
		t.Fatalf("first record wrong: id=%d data=%v err=%v", id1, d1, err)
	}
	id2, d2, err := rdr.Next()
	if err != nil || id2 != 0x02 || d2[0] != 0xBB {
		t.Fatalf("second record wrong: id=%d data=%v err=%v", id2, d2, err)
	}
}

// ── StringTable ───────────────────────────────────────────────────────────────

// buildSSTBytes constructs a minimal in-memory sharedStrings.bin containing
// the given strings, starting with an SST record and ending with SST_END.
//
// SST  record id = 0x019F → encoded as two bytes: 0x9F 0x83 (continuation)
// Actually for testing we use the record package encoder which handles IDs < 128.
// But SST = 0x019F needs two bytes.  We build the raw bytes manually.
func buildSSTBytes(strs []string) []byte {
	var buf bytes.Buffer

	// Helper: write a variable-length record ID (up to 2 bytes).
	// The BIFF12 ID encoding: each byte except the last has bit 7 set (continuation).
	// Value is accumulated as: v += byte << (8*i).
	// IDs are designed so that multi-byte ones always have bit7 set in the low byte.
	writeID := func(id int) {
		if id < 0x80 {
			buf.WriteByte(byte(id))
		} else {
			// Write the low byte (already has bit 7 set for valid BIFF12 IDs).
			buf.WriteByte(byte(id & 0xFF))
			// Write the high byte (no continuation bit needed for 2-byte IDs).
			buf.WriteByte(byte(id >> 8))
		}
	}

	// Helper: write a variable-length length field
	writeLen := func(n int) {
		for {
			b := n & 0x7F
			n >>= 7
			if n > 0 {
				buf.WriteByte(byte(b) | 0x80)
			} else {
				buf.WriteByte(byte(b))
				break
			}
		}
	}

	// Helper: encode a string as 4-byte count + UTF-16LE body
	encStr := func(s string) []byte {
		runes := []rune(s)
		var sb bytes.Buffer
		_ = binary.Write(&sb, binary.LittleEndian, uint32(len(runes)))
		for _, r := range runes {
			_ = binary.Write(&sb, binary.LittleEndian, uint16(r))
		}
		return sb.Bytes()
	}

	// SST record (id=0x019F): count(4) + uniqueCount(4)
	sstPayload := make([]byte, 8)
	binary.LittleEndian.PutUint32(sstPayload[0:], uint32(len(strs)))
	binary.LittleEndian.PutUint32(sstPayload[4:], uint32(len(strs)))
	writeID(0x019F)
	writeLen(len(sstPayload))
	buf.Write(sstPayload)

	// SI records (id=0x0013)
	for _, s := range strs {
		payload := append([]byte{0x00}, encStr(s)...) // flag byte + string
		writeID(0x0013)
		writeLen(len(payload))
		buf.Write(payload)
	}

	// SST_END record (id=0x01A0)
	writeID(0x01A0)
	writeLen(0)

	return buf.Bytes()
}

func TestStringTable(t *testing.T) {
	strs := []string{"hello", "world", "foo"}
	data := buildSSTBytes(strs)

	st, err := stringtable.New(bytes.NewReader(data))
	if err != nil {
		t.Fatalf("StringTable.New: %v", err)
	}
	if st.Len() != len(strs) {
		t.Fatalf("Len() = %d, want %d", st.Len(), len(strs))
	}
	for i, want := range strs {
		if got := st.Get(i); got != want {
			t.Errorf("Get(%d) = %q, want %q", i, got, want)
		}
	}
}

// ── Workbook (in-memory ZIP) ───────────────────────────────────────────────────

// buildMinimalXLSB constructs a minimal valid .xlsb ZIP in memory with a single
// sheet named "TestSheet" containing one row: [42.0, "hello"].
func buildMinimalXLSB(t *testing.T) []byte {
	t.Helper()

	// ── helpers ───────────────────────────────────────────────────────────────

	writeID := func(buf *bytes.Buffer, id int) {
		if id < 0x80 {
			buf.WriteByte(byte(id))
		} else {
			buf.WriteByte(byte(id & 0xFF))
			buf.WriteByte(byte(id >> 8))
		}
	}
	writeLen := func(buf *bytes.Buffer, n int) {
		for {
			b := n & 0x7F
			n >>= 7
			if n > 0 {
				buf.WriteByte(byte(b) | 0x80)
			} else {
				buf.WriteByte(byte(b))
				break
			}
		}
	}
	writeRec := func(buf *bytes.Buffer, id int, payload []byte) {
		writeID(buf, id)
		writeLen(buf, len(payload))
		buf.Write(payload)
	}
	encStr := func(s string) []byte {
		runes := []rune(s)
		var sb bytes.Buffer
		_ = binary.Write(&sb, binary.LittleEndian, uint32(len(runes)))
		for _, r := range runes {
			_ = binary.Write(&sb, binary.LittleEndian, uint16(r))
		}
		return sb.Bytes()
	}
	le32 := func(v uint32) []byte {
		b := make([]byte, 4)
		binary.LittleEndian.PutUint32(b, v)
		return b
	}

	// ── xl/workbook.bin ───────────────────────────────────────────────────────

	var wb bytes.Buffer
	// WORKBOOK start
	writeRec(&wb, 0x0183, nil)
	// SHEETS start
	writeRec(&wb, 0x018F, nil)

	// SHEET record: skip(4) + sheetId(4) + relId(string) + name(string)
	var sheetRec bytes.Buffer
	sheetRec.Write(le32(0))             // 4 unknown bytes (state flags)
	sheetRec.Write(le32(1))             // sheetId = 1
	sheetRec.Write(encStr("rId1"))      // relId
	sheetRec.Write(encStr("TestSheet")) // name
	writeRec(&wb, 0x019C, sheetRec.Bytes())

	// SHEETS end
	writeRec(&wb, 0x0190, nil)
	// WORKBOOK end
	writeRec(&wb, 0x0184, nil)

	// ── xl/sharedStrings.bin ──────────────────────────────────────────────────

	sstBuf := bytes.NewBuffer(buildSSTBytes([]string{"hello"}))

	// ── xl/worksheets/sheet1.bin ──────────────────────────────────────────────

	var ws bytes.Buffer
	// WORKSHEET start
	writeRec(&ws, 0x0181, nil)
	// DIMENSION: r1=0, r2=0, c1=0, c2=1
	var dimPayload bytes.Buffer
	dimPayload.Write(le32(0)) // r1
	dimPayload.Write(le32(0)) // r2
	dimPayload.Write(le32(0)) // c1
	dimPayload.Write(le32(1)) // c2
	writeRec(&ws, 0x0194, dimPayload.Bytes())
	// SHEETDATA start
	writeRec(&ws, 0x0191, nil)

	// ROW record: r=0
	writeRec(&ws, 0x0000, le32(0))

	// FLOAT cell at col 0: col(4) + style(4) + double(8) = 42.0
	var floatCell bytes.Buffer
	floatCell.Write(le32(0)) // col
	floatCell.Write(le32(0)) // style
	var f64buf [8]byte
	binary.LittleEndian.PutUint64(f64buf[:], 0x4045000000000000) // 42.0
	floatCell.Write(f64buf[:])
	writeRec(&ws, 0x0005, floatCell.Bytes()) // FLOAT

	// STRING cell at col 1: col(4) + style(4) + index(4) = shared string 0 = "hello"
	var strCell bytes.Buffer
	strCell.Write(le32(1))                 // col
	strCell.Write(le32(0))                 // style
	strCell.Write(le32(0))                 // shared string index 0
	writeRec(&ws, 0x0007, strCell.Bytes()) // STRING

	// SHEETDATA end
	writeRec(&ws, 0x0192, nil)
	// WORKSHEET end
	writeRec(&ws, 0x0182, nil)

	// ── assemble ZIP ─────────────────────────────────────────────────────────

	var zipBuf bytes.Buffer
	zw := zip.NewWriter(&zipBuf)

	addFile := func(name string, data []byte) {
		t.Helper()
		f, err := zw.Create(name)
		if err != nil {
			t.Fatalf("zip create %s: %v", name, err)
		}
		if _, err := f.Write(data); err != nil {
			t.Fatalf("zip write %s: %v", name, err)
		}
	}

	// Workbook relationship file
	relsXML := `<?xml version="1.0" encoding="UTF-8"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.bin"/>` +
		`</Relationships>`
	addFile("xl/_rels/workbook.bin.rels", []byte(relsXML))
	addFile("xl/workbook.bin", wb.Bytes())
	addFile("xl/sharedStrings.bin", sstBuf.Bytes())
	addFile("xl/worksheets/sheet1.bin", ws.Bytes())

	if err := zw.Close(); err != nil {
		t.Fatalf("zip close: %v", err)
	}
	return zipBuf.Bytes()
}

func TestWorkbookSheets(t *testing.T) {
	data := buildMinimalXLSB(t)
	wb, err := workbook.OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer wb.Close()

	sheets := wb.Sheets()
	if len(sheets) != 1 {
		t.Fatalf("len(Sheets) = %d, want 1", len(sheets))
	}
	if sheets[0] != "TestSheet" {
		t.Errorf("Sheets()[0] = %q, want %q", sheets[0], "TestSheet")
	}
}

func TestWorkbookSheetByName(t *testing.T) {
	data := buildMinimalXLSB(t)
	wb, err := workbook.OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer wb.Close()

	// Case-insensitive lookup
	_, err = wb.SheetByName("testsheet")
	if err != nil {
		t.Errorf("SheetByName (lower) unexpected error: %v", err)
	}

	_, err = wb.SheetByName("nonexistent")
	if err == nil {
		t.Error("expected error for missing sheet, got nil")
	}
}

func TestWorkbookSheetRows(t *testing.T) {
	data := buildMinimalXLSB(t)
	wb, err := workbook.OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer wb.Close()

	sheet, err := wb.Sheet(1)
	if err != nil {
		t.Fatalf("Sheet(1): %v", err)
	}

	var allRows [][]interface{}
	for row := range sheet.Rows(true) {
		rowVals := make([]interface{}, len(row))
		for i, c := range row {
			rowVals[i] = c.V
		}
		allRows = append(allRows, rowVals)
	}

	if len(allRows) != 1 {
		t.Fatalf("got %d rows, want 1", len(allRows))
	}
	row := allRows[0]
	if len(row) < 2 {
		t.Fatalf("row has %d cells, want ≥2", len(row))
	}
	if v, ok := row[0].(float64); !ok || v != 42.0 {
		t.Errorf("cell[0] = %v (%T), want float64(42.0)", row[0], row[0])
	}
	if v, ok := row[1].(string); !ok || v != "hello" {
		t.Errorf("cell[1] = %v (%T), want string(\"hello\")", row[1], row[1])
	}
}

func TestWorkbookSheetIndexOutOfRange(t *testing.T) {
	data := buildMinimalXLSB(t)
	wb, err := workbook.OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer wb.Close()

	if _, err := wb.Sheet(0); err == nil {
		t.Error("expected error for Sheet(0)")
	}
	if _, err := wb.Sheet(2); err == nil {
		t.Error("expected error for Sheet(2) when only 1 sheet exists")
	}
}

// ── ConvertDate edge cases ─────────────────────────────────────────────────────

func TestConvertDateNaN(t *testing.T) {
	_, err := xlsb.ConvertDate(math.Inf(1)) // +Inf
	if err == nil {
		t.Error("expected error for Inf input")
	}
}

// TestRecordReaderReadFloat tests the packed-float (NUM cell) decoding.
// The integer-path branch: raw=6 (0b110 → bit1=1 → integer path, val = 6>>2 = 1; bit0=0 → no /100) → 1.0
func TestRecordReaderReadFloat(t *testing.T) {
	// raw int32 = 6 (binary 110): bit1 set → integer path, val = 6>>2 = 1; bit0 clear → no /100
	var buf [4]byte
	binary.LittleEndian.PutUint32(buf[:], 6)
	rr := record.NewRecordReader(buf[:])
	v, err := rr.ReadFloat()
	if err != nil {
		t.Fatal(err)
	}
	if v != 1.0 {
		t.Errorf("ReadFloat() = %v, want 1.0", v)
	}
}

// TestStringContains verifies the UTF-16LE decoder handles multi-char strings.
func TestRecordReaderUnicode(t *testing.T) {
	// "café" in UTF-16LE: c=0x63, a=0x61, f=0x66, é=0xE9 (all in BMP)
	word := "café"
	runes := []rune(word)
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, uint32(len(runes)))
	for _, r := range runes {
		_ = binary.Write(&buf, binary.LittleEndian, uint16(r))
	}
	rr := record.NewRecordReader(buf.Bytes())
	got, err := rr.ReadString()
	if err != nil {
		t.Fatal(err)
	}
	if got != word {
		t.Errorf("got %q, want %q", got, word)
	}
}

// ── MergeCell satellite values ────────────────────────────────────────────────

// buildMergeXLSB constructs an in-memory .xlsb with two merge regions:
//
//   - Vertical merge A1:A3 (rows 0–2, col 0) — anchor value "Grade"
//   - Horizontal merge A5:C5 (row 4, cols 0–2) — anchor value "Header"
//
// Only the anchor cell record is written for each region; satellite cells have
// no record, matching the actual .xlsb on-disk layout.
func buildMergeXLSB(t *testing.T) []byte {
	t.Helper()

	writeID := func(buf *bytes.Buffer, id int) {
		if id < 0x80 {
			buf.WriteByte(byte(id))
		} else {
			buf.WriteByte(byte(id & 0xFF))
			buf.WriteByte(byte(id >> 8))
		}
	}
	writeLen := func(buf *bytes.Buffer, n int) {
		for {
			b := n & 0x7F
			n >>= 7
			if n > 0 {
				buf.WriteByte(byte(b) | 0x80)
			} else {
				buf.WriteByte(byte(b))
				break
			}
		}
	}
	writeRec := func(buf *bytes.Buffer, id int, payload []byte) {
		writeID(buf, id)
		writeLen(buf, len(payload))
		buf.Write(payload)
	}
	encStr := func(s string) []byte {
		runes := []rune(s)
		var sb bytes.Buffer
		_ = binary.Write(&sb, binary.LittleEndian, uint32(len(runes)))
		for _, r := range runes {
			_ = binary.Write(&sb, binary.LittleEndian, uint16(r))
		}
		return sb.Bytes()
	}
	le32 := func(v uint32) []byte {
		b := make([]byte, 4)
		binary.LittleEndian.PutUint32(b, v)
		return b
	}
	// mergeRec encodes a MERGE_CELL record: r1, r2, c1, c2 (all uint32).
	mergeRec := func(r1, r2, c1, c2 uint32) []byte {
		var p bytes.Buffer
		p.Write(le32(r1))
		p.Write(le32(r2))
		p.Write(le32(c1))
		p.Write(le32(c2))
		return p.Bytes()
	}
	// formulaStringCell encodes a FormulaString cell: col(4)+style(4)+string.
	formulaStringCell := func(col uint32, s string) []byte {
		var p bytes.Buffer
		p.Write(le32(col))
		p.Write(le32(0)) // style
		p.Write(encStr(s))
		return p.Bytes()
	}

	// ── xl/workbook.bin ───────────────────────────────────────────────────────

	var wb bytes.Buffer
	writeRec(&wb, 0x0183, nil) // WORKBOOK start
	writeRec(&wb, 0x018F, nil) // SHEETS start
	var sheetRec bytes.Buffer
	sheetRec.Write(le32(0))
	sheetRec.Write(le32(1))
	sheetRec.Write(encStr("rId1"))
	sheetRec.Write(encStr("MergeSheet"))
	writeRec(&wb, 0x019C, sheetRec.Bytes())
	writeRec(&wb, 0x0190, nil) // SHEETS end
	writeRec(&wb, 0x0184, nil) // WORKBOOK end

	// ── xl/worksheets/sheet1.bin ──────────────────────────────────────────────

	var ws bytes.Buffer
	writeRec(&ws, 0x0181, nil) // WORKSHEET start

	// DIMENSION: rows 0–4, cols 0–2
	var dimPay bytes.Buffer
	dimPay.Write(le32(0)) // r1
	dimPay.Write(le32(4)) // r2
	dimPay.Write(le32(0)) // c1
	dimPay.Write(le32(2)) // c2
	writeRec(&ws, 0x0194, dimPay.Bytes())

	writeRec(&ws, 0x0191, nil) // SHEETDATA start

	// Row 0: anchor of vertical merge — FormulaString "Grade" at col 0
	writeRec(&ws, 0x0000, le32(0))
	writeRec(&ws, 0x0008, formulaStringCell(0, "Grade")) // FormulaString = 0x0008

	// Row 1: satellite of vertical merge — no cell records written
	writeRec(&ws, 0x0000, le32(1))

	// Row 2: satellite of vertical merge — no cell records written
	writeRec(&ws, 0x0000, le32(2))

	// Row 4: anchor of horizontal merge — FormulaString "Header" at col 0
	writeRec(&ws, 0x0000, le32(4))
	writeRec(&ws, 0x0008, formulaStringCell(0, "Header")) // FormulaString = 0x0008

	writeRec(&ws, 0x0192, nil) // SHEETDATA end

	// MERGE_CELL records (appear after SHEETDATA in the stream)
	// Vertical merge: A1:A3 → rows 0–2, col 0
	writeRec(&ws, 0x00E5, mergeRec(0, 2, 0, 0)) // MergeCell = 0x00E5
	// Horizontal merge: A5:C5 → row 4, cols 0–2
	writeRec(&ws, 0x00E5, mergeRec(4, 4, 0, 2))

	writeRec(&ws, 0x0182, nil) // WORKSHEET end

	// ── assemble ZIP ─────────────────────────────────────────────────────────

	var zipBuf bytes.Buffer
	zw := zip.NewWriter(&zipBuf)
	addFile := func(name string, data []byte) {
		t.Helper()
		f, err := zw.Create(name)
		if err != nil {
			t.Fatalf("zip create %s: %v", name, err)
		}
		if _, err := f.Write(data); err != nil {
			t.Fatalf("zip write %s: %v", name, err)
		}
	}
	relsXML := `<?xml version="1.0" encoding="UTF-8"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.bin"/>` +
		`</Relationships>`
	addFile("xl/_rels/workbook.bin.rels", []byte(relsXML))
	addFile("xl/workbook.bin", wb.Bytes())
	addFile("xl/worksheets/sheet1.bin", ws.Bytes())
	if err := zw.Close(); err != nil {
		t.Fatalf("zip close: %v", err)
	}
	return zipBuf.Bytes()
}

// TestWorksheetVerticalMergeSatellitesEmpty verifies that satellite cells in a
// vertical merge region return nil — not the anchor cell's value — matching the
// behaviour of excelize's GetRows.  This is a regression test for the bug where
// go-xlsb was incorrectly propagating the anchor value into non-anchor rows.
func TestWorksheetVerticalMergeSatellitesEmpty(t *testing.T) {
	data := buildMergeXLSB(t)
	wb, err := workbook.OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer wb.Close()

	sheet, err := wb.Sheet(1)
	if err != nil {
		t.Fatalf("Sheet(1): %v", err)
	}

	// Collect all rows (dense mode so every row index 0–4 is present).
	rows := make(map[int][]worksheet.Cell)
	for row := range sheet.Rows(false) {
		if len(row) > 0 {
			rows[row[0].R] = row
		}
	}

	tests := []struct {
		name    string
		rowIdx  int
		colIdx  int
		wantVal any
		desc    string
	}{
		// ── vertical merge A1:A3 ─────────────────────────────────────────────
		{
			name:   "vertical anchor row 0 col 0",
			rowIdx: 0, colIdx: 0,
			wantVal: "Grade",
			desc:    "anchor cell must carry the value",
		},
		{
			name:   "vertical satellite row 1 col 0",
			rowIdx: 1, colIdx: 0,
			wantVal: nil,
			desc:    "first satellite must be nil, not propagated",
		},
		{
			name:   "vertical satellite row 2 col 0",
			rowIdx: 2, colIdx: 0,
			wantVal: nil,
			desc:    "second satellite must be nil, not propagated",
		},
		// ── horizontal merge A5:C5 ───────────────────────────────────────────
		{
			name:   "horizontal anchor row 4 col 0",
			rowIdx: 4, colIdx: 0,
			wantVal: "Header",
			desc:    "anchor cell must carry the value",
		},
		{
			name:   "horizontal satellite row 4 col 1",
			rowIdx: 4, colIdx: 1,
			wantVal: nil,
			desc:    "horizontal satellite col 1 must be nil",
		},
		{
			name:   "horizontal satellite row 4 col 2",
			rowIdx: 4, colIdx: 2,
			wantVal: nil,
			desc:    "horizontal satellite col 2 must be nil",
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			row, ok := rows[tc.rowIdx]
			if !ok {
				t.Fatalf("row %d not found in output", tc.rowIdx)
			}
			if tc.colIdx >= len(row) {
				t.Fatalf("row %d has only %d cells; want col %d", tc.rowIdx, len(row), tc.colIdx)
			}
			got := row[tc.colIdx].V
			if got != tc.wantVal {
				t.Errorf("%s: cell[%d][%d].V = %v (%T), want %v — %s",
					tc.name, tc.rowIdx, tc.colIdx, got, got, tc.wantVal, tc.desc)
			}
		})
	}
}

// Compile-time check: the top-level xlsb package is importable and ConvertDate
// is accessible.
var _ = xlsb.ConvertDate

// ── IsDateFormat ──────────────────────────────────────────────────────────────

func TestIsDateFormat(t *testing.T) {
	tests := []struct {
		name      string
		id        int
		formatStr string
		want      bool
	}{
		// ── built-in date IDs ─────────────────────────────────────────────────
		{name: "built-in 14 (m/d/yy)", id: 14, want: true},
		{name: "built-in 15", id: 15, want: true},
		{name: "built-in 16", id: 16, want: true},
		{name: "built-in 17", id: 17, want: true},
		{name: "built-in 22 (m/d/yy h:mm)", id: 22, want: true},
		{name: "built-in 27", id: 27, want: true},
		{name: "built-in 36", id: 36, want: true},
		{name: "built-in 45", id: 45, want: true},
		{name: "built-in 46", id: 46, want: true},
		{name: "built-in 47", id: 47, want: true},
		{name: "built-in 50", id: 50, want: true},
		{name: "built-in 58", id: 58, want: true},
		// ── built-in non-date IDs ─────────────────────────────────────────────
		{name: "built-in 0 (General)", id: 0, want: false},
		{name: "built-in 1 (0)", id: 1, want: false},
		{name: "built-in 4 (0.00)", id: 4, want: false},
		{name: "built-in 11 (0.00E+00)", id: 11, want: false},
		{name: "built-in 49 (@)", id: 49, want: false},
		{name: "boundary 13", id: 13, want: false},
		{name: "boundary 18", id: 18, want: false},
		{name: "boundary 23", id: 23, want: false},
		{name: "boundary 163", id: 163, want: false},
		// ── custom format IDs (>= 164) ────────────────────────────────────────
		{name: "custom yyyy-mm-dd", id: 164, formatStr: "yyyy-mm-dd", want: true},
		{name: "custom dd/mm/yyyy hh:mm", id: 165, formatStr: "dd/mm/yyyy hh:mm", want: true},
		{name: "custom numeric 0.00", id: 166, formatStr: "0.00", want: false},
		{name: "custom text @", id: 167, formatStr: "@", want: false},
		// d inside double quotes must not trigger
		{name: "custom quoted d", id: 168, formatStr: `"date"0.00`, want: false},
		// y inside square brackets (locale) must not trigger
		{name: "custom bracketed y", id: 169, formatStr: `[$-409]0.00`, want: false},
		// uppercase variants
		{name: "custom YYYY", id: 170, formatStr: "YYYY", want: true},
		{name: "custom MM", id: 171, formatStr: "MM", want: true},
		{name: "custom DD", id: 172, formatStr: "DD", want: true},
		{name: "custom HH", id: 173, formatStr: "HH:MM", want: true},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got := xlsb.IsDateFormat(tc.id, tc.formatStr)
			if got != tc.want {
				t.Errorf("IsDateFormat(%d, %q) = %v, want %v", tc.id, tc.formatStr, got, tc.want)
			}
		})
	}
}

// ── Cell.Style ────────────────────────────────────────────────────────────────

// buildMinimalXLSBWithStyle is like buildMinimalXLSB but writes style=wantStyle
// into the FLOAT cell at column 0 so we can verify Cell.Style is populated.
func buildMinimalXLSBWithStyle(t *testing.T, wantStyle uint32) []byte {
	t.Helper()

	writeID := func(buf *bytes.Buffer, id int) {
		if id < 0x80 {
			buf.WriteByte(byte(id))
		} else {
			buf.WriteByte(byte(id & 0xFF))
			buf.WriteByte(byte(id >> 8))
		}
	}
	writeLen := func(buf *bytes.Buffer, n int) {
		for {
			b := n & 0x7F
			n >>= 7
			if n > 0 {
				buf.WriteByte(byte(b) | 0x80)
			} else {
				buf.WriteByte(byte(b))
				break
			}
		}
	}
	writeRec := func(buf *bytes.Buffer, id int, payload []byte) {
		writeID(buf, id)
		writeLen(buf, len(payload))
		buf.Write(payload)
	}
	encStr := func(s string) []byte {
		runes := []rune(s)
		var sb bytes.Buffer
		_ = binary.Write(&sb, binary.LittleEndian, uint32(len(runes)))
		for _, r := range runes {
			_ = binary.Write(&sb, binary.LittleEndian, uint16(r))
		}
		return sb.Bytes()
	}
	le32 := func(v uint32) []byte {
		b := make([]byte, 4)
		binary.LittleEndian.PutUint32(b, v)
		return b
	}

	// xl/workbook.bin
	var wb bytes.Buffer
	writeRec(&wb, 0x0183, nil)
	writeRec(&wb, 0x018F, nil)
	var sheetRec bytes.Buffer
	sheetRec.Write(le32(0))
	sheetRec.Write(le32(1))
	sheetRec.Write(encStr("rId1"))
	sheetRec.Write(encStr("Sheet1"))
	writeRec(&wb, 0x019C, sheetRec.Bytes())
	writeRec(&wb, 0x0190, nil)
	writeRec(&wb, 0x0184, nil)

	// xl/worksheets/sheet1.bin  — FLOAT cell at col 0 with the requested style
	var ws bytes.Buffer
	writeRec(&ws, 0x0181, nil)
	var dimPay bytes.Buffer
	dimPay.Write(le32(0))
	dimPay.Write(le32(0))
	dimPay.Write(le32(0))
	dimPay.Write(le32(0))
	writeRec(&ws, 0x0194, dimPay.Bytes())
	writeRec(&ws, 0x0191, nil)
	writeRec(&ws, 0x0000, le32(0))
	var floatCell bytes.Buffer
	floatCell.Write(le32(0))         // col
	floatCell.Write(le32(wantStyle)) // style
	var f64buf [8]byte
	binary.LittleEndian.PutUint64(f64buf[:], 0x4045000000000000) // 42.0
	floatCell.Write(f64buf[:])
	writeRec(&ws, 0x0005, floatCell.Bytes())
	writeRec(&ws, 0x0192, nil)
	writeRec(&ws, 0x0182, nil)

	var zipBuf bytes.Buffer
	zw := zip.NewWriter(&zipBuf)
	addFile := func(name string, data []byte) {
		t.Helper()
		f, err := zw.Create(name)
		if err != nil {
			t.Fatalf("zip create %s: %v", name, err)
		}
		if _, err := f.Write(data); err != nil {
			t.Fatalf("zip write %s: %v", name, err)
		}
	}
	relsXML := `<?xml version="1.0" encoding="UTF-8"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.bin"/>` +
		`</Relationships>`
	addFile("xl/_rels/workbook.bin.rels", []byte(relsXML))
	addFile("xl/workbook.bin", wb.Bytes())
	addFile("xl/worksheets/sheet1.bin", ws.Bytes())
	if err := zw.Close(); err != nil {
		t.Fatalf("zip close: %v", err)
	}
	return zipBuf.Bytes()
}

func TestCellStyleIndex(t *testing.T) {
	const wantStyle = uint32(7)
	data := buildMinimalXLSBWithStyle(t, wantStyle)
	wb, err := workbook.OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer wb.Close()

	sheet, err := wb.Sheet(1)
	if err != nil {
		t.Fatalf("Sheet(1): %v", err)
	}

	var found bool
	for row := range sheet.Rows(true) {
		for _, cell := range row {
			if cell.C == 0 {
				found = true
				if cell.Style != int(wantStyle) {
					t.Errorf("cell.Style = %d, want %d", cell.Style, wantStyle)
				}
			}
		}
	}
	if !found {
		t.Error("did not find any cell at column 0")
	}
}

// ── IsDateCell ────────────────────────────────────────────────────────────────

// buildStylesBin constructs a minimal xl/styles.bin BIFF12 stream containing:
//   - One BrtFmt record: id=164, format="yyyy-mm-dd"
//   - CellXfs section with three BrtXF records:
//     xf[0]: numFmtId=14  (built-in date)   → IsDateCell(0) == true
//     xf[1]: numFmtId=164 (custom date)     → IsDateCell(1) == true
//     xf[2]: numFmtId=0   (General/normal)  → IsDateCell(2) == false
func buildStylesBin(t *testing.T) []byte {
	t.Helper()

	writeID := func(buf *bytes.Buffer, id int) {
		if id < 0x80 {
			buf.WriteByte(byte(id))
		} else {
			buf.WriteByte(byte(id & 0xFF))
			buf.WriteByte(byte(id >> 8))
		}
	}
	writeLen := func(buf *bytes.Buffer, n int) {
		for {
			b := n & 0x7F
			n >>= 7
			if n > 0 {
				buf.WriteByte(byte(b) | 0x80)
			} else {
				buf.WriteByte(byte(b))
				break
			}
		}
	}
	writeRec := func(buf *bytes.Buffer, id int, payload []byte) {
		writeID(buf, id)
		writeLen(buf, len(payload))
		buf.Write(payload)
	}
	le16 := func(v uint16) []byte {
		b := make([]byte, 2)
		binary.LittleEndian.PutUint16(b, v)
		return b
	}
	encStr := func(s string) []byte {
		runes := []rune(s)
		var sb bytes.Buffer
		_ = binary.Write(&sb, binary.LittleEndian, uint32(len(runes)))
		for _, r := range runes {
			_ = binary.Write(&sb, binary.LittleEndian, uint16(r))
		}
		return sb.Bytes()
	}
	// BrtXF payload: ixfe(2) + numFmtId(2) + fontId(2) + fillId(2) + borderId(2) + flags(4)
	// We only need the first four bytes correct; the rest can be zeros.
	makeXF := func(numFmtID uint16) []byte {
		var p bytes.Buffer
		p.Write(le16(0))         // ixfe
		p.Write(le16(numFmtID))  // numFmtId
		p.Write(make([]byte, 8)) // fontId + fillId + borderId + flags (zeros)
		return p.Bytes()
	}

	var buf bytes.Buffer

	// StyleSheet start (0x0296)
	writeRec(&buf, 0x0296, nil)

	// BrtFmt record: numFmtId=164 + "yyyy-mm-dd"
	var fmtPay bytes.Buffer
	fmtPay.Write(le16(164))
	fmtPay.Write(encStr("yyyy-mm-dd"))
	writeRec(&buf, 0x002C, fmtPay.Bytes())

	// CellXfs start (0x04E9)
	writeRec(&buf, 0x04E9, nil)
	// xf[0]: numFmtId=14 (built-in date)
	writeRec(&buf, 0x002F, makeXF(14))
	// xf[1]: numFmtId=164 (custom date)
	writeRec(&buf, 0x002F, makeXF(164))
	// xf[2]: numFmtId=0 (General)
	writeRec(&buf, 0x002F, makeXF(0))
	// CellXfs end (0x04EA)
	writeRec(&buf, 0x04EA, nil)

	// StyleSheet end (0x0297)
	writeRec(&buf, 0x0297, nil)

	return buf.Bytes()
}

// buildXLSBWithStylesBin builds a minimal .xlsb ZIP that includes xl/styles.bin.
func buildXLSBWithStylesBin(t *testing.T) []byte {
	t.Helper()

	writeID := func(buf *bytes.Buffer, id int) {
		if id < 0x80 {
			buf.WriteByte(byte(id))
		} else {
			buf.WriteByte(byte(id & 0xFF))
			buf.WriteByte(byte(id >> 8))
		}
	}
	writeLen := func(buf *bytes.Buffer, n int) {
		for {
			b := n & 0x7F
			n >>= 7
			if n > 0 {
				buf.WriteByte(byte(b) | 0x80)
			} else {
				buf.WriteByte(byte(b))
				break
			}
		}
	}
	writeRec := func(buf *bytes.Buffer, id int, payload []byte) {
		writeID(buf, id)
		writeLen(buf, len(payload))
		buf.Write(payload)
	}
	encStr := func(s string) []byte {
		runes := []rune(s)
		var sb bytes.Buffer
		_ = binary.Write(&sb, binary.LittleEndian, uint32(len(runes)))
		for _, r := range runes {
			_ = binary.Write(&sb, binary.LittleEndian, uint16(r))
		}
		return sb.Bytes()
	}
	le32 := func(v uint32) []byte {
		b := make([]byte, 4)
		binary.LittleEndian.PutUint32(b, v)
		return b
	}

	// xl/workbook.bin
	var wb bytes.Buffer
	writeRec(&wb, 0x0183, nil)
	writeRec(&wb, 0x018F, nil)
	var sheetRec bytes.Buffer
	sheetRec.Write(le32(0))
	sheetRec.Write(le32(1))
	sheetRec.Write(encStr("rId1"))
	sheetRec.Write(encStr("Sheet1"))
	writeRec(&wb, 0x019C, sheetRec.Bytes())
	writeRec(&wb, 0x0190, nil)
	writeRec(&wb, 0x0184, nil)

	// xl/worksheets/sheet1.bin — minimal, no cells needed
	var ws bytes.Buffer
	writeRec(&ws, 0x0181, nil)
	writeRec(&ws, 0x0191, nil)
	writeRec(&ws, 0x0192, nil)
	writeRec(&ws, 0x0182, nil)

	var zipBuf bytes.Buffer
	zw := zip.NewWriter(&zipBuf)
	addFile := func(name string, data []byte) {
		t.Helper()
		f, err := zw.Create(name)
		if err != nil {
			t.Fatalf("zip create %s: %v", name, err)
		}
		if _, err := f.Write(data); err != nil {
			t.Fatalf("zip write %s: %v", name, err)
		}
	}
	relsXML := `<?xml version="1.0" encoding="UTF-8"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.bin"/>` +
		`</Relationships>`
	addFile("xl/_rels/workbook.bin.rels", []byte(relsXML))
	addFile("xl/workbook.bin", wb.Bytes())
	addFile("xl/styles.bin", buildStylesBin(t))
	addFile("xl/worksheets/sheet1.bin", ws.Bytes())
	if err := zw.Close(); err != nil {
		t.Fatalf("zip close: %v", err)
	}
	return zipBuf.Bytes()
}

// ── FormatValue (numfmt package) ──────────────────────────────────────────────

func TestFormatValueGeneral(t *testing.T) {
	tests := []struct {
		name string
		v    any
		want string
	}{
		{"nil", nil, ""},
		{"bool true", true, "TRUE"},
		{"bool false", false, "FALSE"},
		{"string passthrough", "hello", "hello"},
		{"integer float", float64(42), "42"},
		{"fractional float", float64(3.14), "3.14"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got := numfmt.FormatValue(tc.v, 0, "", false)
			if got != tc.want {
				t.Errorf("FormatValue(%v) = %q, want %q", tc.v, got, tc.want)
			}
		})
	}
}

// TestFormatValueDecimal covers Cat F: decimal precision formatting.
func TestFormatValueDecimal(t *testing.T) {
	tests := []struct {
		name   string
		v      float64
		fmtStr string
		want   string
	}{
		{"0.00 with 303.6", 303.6, "0.00", "303.60"},
		{"0.00 with zero", 0.0, "0.00", "0.00"},
		{"0.## trims trailing zero", 1.5, "0.##", "1.5"},
		{"0 integer only", 42.9, "0", "43"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got := numfmt.FormatValue(tc.v, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueLiteralPrefix covers Cat D: literal prefix + number.
func TestFormatValueLiteralPrefix(t *testing.T) {
	tests := []struct {
		name   string
		v      float64
		fmtStr string
		want   string
	}{
		{"E prefix", 40013205, `"E"0`, "E40013205"},
		{"unit suffix kg", 18000, `0" kg"`, "18000 kg"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got := numfmt.FormatValue(tc.v, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValuePercent covers percent scaling.
func TestFormatValuePercent(t *testing.T) {
	tests := []struct {
		name   string
		v      float64
		fmtStr string
		want   string
	}{
		{"0% with 0.75", 0.75, "0%", "75%"},
		{"0.00% with 0.1234", 0.1234, "0.00%", "12.34%"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got := numfmt.FormatValue(tc.v, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueDateLocale covers Cat B: built-in locale date (numFmtID=14, m/d/yy).
func TestFormatValueDateLocale(t *testing.T) {
	// Excel serial 45412 = 2024-04-30 (1900 system).
	// m/d/yy → "4/30/24"
	got := numfmt.FormatValue(float64(45412), 14, "", false)
	want := "4/30/24"
	if got != want {
		t.Errorf("FormatValue(45412, 14) = %q, want %q", got, want)
	}
}

// TestFormatValueElapsed covers Cat C: elapsed/duration format ([h]:mm:ss).
func TestFormatValueElapsed(t *testing.T) {
	// Serial 0.270833... = 6.5 hours = 6:30:00
	serial := 6.5 / 24.0
	got := numfmt.FormatValue(serial, 46, "", false) // built-in 46 = [h]:mm:ss
	want := "6:30:00"
	if got != want {
		t.Errorf("FormatValue(elapsed) = %q, want %q", got, want)
	}
}

// TestFormatValueDateLong covers Cat E: long-form datetime (DDDD DD/MM/YYYY).
func TestFormatValueDateLong(t *testing.T) {
	// Excel serial 45285 = 2023-12-25 (Monday).
	// "DDDD DD/MM/YYYY" → "Monday 25/12/2023"
	got := numfmt.FormatValue(float64(45285), 164, "DDDD DD/MM/YYYY", false)
	want := "Monday 25/12/2023"
	if got != want {
		t.Errorf("FormatValue(45285, DDDD DD/MM/YYYY) = %q, want %q", got, want)
	}
}

// TestFormatValueDateShort covers Cat G: short date formats.
func TestFormatValueDateShort(t *testing.T) {
	tests := []struct {
		name   string
		serial float64
		fmtStr string
		want   string
	}{
		// 45119 = 2023-07-12
		{"DD-MMM", 45119, "DD-MMM", "12-Jul"},
		// 45367 = 2024-03-16
		{"MM-DD-YY", 45367, "MM-DD-YY", "03-16-24"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got := numfmt.FormatValue(tc.serial, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.serial, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueNegativeSections covers multi-section formats (positive/negative).
func TestFormatValueNegativeSections(t *testing.T) {
	tests := []struct {
		name   string
		v      float64
		fmtStr string
		want   string
	}{
		{"positive section", 42.5, "0.00;(0.00)", "42.50"},
		{"negative section parentheses", -42.5, "0.00;(0.00)", "(42.50)"},
		{"zero falls to positive section (2-section)", 0.0, "0.00;(0.00)", "0.00"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got := numfmt.FormatValue(tc.v, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueFallback verifies that FormatValue never silently drops a
// numeric value when the format string produces no renderable output.
func TestFormatValueFallback(t *testing.T) {
	tests := []struct {
		name    string
		v       float64
		fmtID   int
		fmtStr  string
		wantNot string // result must NOT be this
		desc    string
	}{
		{
			// A custom format string consisting only of a colour token —
			// renderNumber produces no output from its token walk and must
			// fall back to renderGeneral.
			name:    "number: colour-only format",
			v:       42.5,
			fmtID:   164,
			fmtStr:  "[Red]",
			wantNot: "",
			desc:    "colour-only format must not return empty string",
		},
		{
			// A date-typed format string (ID 164 with 'D' present so
			// isDateFormat returns true) that after isDateFormat detection
			// produces only colour tokens — renderDateTime produces no
			// calendar output and must fall back to renderGeneral.
			name:    "date: colour-only format",
			v:       45285, // 2023-12-25
			fmtID:   164,
			fmtStr:  "[Red]D",
			wantNot: "",
			desc:    "date colour-only format must not return empty string",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.v, tc.fmtID, tc.fmtStr, false)
			if got == tc.wantNot {
				t.Errorf("FormatValue(%v, %q): %s; got %q", tc.v, tc.fmtStr, tc.desc, got)
			}
		})
	}
}

// ── WorkbookFormatCell end-to-end ─────────────────────────────────────────────

// TestWorkbookFormatCell builds an in-memory .xlsb with styles.bin and verifies
// that wb.FormatCell renders cells correctly for each XF style entry.
func TestWorkbookFormatCell(t *testing.T) {
	data := buildXLSBWithStylesBin(t)
	wb, err := workbook.OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer wb.Close()

	// The styles.bin built by buildStylesBin has:
	//   xf[0]: numFmtId=14  (built-in date m/d/yy)
	//   xf[1]: numFmtId=164 (custom "yyyy-mm-dd")
	//   xf[2]: numFmtId=0   (General)

	// Excel serial 45285 = 2023-12-25.
	serial := float64(45285)

	tests := []struct {
		name     string
		v        any
		styleIdx int
		want     string
	}{
		{"date built-in m/d/yy", serial, 0, "12/25/23"},
		{"date custom yyyy-mm-dd", serial, 1, "2023-12-25"},
		{"general float", serial, 2, "45285"},
		{"nil value", nil, 0, ""},
		{"string passthrough", "hello", 2, "hello"},
		{"out-of-range style falls back", serial, 99, "45285"},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got := wb.FormatCell(tc.v, tc.styleIdx)
			if got != tc.want {
				t.Errorf("FormatCell(%v, %d) = %q, want %q", tc.v, tc.styleIdx, got, tc.want)
			}
		})
	}
}

// TestWorksheetFormatCell verifies that ws.FormatCell delegates to the injected
// formatFn (i.e. wb.FormatCell) and produces the same result.
func TestWorksheetFormatCell(t *testing.T) {
	data := buildXLSBWithStylesBin(t)
	wb, err := workbook.OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer wb.Close()

	sheet, err := wb.Sheet(1)
	if err != nil {
		t.Fatalf("Sheet(1): %v", err)
	}

	// Excel serial 45285 = 2023-12-25; xf[1] = "yyyy-mm-dd".
	cell := worksheet.Cell{V: float64(45285), Style: 1}
	got := sheet.FormatCell(cell)
	want := "2023-12-25"
	if got != want {
		t.Errorf("ws.FormatCell = %q, want %q", got, want)
	}
}

// TestWorkbookStylesTablePopulated verifies that wb.Styles is populated after
// parsing a workbook that includes xl/styles.bin.
func TestWorkbookStylesTablePopulated(t *testing.T) {
	data := buildXLSBWithStylesBin(t)
	wb, err := workbook.OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer wb.Close()

	if len(wb.Styles) != 3 {
		t.Fatalf("len(wb.Styles) = %d, want 3", len(wb.Styles))
	}
	// xf[0]: numFmtId=14 (built-in date)
	if !wb.Styles.IsDate(0) {
		t.Error("Styles.IsDate(0) = false, want true (numFmtId=14)")
	}
	// xf[1]: numFmtId=164 (custom yyyy-mm-dd)
	if !wb.Styles.IsDate(1) {
		t.Error("Styles.IsDate(1) = false, want true (custom yyyy-mm-dd)")
	}
	// xf[2]: numFmtId=0 (General) — not a date
	if wb.Styles.IsDate(2) {
		t.Error("Styles.IsDate(2) = true, want false (numFmtId=0 General)")
	}
	// out-of-range
	if wb.Styles.IsDate(99) {
		t.Error("Styles.IsDate(99) = true, want false (out of range)")
	}
}

// ── Date1904 ──────────────────────────────────────────────────────────────────

// buildMinimalXLSBWithDate1904 builds a minimal .xlsb whose workbook.bin
// optionally contains a BrtWbProp record (id=0x0199) with the f1904DateSystem
// bit (bit 3 = 0x08) set or cleared.
func buildMinimalXLSBWithDate1904(t *testing.T, date1904 bool) []byte {
	t.Helper()

	writeID := func(buf *bytes.Buffer, id int) {
		if id < 0x80 {
			buf.WriteByte(byte(id))
		} else {
			buf.WriteByte(byte(id & 0xFF))
			buf.WriteByte(byte(id >> 8))
		}
	}
	writeLen := func(buf *bytes.Buffer, n int) {
		for {
			b := n & 0x7F
			n >>= 7
			if n > 0 {
				buf.WriteByte(byte(b) | 0x80)
			} else {
				buf.WriteByte(byte(b))
				break
			}
		}
	}
	writeRec := func(buf *bytes.Buffer, id int, payload []byte) {
		writeID(buf, id)
		writeLen(buf, len(payload))
		buf.Write(payload)
	}
	encStr := func(s string) []byte {
		runes := []rune(s)
		var sb bytes.Buffer
		_ = binary.Write(&sb, binary.LittleEndian, uint32(len(runes)))
		for _, r := range runes {
			_ = binary.Write(&sb, binary.LittleEndian, uint16(r))
		}
		return sb.Bytes()
	}
	le32 := func(v uint32) []byte {
		b := make([]byte, 4)
		binary.LittleEndian.PutUint32(b, v)
		return b
	}

	var wb bytes.Buffer
	writeRec(&wb, 0x0183, nil) // WORKBOOK start

	// BrtWbProp (0x0199): flags uint32; bit 3 = f1904DateSystem
	var flags uint32
	if date1904 {
		flags = 0x08
	}
	writeRec(&wb, 0x0199, le32(flags))

	writeRec(&wb, 0x018F, nil) // SHEETS start
	var sheetRec bytes.Buffer
	sheetRec.Write(le32(0))
	sheetRec.Write(le32(1))
	sheetRec.Write(encStr("rId1"))
	sheetRec.Write(encStr("Sheet1"))
	writeRec(&wb, 0x019C, sheetRec.Bytes())
	writeRec(&wb, 0x0190, nil) // SHEETS end
	writeRec(&wb, 0x0184, nil) // WORKBOOK end

	// minimal worksheet
	var ws bytes.Buffer
	writeRec(&ws, 0x0181, nil)
	writeRec(&ws, 0x0191, nil)
	writeRec(&ws, 0x0192, nil)
	writeRec(&ws, 0x0182, nil)

	var zipBuf bytes.Buffer
	zw := zip.NewWriter(&zipBuf)
	addFile := func(name string, data []byte) {
		t.Helper()
		f, err := zw.Create(name)
		if err != nil {
			t.Fatalf("zip create %s: %v", name, err)
		}
		if _, err := f.Write(data); err != nil {
			t.Fatalf("zip write %s: %v", name, err)
		}
	}
	relsXML := `<?xml version="1.0" encoding="UTF-8"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.bin"/>` +
		`</Relationships>`
	addFile("xl/_rels/workbook.bin.rels", []byte(relsXML))
	addFile("xl/workbook.bin", wb.Bytes())
	addFile("xl/worksheets/sheet1.bin", ws.Bytes())
	if err := zw.Close(); err != nil {
		t.Fatalf("zip close: %v", err)
	}
	return zipBuf.Bytes()
}

func TestWorkbookDate1904(t *testing.T) {
	tests := []struct {
		name     string
		date1904 bool
		want     bool
	}{
		{"flag set (1904 system)", true, true},
		{"flag clear (1900 system)", false, false},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			data := buildMinimalXLSBWithDate1904(t, tc.date1904)
			wb, err := workbook.OpenReader(bytes.NewReader(data), int64(len(data)))
			if err != nil {
				t.Fatalf("OpenReader: %v", err)
			}
			defer wb.Close()
			if wb.Date1904 != tc.want {
				t.Errorf("Date1904 = %v, want %v", wb.Date1904, tc.want)
			}
		})
	}
}

// ── ErrorCellStrings ──────────────────────────────────────────────────────────

// buildErrCellXLSB constructs a minimal .xlsb in memory with one sheet.
// The sheet contains two error cells per test case: one BoolErr (0x0003) and
// one FormulaBoolErr (0x000B), each carrying the given error byte code.
// The cells are placed at column 0 (BoolErr) and column 1 (FormulaBoolErr) in row 0.
func buildErrCellXLSB(t *testing.T, errCode byte) []byte {
	t.Helper()

	writeID := func(buf *bytes.Buffer, id int) {
		if id < 0x80 {
			buf.WriteByte(byte(id))
		} else {
			buf.WriteByte(byte(id & 0xFF))
			buf.WriteByte(byte(id >> 8))
		}
	}
	writeLen := func(buf *bytes.Buffer, n int) {
		for {
			b := n & 0x7F
			n >>= 7
			if n > 0 {
				buf.WriteByte(byte(b) | 0x80)
			} else {
				buf.WriteByte(byte(b))
				break
			}
		}
	}
	writeRec := func(buf *bytes.Buffer, id int, payload []byte) {
		writeID(buf, id)
		writeLen(buf, len(payload))
		buf.Write(payload)
	}
	encStr := func(s string) []byte {
		runes := []rune(s)
		var sb bytes.Buffer
		_ = binary.Write(&sb, binary.LittleEndian, uint32(len(runes)))
		for _, r := range runes {
			_ = binary.Write(&sb, binary.LittleEndian, uint16(r))
		}
		return sb.Bytes()
	}
	le32 := func(v uint32) []byte {
		b := make([]byte, 4)
		binary.LittleEndian.PutUint32(b, v)
		return b
	}

	// ── xl/workbook.bin ───────────────────────────────────────────────────────

	var wb bytes.Buffer
	writeRec(&wb, 0x0183, nil) // WORKBOOK start
	writeRec(&wb, 0x018F, nil) // SHEETS start

	var sheetRec bytes.Buffer
	sheetRec.Write(le32(0))            // state flags (visible)
	sheetRec.Write(le32(1))            // sheetId
	sheetRec.Write(encStr("rId1"))     // relId
	sheetRec.Write(encStr("ErrSheet")) // name
	writeRec(&wb, 0x019C, sheetRec.Bytes())

	writeRec(&wb, 0x0190, nil) // SHEETS end
	writeRec(&wb, 0x0184, nil) // WORKBOOK end

	// ── xl/worksheets/sheet1.bin ──────────────────────────────────────────────

	var ws bytes.Buffer
	writeRec(&ws, 0x0181, nil) // WORKSHEET start

	// DIMENSION: r1=0, r2=0, c1=0, c2=1
	var dimPay bytes.Buffer
	dimPay.Write(le32(0))
	dimPay.Write(le32(0))
	dimPay.Write(le32(0))
	dimPay.Write(le32(1))
	writeRec(&ws, 0x0194, dimPay.Bytes())

	writeRec(&ws, 0x0191, nil) // SHEETDATA start

	// ROW r=0
	writeRec(&ws, 0x0000, le32(0))

	// BoolErr cell at col 0: col(4) + style(4) + errCode(1)
	var boolErrPay bytes.Buffer
	boolErrPay.Write(le32(0)) // col 0
	boolErrPay.Write(le32(0)) // style 0
	boolErrPay.WriteByte(errCode)
	writeRec(&ws, 0x0003, boolErrPay.Bytes()) // BoolErr = 0x0003

	// FormulaBoolErr cell at col 1: col(4) + style(4) + errCode(1)
	var formulaErrPay bytes.Buffer
	formulaErrPay.Write(le32(1)) // col 1
	formulaErrPay.Write(le32(0)) // style 0
	formulaErrPay.WriteByte(errCode)
	writeRec(&ws, 0x000B, formulaErrPay.Bytes()) // FormulaBoolErr = 0x000B

	writeRec(&ws, 0x0192, nil) // SHEETDATA end
	writeRec(&ws, 0x0182, nil) // WORKSHEET end

	// ── assemble ZIP ─────────────────────────────────────────────────────────

	var zipBuf bytes.Buffer
	zw := zip.NewWriter(&zipBuf)

	addFile := func(name string, data []byte) {
		t.Helper()
		f, err := zw.Create(name)
		if err != nil {
			t.Fatalf("zip create %s: %v", name, err)
		}
		if _, err := f.Write(data); err != nil {
			t.Fatalf("zip write %s: %v", name, err)
		}
	}

	relsXML := `<?xml version="1.0" encoding="UTF-8"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.bin"/>` +
		`</Relationships>`
	addFile("xl/_rels/workbook.bin.rels", []byte(relsXML))
	addFile("xl/workbook.bin", wb.Bytes())
	addFile("xl/worksheets/sheet1.bin", ws.Bytes())

	if err := zw.Close(); err != nil {
		t.Fatalf("zip close: %v", err)
	}
	return zipBuf.Bytes()
}

func TestErrorCellStrings(t *testing.T) {
	tests := []struct {
		name    string
		errCode byte
		want    string
	}{
		{"null", 0x00, "#NULL!"},
		{"div0", 0x07, "#DIV/0!"},
		{"value", 0x0F, "#VALUE!"},
		{"ref", 0x17, "#REF!"},
		{"name", 0x1D, "#NAME?"},
		{"num", 0x24, "#NUM!"},
		{"na", 0x2A, "#N/A"},
		{"getting_data", 0x2B, "#GETTING_DATA"},
		{"unknown_fallback", 0xFF, "0xff"},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			data := buildErrCellXLSB(t, tc.errCode)
			wb, err := workbook.OpenReader(bytes.NewReader(data), int64(len(data)))
			if err != nil {
				t.Fatalf("OpenReader: %v", err)
			}
			defer wb.Close()

			ws, err := wb.Sheet(1)
			if err != nil {
				t.Fatalf("Sheet(1): %v", err)
			}

			var row []worksheet.Cell
			for r := range ws.Rows(true) {
				row = r
				break
			}
			if row == nil {
				t.Fatal("no rows returned")
			}

			// col 0: BoolErr
			if got, ok := row[0].V.(string); !ok {
				t.Errorf("BoolErr V type = %T, want string", row[0].V)
			} else if got != tc.want {
				t.Errorf("BoolErr V = %q, want %q", got, tc.want)
			}

			// col 1: FormulaBoolErr
			if got, ok := row[1].V.(string); !ok {
				t.Errorf("FormulaBoolErr V type = %T, want string", row[1].V)
			} else if got != tc.want {
				t.Errorf("FormulaBoolErr V = %q, want %q", got, tc.want)
			}
		})
	}
}

// ── SheetVisibility ───────────────────────────────────────────────────────────

// buildVisibilityXLSB constructs a minimal .xlsb with three sheets whose
// hsState flags are 0 (visible), 1 (hidden), and 2 (veryHidden) respectively.
func buildVisibilityXLSB(t *testing.T) []byte {
	t.Helper()

	writeID := func(buf *bytes.Buffer, id int) {
		if id < 0x80 {
			buf.WriteByte(byte(id))
		} else {
			buf.WriteByte(byte(id & 0xFF))
			buf.WriteByte(byte(id >> 8))
		}
	}
	writeLen := func(buf *bytes.Buffer, n int) {
		for {
			b := n & 0x7F
			n >>= 7
			if n > 0 {
				buf.WriteByte(byte(b) | 0x80)
			} else {
				buf.WriteByte(byte(b))
				break
			}
		}
	}
	writeRec := func(buf *bytes.Buffer, id int, payload []byte) {
		writeID(buf, id)
		writeLen(buf, len(payload))
		buf.Write(payload)
	}
	encStr := func(s string) []byte {
		runes := []rune(s)
		var sb bytes.Buffer
		_ = binary.Write(&sb, binary.LittleEndian, uint32(len(runes)))
		for _, r := range runes {
			_ = binary.Write(&sb, binary.LittleEndian, uint16(r))
		}
		return sb.Bytes()
	}
	le32 := func(v uint32) []byte {
		b := make([]byte, 4)
		binary.LittleEndian.PutUint32(b, v)
		return b
	}

	// ── xl/workbook.bin ───────────────────────────────────────────────────────

	type sheetDef struct {
		name    string
		relID   string
		sheetID uint32
		hsState uint32 // low 2 bits only
	}
	sheets := []sheetDef{
		{"Sheet1", "rId1", 1, 0}, // visible
		{"Sheet2", "rId2", 2, 1}, // hidden
		{"Sheet3", "rId3", 3, 2}, // veryHidden
	}

	var wb bytes.Buffer
	writeRec(&wb, 0x0183, nil) // WORKBOOK start
	writeRec(&wb, 0x018F, nil) // SHEETS start

	for _, s := range sheets {
		var sheetRec bytes.Buffer
		sheetRec.Write(le32(s.hsState)) // flags: low 2 bits = hsState
		sheetRec.Write(le32(s.sheetID)) // sheetId
		sheetRec.Write(encStr(s.relID)) // relId
		sheetRec.Write(encStr(s.name))  // name
		writeRec(&wb, 0x019C, sheetRec.Bytes())
	}

	writeRec(&wb, 0x0190, nil) // SHEETS end
	writeRec(&wb, 0x0184, nil) // WORKBOOK end

	// ── minimal worksheet binary (reused for all three sheets) ────────────────

	var ws bytes.Buffer
	writeRec(&ws, 0x0181, nil) // WORKSHEET start
	writeRec(&ws, 0x0191, nil) // SHEETDATA start
	writeRec(&ws, 0x0192, nil) // SHEETDATA end
	writeRec(&ws, 0x0182, nil) // WORKSHEET end
	wsBytes := ws.Bytes()

	// ── assemble ZIP ─────────────────────────────────────────────────────────

	var zipBuf bytes.Buffer
	zw := zip.NewWriter(&zipBuf)

	addFile := func(name string, data []byte) {
		t.Helper()
		f, err := zw.Create(name)
		if err != nil {
			t.Fatalf("zip create %s: %v", name, err)
		}
		if _, err := f.Write(data); err != nil {
			t.Fatalf("zip write %s: %v", name, err)
		}
	}

	relsXML := `<?xml version="1.0" encoding="UTF-8"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.bin"/>` +
		`<Relationship Id="rId2" Type="worksheet" Target="worksheets/sheet2.bin"/>` +
		`<Relationship Id="rId3" Type="worksheet" Target="worksheets/sheet3.bin"/>` +
		`</Relationships>`
	addFile("xl/_rels/workbook.bin.rels", []byte(relsXML))
	addFile("xl/workbook.bin", wb.Bytes())
	addFile("xl/worksheets/sheet1.bin", wsBytes)
	addFile("xl/worksheets/sheet2.bin", wsBytes)
	addFile("xl/worksheets/sheet3.bin", wsBytes)

	if err := zw.Close(); err != nil {
		t.Fatalf("zip close: %v", err)
	}
	return zipBuf.Bytes()
}

func TestSheetVisibility(t *testing.T) {
	data := buildVisibilityXLSB(t)
	wb, err := workbook.OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer wb.Close()

	t.Run("SheetVisible known sheets", func(t *testing.T) {
		tests := []struct {
			name string
			want bool
		}{
			{"Sheet1", true},
			{"Sheet2", false},
			{"Sheet3", false},
		}
		for _, tc := range tests {
			if got := wb.SheetVisible(tc.name); got != tc.want {
				t.Errorf("SheetVisible(%q) = %v, want %v", tc.name, got, tc.want)
			}
		}
	})

	t.Run("SheetVisible unknown name returns false", func(t *testing.T) {
		if wb.SheetVisible("nonexistent") {
			t.Error("SheetVisible(\"nonexistent\") = true, want false")
		}
	})

	t.Run("SheetVisible case-insensitive", func(t *testing.T) {
		if !wb.SheetVisible("sheet1") {
			t.Error("SheetVisible(\"sheet1\") = false, want true")
		}
		if wb.SheetVisible("SHEET2") {
			t.Error("SheetVisible(\"SHEET2\") = true, want false")
		}
	})

	t.Run("SheetVisibility raw levels", func(t *testing.T) {
		tests := []struct {
			name string
			want int
		}{
			{"Sheet1", workbook.SheetVisible},
			{"Sheet2", workbook.SheetHidden},
			{"Sheet3", workbook.SheetVeryHidden},
			{"nonexistent", -1},
		}
		for _, tc := range tests {
			if got := wb.SheetVisibility(tc.name); got != tc.want {
				t.Errorf("SheetVisibility(%q) = %d, want %d", tc.name, got, tc.want)
			}
		}
	})
}

// ── ConvertDateEx ─────────────────────────────────────────────────────────────

func TestConvertDateEx(t *testing.T) {
	tests := []struct {
		name     string
		input    float64
		date1904 bool
		want     time.Time
		wantErr  bool
	}{
		// ── 1900 system (date1904=false): delegates to ConvertDate ────────────
		{
			name:     "1900: serial 0 gives 1900-01-01",
			input:    0,
			date1904: false,
			want:     time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC),
		},
		{
			name:     "1900: pyxlsb example 41235.45578",
			input:    41235.45578,
			date1904: false,
			want:     time.Date(2012, 11, 22, 10, 56, 19, 0, time.UTC),
		},
		// ── 1904 system (date1904=true) ───────────────────────────────────────
		{
			name:     "1904: serial 0 gives 1904-01-01",
			input:    0,
			date1904: true,
			want:     time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC),
		},
		{
			name:     "1904: serial 0 with time component",
			input:    0.5,
			date1904: true,
			want:     time.Date(1904, 1, 1, 12, 0, 0, 0, time.UTC),
		},
		{
			name:     "1904: serial 1 gives 1904-01-02",
			input:    1,
			date1904: true,
			want:     time.Date(1904, 1, 2, 0, 0, 0, 0, time.UTC),
		},
		{
			name:     "1904: serial 365 gives 1904-12-31",
			input:    365,
			date1904: true,
			want:     time.Date(1904, 12, 31, 0, 0, 0, 0, time.UTC),
		},
		{
			// 1904-system serial 39813 = 1904-01-01 + 39813 days = 2013-01-01.
			// Cross-check: 1900-system serial 41275 also = 2013-01-01 (39813 + 1462 = 41275).
			name:     "1904: serial 39813 gives 2013-01-01",
			input:    39813,
			date1904: true,
			want:     time.Date(2013, 1, 1, 0, 0, 0, 0, time.UTC),
		},
		// ── error cases ───────────────────────────────────────────────────────
		{
			name:     "1904: NaN returns error",
			input:    math.NaN(),
			date1904: true,
			wantErr:  true,
		},
		{
			name:     "1904: +Inf returns error",
			input:    math.Inf(1),
			date1904: true,
			wantErr:  true,
		},
		{
			name:     "1904: negative serial returns error",
			input:    -1,
			date1904: true,
			wantErr:  true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got, err := xlsb.ConvertDateEx(tc.input, tc.date1904)
			if tc.wantErr {
				if err == nil {
					t.Fatalf("expected error, got nil (result=%v)", got)
				}
				return
			}
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if !got.Equal(tc.want) {
				t.Errorf("ConvertDateEx(%v, %v) = %v, want %v", tc.input, tc.date1904, got, tc.want)
			}
		})
	}
}
