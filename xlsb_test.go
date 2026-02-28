package xlsb_test

// Unit tests for the go-xlsb library.
//
// The tests are intentionally self-contained: they build all binary fixtures
// in memory so no external .xlsb file is required.

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"fmt"
	"math"
	"testing"
	"time"

	"github.com/TsubasaBE/go-xlsb"
	"github.com/TsubasaBE/go-xlsb/numfmt"
	"github.com/TsubasaBE/go-xlsb/record"
	"github.com/TsubasaBE/go-xlsb/stringtable"
	"github.com/TsubasaBE/go-xlsb/styles"
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
		{name: "built-in 22 (m/d/yy hh:mm)", id: 22, want: true},
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
	// mm-dd-yy (excelize ground truth for built-in ID 14) → "04-30-24"
	got := numfmt.FormatValue(float64(45412), 14, "", false)
	want := "04-30-24"
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
		{"date built-in m/d/yy -> mm-dd-yy", serial, 0, "12-25-23"},
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

// ── Batch 1 tests ─────────────────────────────────────────────────────────────

// TestFormatValueTextSection verifies that the fourth section of a multi-section
// format string is applied to string cell values, with "@" substituted by the
// cell value and surrounding literals emitted.
func TestFormatValueTextSection(t *testing.T) {
	tests := []struct {
		name   string
		v      string
		fmtStr string
		want   string
	}{
		{
			// Four-section format: positive;negative;zero;text.
			// The text section prefixes the cell value with a literal.
			// Note: nfp v0.0.1 has a known quirk with bracket literals inside
			// quoted strings; use angle-bracket literals instead.
			name:   "@ in fourth section wraps value",
			v:      "hello",
			fmtStr: `0;-0;0;">> "@" <<"`,
			want:   ">> hello <<",
		},
		{
			// Plain "@" as the entire format — shortcut path, return as-is.
			name:   "bare @ format returns value unchanged",
			v:      "world",
			fmtStr: "@",
			want:   "world",
		},
		{
			// "General" format — string passthrough.
			name:   "General format returns value unchanged",
			v:      "test",
			fmtStr: "General",
			want:   "test",
		},
		{
			// Three-section format — no text section; string returned as-is.
			name:   "three sections: no text section, passthrough",
			v:      "abc",
			fmtStr: `0;-0;0`,
			want:   "abc",
		},
		{
			// Text section with literal prefix only (no @ token).
			// nfp emits a literal for the quoted prefix.  With no @ token, the
			// section still has content, so we emit the literals.
			name:   "text section literal prefix no placeholder",
			v:      "xyz",
			fmtStr: `0;-0;0;"prefix"`,
			want:   "prefix",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.v, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%q, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueAlignmentToken verifies that _x alignment tokens are rendered
// as a single space in numeric, datetime, and text sections.
func TestFormatValueAlignmentToken(t *testing.T) {
	tests := []struct {
		name   string
		v      any
		fmtID  int
		fmtStr string
		want   string
	}{
		{
			// Accounting-style format uses _) at the end to align with
			// negative-parenthesis formats.  The _) should produce one space.
			name:   "numeric: trailing _) emits space",
			v:      float64(1234),
			fmtID:  164,
			fmtStr: "#,##0_)",
			want:   "1,234 ",
		},
		{
			// Built-in ID 37: (#,##0_);(#,##0).
			// Section 0 is "(#,##0_)" — the "(" is a literal, so positive values
			// render as "(500 " (open paren + number + alignment space).
			name:   "built-in 37 positive: literal paren plus trailing space",
			v:      float64(500),
			fmtID:  37,
			fmtStr: "",
			want:   "(500 ",
		},
		{
			// Built-in ID 37 negative value: rendered with parentheses, no trailing space.
			name:   "built-in 37 negative parentheses",
			v:      float64(-500),
			fmtID:  37,
			fmtStr: "",
			want:   "(500)",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.v, tc.fmtID, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, id=%d, %q) = %q, want %q", tc.v, tc.fmtID, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueRepeatsChar verifies that *x repeat-char tokens are rendered
// as a single instance of the character in numeric, datetime, and text sections.
func TestFormatValueRepeatsChar(t *testing.T) {
	tests := []struct {
		name   string
		v      any
		fmtID  int
		fmtStr string
		want   string
	}{
		{
			// *- before the number: fills column with '-', but in plain-text we
			// emit exactly one '-'.
			name:   "numeric: *- prefix emits one dash",
			v:      float64(42),
			fmtID:  164,
			fmtStr: "*-0",
			want:   "-42",
		},
		{
			// "0 *-": space literal then *- repeat-char.  nfp parses *- as
			// RepeatsChar with TValue="-".  Plain-text rendering emits one "-".
			name:   "numeric: trailing *- emits one dash",
			v:      float64(7),
			fmtID:  164,
			fmtStr: `0 *-`,
			want:   "7 -",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.v, tc.fmtID, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueCurrencySymbol verifies that [$symbol-locale] currency language
// tokens emit only the symbol portion (not the raw bracket expression).
func TestFormatValueCurrencySymbol(t *testing.T) {
	tests := []struct {
		name   string
		v      float64
		fmtStr string
		want   string
	}{
		{
			// Euro symbol with German locale: [$€-407]#,##0.00
			name:   "euro symbol with locale",
			v:      1234.5,
			fmtStr: `[$€-407]#,##0.00`,
			want:   "€1,234.50",
		},
		{
			// Dollar sign with US locale: [$USD-409] #,##0
			name:   "USD symbol with locale",
			v:      1000,
			fmtStr: `[$USD-409] #,##0`,
			want:   "USD 1,000",
		},
		{
			// Plain $ (not a currency-language token, just a literal) — verify
			// existing literal handling still works alongside.
			name:   "plain dollar literal",
			v:      50,
			fmtStr: `"$"0`,
			want:   "$50",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.v, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestBuiltInNumFmtIDs verifies that all expected built-in format IDs are
// present in styles.BuiltInNumFmt and hold the correct canonical strings as
// defined by ECMA-376 and confirmed against excelize.
func TestBuiltInNumFmtIDs(t *testing.T) {
	tests := []struct {
		id   int
		want string
	}{
		// IDs 5–8: currency variants (added in Batch 1)
		{5, `($#,##0_);($#,##0)`},
		{6, `($#,##0_);[Red]($#,##0)`},
		{7, `($#,##0.00_);($#,##0.00)`},
		{8, `($#,##0.00_);[Red]($#,##0.00)`},
		// IDs 41–44: accounting formats (added in Batch 1)
		{41, `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)`},
		{42, `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)`},
		{43, `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`},
		{44, `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)`},
		// ID 14: excelize canonical value is "mm-dd-yy" (lowercase), not "MM-DD-YY"
		{14, "mm-dd-yy"},
		// IDs fixed in Batch 1
		{22, "m/d/yy hh:mm"},
		{37, `(#,##0_);(#,##0)`},
		{38, `(#,##0_);[Red](#,##0)`},
		{39, `(#,##0.00_);(#,##0.00)`},
		{40, `(#,##0.00_);[Red](#,##0.00)`},
		{47, "mm:ss.0"},
	}
	for _, tc := range tests {
		t.Run(fmt.Sprintf("ID_%d", tc.id), func(t *testing.T) {
			t.Helper()
			got, ok := styles.BuiltInNumFmt[tc.id]
			if !ok {
				t.Fatalf("BuiltInNumFmt[%d] not found", tc.id)
			}
			if got != tc.want {
				t.Errorf("BuiltInNumFmt[%d] = %q, want %q", tc.id, got, tc.want)
			}
		})
	}
}

// ── Batch 2 tests ─────────────────────────────────────────────────────────────

// TestFormatValueScientific verifies E+/E- scientific and engineering notation.
func TestFormatValueScientific(t *testing.T) {
	tests := []struct {
		name   string
		v      float64
		fmtStr string
		want   string
	}{
		// Standard scientific: one integer digit.
		{
			name:   "positive standard scientific",
			v:      12345.6789,
			fmtStr: "0.00E+00",
			want:   "1.23E+04",
		},
		{
			name:   "small positive standard scientific",
			v:      0.000123,
			fmtStr: "0.00E+00",
			want:   "1.23E-04",
		},
		{
			name:   "negative standard scientific",
			v:      -9876.5,
			fmtStr: "0.00E+00",
			want:   "-9.88E+03",
		},
		{
			name:   "zero standard scientific",
			v:      0,
			fmtStr: "0.00E+00",
			want:   "0.00E+00",
		},
		// Built-in ID 11 format string used directly.
		{
			name:   "built-in 11: 0.00E+00",
			v:      1234.567,
			fmtStr: "0.00E+00",
			want:   "1.23E+03",
		},
		// Engineering notation: 3 integer digits, exponent multiple of 3.
		{
			name:   "engineering ##0.0E+0 large",
			v:      12345678,
			fmtStr: "##0.0E+0",
			want:   "12.3E+6",
		},
		{
			name:   "engineering ##0.0E+0 medium",
			v:      1234,
			fmtStr: "##0.0E+0",
			want:   "1.2E+3",
		},
		{
			name:   "engineering ##0.0E+0 small",
			v:      0.000456,
			fmtStr: "##0.0E+0",
			want:   "456.0E-6",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.v, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueFraction verifies mixed-number fraction rendering.
func TestFormatValueFraction(t *testing.T) {
	tests := []struct {
		name   string
		v      float64
		fmtStr string
		want   string
	}{
		// Single-digit denominator (max 9).
		{
			name:   "# ?/?: 1.75 → 1 3/4",
			v:      1.75,
			fmtStr: "# ?/?",
			want:   "1 3/4",
		},
		{
			name:   "# ?/?: 1.333333 → 1 1/3",
			v:      1.333333,
			fmtStr: "# ?/?",
			want:   "1 1/3",
		},
		{
			name:   "# ?/?: 0.5 → space-padded integer then 1/2",
			v:      0.5,
			fmtStr: "# ?/?",
			want:   " 1/2",
		},
		{
			name:   "# ?/?: exact integer suppresses fraction",
			v:      2.0,
			fmtStr: "# ?/?",
			want:   "2    ",
		},
		// Double-digit denominator (max 99).
		{
			name:   "# ??/??: pi approx → 3 14/99",
			v:      3.141592653589793,
			fmtStr: "# ??/??",
			want:   "3 14/99",
		},
		{
			name:   "# ??/??: exact integer suppresses fraction",
			v:      3.0,
			fmtStr: "# ??/??",
			want:   "3      ",
		},
		{
			name:   "# ??/??: 0.1 → space-padded",
			v:      0.1,
			fmtStr: "# ??/??",
			want:   "  1/10",
		},
		// Negative value.
		{
			name:   "negative fraction (single section → minus prefix)",
			v:      -1.75,
			fmtStr: "# ?/?",
			want:   "-1 3/4",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.v, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueDigitalPlaceHolder verifies '?' space-padded digit placeholder.
func TestFormatValueDigitalPlaceHolder(t *testing.T) {
	tests := []struct {
		name   string
		v      float64
		fmtStr string
		want   string
	}{
		{
			// "0.??" — decimal with optional trailing digits, space-padded.
			// 1.5 → "1.5" (one digit after decimal; second ? is trimmed).
			name:   "0.?? trims trailing zero",
			v:      1.5,
			fmtStr: "0.??",
			want:   "1.5",
		},
		{
			// Integer value: decimal point suppressed when no fractional part.
			name:   "0.?? integer: no decimal",
			v:      1.0,
			fmtStr: "0.??",
			want:   "1",
		},
		{
			// "?.??" — nfp quirk: leading ? merged into DecimalPoint token.
			// Value 3.14: int part 3, frac "14".
			name:   "?.?? integer plus two decimal places",
			v:      3.14,
			fmtStr: "?.??",
			want:   "3.14",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.v, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// ── Batch 3 tests ─────────────────────────────────────────────────────────────

// TestFormatValueThousandsScaling verifies that trailing commas in a number
// format string scale the value by 1,000 per comma (Excel convention).
func TestFormatValueThousandsScaling(t *testing.T) {
	t.Helper()
	tests := []struct {
		name   string
		v      float64
		fmtStr string
		want   string
	}{
		{
			// Single trailing comma: divide by 1,000 → display in thousands.
			name:   "single trailing comma scale by 1000",
			v:      1234567,
			fmtStr: "#,##0,",
			want:   "1,235",
		},
		{
			// Two trailing commas: divide by 1,000,000 → display in millions.
			name:   "two trailing commas scale by 1000000",
			v:      1234567890,
			fmtStr: "#,##0,,",
			want:   "1,235",
		},
		{
			// Thousands separator AND one scaling comma.
			// 5,000,000 / 1,000 = 5,000 → displayed as "5,000M"
			name:   "thousands separator with one scaling comma",
			v:      5000000,
			fmtStr: `#,##0,"M"`,
			want:   "5,000M",
		},
		{
			// Plain format, no trailing comma: no scaling.
			name:   "no trailing comma no scaling",
			v:      1234567,
			fmtStr: "#,##0",
			want:   "1,234,567",
		},
		{
			// Two decimal places with one trailing comma.
			name:   "decimal places with trailing comma",
			v:      1500000,
			fmtStr: `#,##0.0,`,
			want:   "1,500.0",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.v, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueMilliseconds verifies that sub-second digits are rendered
// correctly when a format contains ".0", ".00", or ".000" after seconds.
// The built-in format ID 47 is "mm:ss.0".
func TestFormatValueMilliseconds(t *testing.T) {
	t.Helper()
	// Serial 0.5 = noon (12:00:00.000).
	// Serial for 00:01:02.750 = (1*60 + 2 + 0.75) / 86400
	serial62750ms := (float64(62) + 0.75) / 86400.0
	// Serial for 00:00:05.123
	serial5123ms := (float64(5) + 0.123) / 86400.0

	tests := []struct {
		name     string
		serial   float64
		fmtStr   string
		numFmtID int
		want     string
	}{
		{
			name:     "mm:ss.0 one decisecond digit",
			serial:   serial62750ms,
			fmtStr:   "mm:ss.0",
			numFmtID: 164,
			want:     "01:02.8", // 0.75 s → rounds to 0.8 with 1 digit
		},
		{
			name:     "mm:ss.000 three millisecond digits",
			serial:   serial5123ms,
			fmtStr:   "mm:ss.000",
			numFmtID: 164,
			// mm before ss is now correctly disambiguated as minutes (Batch 4 fix).
			want: "00:05.123",
		},
		{
			name:     "h:mm:ss.00 two centisecond digits",
			serial:   serial62750ms,
			fmtStr:   "h:mm:ss.00",
			numFmtID: 164,
			want:     "0:01:02.75",
		},
		{
			// Exact whole second: milliseconds should be .0
			name:     "whole second produces .0",
			serial:   float64(62) / 86400.0,
			fmtStr:   "mm:ss.0",
			numFmtID: 164,
			want:     "01:02.0",
		},
		{
			// Built-in ID 47 = "mm:ss.0"
			name:     "built-in ID 47",
			serial:   serial62750ms,
			fmtStr:   "",
			numFmtID: 47,
			want:     "01:02.8",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.serial, tc.numFmtID, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(serial=%v, id=%d, %q) = %q, want %q",
					tc.serial, tc.numFmtID, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestRenderGeneralLargeNumbers verifies that renderGeneral switches to E+
// scientific notation for values with magnitude >= 1e11.
func TestRenderGeneralLargeNumbers(t *testing.T) {
	t.Helper()
	tests := []struct {
		name string
		v    float64
		want string
	}{
		{
			name: "small integer no scientific",
			v:    12345,
			want: "12345",
		},
		{
			name: "just below 1e11 threshold",
			v:    99999999999,
			want: "99999999999",
		},
		{
			name: "exactly 1e11 uses scientific",
			v:    1e11,
			want: "1E+11",
		},
		{
			name: "1.23456e12 uses scientific",
			v:    1.23456e12,
			want: "1.23456E+12",
		},
		{
			name: "negative large uses scientific",
			v:    -1.5e13,
			want: "-1.5E+13",
		},
		{
			name: "fractional below threshold stays G10",
			v:    1234567.89,
			want: "1234567.89",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			// Use numFmtID=0, fmtStr="" to force renderGeneral path.
			got := numfmt.FormatValue(tc.v, 0, "", false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, General) = %q, want %q", tc.v, got, tc.want)
			}
		})
	}
}

// ── Batch 4 tests ─────────────────────────────────────────────────────────────

// TestFormatValueMmmmm verifies that the MMMMM token renders the single-letter
// month initial for each calendar month (J F M A M J J A S O N D).
func TestFormatValueMmmmm(t *testing.T) {
	t.Helper()
	// Serials for the 1st of each month in 2023 (1900 date system).
	// Jan 1 2023 = serial 44927; subsequent months follow.
	tests := []struct {
		name   string
		serial float64
		want   string
	}{
		{"January", 44927, "J"},
		{"February", 44958, "F"},
		{"March", 44986, "M"},
		{"April", 45017, "A"},
		{"May", 45047, "M"},
		{"June", 45078, "J"},
		{"July", 45108, "J"},
		{"August", 45139, "A"},
		{"September", 45170, "S"},
		{"October", 45200, "O"},
		{"November", 45231, "N"},
		{"December", 45261, "D"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.serial, 164, "MMMMM", false)
			if got != tc.want {
				t.Errorf("FormatValue(MMMMM, %v) = %q, want %q", tc.serial, got, tc.want)
			}
		})
	}
}

// TestFormatValueMBeforeSS verifies that "m" immediately before "ss" is
// treated as minutes (not month) even with no preceding hour token.
func TestFormatValueMBeforeSS(t *testing.T) {
	t.Helper()
	// Serial for exactly 1 minute 2 seconds past midnight on 2023-12-25.
	// Fractional day = 62 / 86400.
	serial := float64(45285) + 62.0/86400.0
	got := numfmt.FormatValue(serial, 164, "m:ss", false)
	want := "1:02"
	if got != want {
		t.Errorf("FormatValue(m:ss, %v) = %q, want %q", serial, got, want)
	}
}

// TestFormatValueFixedDenominator verifies that fixed-denominator fraction
// formats (e.g. "# ?/4", "# ?/8") produce the correct numerator.
func TestFormatValueFixedDenominator(t *testing.T) {
	t.Helper()
	tests := []struct {
		name   string
		v      float64
		fmtStr string
		want   string
	}{
		{
			name:   "# ?/4: 1.75 → 1 3/4",
			v:      1.75,
			fmtStr: "# ?/4",
			want:   "1 3/4",
		},
		{
			name:   "# ?/8: 1.125 → 1 1/8",
			v:      1.125,
			fmtStr: "# ?/8",
			want:   "1 1/8",
		},
		{
			name:   "# ?/4: exact integer 2 shows fixed denominator",
			v:      2.0,
			fmtStr: "# ?/4",
			want:   "2  /4",
		},
		{
			name:   "# ?/8: 0.25 → 2/8",
			v:      0.25,
			fmtStr: "# ?/8",
			want:   " 2/8",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.v, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueBuiltInIDs27to36 verifies that built-in number format IDs 27–36
// produce a date-like string (not the raw serial number) for a known date serial.
func TestFormatValueBuiltInIDs27to36(t *testing.T) {
	t.Helper()
	// Serial 45285 = 2023-12-25 in the 1900 date system.
	serial := float64(45285)
	rawSerial := "45285"
	for id := 27; id <= 36; id++ {
		id := id
		t.Run(fmt.Sprintf("ID%d", id), func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(serial, id, "", false)
			if got == rawSerial {
				t.Errorf("FormatValue(%v, ID=%d) = %q: want a formatted date, not the raw serial", serial, id, got)
			}
		})
	}
}

// TestFormatValueBuiltInIDs50to58 verifies that built-in number format IDs 50–58
// produce a date-like string (not the raw serial number) for a known date serial.
func TestFormatValueBuiltInIDs50to58(t *testing.T) {
	t.Helper()
	// Serial 45285 = 2023-12-25 in the 1900 date system.
	serial := float64(45285)
	rawSerial := "45285"
	for id := 50; id <= 58; id++ {
		id := id
		t.Run(fmt.Sprintf("ID%d", id), func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(serial, id, "", false)
			if got == rawSerial {
				t.Errorf("FormatValue(%v, ID=%d) = %q: want a formatted date, not the raw serial", serial, id, got)
			}
		})
	}
}

// TestFormatValueB1B2Tokens verifies that B1/B2 Buddhist Era / Gregorian
// calendar mode tokens are silently ignored (matching excelize behaviour)
// while the rest of the format string still renders correctly.
func TestFormatValueB1B2Tokens(t *testing.T) {
	t.Helper()
	// Serial 45285 = 2023-12-25. "B2" should be ignored; date renders normally.
	tests := []struct {
		name   string
		fmtStr string
		want   string
	}{
		{
			name:   "B2 prefix ignored, YYYY/MM/DD renders",
			fmtStr: "B2YYYY/MM/DD",
			want:   "2023/12/25",
		},
		{
			name:   "B1 prefix ignored, YYYY/MM/DD renders",
			fmtStr: "B1YYYY/MM/DD",
			want:   "2023/12/25",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(float64(45285), 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(45285, %q) = %q, want %q", tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueNegativeSectionNoWrapper verifies that a two-section format
// like "0;0" (negative section contains no sign wrapper) still prepends a
// minus sign for negative values.
func TestFormatValueNegativeSectionNoWrapper(t *testing.T) {
	t.Helper()
	tests := []struct {
		name   string
		v      float64
		fmtStr string
		want   string
	}{
		{
			name:   "0;0: negative value gets minus",
			v:      -5,
			fmtStr: "0;0",
			want:   "-5",
		},
		{
			name:   "0;0: positive value no minus",
			v:      5,
			fmtStr: "0;0",
			want:   "5",
		},
		{
			name:   "parenthesis section: no extra minus",
			v:      -5,
			fmtStr: `0;(0)`,
			want:   "(5)",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.v, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.v, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// ── @ format on numeric cells (ID 49) ─────────────────────────────────────────

// TestFormatValueAtFormatOnNumbers verifies that format ID 49 ("@") applied to
// a numeric cell renders the number in General style (not as a string literal).
// Excel's behaviour for @ on a number is to display the raw numeric value
// identically to General format.
func TestFormatValueAtFormatOnNumbers(t *testing.T) {
	tests := []struct {
		name string
		v    float64
		want string
	}{
		{"integer", 42, "42"},
		{"negative integer", -7, "-7"},
		{"zero", 0, "0"},
		{"decimal", 3.14, "3.14"},
		{"large integer", 1234567, "1234567"},
		{"scientific threshold", 1e11, "1E+11"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			// Use built-in ID 49 ("@") with no custom format string.
			got := numfmt.FormatValue(tc.v, 49, "", false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, 49, \"\") = %q, want %q", tc.v, got, tc.want)
			}
			// Also test explicit "@" as custom format string.
			got2 := numfmt.FormatValue(tc.v, 164, "@", false)
			if got2 != tc.want {
				t.Errorf("FormatValue(%v, 164, \"@\") = %q, want %q", tc.v, got2, tc.want)
			}
		})
	}
}

// ── renderElapsed large serial (int64 overflow guard) ─────────────────────────

// TestFormatValueElapsedLargeSerial verifies that elapsed-time formats ([h]:mm:ss)
// produce correct output for serials that would overflow a 32-bit int (> ~89 days).
func TestFormatValueElapsedLargeSerial(t *testing.T) {
	tests := []struct {
		name   string
		serial float64 // fractional days
		fmtStr string
		want   string
	}{
		// 10000 hours = 416.666... days; [h] must not overflow int32 (max ~2147 hours on int32).
		{
			name:   "10000 hours elapsed",
			serial: 10000.0 / 24.0,
			fmtStr: "[h]:mm:ss",
			want:   "10000:00:00",
		},
		// 50000 hours = 2083.333... days; exceeds int32 range (32767 hours).
		{
			name:   "50000 hours elapsed",
			serial: 50000.0 / 24.0,
			fmtStr: "[h]:mm:ss",
			want:   "50000:00:00",
		},
		// 100000 hours.
		{
			name:   "100000 hours elapsed",
			serial: 100000.0 / 24.0,
			fmtStr: "[h]:mm:ss",
			want:   "100000:00:00",
		},
		// 1.5 hours = 90 minutes elapsed, expressed as [mm]:ss.
		{
			name:   "90 minutes elapsed [mm]:ss",
			serial: 1.5 / 24.0,
			fmtStr: "[mm]:ss",
			want:   "90:00",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.serial, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.serial, tc.fmtStr, got, tc.want)
			}
		})
	}
}

// ── isDateFormat: s/S custom scan ─────────────────────────────────────────────

// TestIsDateFormatSecondsOnly verifies that custom formats containing only 's'/'S'
// (seconds) tokens are correctly detected as date/time formats.
func TestIsDateFormatSecondsOnly(t *testing.T) {
	tests := []struct {
		name      string
		id        int
		formatStr string
		want      bool
	}{
		{"custom ss", 164, "ss", true},
		{"custom SS", 165, "SS", true},
		{"custom s", 166, "s", true},
		{"custom S", 167, "S", true},
		// s inside double quotes must not trigger
		{"quoted s", 168, `"seconds"0`, false},
		// s inside brackets must not trigger
		{"bracketed s", 169, `[$-409]0.00`, false},
		// pure number format with no date chars — must remain false
		{"numeric 0.00", 170, "0.00", false},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := xlsb.IsDateFormat(tc.id, tc.formatStr)
			if got != tc.want {
				t.Errorf("IsDateFormat(%d, %q) = %v, want %v", tc.id, tc.formatStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueSecondsOnlyFormat verifies that a pure-seconds custom format
// actually renders the time component (not a raw number) when applied to a serial.
func TestFormatValueSecondsOnlyFormat(t *testing.T) {
	// Serial 0.5 = noon = 43200 seconds into the day.
	// Format "ss" should render the seconds component: "00" (noon is exactly on a
	// minute boundary, so seconds = 0).
	// Serial 0.5 + 30/86400 = noon + 30 seconds.
	serial := 0.5 + 30.0/86400.0
	got := numfmt.FormatValue(serial, 164, "ss", false)
	want := "30"
	if got != want {
		t.Errorf("FormatValue(%v, \"ss\") = %q, want %q", serial, got, want)
	}
}

// ── Chinese AM/PM (上午/下午) ──────────────────────────────────────────────────

// TestBuiltInNumFmtID20and21 verifies that built-in format IDs 20 and 21 produce
// zero-padded hours (hh), matching Excel and excelize behaviour.
func TestBuiltInNumFmtID20and21(t *testing.T) {
	t.Helper()
	// Serial for 09:05:03 on any day: 9 hours + 5 min + 3 sec = 32703 seconds.
	// fractional day = 32703 / 86400
	serial := float64(32703) / 86400.0

	tests := []struct {
		name     string
		numFmtID int
		want     string
	}{
		// ID 20 = "hh:mm" — zero-padded hours and minutes.
		{
			name:     "ID 20 hh:mm zero-padded",
			numFmtID: 20,
			want:     "09:05",
		},
		// ID 21 = "hh:mm:ss" — zero-padded hours, minutes, and seconds.
		{
			name:     "ID 21 hh:mm:ss zero-padded",
			numFmtID: 21,
			want:     "09:05:03",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(serial, tc.numFmtID, "", false)
			if got != tc.want {
				t.Errorf("FormatValue(serial=%v, ID=%d) = %q, want %q", serial, tc.numFmtID, got, tc.want)
			}
		})
	}
}

// TestFormatValueEraTokens verifies that E/EE era tokens fall back to the
// Gregorian year for non-CJK locales (matching excelize behaviour), and that
// G/GG/GGG era-name tokens produce no output.
func TestFormatValueEraTokens(t *testing.T) {
	t.Helper()
	// Serial 45285 = 2023-12-25 in the 1900 date system.
	serial := float64(45285)

	tests := []struct {
		name   string
		fmtStr string
		want   string
	}{
		// E alone — Gregorian year as integer (matches excelize).
		{
			name:   "E token: Gregorian year",
			fmtStr: "E",
			want:   "2023",
		},
		// EE — same as E for western locales.
		{
			name:   "EE token: Gregorian year",
			fmtStr: "EE",
			want:   "2023",
		},
		// E combined with other date tokens — only the year part of E, the
		// rest renders normally.
		{
			name:   "E/MM/DD combined",
			fmtStr: "E/MM/DD",
			want:   "2023/12/25",
		},
		// G token — silent for western locales; remaining tokens still render.
		{
			name:   "G token: silent, date still renders",
			fmtStr: "GYYYY",
			want:   "2023",
		},
		// GGG token — also silent.
		{
			name:   "GGG token: silent, date still renders",
			fmtStr: "GGGYYYY",
			want:   "2023",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(serial, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(45285, %q) = %q, want %q", tc.fmtStr, got, tc.want)
			}
		})
	}
}

// TestFormatValueChineseAmPm verifies that the Chinese AM/PM token (上午/下午)
// is rendered correctly in date/time format strings.
func TestFormatValueChineseAmPm(t *testing.T) {
	tests := []struct {
		name   string
		serial float64 // fractional days
		fmtStr string
		want   string
	}{
		// Serial 0.25 = 06:00 AM (6/24 of a day).
		{
			name:   "morning 上午",
			serial: 6.0 / 24.0,
			fmtStr: `上午/下午hh"時"mm"分"`,
			want:   "上午06時00分",
		},
		// Serial 13/24 = 13:00 = 1 PM.
		{
			name:   "afternoon 下午",
			serial: 13.0 / 24.0,
			fmtStr: `上午/下午hh"時"mm"分"`,
			want:   "下午01時00分",
		},
		// Midnight (serial 0 = 00:00) → 上午.
		{
			name:   "midnight 上午",
			serial: 0.0,
			fmtStr: `上午/下午hh"時"`,
			want:   "上午12時",
		},
		// Noon (serial 0.5 = 12:00) → 下午.
		{
			name:   "noon 下午",
			serial: 0.5,
			fmtStr: `上午/下午hh"時"`,
			want:   "下午12時",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()
			got := numfmt.FormatValue(tc.serial, 164, tc.fmtStr, false)
			if got != tc.want {
				t.Errorf("FormatValue(%v, %q) = %q, want %q", tc.serial, tc.fmtStr, got, tc.want)
			}
		})
	}
}
