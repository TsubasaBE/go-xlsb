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
	"github.com/TsubasaBE/go-xlsb/record"
	"github.com/TsubasaBE/go-xlsb/stringtable"
	"github.com/TsubasaBE/go-xlsb/workbook"
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

// Compile-time check: the top-level xlsb package is importable and ConvertDate
// is accessible.
var _ = xlsb.ConvertDate
