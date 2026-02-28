package xlsb_test

// xlsb_testhelpers_test.go — package-level BIFF12 encoding helpers shared by
// all fixture-builder functions in xlsb_test.go.
//
// These replace the identical closure declarations that previously appeared
// inside each of the eight buildXxx functions.

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"testing"
)

// biff12WriteID writes a BIFF12 record ID to buf using the variable-length
// continuation-bit encoding described in MS-XLSB §2.1.4.
//
// Each byte's MSB (bit 7) acts as a continuation flag: 1 means "more bytes
// follow", 0 means "this is the last byte".  The full 8-bit value of each
// byte contributes to the ID via simple byte-shift accumulation (not 7-bit
// stripping), mirroring record/reader.go readID.
//
// Encoding rules:
//   - IDs 0x00–0x7F fit in one byte (MSB clear → no continuation).
//   - IDs 0x80–0x7FFF fit in two bytes: the first byte must have MSB set
//     (continuation), so bits 0–6 carry the low 7 bits of the ID with bit 7
//     forced to 1 via OR 0x80; the second byte carries the remaining high
//     bits with MSB clear.
//   - IDs 0x8000–0x7FFFFF fit in three bytes (same pattern, not yet needed
//     for any defined BIFF12 constant but included for correctness).
func biff12WriteID(buf *bytes.Buffer, id int) {
	for {
		b := id & 0xFF
		id >>= 8
		if id > 0 {
			// More bytes follow: set the continuation bit.
			buf.WriteByte(byte(b) | 0x80)
		} else {
			// Last byte: MSB must be clear.
			buf.WriteByte(byte(b) &^ 0x80)
			break
		}
	}
}

// biff12WriteLen writes a BIFF12 variable-length record size to buf using
// the standard base-128 (LEB128) encoding.
func biff12WriteLen(buf *bytes.Buffer, n int) {
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

// biff12WriteRec writes a complete BIFF12 record (ID + length + payload) to buf.
func biff12WriteRec(buf *bytes.Buffer, id int, payload []byte) {
	biff12WriteID(buf, id)
	biff12WriteLen(buf, len(payload))
	buf.Write(payload)
}

// biff12EncStr encodes s as a BIFF12 XLWideString:
// a uint32 character-count followed by the UTF-16LE code units.
func biff12EncStr(s string) []byte {
	runes := []rune(s)
	var sb bytes.Buffer
	_ = binary.Write(&sb, binary.LittleEndian, uint32(len(runes)))
	for _, r := range runes {
		_ = binary.Write(&sb, binary.LittleEndian, uint16(r))
	}
	return sb.Bytes()
}

// biff12Le32 returns the little-endian 4-byte encoding of v.
func biff12Le32(v uint32) []byte {
	b := make([]byte, 4)
	binary.LittleEndian.PutUint32(b, v)
	return b
}

// biff12Le16 returns the little-endian 2-byte encoding of v.
func biff12Le16(v uint16) []byte {
	b := make([]byte, 2)
	binary.LittleEndian.PutUint16(b, v)
	return b
}

// zipAddFile writes data as a new entry named name into zw.
// It calls t.Fatalf on any error.
func zipAddFile(t *testing.T, zw *zip.Writer, name string, data []byte) {
	t.Helper()
	f, err := zw.Create(name)
	if err != nil {
		t.Fatalf("zip create %s: %v", name, err)
	}
	if _, err := f.Write(data); err != nil {
		t.Fatalf("zip write %s: %v", name, err)
	}
}
