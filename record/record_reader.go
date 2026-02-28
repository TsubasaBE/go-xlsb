// Package record provides low-level BIFF12 record parsing primitives.
package record

import (
	"encoding/binary"
	"fmt"
	"io"
	"math"
	"unicode/utf16"
)

// RecordReader wraps the raw bytes of a single BIFF12 record and provides
// typed read helpers that mirror Python's RecordReader class.
type RecordReader struct {
	data []byte
	pos  int
}

// NewRecordReader creates a RecordReader over the given byte slice.
func NewRecordReader(data []byte) *RecordReader {
	return &RecordReader{data: data}
}

// remaining returns the number of unread bytes.
func (r *RecordReader) remaining() int {
	return len(r.data) - r.pos
}

// Skip advances the read position by n bytes.
func (r *RecordReader) Skip(n int) error {
	if n < 0 {
		return fmt.Errorf("record: skip count %d is negative", n)
	}
	if r.remaining() < n {
		return fmt.Errorf("record: skip %d bytes but only %d remain", n, r.remaining())
	}
	r.pos += n
	return nil
}

// Read reads exactly len(p) bytes into p.
func (r *RecordReader) Read(p []byte) error {
	n := len(p)
	if r.remaining() < n {
		return io.ErrUnexpectedEOF
	}
	copy(p, r.data[r.pos:r.pos+n])
	r.pos += n
	return nil
}

// ReadUint8 reads one unsigned byte.
func (r *RecordReader) ReadUint8() (uint8, error) {
	if r.remaining() < 1 {
		return 0, io.ErrUnexpectedEOF
	}
	v := r.data[r.pos]
	r.pos++
	return v, nil
}

// ReadUint16 reads a little-endian uint16.
func (r *RecordReader) ReadUint16() (uint16, error) {
	if r.remaining() < 2 {
		return 0, io.ErrUnexpectedEOF
	}
	v := binary.LittleEndian.Uint16(r.data[r.pos:])
	r.pos += 2
	return v, nil
}

// ReadUint32 reads a little-endian uint32.
func (r *RecordReader) ReadUint32() (uint32, error) {
	if r.remaining() < 4 {
		return 0, io.ErrUnexpectedEOF
	}
	v := binary.LittleEndian.Uint32(r.data[r.pos:])
	r.pos += 4
	return v, nil
}

// ReadInt32 reads a little-endian int32.
func (r *RecordReader) ReadInt32() (int32, error) {
	u, err := r.ReadUint32()
	return int32(u), err
}

// ReadFloat reads a 4-byte packed numeric value used by NUM records.
//
// The encoding mirrors the Python read_float implementation:
//   - Bits 0 and 1 are flag bits.
//   - If bit 1 is set the value is a scaled integer (>> 2).
//   - Otherwise the lower 4 bytes are the high word of a double (low word = 0).
//   - If bit 0 is set the final value is divided by 100.
func (r *RecordReader) ReadFloat() (float64, error) {
	raw, err := r.ReadInt32()
	if err != nil {
		return 0, err
	}

	var v float64
	if raw&0x02 != 0 {
		// Arithmetic right-shift by 2, exactly as Python does: `float(intval >> 2)`.
		// Python's >> on a signed int is arithmetic (sign-extending), and Go's >> on
		// int32 is also arithmetic, so this matches for both positive and negative values.
		// The previous uint32 detour was incorrect: it produced wrong results for any
		// negative raw value (e.g. raw=-4 → Python gives -1, uint32 path gives 1073741823).
		v = float64(raw >> 2)
	} else {
		// Reconstruct a double: low 32 bits = 0, high 32 bits = raw & 0xFFFFFFFC
		hi := uint32(raw) & 0xFFFFFFFC
		bits := uint64(hi) << 32
		v = math.Float64frombits(bits)
	}
	if raw&0x01 != 0 {
		v /= 100
	}
	return v, nil
}

// ReadDouble reads a little-endian IEEE-754 double (8 bytes).
func (r *RecordReader) ReadDouble() (float64, error) {
	if r.remaining() < 8 {
		return 0, io.ErrUnexpectedEOF
	}
	bits := binary.LittleEndian.Uint64(r.data[r.pos:])
	r.pos += 8
	return math.Float64frombits(bits), nil
}

// ReadString reads a 4-byte little-endian character count followed by that
// many UTF-16LE code units and decodes them to a Go string.
func (r *RecordReader) ReadString() (string, error) {
	charCount, err := r.ReadUint32()
	if err != nil {
		return "", err
	}
	// Guard against overflow: charCount*2 must fit in int on all platforms.
	// math.MaxInt/2 as a uint32-compatible bound: cap at 0x3FFFFFFF (~1 billion chars).
	const maxChars = 0x3FFFFFFF
	if charCount > maxChars {
		return "", fmt.Errorf("record: string length %d is too large", charCount)
	}
	byteCount := int(charCount) * 2
	if r.remaining() < byteCount {
		return "", io.ErrUnexpectedEOF
	}
	raw := r.data[r.pos : r.pos+byteCount]
	r.pos += byteCount
	return decodeUTF16LE(raw), nil
}

// decodeUTF16LE converts a byte slice of UTF-16 little-endian code units into
// a UTF-8 Go string. Invalid code units are replaced with the Unicode
// replacement character (U+FFFD), matching Python's errors='replace' behaviour.
func decodeUTF16LE(b []byte) string {
	if len(b) == 0 {
		return ""
	}
	// Pair bytes into uint16 code units.
	n := len(b) / 2
	u16 := make([]uint16, n)
	for i := range n {
		u16[i] = binary.LittleEndian.Uint16(b[i*2:])
	}
	runes := utf16.Decode(u16)
	return string(runes)
}
