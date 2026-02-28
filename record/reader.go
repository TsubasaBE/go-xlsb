package record

import (
	"fmt"
	"io"
)

// Reader iterates over BIFF12 records from an io.ReadSeeker.  Each call to
// Next returns the record type ID, the raw payload bytes, and any error.
//
// Record IDs and lengths are both variable-length encoded:
//   - ID:  up to 4 continuation bytes; the MSB of each byte signals more bytes.
//   - Len: up to 4 bytes of 7-bit little-endian chunks (standard LEB-128).
type Reader struct {
	r io.ReadSeeker
}

// NewReader wraps an io.ReadSeeker for BIFF12 record iteration.
func NewReader(r io.ReadSeeker) *Reader {
	return &Reader{r: r}
}

// Tell returns the current byte offset within the underlying stream.
func (r *Reader) Tell() (int64, error) {
	return r.r.Seek(0, io.SeekCurrent)
}

// Seek repositions the stream.  whence follows the io.Seek* constants.
func (r *Reader) Seek(offset int64, whence int) (int64, error) {
	return r.r.Seek(offset, whence)
}

// readID reads a variable-length record type ID (1–4 bytes).
// The continuation bit is the MSB (bit 7) of each byte; once a byte has
// bit 7 clear, reading stops.  Each byte contributes its 8 bits at increasing
// byte positions (simple byte-shift accumulation, NOT 7-bit stripping).
// Returns an error if the 4th byte still has the continuation bit set (the
// stream would otherwise become misaligned).
//
// Accumulation is done into uint32 to prevent signed-integer overflow on
// 32-bit platforms (where int is 32 bits).  The returned int is safe to use
// for record-ID switch statements because all defined BIFF12 IDs fit in uint16.
func (r *Reader) readID() (int, error) {
	buf := [1]byte{}
	var v uint32
	for i := range 4 {
		_, err := io.ReadFull(r.r, buf[:])
		if err != nil {
			return 0, err
		}
		b := uint32(buf[0])
		v += b << (8 * i)
		if b&0x80 == 0 {
			return int(v), nil
		}
		if i == 3 {
			return 0, fmt.Errorf("record: ID continuation bit set on 4th byte (stream corrupt)")
		}
	}
	// Unreachable: the loop always returns inside the body for i==3.
	panic("record: readID: unreachable")
}

// readLen reads a variable-length record length (1–4 bytes) encoded as
// 7-bit little-endian chunks (LEB-128 without sign extension).
// Returns an error if the 4th byte still has the continuation bit set.
//
// Accumulation is done into uint32 (matching readID) to avoid any
// signed-integer behaviour on 32-bit platforms where int is 32 bits.
// The returned int is safe because the maxRecordLen guard in Next() ensures
// the value is always within int range on all platforms.
func (r *Reader) readLen() (int, error) {
	buf := [1]byte{}
	var v uint32
	for i := range 4 {
		_, err := io.ReadFull(r.r, buf[:])
		if err != nil {
			return 0, err
		}
		b := uint32(buf[0])
		v += (b & 0x7F) << (7 * uint32(i))
		if b&0x80 == 0 {
			return int(v), nil
		}
		if i == 3 {
			return 0, fmt.Errorf("record: length continuation bit set on 4th byte (stream corrupt)")
		}
	}
	// Unreachable: the loop always returns inside the body for i==3.
	panic("record: readLen: unreachable")
}

// Next reads the next record from the stream.
// Returns (recID, data, nil) on success, or (0, nil, io.EOF) at end of stream.
// A truncated stream (record ID found but length or payload missing) returns a
// non-EOF error rather than silently masking data corruption as end-of-file.
func (r *Reader) Next() (recID int, data []byte, err error) {
	recID, err = r.readID()
	if err != nil {
		if err == io.EOF || err == io.ErrUnexpectedEOF {
			return 0, nil, io.EOF
		}
		return 0, nil, fmt.Errorf("record: reading ID: %w", err)
	}

	recLen, err := r.readLen()
	if err != nil {
		// EOF here means the stream was truncated after the record ID — that is
		// always a corruption, not a clean end-of-file.
		return 0, nil, fmt.Errorf("record: reading length after ID 0x%X: %w", recID, err)
	}

	// Guard against corrupt length fields that would cause multi-hundred-MB
	// allocations.  No legitimate BIFF12 record exceeds 10 MB.
	const maxRecordLen = 10 * 1024 * 1024 // 10 MiB
	if recLen > maxRecordLen {
		return 0, nil, fmt.Errorf("record: payload length %d for ID 0x%X exceeds %d byte limit (stream corrupt)", recLen, recID, maxRecordLen)
	}

	if recLen == 0 {
		return recID, nil, nil
	}

	data = make([]byte, recLen)
	if _, err = io.ReadFull(r.r, data); err != nil {
		return 0, nil, fmt.Errorf("record: reading %d payload bytes for ID 0x%X: %w", recLen, recID, err)
	}
	return recID, data, nil
}
