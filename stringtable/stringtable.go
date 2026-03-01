// Package stringtable parses the xl/sharedStrings.bin part of an .xlsb file
// and provides indexed access to the shared string values.
package stringtable

import (
	"bytes"
	"fmt"
	"io"

	"github.com/TsubasaBE/go-xlsb/biff12"
	"github.com/TsubasaBE/go-xlsb/record"
)

// StringTable holds the shared strings parsed from xl/sharedStrings.bin.
type StringTable struct {
	strings []string
}

// New reads all shared string entries from r and returns a populated
// StringTable.  r must be positioned at the start of the sharedStrings.bin
// payload (i.e. the SST record stream).
func New(r io.ReadSeeker) (*StringTable, error) {
	st := &StringTable{}
	rdr := record.NewReader(r)
	for {
		recID, data, err := rdr.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("stringtable: %w", err)
		}

		switch recID {
		case biff12.Si:
			s, err := parseSI(data)
			if err != nil {
				// Treat malformed SI as empty string rather than aborting.
				s = ""
			}
			st.strings = append(st.strings, s)
		case biff12.SstEnd:
			return st, nil
		}
	}
	return st, nil
}

// Get returns the shared string at index idx.  It panics if idx is out of
// range, matching the behaviour of a slice index.
func (st *StringTable) Get(idx int) string {
	return st.strings[idx]
}

// Len returns the total number of shared strings loaded.
func (st *StringTable) Len() int {
	return len(st.strings)
}

// parseSI decodes a single SI (string instance) record payload.
//
// BrtSi layout per MS-XLSB §2.4.726:
//
//	flags    uint8   — bit 0: fRichStr (rich-text run data follows string)
//	                   bit 1: fExtStr  (phonetic/extended data follows)
//	if fRichStr: crun uint32  — number of rich-text run entries
//	if fExtStr:  sz   uint32  — byte size of phonetic data
//	string   XLWideString    — the actual text (4-byte char count + UTF-16LE)
//
// (Rich-text run records and phonetic data follow the string in the record
// stream, not the record payload, so we do not need to skip them here.)
func parseSI(data []byte) (string, error) {
	if len(data) == 0 {
		return "", nil
	}
	rr := record.NewRecordReader(data)

	// Read the flag byte.
	flags, err := rr.ReadUint8()
	if err != nil {
		// No data at all — treat as empty string.
		return "", nil
	}
	fRichStr := (flags & 0x01) != 0
	fExtStr := (flags & 0x02) != 0

	// If fRichStr is set, a 4-byte crun count follows before the string.
	if fRichStr {
		if _, err := rr.ReadUint32(); err != nil {
			return "", fmt.Errorf("parseSI: read crun: %w", err)
		}
	}
	// If fExtStr is set, a 4-byte phonetic-data size follows.
	if fExtStr {
		if _, err := rr.ReadUint32(); err != nil {
			return "", fmt.Errorf("parseSI: read extStr size: %w", err)
		}
	}

	s, err := rr.ReadString()
	if err != nil {
		return "", fmt.Errorf("parseSI: %w", err)
	}
	return s, nil
}

// NewFromBytes is a convenience wrapper that builds a StringTable from an
// in-memory byte slice (useful in tests).
func NewFromBytes(b []byte) (*StringTable, error) {
	return New(bytes.NewReader(b))
}
