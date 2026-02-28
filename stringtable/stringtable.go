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
// The Python handler does:
//
//	reader.skip(1)          # skip one flag byte
//	val = reader.read_string()
func parseSI(data []byte) (string, error) {
	if len(data) == 0 {
		return "", nil
	}
	rr := record.NewRecordReader(data)
	if err := rr.Skip(1); err != nil {
		// No data beyond the flag byte — treat as empty string.
		return "", nil
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
