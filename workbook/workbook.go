// Package workbook opens and parses an .xlsb workbook file (a ZIP archive).
package workbook

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strings"

	"github.com/TsubasaBE/go-xlsb/biff12"
	"github.com/TsubasaBE/go-xlsb/record"
	"github.com/TsubasaBE/go-xlsb/stringtable"
	"github.com/TsubasaBE/go-xlsb/worksheet"
)

// sheetEntry holds the display name and the zip-internal path target for one
// worksheet.
type sheetEntry struct {
	name   string
	target string // e.g. "worksheets/sheet1.bin"
}

// Workbook represents an open .xlsb workbook.
type Workbook struct {
	zr          *zip.ReadCloser // non-nil when opened by file name
	zf          *zip.Reader     // always non-nil
	sheets      []sheetEntry
	stringTable *stringtable.StringTable
}

// Open opens the named .xlsb file and parses its workbook metadata.
// The caller must call Close on the returned Workbook when done to release the
// underlying file handle.
func Open(name string) (*Workbook, error) {
	rc, err := zip.OpenReader(name)
	if err != nil {
		return nil, fmt.Errorf("workbook: open %q: %w", name, err)
	}
	wb := &Workbook{zr: rc, zf: &rc.Reader}
	if err := wb.parse(); err != nil {
		_ = rc.Close()
		return nil, err
	}
	return wb, nil
}

// OpenReader parses an .xlsb workbook from an in-memory ReaderAt.
// size must be the total byte size of the ZIP data.
func OpenReader(r io.ReaderAt, size int64) (*Workbook, error) {
	zf, err := zip.NewReader(r, size)
	if err != nil {
		return nil, fmt.Errorf("workbook: open reader: %w", err)
	}
	wb := &Workbook{zf: zf}
	if err := wb.parse(); err != nil {
		return nil, err
	}
	return wb, nil
}

// Sheets returns the display names of all worksheets in order.
func (wb *Workbook) Sheets() []string {
	names := make([]string, len(wb.sheets))
	for i, s := range wb.sheets {
		names[i] = s.name
	}
	return names
}

// Sheet returns the worksheet at the given 1-based index.
// Index 1 refers to the first sheet. An out-of-range index returns a non-nil
// error describing the valid range.
func (wb *Workbook) Sheet(idx int) (*worksheet.Worksheet, error) {
	if idx < 1 || idx > len(wb.sheets) {
		return nil, fmt.Errorf("workbook: sheet index %d out of range [1, %d]", idx, len(wb.sheets))
	}
	return wb.openSheet(wb.sheets[idx-1])
}

// SheetByName returns the worksheet with the given name (case-insensitive).
// It returns a non-nil error if no sheet with that name exists.
func (wb *Workbook) SheetByName(name string) (*worksheet.Worksheet, error) {
	lower := strings.ToLower(name)
	for _, s := range wb.sheets {
		if strings.ToLower(s.name) == lower {
			return wb.openSheet(s)
		}
	}
	return nil, fmt.Errorf("workbook: sheet %q not found", name)
}

// Close releases the underlying ZIP file handle.
// It is a no-op when the workbook was opened via OpenReader (no file handle to
// release), and always returns nil in that case.
func (wb *Workbook) Close() error {
	if wb.zr != nil {
		return wb.zr.Close()
	}
	return nil
}

// ── internal ─────────────────────────────────────────────────────────────────

// parse reads workbook.bin and sharedStrings.bin (if present).
func (wb *Workbook) parse() error {
	if err := wb.parseWorkbook(); err != nil {
		return err
	}
	if err := wb.parseSharedStrings(); err != nil {
		return err
	}
	return nil
}

// parseWorkbook reads xl/_rels/workbook.bin.rels (XML) and xl/workbook.bin
// to build the sheet list.
func (wb *Workbook) parseWorkbook() error {
	// Step 1: load relationship ID → target map from the .rels XML.
	rels, err := wb.readRels("xl/_rels/workbook.bin.rels")
	if err != nil {
		return fmt.Errorf("workbook: parse rels: %w", err)
	}

	// Step 2: read workbook.bin record stream.
	data, err := wb.readZipEntry("xl/workbook.bin")
	if err != nil {
		return fmt.Errorf("workbook: read workbook.bin: %w", err)
	}

	rdr := record.NewReader(bytes.NewReader(data))
	for {
		recID, recData, err := rdr.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			return fmt.Errorf("workbook: %w", err)
		}

		if recID == biff12.Sheet {
			entry, err := parseSheetRecord(recData, rels)
			if err != nil {
				return fmt.Errorf("workbook: parse SHEET record: %w", err)
			}
			wb.sheets = append(wb.sheets, entry)
		} else if recID == biff12.SheetsEnd {
			break
		}
	}
	return nil
}

// parseSharedStrings reads xl/sharedStrings.bin if it exists.
func (wb *Workbook) parseSharedStrings() error {
	data, err := wb.readZipEntry("xl/sharedStrings.bin")
	if err != nil {
		// File is optional — no shared strings in this workbook.
		return nil
	}
	st, err := stringtable.New(bytes.NewReader(data))
	if err != nil {
		return fmt.Errorf("workbook: shared strings: %w", err)
	}
	wb.stringTable = st
	return nil
}

// openSheet reads the binary data for the given sheet entry and returns a
// ready-to-use Worksheet.
func (wb *Workbook) openSheet(entry sheetEntry) (*worksheet.Worksheet, error) {
	// Resolve "worksheets/sheet1.bin" → "xl/worksheets/sheet1.bin".
	// Absolute targets (starting with "/") are used as-is after stripping the
	// leading slash; relative targets are prefixed with "xl/".
	target := strings.TrimPrefix(entry.target, "/")
	var zipPath string
	if strings.HasPrefix(target, "xl/") {
		zipPath = target
	} else {
		zipPath = "xl/" + target
	}

	data, err := wb.readZipEntry(zipPath)
	if err != nil {
		return nil, fmt.Errorf("workbook: open sheet %q: %w", entry.name, err)
	}

	// Attempt to load the sheet .rels file (optional; needed for hyperlinks).
	lastSlash := strings.LastIndex(zipPath, "/")
	relsPath := zipPath[:lastSlash+1] + "_rels/" + zipPath[lastSlash+1:] + ".rels"
	relsData, _ := wb.readZipEntry(relsPath) // ignore error — it's optional

	return worksheet.New(entry.name, data, relsData, wb.stringTable)
}

// readZipEntry reads the full contents of a named entry from the ZIP archive.
func (wb *Workbook) readZipEntry(name string) ([]byte, error) {
	for _, f := range wb.zf.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			data, readErr := io.ReadAll(rc)
			closeErr := rc.Close()
			if readErr != nil {
				return nil, readErr
			}
			// Propagate decompressor checksum / close errors even when the read
			// appeared to succeed (e.g. truncated gzip stream).
			if closeErr != nil {
				return nil, closeErr
			}
			return data, nil
		}
	}
	return nil, fmt.Errorf("%q not found in archive", name)
}

// readRels parses a .rels XML file and returns a map of Id → Target.
func (wb *Workbook) readRels(name string) (map[string]string, error) {
	data, err := wb.readZipEntry(name)
	if err != nil {
		return nil, err
	}
	return parseRelsXML(data)
}

// ── XML relationship parsing ──────────────────────────────────────────────────

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

// ── SHEET record parsing ───────────────────────────────────────────────────────

// parseSheetRecord decodes a SHEET record payload.
//
// Python layout:
//
//	reader.skip(4)           # 4 unknown bytes
//	sheetid = reader.read_int()
//	relid   = reader.read_string()
//	name    = reader.read_string()
func parseSheetRecord(data []byte, rels map[string]string) (sheetEntry, error) {
	rr := record.NewRecordReader(data)

	if err := rr.Skip(4); err != nil {
		return sheetEntry{}, fmt.Errorf("skip state flags: %w", err)
	}
	if _, err := rr.ReadUint32(); err != nil { // sheetId — not used by us
		return sheetEntry{}, fmt.Errorf("read sheetId: %w", err)
	}
	relID, err := rr.ReadString()
	if err != nil {
		return sheetEntry{}, fmt.Errorf("read relId: %w", err)
	}
	name, err := rr.ReadString()
	if err != nil {
		return sheetEntry{}, fmt.Errorf("read sheet name: %w", err)
	}

	target, ok := rels[relID]
	if !ok {
		return sheetEntry{}, fmt.Errorf("no relationship found for rId %q", relID)
	}
	return sheetEntry{name: name, target: target}, nil
}
