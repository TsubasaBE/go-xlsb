# AGENTS.md — go-xlsb

This file documents conventions and workflows for agentic coding assistants operating
in this repository. Keep it up-to-date when conventions change.

---

## Repository Overview

**Module:** `github.com/TsubasaBE/go-xlsb`  
**Language:** Go (`go 1.25` in `go.mod`; minimum runtime required is Go 1.22)  
**External dependencies:** `github.com/xuri/nfp` (number format token parser)  
**License:** LGPL-3.0-or-later  

A pure-Go reader for the Excel Binary Workbook (`.xlsb`) format, ported from the
Python library [pyxlsb](https://github.com/willtrnr/pyxlsb) by William Turner.

### Package layout

```
go-xlsb/           ← root package `xlsb`; public API (Open, OpenReader, ConvertDate)
├── biff12/        ← BIFF12 record-ID constants (pure const table, no logic)
├── internal/
│   ├── dateformat/ ← isDateFormat helper shared by workbook/ and numfmt/
│   └── rels/       ← shared XML relationship types (xmlRelationships, xmlRelationship)
├── numfmt/        ← number format rendering (FormatValue)
├── record/        ← low-level BIFF12 stream reader + typed field reader
├── stringtable/   ← shared-string table (SST) parser
├── styles/        ← styles.bin parser; XFStyle, StyleTable, BuiltInNumFmt
├── workbook/      ← ZIP opener, sheet list, relationship XML
├── worksheet/     ← per-sheet parser + Rows() iterator
├── xls/           ← real .xlsb fixtures (gitignored); used by integration tests
├── _probe.go      ← //go:build ignore developer CLI tool; never compiled normally
├── xlsb.go        ← root package implementation
├── xlsb_test.go   ← unit tests (package xlsb_test)
└── xlsb_integration_test.go  ← integration tests against real .xlsb files
```

---

## Build, Vet, and Test Commands

```bash
# Build all packages
go build ./...

# Run all tests (unit + integration; integration tests skip if xls/ files are absent)
go test ./...

# Run all tests with verbose output
go test -v ./...

# Run only unit tests (no real files needed)
go test -run ^Test .

# Run a single named test (supports regex; -run is anchored with ^...$)
go test -run ^TestConvertDate$ .
go test -run ^TestRecordReaderReadUint8$ .
go test -run ^TestWorkbookSheets$ .

# Run tests matching a prefix
go test -run ^TestRecord .

# Run a subtest (parent/subtest_name)
go test -run ^TestConvertDate/negative_serial .

# Run integration tests only (require real .xlsb files in xls/)
go test -run ^TestReal .
go test -run ^TestMaand .
go test -run ^TestPlanning .

# Run with race detector
go test -race ./...

# Static analysis (always run before committing)
go vet ./...

# Format all Go source files
gofmt -l -w .
# or
gofmt -d .   # diff only
```

No `Makefile`, `taskfile.yml`, or custom build scripts exist — use plain `go` toolchain
commands only.

---

## Testing Conventions

### Unit tests (`xlsb_test.go`)

- **Package:** `xlsb_test` — black-box testing only. Internal packages may be imported
  directly by their full module path when needed.
- **In-memory fixtures only.** Never write `.xlsb` files to disk. Use helper functions
  such as `buildMinimalXLSB(t)`, `buildSSTBytes(strs)`, `encodeRecord(id, payload)`.
- Sub-packages (`biff12`, `record`, `stringtable`, `workbook`, `worksheet`) have no
  test files of their own; they are exercised through this top-level file.

### Integration tests (`xlsb_integration_test.go`)

- Also `package xlsb_test`. Tests use real `.xlsb` files from the `xls/` directory
  (gitignored). Every test must call `t.Skip()` gracefully when the file is absent so
  the suite is safe to run in CI without the fixture files.
- Use the `xlsbPath(t, name)` helper (skips if absent) and `openXLSB(t, name)` (opens
  workbook and registers `t.Cleanup` for `Close()`).

### General rules

- **Table-driven subtests** (`t.Run`) for any function with multiple input/output cases.
  Table struct fields follow the pattern `name, input, want, wantErr`.
- **Flat tests** (no subtests) for single-behaviour unit tests.
- Helper functions must call `t.Helper()` so failure lines point to the caller.
- Use `t.Fatal` / `t.Fatalf` for setup or prerequisite failures (the test cannot
  proceed). Use `t.Error` / `t.Errorf` for assertion failures (continue checking).
- Always `defer wb.Close()` immediately after opening a workbook in a test.

### Test naming

| Target | Convention | Example |
|---|---|---|
| Type method | `Test<Type><Method>` | `TestRecordReaderReadUint8` |
| Top-level function | `Test<FunctionName>` | `TestConvertDate` |
| Edge case variant | `Test<FunctionName><EdgeCase>` | `TestConvertDateNaN` |

---

## Code Style

### Formatting

- All code must be `gofmt`-formatted. No configuration needed; use default `gofmt`.
- Line length is not rigidly enforced, but prefer staying under ~100 characters.

### Imports

Two-block import style, separated by a blank line:

```go
import (
    // 1. Standard library — alphabetical
    "encoding/binary"
    "fmt"
    "io"

    // 2. Internal module packages — alphabetical
    "github.com/TsubasaBE/go-xlsb/biff12"
    "github.com/TsubasaBE/go-xlsb/record"
)
```

There is one external dependency (`github.com/xuri/nfp`). If additional third-party
packages are ever added, they form a third import block between stdlib and internal.

### Naming

| Entity | Style | Examples |
|---|---|---|
| Packages | lowercase, single word | `biff12`, `record`, `stringtable` |
| Exported types | PascalCase | `Workbook`, `Cell`, `RecordReader` |
| Unexported types | camelCase | `sheetEntry`, `internalCell` |
| Exported functions | PascalCase | `Open`, `ConvertDate`, `NewReader` |
| Unexported functions | camelCase | `parseSheetRecord`, `readZipEntry` |
| Constants | PascalCase (exported), camelCase (unexported) | `Version`, `maxRecordLen` |
| Receiver names | Short abbreviation of type | `wb` (*Workbook), `ws` (*Worksheet), `r` (*Reader), `rr` (*RecordReader), `st` (*StringTable) |
| Test helpers | camelCase | `buildMinimalXLSB`, `encodeRecord` |

### Error handling

1. **Wrap with context using `fmt.Errorf` and `%w`** for all non-EOF errors:
   ```go
   return nil, fmt.Errorf("workbook: open %q: %w", name, err)
   ```
   Error strings always begin with the package name: `"<package>: <context>: %w"`.

2. **Direct `== io.EOF` comparison** (not `errors.Is`) for BIFF12 stream loops where
   `io.EOF` is a normal loop-termination sentinel, not an error condition:
   ```go
   if err == io.EOF { break }
   ```

3. **Graceful degradation** for optional or malformed data — prefer an empty/zero
   value over aborting when the spec allows it:
   ```go
   // Treat malformed SI record as empty string rather than aborting.
   s = ""
   ```

4. **Bounds guards before allocations** to prevent OOM from corrupt input:
   ```go
   const maxRecordLen = 10 * 1024 * 1024
   if recLen > maxRecordLen {
       return nil, fmt.Errorf("record: length %d exceeds limit", recLen)
   }
   ```

5. **`panic("unreachable")`** only after exhaustive switches or loops where the
   compiler cannot prove termination — never for recoverable errors.

6. **Blank identifier `_`** only for intentionally ignored cleanup errors:
   ```go
   _ = rc.Close()
   ```

### Types

- Use `any` for the polymorphic cell value `Cell.V` (may be `nil`, `string`,
  `float64`, or `bool`).
- Raw binary reads use the smallest correct unsigned type (`uint8`, `uint16`,
  `uint32`), then cast to `int` immediately for subsequent logic.
- `[2]int` as a map key is acceptable for compact coordinate lookups (e.g. hyperlinks).
- Use `io.ReadSeeker` when the reader must support seeking (worksheet two-pass parse).
- Use `io.ReaderAt` for the public API where the caller provides random-access data
  (e.g. `bytes.NewReader`, `os.File`).
- Range-over-func (`func(yield func([]Cell) bool)`) is the iterator pattern for
  `Rows()`. This requires Go 1.22+.
- Integer range (`for i := range N`) is preferred over `for i := 0; i < N; i++`.

### Comments

- Every package must have a `// Package <name> ...` doc comment.
- Every exported symbol must have a doc comment starting with the symbol name.
- Unexported functions get doc comments when the logic is non-obvious.
- Inline comments explain *why*, not *what*.
- Long files use section-divider comments to group related code:
  ```go
  // ── record decoders ────────────────────────────────────────────────────────
  ```

---

## README Conventions

- No emojis anywhere in `README.md`.
- New sections should match the register and style of the existing content.

---

## Architecture Notes

- **Two-pass worksheet parsing:** `worksheet.New()` pre-scans the stream for metadata
  (dimension, column widths, hyperlinks, merge cells), recording the byte offset where
  row data begins (`dataOffset`). Every call to `Rows()` seeks back to `dataOffset` and
  streams rows on demand. Keep this contract intact when modifying worksheet parsing.
- **`Rows(sparse bool)`:** when `sparse` is false (dense mode), every row is padded to
  the full sheet width. Only the anchor cell of a merged region carries a value;
  satellite cells (other rows/columns of the merge) are emitted as zero-value `Cell`
  entries. Use `MergeCells` in application code to propagate anchor values if needed.
- **No global state.** All state is held in struct fields.
- **Zero allocations on the hot path** is a goal. Avoid unnecessary heap allocations
  inside `Rows()` loops.
- **The `biff12` package** is a pure constant table — no logic, no init side-effects.
- **`internal/rels` package:** the `xmlRelationships` and `xmlRelationship` types and
  the `parseRelsXML` function live in `internal/rels/rels.go` (`package rels`). Both
  `workbook/workbook.go` and `worksheet/worksheet.go` import this shared package.
- **`_probe.go`** has `//go:build ignore` and is `package main`. It is a developer
  convenience tool for dumping sheet metadata from files in `xls/`. It is never
  compiled as part of the library and should not be modified during normal development.
