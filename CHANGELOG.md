# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.2] - 2026-02-28

### Fixed

- `record.Reader.readLen`: changed LEB-128 accumulator and byte variable from
  `int` to `uint32`, matching the existing `readID` implementation.  On 32-bit
  platforms the behaviour was already correct for all valid input, but using a
  signed architecture-dependent type was inconsistent and silently relied on the
  `maxRecordLen` guard keeping values in range.
- `worksheet.parseColRecord`: added a `maxStyleIndex` (`0x7FFFFFFF`) bounds
  check on the raw `uint32` style index before casting to `int`.  A corrupt
  value above `math.MaxInt32` previously produced a different `int` result on
  32-bit vs 64-bit platforms; it now returns a descriptive error on all
  platforms.

## [1.0.1] - 2026-02-28

### Added

- Updated godoc comments

## [1.0.0] - 2026-02-28

### Added

- Merge-cell value propagation in `Rows()`: non-anchor cells in a vertical merge
  now carry the anchor cell's value, matching the visual appearance in Excel.
  Previously, only the top-left anchor cell of a merged region held a value;
  all other rows in the merge returned `nil`.

## [0.2.0] - 2026-02-28

### Added

- Initial public release of the `go-xlsb` library.
- Pure-Go reader for Microsoft Excel Binary Workbook (`.xlsb`) files with no cgo dependency.
- `Open` function to open `.xlsb` files by path.
- `Workbook` type with `Sheets()` and `Sheet(n)` methods for workbook navigation.
- Row and cell iteration via `sheet.Rows()`.
- `ConvertDate` helper to convert Excel serial date floats to `time.Time`.
- Exported `Version` constant (`"0.2.0"`).
- Internal packages: `biff12`, `record`, `stringtable`, `workbook`, `worksheet`.
- Parsing and exposure of `MergeCells` (`[]MergeArea`) on `Worksheet`, including
  `MergeArea` type with `R`, `C`, `H`, `W` fields for the merged range.
- Integration test suite (`xlsb_integration_test.go`) covering real `.xlsb` fixtures.
- Developer probe tool (`_probe.go`, build-ignored).

### Removed

- `docs/official_microsoft_xlsb_format_documentation.pdf` removed from the
  repository; added to `.gitignore`.

[Unreleased]: https://github.com/TsubasaBE/go-xlsb/compare/v1.0.2...HEAD
[1.0.2]: https://github.com/TsubasaBE/go-xlsb/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/TsubasaBE/go-xlsb/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/TsubasaBE/go-xlsb/compare/v0.2.0...v1.0.0
[0.2.0]: https://github.com/TsubasaBE/go-xlsb/releases/tag/v0.2.0
