# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/TsubasaBE/go-xlsb/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/TsubasaBE/go-xlsb/releases/tag/v0.2.0
