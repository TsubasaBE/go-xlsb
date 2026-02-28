// Package biff12 contains all BIFF12 record-type constants used in the .xlsb
// format (Office Open XML Binary format).
package biff12

// Record IDs are encoded as variable-length integers inside the binary stream.
// All values below match the constants defined by the ECMA-376 specification
// and the pyxlsb reference implementation.
const (
	// ── Workbook records ──────────────────────────────────────────────────────
	DefinedName           = 0x0027
	FileVersion           = 0x0180
	Workbook              = 0x0183
	WorkbookEnd           = 0x0184
	BookViews             = 0x0187
	BookViewsEnd          = 0x0188
	Sheets                = 0x018F
	SheetsEnd             = 0x0190
	WorkbookPr            = 0x0199
	Sheet                 = 0x019C
	CalcPr                = 0x019D
	WorkbookView          = 0x019E
	ExternalReferences    = 0x02E1
	ExternalReferencesEnd = 0x02E2
	ExternalReference     = 0x02E3
	WebPublishing         = 0x04A9

	// ── Worksheet records ─────────────────────────────────────────────────────
	Row                      = 0x0000
	Blank                    = 0x0001
	Num                      = 0x0002
	BoolErr                  = 0x0003
	Bool                     = 0x0004
	Float                    = 0x0005
	String                   = 0x0007
	FormulaString            = 0x0008
	FormulaFloat             = 0x0009
	FormulaBool              = 0x000A
	FormulaBoolErr           = 0x000B
	Col                      = 0x003C
	Worksheet                = 0x0181
	WorksheetEnd             = 0x0182
	SheetViews               = 0x0185
	SheetViewsEnd            = 0x0186
	SheetView                = 0x0189
	SheetViewEnd             = 0x018A
	SheetData                = 0x0191
	SheetDataEnd             = 0x0192
	SheetPr                  = 0x0193
	Dimension                = 0x0194
	Selection                = 0x0198
	Cols                     = 0x0386
	ColsEnd                  = 0x0387
	ConditionalFormatting    = 0x03CD
	ConditionalFormattingEnd = 0x03CE
	CfRule                   = 0x03CF
	CfRuleEnd                = 0x03D0
	IconSet                  = 0x03D1
	IconSetEnd               = 0x03D2
	DataBar                  = 0x03D3
	DataBarEnd               = 0x03D4
	ColorScale               = 0x03D5
	ColorScaleEnd            = 0x03D6
	Cfvo                     = 0x03D7
	PageMargins              = 0x03DC
	PrintOptions             = 0x03DD
	PageSetup                = 0x03DE
	HeaderFooter             = 0x03DF
	SheetFormatPr            = 0x03E5
	Hyperlink                = 0x03EE
	Drawing                  = 0x04A6
	LegacyDrawing            = 0x04A7
	Color                    = 0x04B4
	OleObjects               = 0x04FE
	OleObject                = 0x04FF
	OleObjectsEnd            = 0x0580
	TableParts               = 0x0594
	TablePart                = 0x0595
	TablePartsEnd            = 0x0596

	// ── SharedStrings records ─────────────────────────────────────────────────
	Si     = 0x0013
	Sst    = 0x019F
	SstEnd = 0x01A0

	// ── Styles records ────────────────────────────────────────────────────────
	Font            = 0x002B
	Fill            = 0x002D
	Border          = 0x002E
	Xf              = 0x002F
	CellStyle       = 0x0030
	StyleSheet      = 0x0296
	StyleSheetEnd   = 0x0297
	Colors          = 0x03D9
	ColorsEnd       = 0x03DA
	Dxfs            = 0x03F9
	DxfsEnd         = 0x03FA
	TableStyles     = 0x03FC
	TableStylesEnd  = 0x03FD
	Fills           = 0x04DB
	FillsEnd        = 0x04DC
	Fonts           = 0x04E3
	FontsEnd        = 0x04E4
	Borders         = 0x04E5
	BordersEnd      = 0x04E6
	CellXfs         = 0x04E9
	CellXfsEnd      = 0x04EA
	CellStyles      = 0x04EB
	CellStylesEnd   = 0x04EC
	CellStyleXfs    = 0x04F2
	CellStyleXfsEnd = 0x04F3

	// ── Comment records ───────────────────────────────────────────────────────
	Comments       = 0x04F4
	CommentsEnd    = 0x04F5
	Authors        = 0x04F6
	AuthorsEnd     = 0x04F7
	Author         = 0x04F8
	CommentList    = 0x04F9
	CommentListEnd = 0x04FA
	Comment        = 0x04FB
	CommentEnd     = 0x04FC
	Text           = 0x04FD

	// ── Table records ─────────────────────────────────────────────────────────
	MergeCells    = 0x01B1
	MergeCellsEnd = 0x01B2
	MergeCell     = 0x01B0

	AutoFilter      = 0x01A1
	AutoFilterEnd   = 0x01A2
	FilterColumn    = 0x01A3
	FilterColumnEnd = 0x01A4
	Filters         = 0x01A5
	FiltersEnd      = 0x01A6
	Filter          = 0x01A7
	Table           = 0x02D7
	TableEnd        = 0x02D8
	TableColumns    = 0x02D9
	TableColumnsEnd = 0x02DA
	TableColumn     = 0x02DB
	TableColumnEnd  = 0x02DC
	TableStyleInfo  = 0x0481
	SortState       = 0x0492
	SortCondition   = 0x0494
	SortStateEnd    = 0x0495

	// ── QueryTable records ────────────────────────────────────────────────────
	QueryTable           = 0x03BF
	QueryTableEnd        = 0x03C0
	QueryTableRefresh    = 0x03C1
	QueryTableRefreshEnd = 0x03C2
	QueryTableFields     = 0x03C7
	QueryTableFieldsEnd  = 0x03C8
	QueryTableField      = 0x03C9
	QueryTableFieldEnd   = 0x03CA

	// ── Connection records ────────────────────────────────────────────────────
	Connections    = 0x03AD
	ConnectionsEnd = 0x03AE
	Connection     = 0x01C9
	ConnectionEnd  = 0x01CA
	DbPr           = 0x01CB
	DbPrEnd        = 0x01CC
)
