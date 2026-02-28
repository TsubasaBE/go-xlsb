// Package biff12 contains all BIFF12 record-type constants used in the .xlsb
// format (Office Open XML Binary format).
package biff12

// Record IDs are encoded as variable-length integers inside the binary stream.
// All values below match the constants defined by the ECMA-376 specification
// and the pyxlsb reference implementation.
const (
	// ── Workbook records ──────────────────────────────────────────────────────

	// DefinedName marks a defined name entry in the workbook
	// (ECMA-376 §2.4.107, record ID 0x0027).
	DefinedName = 0x0027

	// FileVersion records the application version that last saved the file
	// (ECMA-376 §2.4.173, record ID 0x0180).
	FileVersion = 0x0180

	// Workbook marks the start of the workbook stream
	// (ECMA-376 §2.4.823, record ID 0x0183).
	Workbook = 0x0183

	// WorkbookEnd marks the end of the workbook stream
	// (ECMA-376 §2.4.824, record ID 0x0184).
	WorkbookEnd = 0x0184

	// BookViews marks the start of the book-views collection
	// (ECMA-376 §2.4.28, record ID 0x0187).
	BookViews = 0x0187

	// BookViewsEnd marks the end of the book-views collection
	// (ECMA-376 §2.4.29, record ID 0x0188).
	BookViewsEnd = 0x0188

	// Sheets marks the start of the sheet-list collection
	// (ECMA-376 §2.4.723, record ID 0x018F).
	Sheets = 0x018F

	// SheetsEnd marks the end of the sheet-list collection
	// (ECMA-376 §2.4.724, record ID 0x0190).
	SheetsEnd = 0x0190

	// WorkbookPr carries workbook-level properties such as the date system
	// (ECMA-376 §2.4.822, record ID 0x0199).
	WorkbookPr = 0x0199

	// Sheet records a single worksheet entry (name, relationship ID, sheet ID)
	// inside the Sheets collection (ECMA-376 §2.4.720, record ID 0x019C).
	Sheet = 0x019C

	// CalcPr carries calculation properties for the workbook
	// (ECMA-376 §2.4.42, record ID 0x019D).
	CalcPr = 0x019D

	// WorkbookView records global workbook-view settings (active sheet, scroll
	// position, zoom) (ECMA-376 §2.4.825, record ID 0x019E).
	WorkbookView = 0x019E

	// ExternalReferences marks the start of the external-references collection
	// (ECMA-376 §2.4.148, record ID 0x02E1).
	ExternalReferences = 0x02E1

	// ExternalReferencesEnd marks the end of the external-references collection
	// (ECMA-376 §2.4.149, record ID 0x02E2).
	ExternalReferencesEnd = 0x02E2

	// ExternalReference records a single external workbook reference
	// (ECMA-376 §2.4.147, record ID 0x02E3).
	ExternalReference = 0x02E3

	// WebPublishing carries web-publishing properties for the workbook
	// (ECMA-376 §2.4.803, record ID 0x04A9).
	WebPublishing = 0x04A9

	// ── Worksheet records ─────────────────────────────────────────────────────

	// Row marks the start of a row of cells in the worksheet; the payload
	// contains the 0-based row index (ECMA-376 §2.4.657, record ID 0x0000).
	Row = 0x0000

	// Blank records a blank (empty) cell with a style index but no value
	// (ECMA-376 §2.4.20, record ID 0x0001).
	Blank = 0x0001

	// Num records a cell whose value is a packed numeric (RKNumber) type
	// (ECMA-376 §2.4.565, record ID 0x0002).
	Num = 0x0002

	// BoolErr records a cell whose value is an error code
	// (ECMA-376 §2.4.30, record ID 0x0003).
	BoolErr = 0x0003

	// Bool records a cell whose value is a boolean
	// (ECMA-376 §2.4.27, record ID 0x0004).
	Bool = 0x0004

	// Float records a cell whose value is an IEEE-754 double-precision float
	// (ECMA-376 §2.4.175, record ID 0x0005).
	Float = 0x0005

	// String records a cell whose value is an index into the shared-string table
	// (ECMA-376 §2.4.752, record ID 0x0007).
	String = 0x0007

	// FormulaString records a formula cell whose cached result is a string
	// (ECMA-376 §2.4.198, record ID 0x0008).
	FormulaString = 0x0008

	// FormulaFloat records a formula cell whose cached result is a
	// double-precision float (ECMA-376 §2.4.196, record ID 0x0009).
	FormulaFloat = 0x0009

	// FormulaBool records a formula cell whose cached result is a boolean
	// (ECMA-376 §2.4.194, record ID 0x000A).
	FormulaBool = 0x000A

	// FormulaBoolErr records a formula cell whose cached result is an error
	// (ECMA-376 §2.4.195, record ID 0x000B).
	FormulaBoolErr = 0x000B

	// Col records a column-definition entry (width, style, range)
	// (ECMA-376 §2.4.60, record ID 0x003C).
	Col = 0x003C

	// Worksheet marks the start of a worksheet binary part
	// (ECMA-376 §2.4.816, record ID 0x0181).
	Worksheet = 0x0181

	// WorksheetEnd marks the end of a worksheet binary part
	// (ECMA-376 §2.4.817, record ID 0x0182).
	WorksheetEnd = 0x0182

	// SheetViews marks the start of the sheet-views collection
	// (ECMA-376 §2.4.736, record ID 0x0185).
	SheetViews = 0x0185

	// SheetViewsEnd marks the end of the sheet-views collection
	// (ECMA-376 §2.4.737, record ID 0x0186).
	SheetViewsEnd = 0x0186

	// SheetView marks the start of a single sheet-view record
	// (ECMA-376 §2.4.734, record ID 0x0189).
	SheetView = 0x0189

	// SheetViewEnd marks the end of a single sheet-view record
	// (ECMA-376 §2.4.735, record ID 0x018A).
	SheetViewEnd = 0x018A

	// SheetData marks the start of the cell-data section of a worksheet
	// (ECMA-376 §2.4.710, record ID 0x0191).
	SheetData = 0x0191

	// SheetDataEnd marks the end of the cell-data section of a worksheet
	// (ECMA-376 §2.4.711, record ID 0x0192).
	SheetDataEnd = 0x0192

	// SheetPr carries sheet-level properties (tab colour, outline settings, etc.)
	// (ECMA-376 §2.4.715, record ID 0x0193).
	SheetPr = 0x0193

	// Dimension records the used-range of the worksheet (first/last row and
	// column, 0-based) (ECMA-376 §2.4.114, record ID 0x0194).
	Dimension = 0x0194

	// Selection records the currently selected cell range in a sheet view
	// (ECMA-376 §2.4.690, record ID 0x0198).
	Selection = 0x0198

	// Cols marks the start of the column-definitions collection
	// (ECMA-376 §2.4.61, record ID 0x0386).
	Cols = 0x0386

	// ColsEnd marks the end of the column-definitions collection
	// (ECMA-376 §2.4.62, record ID 0x0387).
	ColsEnd = 0x0387

	// ConditionalFormatting marks the start of a conditional-formatting block
	// (ECMA-376 §2.4.80, record ID 0x03CD).
	ConditionalFormatting = 0x03CD

	// ConditionalFormattingEnd marks the end of a conditional-formatting block
	// (ECMA-376 §2.4.81, record ID 0x03CE).
	ConditionalFormattingEnd = 0x03CE

	// CfRule marks the start of a single conditional-formatting rule
	// (ECMA-376 §2.4.50, record ID 0x03CF).
	CfRule = 0x03CF

	// CfRuleEnd marks the end of a single conditional-formatting rule
	// (ECMA-376 §2.4.51, record ID 0x03D0).
	CfRuleEnd = 0x03D0

	// IconSet marks the start of an icon-set conditional-formatting rule
	// (ECMA-376 §2.4.255, record ID 0x03D1).
	IconSet = 0x03D1

	// IconSetEnd marks the end of an icon-set conditional-formatting rule
	// (ECMA-376 §2.4.256, record ID 0x03D2).
	IconSetEnd = 0x03D2

	// DataBar marks the start of a data-bar conditional-formatting rule
	// (ECMA-376 §2.4.106, record ID 0x03D3).
	DataBar = 0x03D3

	// DataBarEnd marks the end of a data-bar conditional-formatting rule
	// (ECMA-376 §2.4.107, record ID 0x03D4).
	DataBarEnd = 0x03D4

	// ColorScale marks the start of a color-scale conditional-formatting rule
	// (ECMA-376 §2.4.67, record ID 0x03D5).
	ColorScale = 0x03D5

	// ColorScaleEnd marks the end of a color-scale conditional-formatting rule
	// (ECMA-376 §2.4.68, record ID 0x03D6).
	ColorScaleEnd = 0x03D6

	// Cfvo records a conditional-formatting value object (threshold or formula)
	// (ECMA-376 §2.4.52, record ID 0x03D7).
	Cfvo = 0x03D7

	// PageMargins records the page-margin settings for printing
	// (ECMA-376 §2.4.579, record ID 0x03DC).
	PageMargins = 0x03DC

	// PrintOptions records print-related options (gridlines, headings, etc.)
	// (ECMA-376 §2.4.614, record ID 0x03DD).
	PrintOptions = 0x03DD

	// PageSetup records page-setup properties (paper size, orientation, scale)
	// (ECMA-376 §2.4.580, record ID 0x03DE).
	PageSetup = 0x03DE

	// HeaderFooter records header and footer strings for printed pages
	// (ECMA-376 §2.4.226, record ID 0x03DF).
	HeaderFooter = 0x03DF

	// SheetFormatPr records default row-height and column-width properties
	// (ECMA-376 §2.4.712, record ID 0x03E5).
	SheetFormatPr = 0x03E5

	// Hyperlink records a hyperlink associated with a cell range; the payload
	// includes the range coordinates and the relationship ID
	// (ECMA-376 §2.4.247, record ID 0x03EE).
	Hyperlink = 0x03EE

	// Drawing associates an embedded drawing part with the worksheet
	// (ECMA-376 §2.4.122, record ID 0x04A6).
	Drawing = 0x04A6

	// LegacyDrawing associates a legacy VML drawing part with the worksheet
	// (ECMA-376 §2.4.326, record ID 0x04A7).
	LegacyDrawing = 0x04A7

	// Color records a color value used within a conditional-formatting rule
	// (ECMA-376 §2.4.65, record ID 0x04B4).
	Color = 0x04B4

	// OleObjects marks the start of the OLE-objects collection in a worksheet
	// (ECMA-376 §2.4.572, record ID 0x04FE).
	OleObjects = 0x04FE

	// OleObject records a single embedded OLE object
	// (ECMA-376 §2.4.571, record ID 0x04FF).
	OleObject = 0x04FF

	// OleObjectsEnd marks the end of the OLE-objects collection
	// (ECMA-376 §2.4.573, record ID 0x0580).
	OleObjectsEnd = 0x0580

	// TableParts marks the start of the table-parts collection in a worksheet
	// (ECMA-376 §2.4.776, record ID 0x0594).
	TableParts = 0x0594

	// TablePart references a single table part by relationship ID
	// (ECMA-376 §2.4.775, record ID 0x0595).
	TablePart = 0x0595

	// TablePartsEnd marks the end of the table-parts collection
	// (ECMA-376 §2.4.777, record ID 0x0596).
	TablePartsEnd = 0x0596

	// ── SharedStrings records ─────────────────────────────────────────────────

	// Si records a single shared-string item (rich-text or plain text) in the
	// shared-string table (ECMA-376 §2.4.741, record ID 0x0013).
	Si = 0x0013

	// Sst marks the start of the shared-string table stream; the payload
	// includes the total count and unique-string count
	// (ECMA-376 §2.4.753, record ID 0x019F).
	Sst = 0x019F

	// SstEnd marks the end of the shared-string table stream
	// (ECMA-376 §2.4.754, record ID 0x01A0).
	SstEnd = 0x01A0

	// ── Styles records ────────────────────────────────────────────────────────

	// Font records a single font definition in the styles part
	// (ECMA-376 §2.4.181, record ID 0x002B).
	Font = 0x002B

	// NumFmt records a single number-format entry (numFmtId + format string)
	// in the styles part (MS-XLSB §2.4.697 / ECMA-376 §2.4.497, record ID 0x002C).
	NumFmt = 0x002C

	// Fill records a single fill (pattern or gradient) definition
	// (ECMA-376 §2.4.167, record ID 0x002D).
	Fill = 0x002D

	// Border records a single border definition
	// (ECMA-376 §2.4.24, record ID 0x002E).
	Border = 0x002E

	// Xf records a single cell format (XF) entry that combines font, fill,
	// border, and number-format indices (ECMA-376 §2.4.830, record ID 0x002F).
	Xf = 0x002F

	// CellStyle records a named cell style (e.g. "Normal", "Heading 1")
	// (ECMA-376 §2.4.46, record ID 0x0030).
	CellStyle = 0x0030

	// StyleSheet marks the start of the styles-part stream
	// (ECMA-376 §2.4.756, record ID 0x0296).
	StyleSheet = 0x0296

	// StyleSheetEnd marks the end of the styles-part stream
	// (ECMA-376 §2.4.757, record ID 0x0297).
	StyleSheetEnd = 0x0297

	// Colors marks the start of the color-palette collection in the styles part
	// (ECMA-376 §2.4.66, record ID 0x03D9).
	Colors = 0x03D9

	// ColorsEnd marks the end of the color-palette collection
	// (ECMA-376 §2.4.69, record ID 0x03DA).
	ColorsEnd = 0x03DA

	// Dxfs marks the start of the differential-formatting collection
	// (ECMA-376 §2.4.131, record ID 0x03F9).
	Dxfs = 0x03F9

	// DxfsEnd marks the end of the differential-formatting collection
	// (ECMA-376 §2.4.132, record ID 0x03FA).
	DxfsEnd = 0x03FA

	// TableStyles marks the start of the table-styles collection
	// (ECMA-376 §2.4.778, record ID 0x03FC).
	TableStyles = 0x03FC

	// TableStylesEnd marks the end of the table-styles collection
	// (ECMA-376 §2.4.779, record ID 0x03FD).
	TableStylesEnd = 0x03FD

	// Fills marks the start of the fills collection in the styles part
	// (ECMA-376 §2.4.168, record ID 0x04DB).
	Fills = 0x04DB

	// FillsEnd marks the end of the fills collection
	// (ECMA-376 §2.4.169, record ID 0x04DC).
	FillsEnd = 0x04DC

	// Fonts marks the start of the fonts collection in the styles part
	// (ECMA-376 §2.4.182, record ID 0x04E3).
	Fonts = 0x04E3

	// FontsEnd marks the end of the fonts collection
	// (ECMA-376 §2.4.183, record ID 0x04E4).
	FontsEnd = 0x04E4

	// Borders marks the start of the borders collection in the styles part
	// (ECMA-376 §2.4.25, record ID 0x04E5).
	Borders = 0x04E5

	// BordersEnd marks the end of the borders collection
	// (ECMA-376 §2.4.26, record ID 0x04E6).
	BordersEnd = 0x04E6

	// NumFmts marks the start of the number-formats collection in the styles part
	// (MS-XLSB §2.4.698 / ECMA-376 §2.4.498, record ID 0x02C6).
	NumFmts = 0x02C6

	// NumFmtsEnd marks the end of the number-formats collection
	// (MS-XLSB §2.4.699 / ECMA-376 §2.4.499, record ID 0x02C8).
	NumFmtsEnd = 0x02C8

	// CellXfs marks the start of the cell-XF collection (applied cell formats)
	// (ECMA-376 §2.4.48, record ID 0x04E9).
	CellXfs = 0x04E9

	// CellXfsEnd marks the end of the cell-XF collection
	// (ECMA-376 §2.4.49, record ID 0x04EA).
	CellXfsEnd = 0x04EA

	// CellStyles marks the start of the named cell-styles collection
	// (ECMA-376 §2.4.47, record ID 0x04EB).
	CellStyles = 0x04EB

	// CellStylesEnd marks the end of the named cell-styles collection
	// (ECMA-376 §2.4.47, record ID 0x04EC).
	CellStylesEnd = 0x04EC

	// CellStyleXfs marks the start of the cell-style XF collection (master
	// formats that named styles inherit from) (ECMA-376 §2.4.45, record ID 0x04F2).
	CellStyleXfs = 0x04F2

	// CellStyleXfsEnd marks the end of the cell-style XF collection
	// (ECMA-376 §2.4.45, record ID 0x04F3).
	CellStyleXfsEnd = 0x04F3

	// ── Comment records ───────────────────────────────────────────────────────

	// Comments marks the start of the comments collection for a worksheet
	// (ECMA-376 §2.4.77, record ID 0x04F4).
	Comments = 0x04F4

	// CommentsEnd marks the end of the comments collection
	// (ECMA-376 §2.4.78, record ID 0x04F5).
	CommentsEnd = 0x04F5

	// Authors marks the start of the comment-authors list
	// (ECMA-376 §2.4.12, record ID 0x04F6).
	Authors = 0x04F6

	// AuthorsEnd marks the end of the comment-authors list
	// (ECMA-376 §2.4.13, record ID 0x04F7).
	AuthorsEnd = 0x04F7

	// Author records a single comment-author name
	// (ECMA-376 §2.4.11, record ID 0x04F8).
	Author = 0x04F8

	// CommentList marks the start of the list of comment records
	// (ECMA-376 §2.4.75, record ID 0x04F9).
	CommentList = 0x04F9

	// CommentListEnd marks the end of the list of comment records
	// (ECMA-376 §2.4.76, record ID 0x04FA).
	CommentListEnd = 0x04FA

	// Comment marks the start of a single cell comment
	// (ECMA-376 §2.4.73, record ID 0x04FB).
	Comment = 0x04FB

	// CommentEnd marks the end of a single cell comment
	// (ECMA-376 §2.4.74, record ID 0x04FC).
	CommentEnd = 0x04FC

	// Text records the rich-text content of a comment
	// (ECMA-376 §2.4.787, record ID 0x04FD).
	Text = 0x04FD

	// ── Table records ─────────────────────────────────────────────────────────

	// MergeCells marks the start of the merged-cells collection in a worksheet
	// (ECMA-376 §2.4.393, record ID 0x01B1).
	MergeCells = 0x01B1

	// MergeCellsEnd marks the end of the merged-cells collection
	// (ECMA-376 §2.4.394, record ID 0x01B2).
	MergeCellsEnd = 0x01B2

	// MergeCell records a single merged-cell range (first/last row and column)
	// (ECMA-376 §2.4.392, record ID 0x01B0).
	MergeCell = 0x01B0

	// AutoFilter marks the start of an auto-filter definition
	// (ECMA-376 §2.4.14, record ID 0x01A1).
	AutoFilter = 0x01A1

	// AutoFilterEnd marks the end of an auto-filter definition
	// (ECMA-376 §2.4.15, record ID 0x01A2).
	AutoFilterEnd = 0x01A2

	// FilterColumn marks the start of a filter-column definition
	// (ECMA-376 §2.4.170, record ID 0x01A3).
	FilterColumn = 0x01A3

	// FilterColumnEnd marks the end of a filter-column definition
	// (ECMA-376 §2.4.171, record ID 0x01A4).
	FilterColumnEnd = 0x01A4

	// Filters marks the start of the filter-criteria list for a column
	// (ECMA-376 §2.4.172, record ID 0x01A5).
	Filters = 0x01A5

	// FiltersEnd marks the end of the filter-criteria list
	// (ECMA-376 §2.4.173, record ID 0x01A6).
	FiltersEnd = 0x01A6

	// Filter records a single filter criterion (value or wildcard)
	// (ECMA-376 §2.4.169, record ID 0x01A7).
	Filter = 0x01A7

	// Table marks the start of a structured table definition
	// (ECMA-376 §2.4.760, record ID 0x02D7).
	Table = 0x02D7

	// TableEnd marks the end of a structured table definition
	// (ECMA-376 §2.4.761, record ID 0x02D8).
	TableEnd = 0x02D8

	// TableColumns marks the start of the table-columns collection
	// (ECMA-376 §2.4.762, record ID 0x02D9).
	TableColumns = 0x02D9

	// TableColumnsEnd marks the end of the table-columns collection
	// (ECMA-376 §2.4.763, record ID 0x02DA).
	TableColumnsEnd = 0x02DA

	// TableColumn marks the start of a single table-column definition
	// (ECMA-376 §2.4.764, record ID 0x02DB).
	TableColumn = 0x02DB

	// TableColumnEnd marks the end of a single table-column definition
	// (ECMA-376 §2.4.765, record ID 0x02DC).
	TableColumnEnd = 0x02DC

	// TableStyleInfo records the style applied to a structured table
	// (ECMA-376 §2.4.780, record ID 0x0481).
	TableStyleInfo = 0x0481

	// SortState marks the start of a sort-state definition for a table or range
	// (ECMA-376 §2.4.748, record ID 0x0492).
	SortState = 0x0492

	// SortCondition records a single sort-condition (column, order, data type)
	// (ECMA-376 §2.4.747, record ID 0x0494).
	SortCondition = 0x0494

	// SortStateEnd marks the end of a sort-state definition
	// (ECMA-376 §2.4.749, record ID 0x0495).
	SortStateEnd = 0x0495

	// ── QueryTable records ────────────────────────────────────────────────────

	// QueryTable marks the start of a query-table definition (external data)
	// (ECMA-376 §2.4.628, record ID 0x03BF).
	QueryTable = 0x03BF

	// QueryTableEnd marks the end of a query-table definition
	// (ECMA-376 §2.4.629, record ID 0x03C0).
	QueryTableEnd = 0x03C0

	// QueryTableRefresh marks the start of query-table refresh properties
	// (ECMA-376 §2.4.630, record ID 0x03C1).
	QueryTableRefresh = 0x03C1

	// QueryTableRefreshEnd marks the end of query-table refresh properties
	// (ECMA-376 §2.4.631, record ID 0x03C2).
	QueryTableRefreshEnd = 0x03C2

	// QueryTableFields marks the start of the query-table field list
	// (ECMA-376 §2.4.632, record ID 0x03C7).
	QueryTableFields = 0x03C7

	// QueryTableFieldsEnd marks the end of the query-table field list
	// (ECMA-376 §2.4.633, record ID 0x03C8).
	QueryTableFieldsEnd = 0x03C8

	// QueryTableField marks the start of a single query-table field definition
	// (ECMA-376 §2.4.634, record ID 0x03C9).
	QueryTableField = 0x03C9

	// QueryTableFieldEnd marks the end of a single query-table field definition
	// (ECMA-376 §2.4.635, record ID 0x03CA).
	QueryTableFieldEnd = 0x03CA

	// ── Connection records ────────────────────────────────────────────────────

	// Connections marks the start of the external data-connections collection
	// (ECMA-376 §2.4.83, record ID 0x03AD).
	Connections = 0x03AD

	// ConnectionsEnd marks the end of the external data-connections collection
	// (ECMA-376 §2.4.84, record ID 0x03AE).
	ConnectionsEnd = 0x03AE

	// Connection marks the start of a single external data-connection record
	// (ECMA-376 §2.4.82, record ID 0x01C9).
	Connection = 0x01C9

	// ConnectionEnd marks the end of a single external data-connection record
	// (ECMA-376 §2.4.82, record ID 0x01CA).
	ConnectionEnd = 0x01CA

	// DbPr records the database-connection properties (DSN, command text, etc.)
	// (ECMA-376 §2.4.110, record ID 0x01CB).
	DbPr = 0x01CB

	// DbPrEnd marks the end of the database-connection properties
	// (ECMA-376 §2.4.111, record ID 0x01CC).
	DbPrEnd = 0x01CC
)
