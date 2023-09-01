package netexcelgo

type ExcelSheetRow struct {
	Data []string
}

// Description: ExcelSheet
type ExcelSheet struct {

	// Description: Name
	Name string
	Data []ExcelSheetRow
}

// Description: ExcelSheetList
// List of sheets
type ExcelSheetList struct {
	ExcelSheets []ExcelSheet
	FileName    string
}
