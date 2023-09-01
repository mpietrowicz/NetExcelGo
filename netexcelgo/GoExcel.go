package netexcelgo

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func BuildExcel(list ExcelSheetList) {

	f := excelize.NewFile()

	for _, sheet := range list.ExcelSheets {
		f.NewSheet(sheet.Name)
	}

	for _, sheet := range list.ExcelSheets {

		rowDataRow := &sheet.Data
		count := 0
		for _, row := range *rowDataRow {
			rowData := row.Data
			count++
			countInString := fmt.Sprint(count)
			cell := "A" + countInString
			if err := f.SetSheetRow(sheet.Name, cell, &rowData); err != nil {
				fmt.Println(err)
				return
			}
		}
		count = 0
	}

	f.SetActiveSheet(0)

	if err := f.SaveAs(list.FileName + ".xlsx"); err != nil {
		println(err.Error())
	}
}
