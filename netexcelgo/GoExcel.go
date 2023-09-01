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
	count := 0
	for _, sheet := range list.ExcelSheets {
		count++
		countInString := fmt.Sprint(count)
		rowData := &sheet.Data
		if err := f.SetSheetRow(sheet.Name, "A"+countInString, &rowData); err != nil {
			fmt.Println(err)
			return
		}
	}

	f.SetActiveSheet(0)

	if err := f.SaveAs(list.FileName + ".xlsx"); err != nil {
		println(err.Error())
	}
}
