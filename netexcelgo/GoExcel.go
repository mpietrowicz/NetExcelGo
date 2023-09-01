package netexcelgo

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

func GetColumnMaps() map[int8]string {
	masp := make(map[int8]string)
	masp[0] = "A"
	masp[1] = "B"
	masp[2] = "C"
	masp[3] = "D"
	masp[4] = "E"
	masp[5] = "F"
	masp[6] = "G"
	masp[7] = "H"
	masp[8] = "I"
	masp[9] = "J"
	masp[10] = "K"
	masp[11] = "L"
	masp[12] = "M"
	masp[13] = "N"
	masp[14] = "O"
	masp[15] = "P"
	masp[16] = "Q"
	masp[17] = "R"
	masp[18] = "S"
	masp[19] = "T"
	masp[20] = "U"
	masp[21] = "V"
	masp[22] = "W"
	masp[23] = "X"
	masp[24] = "Y"
	masp[25] = "Z"

	return masp
}

func BuildColumnIndex(index int) string {
	msp := GetColumnMaps()
	convertedIndex := int8(index)
	if index < 26 {
		return msp[convertedIndex]
	} else {

		return msp[convertedIndex/26] + msp[convertedIndex%26]
	}

}

func BuildExcel(list *ExcelSheetList) {

	f := excelize.NewFile()

	for _, sheet := range list.ExcelSheets {
		if _, err := f.NewSheet(sheet.Name); err != nil {
			fmt.Println(err)
			return
		}
	}
	for _, sheet := range list.ExcelSheets {

		rowDataRow := &sheet.Data

		rowCount := 0
		for _, row := range *rowDataRow {
			rowData := &row.Data
			columnCount := 0
			for _, cellData := range *rowData {
				columnCount++
				cellLetter := BuildColumnIndex(rowCount)
				cellIndex := columnCount
				cell := cellLetter + fmt.Sprintf("%d", cellIndex)
				if err := f.SetCellValue(sheet.Name, cell, cellData); err != nil {
					fmt.Println(err)
					return
				}
			}
			fmt.Printf("%s  Data: %s \n", sheet.Name, rowData)
			rowCount++

		}
	}

	f.SetActiveSheet(0)

	if err := f.SaveAs(list.FileName + ".xlsx"); err != nil {
		println(err.Error())
	}
}
