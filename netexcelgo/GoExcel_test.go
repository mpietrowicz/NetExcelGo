package netexcelgo

import (
	"os"
	"testing"
)

func TestBuildExcel(t *testing.T) {
	testlist := ExcelSheetList{
		ExcelSheets: []ExcelSheet{
			{
				Name: "Sheet1",
				Data: []ExcelSheetRow{
					{
						Data: []string{"1", "2", "3"},
					},
					{
						Data: []string{"4", "5", "6"},
					},
				},
			},
			{
				Name: "Sheet2",
				Data: []ExcelSheetRow{
					{
						Data: []string{"7", "8", "9"},
					},
					{
						Data: []string{"10", "11", "12"},
					},
				},
			},
		},
		FileName: "test",
	}

	for i := 3; i < 100; i++ {
	}

	BuildExcel(testlist)

	if _, err := os.Stat(testlist.FileName + ".xlsx"); err == nil {
		t.Logf("File exists\n")
		os.Remove(testlist.FileName + ".xlsx")
	} else {
		t.Errorf("File does not exist\n")
	}
}
