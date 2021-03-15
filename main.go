package main

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func main() {
	f := excelize.NewFile()
	// Create a new sheet.
	index := f.NewSheet("Sheet1")
	// Set value of a cell.
	f.SetCellValue("Sheet1", "A1", "BOOKNO")
	f.SetCellValue("Sheet1", "B1", "NAME")
	f.SetCellValue("Sheet1", "C1", "SPOUSE")
	f.SetCellValue("Sheet1", "D1", "MOBILE")
	f.SetCellValue("Sheet1", "E1", "WARD")
	f.SetCellValue("Sheet1", "F1", "UNION")
	f.SetCellValue("Sheet1", "G1", "THANA")
	f.SetCellValue("Sheet1", "H1", "DISTRICT")
	for i := 2; i < 10; i++ {
		f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), "700152445111")
		f.SetCellValue("Sheet1", "B"+strconv.Itoa(i), "KHALEDA")
		f.SetCellValue("Sheet1", "C"+strconv.Itoa(i), "AFZAL")
		f.SetCellValue("Sheet1", "D"+strconv.Itoa(i), "017177000000")
		f.SetCellValue("Sheet1", "E"+strconv.Itoa(i), "2")
		f.SetCellValue("Sheet1", "F"+strconv.Itoa(i), "BARIADHALA")
		f.SetCellValue("Sheet1", "G"+strconv.Itoa(i), "PALASHBARI")
		f.SetCellValue("Sheet1", "H"+strconv.Itoa(i), "DHAKA")
	}
	f.SetCellValue("Sheet1", "B", 100)
	// Set active sheet of the workbook.
	f.SetActiveSheet(index)
	// Save spreadsheet by the given path.
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
