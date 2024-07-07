package internal

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

func Replace(file string, book Book) {
	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Printf("Excel file open failed: %v\n", err)
		return
	}
	for _, w := range book.Worksheets {
		for _, r := range w.Replacements {
			err = f.SetCellValue(w.Sheet, r.Range, r.Value)
			if err != nil {
				fmt.Printf("Cell value error: %v\n", err)
				return
			}
			if r.Color != "" {
				style, err := f.NewStyle(&excelize.Style{
					Font: &excelize.Font{
						Color: r.Color,
					},
				})
				if err != nil {
					fmt.Printf("Font error: %v\n", err)
				}
				err = f.SetCellStyle(w.Sheet, r.Range, r.Range, style)
			}
		}
	}
	// save as

}
