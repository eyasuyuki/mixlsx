package internal

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"path/filepath"
)

type Xlsx struct {
		File *excelize.File
}

type Exel interface {
	OpenFile(filename string) (*Xlsx, error)
	SetSellValue(sheet, cell, value string) error
}

func Replace(file string, book *Book) {
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
				if err = f.SetCellStyle(w.Sheet, r.Range, r.Range, style); err != nil {
					fmt.Printf("Style error: %v\n", err)
				}
			}
		}
	}
	// generate new filename
	ext := filepath.Ext(book.Path)
	n := book.Path[:len(book.Path)-len(ext)] + ".xlsx"
	// save as
	if err = f.SaveAs(n); err != nil {
		fmt.Printf("File save error: %v\n", err)
	}
}
