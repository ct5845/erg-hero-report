package main

import (
	"erg-hero-report/row_hero"
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

func ParseXLSX(filename string) []row_hero.Piece {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

	sheetNames := f.GetSheetList()

	fmt.Printf("Sheet names: %s\n", sheetNames)

	var pieces []row_hero.Piece

	// Iterate over each sheet
	for _, sheet := range sheetNames {
		rows, err := f.GetRows(sheet)
		if err != nil {
			log.Fatal(err)
		}

		piece := row_hero.ParseSheet(sheet, rows)

		pieces = append(pieces, piece)
	}

	return pieces
}
