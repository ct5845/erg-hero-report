package main

import (
	"erg-hero-report/row_hero"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

func checkOutputFilepath(fp string) {
	dir := filepath.Dir(fp)

	if _, err := os.Stat(dir); os.IsNotExist(err) {
		err := os.MkdirAll(dir, os.ModePerm)

		if err != nil {
			log.Fatalf("error creating directory: %v", err)
		}
	}
}

func outputPiecesSummaryDataSheet(f *excelize.File, pieces []row_hero.Piece) {
	sheet := "Summary"
	f.NewSheet(sheet)

	headers := []string{
		"Athlete",
		"ForceCurvePeakForcePos",
		"StrokeLength",
		"ForceCurveSmoothness",
		"IsPrecisionStroke",
	}

	for colIdx, header := range headers {
		cell := fmt.Sprintf("%s1", string(rune('A'+colIdx))) // A1, B1, etc.
		f.SetCellValue(sheet, cell, header)
	}

	for rowIdx, piece := range pieces {
		f.SetCellValue(sheet, fmt.Sprintf("A%d", rowIdx+2), piece.Name)

		f.SetCellFormula(sheet, fmt.Sprintf("B%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!M:M)", piece.Name)) // ForceCurvePeakForcePos

		f.SetCellFormula(sheet, fmt.Sprintf("C%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!K:K)", piece.Name)) // StrokeLength

		f.SetCellFormula(sheet, fmt.Sprintf("D%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!L:L)", piece.Name)) // ForceCurveSmoothness

		f.SetCellFormula(sheet, fmt.Sprintf("E%d", rowIdx+2), fmt.Sprintf("COUNTIFS('%s'!M:M, \">=0.32\", '%s'!M:M, \"<=0.38\") / COUNTIF('%s'!M:M, \"<>\")", piece.Name, piece.Name, piece.Name))
	}
}

func outputPiecesAggregatesDataSheet(f *excelize.File, pieces []row_hero.Piece) {
	sheet := "Aggregated"
	f.NewSheet(sheet)

	headers := []string{
		"Athlete",
		"StrokeRate",
		"Watts",
		"DragFactor",
		"StrokeLength",
		"ForceCurveSmoothness",
		"ForceCurvePeakForcePos",
		"IsPrecisionStroke",
		"Work",
		"PeakDriveForce",
		"AvgDriveForce",
		"DriveTime",
		"RecoveryTime",
		"DistancePerStroke",
	}

	for colIdx, header := range headers {
		cell := fmt.Sprintf("%s1", string(rune('A'+colIdx))) // A1, B1, etc.
		f.SetCellValue(sheet, cell, header)
	}

	for rowIdx, piece := range pieces {
		f.SetCellValue(sheet, fmt.Sprintf("A%d", rowIdx+2), piece.Name)

		// Set formulas for averages
		f.SetCellFormula(sheet, fmt.Sprintf("B%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!I:I)", piece.Name)) // StrokeRate

		f.SetCellFormula(sheet, fmt.Sprintf("C%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!H:H)", piece.Name)) // Watts

		f.SetCellFormula(sheet, fmt.Sprintf("D%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!J:J)", piece.Name)) // DragFactor

		f.SetCellFormula(sheet, fmt.Sprintf("E%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!K:K)", piece.Name)) // StrokeLength

		f.SetCellFormula(sheet, fmt.Sprintf("F%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!L:L)", piece.Name)) // ForceCurveSmoothness

		f.SetCellFormula(sheet, fmt.Sprintf("G%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!M:M)", piece.Name)) // ForceCurvePeakForcePos

		f.SetCellFormula(sheet, fmt.Sprintf("H%d", rowIdx+2), fmt.Sprintf("COUNTIFS('%s'!M:M, \">=0.32\", '%s'!M:M, \"<=0.38\") / COUNTIF('%s'!M:M, \"<>\")", piece.Name, piece.Name, piece.Name))

		f.SetCellFormula(sheet, fmt.Sprintf("I%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!O:O)", piece.Name)) // Work

		f.SetCellFormula(sheet, fmt.Sprintf("J%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!P:P)", piece.Name)) // PeakDriveForce

		f.SetCellFormula(sheet, fmt.Sprintf("K%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!Q:Q)", piece.Name)) // AvgDriveForce

		f.SetCellFormula(sheet, fmt.Sprintf("L%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!R:R)", piece.Name)) // DriveTime

		f.SetCellFormula(sheet, fmt.Sprintf("M%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!S:S)", piece.Name)) // RecoveryTime

		f.SetCellFormula(sheet, fmt.Sprintf("N%d", rowIdx+2),
			fmt.Sprintf("AVERAGE('%s'!T:T)", piece.Name)) // DistancePerStroke
	}
}

func outputPieceDataSheet(f *excelize.File, piece row_hero.Piece) {
	// Create a new sheet with the name
	f.NewSheet(piece.Name)

	// Set headers in the first row
	headers := []string{
		"Name",
		"Date",
		"Piece",
		"Stroke",
		"Time",
		"Distance",
		"Split",
		"Watts",
		"StrokeRate",
		"DragFactor",
		"StrokeLength",
		"ForceCurveSmoothness",
		"ForceCurvePeakForcePos",
		"IsPrecisionStroke",
		"Work",
		"PeakDriveForce",
		"AvgDriveForce",
		"DriveTime",
		"RecoveryTime",
		"DistancePerStroke",
	}

	for colIdx, header := range headers {
		cell := fmt.Sprintf("%s1", string(rune('A'+colIdx))) // A1, B1, etc.
		f.SetCellValue(piece.Name, cell, header)
	}

	// Fill in the rows with data
	for rowIdx, row := range piece.Rows {
		f.SetCellValue(piece.Name, fmt.Sprintf("A%d", rowIdx+2), piece.Name)
		f.SetCellValue(piece.Name, fmt.Sprintf("B%d", rowIdx+2), piece.Date)
		f.SetCellValue(piece.Name, fmt.Sprintf("C%d", rowIdx+2), piece.Piece)
		f.SetCellValue(piece.Name, fmt.Sprintf("D%d", rowIdx+2), row.Stroke)
		f.SetCellValue(piece.Name, fmt.Sprintf("E%d", rowIdx+2), row.Time)
		f.SetCellValue(piece.Name, fmt.Sprintf("F%d", rowIdx+2), row.Distance)
		f.SetCellValue(piece.Name, fmt.Sprintf("G%d", rowIdx+2), row.Split)
		f.SetCellValue(piece.Name, fmt.Sprintf("H%d", rowIdx+2), row.Watts)
		f.SetCellValue(piece.Name, fmt.Sprintf("I%d", rowIdx+2), row.StrokeRate)
		f.SetCellValue(piece.Name, fmt.Sprintf("J%d", rowIdx+2), row.DragFactor)
		f.SetCellValue(piece.Name, fmt.Sprintf("K%d", rowIdx+2), row.StrokeLength)
		f.SetCellValue(piece.Name, fmt.Sprintf("L%d", rowIdx+2), row.ForceCurveSmoothness)
		f.SetCellValue(piece.Name, fmt.Sprintf("M%d", rowIdx+2), row.ForceCurvePeakForcePos)
		f.SetCellValue(piece.Name, fmt.Sprintf("N%d", rowIdx+2), row.IsPrecisionStroke)
		f.SetCellValue(piece.Name, fmt.Sprintf("O%d", rowIdx+2), row.Work)
		f.SetCellValue(piece.Name, fmt.Sprintf("P%d", rowIdx+2), row.PeakDriveForce)
		f.SetCellValue(piece.Name, fmt.Sprintf("Q%d", rowIdx+2), row.AvgDriveForce)
		f.SetCellValue(piece.Name, fmt.Sprintf("R%d", rowIdx+2), row.DriveTime)
		f.SetCellValue(piece.Name, fmt.Sprintf("S%d", rowIdx+2), row.RecoveryTime)
		f.SetCellValue(piece.Name, fmt.Sprintf("T%d", rowIdx+2), row.DistancePerStroke)
	}
}

func OutputXLSX(pieces []row_hero.Piece, fp string) error {
	checkOutputFilepath(fp)

	// Create a new Excel file
	f := excelize.NewFile()

	outputPiecesSummaryDataSheet(f, pieces)

	outputPiecesAggregatesDataSheet(f, pieces)

	// Iterate over the sheets map
	for _, piece := range pieces {
		outputPieceDataSheet(f, piece)
	}

	f.DeleteSheet("Sheet1")

	// Save the file, overwriting if it exists
	err := f.SaveAs(fp)
	if err != nil {
		return fmt.Errorf("error saving file: %v", err)
	}

	fmt.Println("Excel output saved")

	return nil
}
