package row_hero

import (
	"fmt"
	"strconv"
	"strings"
)

type Piece struct {
	Name  string
	Date  string `csv:"Date"`
	Piece string `csv:"Piece"`
	Rows  []Row
}

type Row struct {
	Stroke                 int     `csv:"Stroke #"`
	Time                   string  `csv:"Time"`
	Distance               int     `csv:"Distance"`
	Split                  string  `csv:"Split"`
	Watts                  int     `csv:"Watts"`
	StrokeRate             int     `csv:"Stroke Rate"`
	DragFactor             int     `csv:"Drag Factor"`
	StrokeLength           int     `csv:"Stroke Length"`
	ForceCurveSmoothness   float64 `csv:"Force Curve Smoothness"`
	ForceCurvePeakForcePos float64 `csv:"Force Curve Peak Force Position"`
	IsPrecisionStroke      string  `csv:"Is Precision Stroke"`
	Work                   float64 `csv:"Work (J)"`
	PeakDriveForce         float64 `csv:"Peak Drive Force (N)"`
	AvgDriveForce          float64 `csv:"Average Drive Force (N)"`
	DriveTime              float64 `csv:"Drive Time"`
	RecoveryTime           float64 `csv:"Recovery Time"`
	DistancePerStroke      float64 `csv:"Distance per Stroke"`
}

func parsePercentString(percentString string) (float64, error) {
	percentString = strings.TrimSuffix(percentString, "%")
	pctValue, err := strconv.ParseFloat(percentString, 64)
	if err != nil {
		return 0, fmt.Errorf("invalid percentage value: %s", percentString)
	}
	return pctValue / 100, nil // Convert to decimal
}

func isDataRow(row []string) error {
	if len(row) < 10 {
		return fmt.Errorf("isDataRow: warning - not a valid data row")
	}

	if !strings.HasSuffix(row[9], "%") {
		return fmt.Errorf("isDataRow: warning - peakForcePos not valid format")
	}

	return nil
}

func parseRow(row []string) (Row, error) {
	if err := isDataRow(row); err != nil {
		return Row{}, err
	}

	stroke, _ := strconv.Atoi(row[0])
	time := row[1]
	distance, _ := strconv.Atoi(row[2])
	split := row[3]
	watts, _ := strconv.Atoi(row[4])
	strokeRate, _ := strconv.Atoi(row[5])
	dragFactor, _ := strconv.Atoi(row[6])
	strokeLength, _ := strconv.Atoi(row[7])
	forceCurveSmoothness, err := strconv.ParseFloat(row[8], 64)
	if err != nil {
		return Row{}, err
	}
	forceCurvePeakForPos, err := parsePercentString(row[9])
	if err != nil {
		return Row{}, err
	}
	isPrecisionStroke := row[10]
	work, _ := strconv.ParseFloat(row[11], 64)
	peakDriveForce, _ := strconv.ParseFloat(row[12], 64)
	avgDriveForce, _ := strconv.ParseFloat(row[13], 64)
	driveTime, _ := strconv.ParseFloat(row[14], 64)
	recoveryTime, _ := strconv.ParseFloat(row[15], 64)
	distancePerStroke, _ := strconv.ParseFloat(row[16], 64)

	return Row{
		Stroke:                 stroke,
		Time:                   time,
		Distance:               distance,
		Split:                  split,
		Watts:                  watts,
		StrokeRate:             strokeRate,
		DragFactor:             dragFactor,
		StrokeLength:           strokeLength,
		ForceCurveSmoothness:   forceCurveSmoothness,
		ForceCurvePeakForcePos: forceCurvePeakForPos,
		IsPrecisionStroke:      isPrecisionStroke,
		Work:                   work,
		PeakDriveForce:         peakDriveForce,
		AvgDriveForce:          avgDriveForce,
		DriveTime:              driveTime,
		RecoveryTime:           recoveryTime,
		DistancePerStroke:      distancePerStroke,
	}, nil
}

func ParseSheet(name string, rows [][]string) Piece {
	piece := Piece{
		Name:  name,
		Date:  rows[1][0],
		Piece: rows[1][1],
	}

	skippedRows := 0

	var pieceRows []Row
	for _, row := range rows[3:] {
		if pieceRow, err := parseRow(row); err != nil {
			skippedRows += 1
		} else {
			pieceRows = append(pieceRows, pieceRow)
		}
	}

	if skippedRows >= 1 {
		fmt.Printf("Sheet: %s, Skipped %d rows", name, skippedRows)
	}

	piece.Rows = pieceRows

	return piece
}
