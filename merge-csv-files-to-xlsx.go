package main

import (
	"encoding/csv"
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
)

// Merge multiple CSV files into a single XLSX/Excel file
// Usage: merge-csv-files-to-xlsx output.xlsx Sheet1 sheet1.csv Sheet2 sheet2.csv ...

func main() {
	prog := os.Args[0]

	if len(os.Args) < 4 || len(os.Args)%2 != 0 {
		fmt.Printf("usage: %s <output> <sheet name> <csv path> [[<sheet name> <csv path>] ...]\n", prog)
		os.Exit(1)
	}

	filePath := os.Args[1]
	sheets := os.Args[2:]

	fmt.Printf("filePath: %s\n", filePath)
	fmt.Printf("sheets: %v\n", sheets)

	// Convert CSV files to XLSX sheets
	file := xlsx.NewFile()

	for i, sheetName := range sheets {
		if i%2 != 0 {
			continue
		}

		fmt.Printf("i: %d, sheetName: %s\n", i, sheetName)
		csvPath := sheets[i+1]
		lines, err := parseCSV(csvPath)

		if err != nil {
			fmt.Printf("Error parsing csv: %e\n", err)
		}

		addSheet(file, sheetName, lines)
	}

	err := file.Save(filePath)

	if err != nil {
		fmt.Printf(err.Error())
	}

	fmt.Printf("Excel file saved to: %s\n", filePath)
}

func parseCSV(path string) ([][]string, error) {
	f, err := os.Open(path)

	if err != nil {
		return nil, err
	}

	defer f.Close() // this needs to be after the err check

	lines, err := csv.NewReader(f).ReadAll()

	if err != nil {
		return nil, err
	}

	return lines, nil
}

func addSheet(file *xlsx.File, sheetName string, lines [][]string) error {
	sheet, err := file.AddSheet(sheetName)

	if err != nil {
		return err
	}

	for _, line := range lines {
		row := sheet.AddRow()

		for _, col := range line {
			cell := row.AddCell()

			cell.SetValue(col)
		}
	}

	return nil
}
