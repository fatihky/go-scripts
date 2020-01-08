package main

import (
	"encoding/csv"
	"fmt"
	"github.com/tealeg/xlsx"
	"io"
	"log"
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

	// Create a new Excel file
	file := xlsx.NewFile()

	// Convert CSV files to XLSX sheets
	for i, sheetName := range sheets {
		if i%2 != 0 {
			continue
		}

		fmt.Printf("i: %d, sheetName: %s\n", i, sheetName)
		csvPath := sheets[i+1]

		addSheetFromCsv(file, sheetName, csvPath)
	}

	// Save Excel file
	err := file.Save(filePath)

	// Check errors
	if err != nil {
		fmt.Printf(err.Error())
	}

	fmt.Printf("Excel file saved to: %s\n", filePath)
}

func addSheetFromCsv(file *xlsx.File, sheetName string, csvPath string) {
	// Open CSV file
	f, err := os.Open(csvPath)

	if err != nil {
		log.Fatal(err)
	}

	// Create CSV reader
	reader := csv.NewReader(f)

	// Create sheet
	sheet, err := file.AddSheet(sheetName)

	if err != nil {
		log.Fatal(err)
	}

	// Walk over the lines
	for {
		line, err := reader.Read()

		if err == io.EOF {
			break
		} else if err != nil {
			log.Fatal(err)
		}

		addRow(sheet, line)
	}
}

func addRow(sheet *xlsx.Sheet, line []string) {
	row := sheet.AddRow()

	for _, col := range line {
		cell := row.AddCell()

		cell.SetValue(col)
	}
}
