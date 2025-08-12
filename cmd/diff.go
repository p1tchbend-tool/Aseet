package cmd

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// findHeaderRow scans the first 100 rows to find the header row.
// A header row candidate has no empty cells between the first and last non-empty cell.
// The header row is the candidate with the most non-empty cells.
func findHeaderRow(f *excelize.File, sheetName string) ([]string, int, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, 0, err
	}

	var headerRow []string
	headerRowNum := 0
	maxCells := -1

	numRowsToCheck := 100
	if len(rows) < numRowsToCheck {
		numRowsToCheck = len(rows)
	}

	for i := 0; i < numRowsToCheck; i++ {
		row := rows[i]
		if len(row) == 0 {
			continue
		}

		firstCellIdx := -1
		lastCellIdx := -1

		// Find first and last non-empty cell indices
		for j, cell := range row {
			if cell != "" {
				if firstCellIdx == -1 {
					firstCellIdx = j
				}
				lastCellIdx = j
			}
		}

		if firstCellIdx == -1 { // Row is effectively empty
			continue
		}

		// Check for empty cells between first and last non-empty cells
		isCandidate := true
		sliceToCheck := row[firstCellIdx : lastCellIdx+1]
		for _, cell := range sliceToCheck {
			if cell == "" {
				isCandidate = false
				break
			}
		}

		if !isCandidate {
			continue
		}

		// This is a candidate, check if it's the best one so far
		cellCount := len(sliceToCheck)
		if cellCount > maxCells {
			maxCells = cellCount
			headerRow = row
			headerRowNum = i + 1 // 1-based index
		}
	}

	if headerRowNum == 0 {
		return nil, 0, nil // No suitable header found
	}

	return headerRow, headerRowNum, nil
}

var diffCmd = &cobra.Command{
	Use:   "diff [file1] [file2]",
	Short: "Show the difference in sheet names and header row content between two excel files",
	Long:  `Show the difference in sheet names and header row content between two excel files. This command compares header columns by their content, accounting for additions and deletions. The header row is identified by scanning the first 100 rows. Empty cells in header rows are ignored.`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		file1Path := args[0]
		file2Path := args[1]

		f1, err := excelize.OpenFile(file1Path)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", file1Path, err)
			os.Exit(1)
		}
		defer func() {
			if err := f1.Close(); err != nil {
				fmt.Fprintf(os.Stderr, "Error closing file %s: %v\n", file1Path, err)
			}
		}()

		f2, err := excelize.OpenFile(file2Path)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", file2Path, err)
			os.Exit(1)
		}
		defer func() {
			if err := f2.Close(); err != nil {
				fmt.Fprintf(os.Stderr, "Error closing file %s: %v\n", file2Path, err)
			}
		}()

		sheets1 := f1.GetSheetList()
		sheets2 := f2.GetSheetList()

		map1 := make(map[string]bool)
		for _, s := range sheets1 {
			map1[s] = true
		}

		map2 := make(map[string]bool)
		for _, s := range sheets2 {
			map2[s] = true
		}

		onlyIn1 := []string{}
		for _, s := range sheets1 {
			if !map2[s] {
				onlyIn1 = append(onlyIn1, s)
			}
		}

		onlyIn2 := []string{}
		for _, s := range sheets2 {
			if !map1[s] {
				onlyIn2 = append(onlyIn2, s)
			}
		}

		sheetNameDiff := false
		if len(onlyIn1) > 0 {
			sheetNameDiff = true
			fmt.Printf("Sheets only in %s:\n", file1Path)
			for _, s := range onlyIn1 {
				fmt.Printf("- %s\n", s)
			}
			fmt.Println()
		}

		if len(onlyIn2) > 0 {
			sheetNameDiff = true
			fmt.Printf("Sheets only in %s:\n", file2Path)
			for _, s := range onlyIn2 {
				fmt.Printf("- %s\n", s)
			}
			fmt.Println()
		}

		commonSheets := []string{}
		for _, s := range sheets1 {
			if map2[s] {
				commonSheets = append(commonSheets, s)
			}
		}

		contentDiff := false
		if len(commonSheets) > 0 && (len(onlyIn1) > 0 || len(onlyIn2) > 0) {
			fmt.Println("--- Comparing common sheets ---")
		}

		for _, sheet := range commonSheets {
			row1, rowNum1, err1 := findHeaderRow(f1, sheet)
			if err1 != nil {
				fmt.Fprintf(os.Stderr, "Error reading sheet %s from %s: %v\n", sheet, file1Path, err1)
				continue
			}

			row2, rowNum2, err2 := findHeaderRow(f2, sheet)
			if err2 != nil {
				fmt.Fprintf(os.Stderr, "Error reading sheet %s from %s: %v\n", sheet, file2Path, err2)
				continue
			}

			if row1 == nil && row2 == nil {
				continue
			}

			// Filter out empty cells for comparison
			var r1NonEmpty []string
			for _, cell := range row1 {
				if cell != "" {
					r1NonEmpty = append(r1NonEmpty, cell)
				}
			}
			var r2NonEmpty []string
			for _, cell := range row2 {
				if cell != "" {
					r2NonEmpty = append(r2NonEmpty, cell)
				}
			}

			map1 := make(map[string]int)
			for _, s := range r1NonEmpty {
				map1[s]++
			}

			map2 := make(map[string]int)
			for _, s := range r2NonEmpty {
				map2[s]++
			}

			var onlyInFile1, onlyInFile2 []string

			for val, count1 := range map1 {
				count2 := map2[val]
				if count1 > count2 {
					for i := 0; i < count1-count2; i++ {
						onlyInFile1 = append(onlyInFile1, val)
					}
				}
			}

			for val, count2 := range map2 {
				count1 := map1[val]
				if count2 > count1 {
					for i := 0; i < count2-count1; i++ {
						onlyInFile2 = append(onlyInFile2, val)
					}
				}
			}

			if len(onlyInFile1) > 0 || len(onlyInFile2) > 0 {
				contentDiff = true
				fmt.Printf("Sheet '%s': Header row content mismatch. Comparing %s (Row %d) and %s (Row %d):\n", sheet, file1Path, rowNum1, file2Path, rowNum2)
				if len(onlyInFile1) > 0 {
					fmt.Printf("  Columns only in %s:\n", file1Path)
					for _, s := range onlyInFile1 {
						fmt.Printf("    - %s\n", s)
					}
				}
				if len(onlyInFile2) > 0 {
					fmt.Printf("  Columns only in %s:\n", file2Path)
					for _, s := range onlyInFile2 {
						fmt.Printf("    - %s\n", s)
					}
				}
				fmt.Println()
			}
		}

		if !sheetNameDiff && !contentDiff {
			fmt.Println("The sheet names and header rows are identical in both files.")
		}
	},
}

func init() {
	rootCmd.AddCommand(diffCmd)
}
