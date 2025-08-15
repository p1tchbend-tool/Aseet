package cmd

import (
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"sort"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// findHeaderRow scans the first 100 rows to find the header row.
// A header row candidate has no empty cells between the first and last non-empty cell.
// The header row is the candidate with the most cells.
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

var openFiles bool

var diffCmd = &cobra.Command{
	Use:   "diff [file1] [file2]",
	Short: "Show the difference in sheet names and header row content between two excel files",
	Long:  `Show the difference in sheet names and header row content between two excel files. This command compares header columns by their content, accounting for additions and deletions. The header row is identified by scanning the first 100 rows. Empty cells in header rows are ignored. It also compares data rows cell by cell for columns with matching headers, prioritizing formulas over calculated values.`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		file1Path := args[0]
		file2Path := args[1]

		if filepath.Base(file1Path) == filepath.Base(file2Path) {
			cacheDir, err := os.UserCacheDir()
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error getting user cache dir: %v\n", err)
				os.Exit(1)
			}
			tempDir := filepath.Join(cacheDir, "asheet", "temp")
			if err := os.MkdirAll(tempDir, 0755); err != nil {
				fmt.Fprintf(os.Stderr, "Error creating temp dir: %v\n", err)
				os.Exit(1)
			}

			baseName := filepath.Base(file2Path)
			newFileName := "[REMOTE]" + baseName
			destPath := filepath.Join(tempDir, newFileName)

			sourceFile, err := os.Open(file2Path)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error opening source file for copy: %v\n", err)
				os.Exit(1)
			}
			defer sourceFile.Close()

			destFile, err := os.Create(destPath)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error creating destination file for copy: %v\n", err)
				os.Exit(1)
			}
			defer destFile.Close()

			_, err = io.Copy(destFile, sourceFile)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error copying file: %v\n", err)
				os.Exit(1)
			}

			file2Path = destPath
		}

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

		// Create maps for quick lookup
		sheetMap1 := make(map[string]bool)
		for _, s := range sheets1 {
			sheetMap1[s] = true
		}
		sheetMap2 := make(map[string]bool)
		for _, s := range sheets2 {
			sheetMap2[s] = true
		}

		// Create a combined, unique, sorted list of all sheet names
		allSheetsMap := make(map[string]bool)
		for _, s := range sheets1 {
			allSheetsMap[s] = true
		}
		for _, s := range sheets2 {
			allSheetsMap[s] = true
		}
		var allSheets []string
		for s := range allSheetsMap {
			allSheets = append(allSheets, s)
		}
		sort.Strings(allSheets)

		for _, sheet := range allSheets {
			_, existsIn1 := sheetMap1[sheet]
			_, existsIn2 := sheetMap2[sheet]

			if !existsIn1 {
				fmt.Printf("Sheet '%s' only in %s\n\n", sheet, file2Path)
				continue
			}
			if !existsIn2 {
				fmt.Printf("Sheet '%s' only in %s\n\n", sheet, file1Path)
				continue
			}

			// Common sheet logic
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

			// Compare headers
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

			// Headers are identical, compare data rows
			allRows1, err := f1.GetRows(sheet)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error reading all rows from sheet %s in %s: %v\n", sheet, file1Path, err)
				continue
			}
			allRows2, err := f2.GetRows(sheet)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error reading all rows from sheet %s in %s: %v\n", sheet, file2Path, err)
				continue
			}

			header1Indices := make(map[string]int)
			for i, h := range row1 {
				if h != "" {
					if _, exists := header1Indices[h]; !exists {
						header1Indices[h] = i
					}
				}
			}
			header2Indices := make(map[string]int)
			for i, h := range row2 {
				if h != "" {
					if _, exists := header2Indices[h]; !exists {
						header2Indices[h] = i
					}
				}
			}

			var commonHeaderSlice []string
			for h := range header1Indices {
				if _, ok := header2Indices[h]; ok {
					commonHeaderSlice = append(commonHeaderSlice, h)
				}
			}
			sort.Strings(commonHeaderSlice)

			maxRows := len(allRows1)
			if len(allRows2) > maxRows {
				maxRows = len(allRows2)
			}

			rowContentDiff := false
			for i := 0; i < maxRows; i++ {
				physicalRowNum := i + 1

				if (rowNum1 > 0 && physicalRowNum == rowNum1) || (rowNum2 > 0 && physicalRowNum == rowNum2) {
					continue
				}

				rowHasDiff := false
				var row1Vals, row2Vals []string

				for _, hName := range commonHeaderSlice {
					idx1 := header1Indices[hName]
					idx2 := header2Indices[hName]

					cellName1, _ := excelize.CoordinatesToCellName(idx1+1, physicalRowNum)
					val1, _ := f1.GetCellFormula(sheet, cellName1)
					if val1 == "" {
						val1, _ = f1.GetCellValue(sheet, cellName1)
					}

					cellName2, _ := excelize.CoordinatesToCellName(idx2+1, physicalRowNum)
					val2, _ := f2.GetCellFormula(sheet, cellName2)
					if val2 == "" {
						val2, _ = f2.GetCellValue(sheet, cellName2)
					}

					if val1 != val2 {
						rowHasDiff = true
					}
					row1Vals = append(row1Vals, val1)
					row2Vals = append(row2Vals, val2)
				}

				if rowHasDiff {
					if !rowContentDiff {
						fmt.Printf("Sheet '%s': Found differences in row content:\n", sheet)
						rowContentDiff = true
					}
					row1Str := strings.Join(row1Vals, ", ")
					row2Str := strings.Join(row2Vals, ", ")
					fmt.Printf("  - Row %d: [%s] vs [%s]\n", physicalRowNum, row1Str, row2Str)
				}
			}
			if rowContentDiff {
				fmt.Println()
			}
		}

		if openFiles {
			exec.Command("cmd", "/C", "start", file1Path).Start()
			exec.Command("cmd", "/C", "start", file2Path).Start()
		}
	},
}

func init() {
	rootCmd.AddCommand(diffCmd)
	diffCmd.Flags().BoolVarP(&openFiles, "open", "o", false, "最後に2つのファイルを関連付けられたアプリケーションで開きます。")
}
