package cmd

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var sdRecursive bool
var sdIgnoreCase bool
var sdSheetName string

var sdCmd = &cobra.Command{
	Use:   "sd [search] [replace] [file or directory]",
	Short: "Excelファイルまたはディレクトリ内の文字列を置換します。",
	Long:  `指定されたExcelファイルまたはディレクトリ内の全シートの全セルで、検索文字列を置換文字列に置換します。`,
	Args:  cobra.ExactArgs(3),
	Run: func(cmd *cobra.Command, args []string) {
		search := args[0]
		replace := args[1]
		path := args[2]

		var re *regexp.Regexp
		var err error
		if sdIgnoreCase {
			re, err = regexp.Compile("(?i)" + search)
		} else {
			re, err = regexp.Compile(search)
		}
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error compiling regex: %v\n", err)
			os.Exit(1)
		}

		info, err := os.Stat(path)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error accessing path %s: %v\n", path, err)
			os.Exit(1)
		}

		var filesToProcess []string
		if info.IsDir() {
			if sdRecursive {
				err := filepath.Walk(path, func(p string, info os.FileInfo, err error) error {
					if err != nil {
						return err
					}
					if !info.IsDir() {
						ext := strings.ToLower(filepath.Ext(p))
						if ext == ".xlsx" || ext == ".xlsm" {
							filesToProcess = append(filesToProcess, p)
						}
					}
					return nil
				})
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error walking directory %s: %v\n", path, err)
					os.Exit(1)
				}
			} else {
				entries, err := os.ReadDir(path)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error reading directory %s: %v\n", path, err)
					os.Exit(1)
				}
				for _, entry := range entries {
					if !entry.IsDir() {
						ext := strings.ToLower(filepath.Ext(entry.Name()))
						if ext == ".xlsx" || ext == ".xlsm" {
							filesToProcess = append(filesToProcess, filepath.Join(path, entry.Name()))
						}
					}
				}
			}
		} else {
			filesToProcess = append(filesToProcess, path)
		}

		for _, filePath := range filesToProcess {
			f, err := excelize.OpenFile(filePath)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", filePath, err)
				continue
			}

			var sheetsToProcess []string
			if sdSheetName != "" {
				// Check if sheet exists
				isSheetFound := false
				for _, s := range f.GetSheetList() {
					if s == sdSheetName {
						isSheetFound = true
						break
					}
				}

				if isSheetFound {
					sheetsToProcess = append(sheetsToProcess, sdSheetName)
				} else {
					fmt.Fprintf(os.Stderr, "Error: Sheet '%s' not found in file %s\n", sdSheetName, filePath)
					continue
				}
			} else {
				sheetsToProcess = f.GetSheetList()
			}

			for _, sheetName := range sheetsToProcess {
				rows, err := f.GetRows(sheetName)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error getting rows from sheet %s: %v\n", sheetName, err)
					continue
				}

				for r, row := range rows {
					isRowModified := false
					newRowValues := make([]string, len(row))
					copy(newRowValues, row)

					for c, cellValue := range row {
						cellName, err := excelize.CoordinatesToCellName(c+1, r+1)
						if err != nil {
							fmt.Fprintf(os.Stderr, "Error converting coordinates to cell name for sheet %s, row %d, col %d: %v\n", sheetName, r+1, c+1, err)
							continue
						}

						formula, err := f.GetCellFormula(sheetName, cellName)
						if err == nil && formula != "" {
							// This cell has a formula
							if re.MatchString(formula) {
								isRowModified = true
								newFormula := re.ReplaceAllString(formula, replace)
								if err := f.SetCellFormula(sheetName, cellName, newFormula); err != nil {
									fmt.Fprintf(os.Stderr, "Error setting cell formula for %s on sheet %s: %v\n", cellName, sheetName, err)
									continue
								}

								f.WorkBook.CalcPr.FullCalcOnLoad = true
							}
						} else {
							// This cell does not have a formula, or we couldn't get it.
							// Operate on the cell value.
							if re.MatchString(cellValue) {
								isRowModified = true
								newCellValue := re.ReplaceAllString(cellValue, replace)
								newRowValues[c] = newCellValue
								if err := f.SetCellValue(sheetName, cellName, newCellValue); err != nil {
									fmt.Fprintf(os.Stderr, "Error setting cell value for %s on sheet %s: %v\n", cellName, sheetName, err)
									continue
								}
							}
						}
					}

					if isRowModified {
						fmt.Printf("[Replaced] %s: %s: Row %d\n", filePath, sheetName, r+1)
					}
				}
			}

			if err := f.Save(); err != nil {
				fmt.Fprintf(os.Stderr, "Error saving file %s: %v\n", filePath, err)
				continue
			}

			if err := f.Close(); err != nil {
				fmt.Fprintln(os.Stderr, err)
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(sdCmd)
	sdCmd.Flags().BoolVarP(&sdRecursive, "recursive", "r", false, "サブディレクトリまで再帰的に処理します。")
	sdCmd.Flags().BoolVarP(&sdIgnoreCase, "ignore-case", "i", false, "検索時に大文字小文字を区別しません。")
	sdCmd.Flags().StringVarP(&sdSheetName, "name", "n", "", "指定したシートのセルの値を置換します。")
}
