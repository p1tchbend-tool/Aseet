package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var sdCmd = &cobra.Command{
	Use:   "sd [search] [replace] [file]",
	Short: "Excelファイル内の文字列を置換します。",
	Long:  `指定されたExcelファイルの全シートの全セルで、検索文字列を置換文字列に置換します。`,
	Args:  cobra.ExactArgs(3),
	Run: func(cmd *cobra.Command, args []string) {
		search := args[0]
		replace := args[1]
		filePath := args[2]

		f, err := excelize.OpenFile(filePath)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", filePath, err)
			os.Exit(1)
		}
		defer func() {
			if err := f.Close(); err != nil {
				fmt.Fprintln(os.Stderr, err)
			}
		}()

		for _, sheetName := range f.GetSheetList() {
			rows, err := f.GetRows(sheetName)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error getting rows from sheet %s: %v\n", sheetName, err)
				continue
			}

			for r, row := range rows {
				rowModified := false
				newRowValues := make([]string, len(row))
				copy(newRowValues, row)

				for c, cellValue := range row {
					if strings.Contains(cellValue, search) {
						rowModified = true
						newCellValue := strings.ReplaceAll(cellValue, search, replace)
						newRowValues[c] = newCellValue
						cellName, err := excelize.CoordinatesToCellName(c+1, r+1)
						if err != nil {
							fmt.Fprintf(os.Stderr, "Error converting coordinates to cell name for sheet %s, row %d, col %d: %v\n", sheetName, r+1, c+1, err)
							continue
						}
						if err := f.SetCellValue(sheetName, cellName, newCellValue); err != nil {
							fmt.Fprintf(os.Stderr, "Error setting cell value for %s on sheet %s: %v\n", cellName, sheetName, err)
							continue
						}
					}
				}

				if rowModified {
					fmt.Printf("%s:%s:%d:%s\n", filePath, sheetName, r+1, strings.Join(newRowValues, ","))
				}
			}
		}

		if err := f.Save(); err != nil {
			fmt.Fprintf(os.Stderr, "Error saving file %s: %v\n", filePath, err)
			os.Exit(1)
		}
	},
}

func init() {
	rootCmd.AddCommand(sdCmd)
}
