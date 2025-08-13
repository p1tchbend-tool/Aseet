package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

func printSheetContents(f *excelize.File, sheetName string) error {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return err
	}

	for r, row := range rows {
		var outputCells []string
		for c, originalValue := range row {
			cellName, err := excelize.CoordinatesToCellName(c+1, r+1)
			if err != nil {
				// This is unlikely to happen. Fallback to original value from GetRows.
				outputCells = append(outputCells, originalValue)
				continue
			}

			formula, err := f.GetCellFormula(sheetName, cellName)
			if err == nil && formula != "" {
				outputCells = append(outputCells, formula)
			} else {
				value, err := f.GetCellValue(sheetName, cellName)
				if err != nil {
					// If GetCellValue fails, fallback to original value.
					outputCells = append(outputCells, originalValue)
				} else {
					outputCells = append(outputCells, value)
				}
			}
		}
		fmt.Println(strings.Join(outputCells, ","))
	}
	return nil
}

var all bool
var sheetName string

var catCmd = &cobra.Command{
	Use:   "cat [file]",
	Short: "Excelファイルのシート名を出力します。",
	Long:  `指定されたExcelファイルのシート名を1行ずつ出力します。`,
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		filePath := args[0]

		f, err := excelize.OpenFile(filePath)
		if err != nil {
			fmt.Fprintf(os.Stderr, "ファイル %s を開く際にエラーが発生しました: %v\n", filePath, err)
			os.Exit(1)
		}
		defer func() {
			if err := f.Close(); err != nil {
				fmt.Fprintf(os.Stderr, "ファイル %s を閉じる際にエラーが発生しました: %v\n", filePath, err)
			}
		}()

		if sheetName != "" {
			if err := printSheetContents(f, sheetName); err != nil {
				fmt.Fprintf(os.Stderr, "シート %s の行取得中にエラーが発生しました: %v\n", sheetName, err)
				os.Exit(1)
			}
		} else if all {
			for _, sheet := range f.GetSheetList() {
				fmt.Println("")
				fmt.Println("================================================================================")
				fmt.Println(sheet)
				fmt.Println("================================================================================")
				if err := printSheetContents(f, sheet); err != nil {
					fmt.Fprintf(os.Stderr, "シート %s の行取得中にエラーが発生しました: %v\n", sheet, err)
					continue
				}
			}
		} else {
			for _, sheet := range f.GetSheetList() {
				fmt.Println(sheet)
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(catCmd)
	catCmd.Flags().BoolVarP(&all, "all", "a", false, "すべてのシートのセルの値をカンマ区切りで表示します。")
	catCmd.Flags().StringVarP(&sheetName, "name", "n", "", "指定したシートのセルの値をカンマ区切りで表示します。")
}
