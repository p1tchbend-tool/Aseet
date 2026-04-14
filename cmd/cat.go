package cmd

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var all bool
var catFormula bool
var sheetName string

var catCmd = &cobra.Command{
	Use:   "cat [file]",
	Short: "Output sheet names or cell values of an Excel file",
	Long:  `When executed without options, outputs the sheet names of the specified Excel file line by line.`,
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		filePath := args[0]

		// Excelファイルを開く
		f, err := excelize.OpenFile(filePath)
		if err != nil {
			fmt.Printf("Error opening file %s\n", filePath)
			os.Exit(1)
		}
		defer f.Close()

		// 特定のシート名が指定された場合の処理
		if sheetName != "" {
			data, err := getSheetData(f, sheetName, catFormula)
			if err != nil {
				fmt.Printf("Error reading sheet %s\n", sheetName)
				os.Exit(1)
			}

			results := []sheetResult{
				{
					title:   sheetName,
					cells:   data,
					isTable: true,
				},
			}

			// TUIアプリケーションを実行する
			if err := displayFileTui(results); err != nil {
				fmt.Printf("Error running TUI: %v\n", err)
				os.Exit(1)
			}

		} else if all {
			// --allオプションが指定された場合、全シートをTUIで表示する
			var results []sheetResult

			// 全シートの内容を取得してスライスに保存する
			for _, sheet := range f.GetSheetList() {
				data, err := getSheetData(f, sheet, catFormula)
				if err != nil {
					fmt.Printf("Error reading sheet %s\n", sheet)
					continue
				}
				results = append(results, sheetResult{
					title:   sheet,
					cells:   data,
					isTable: true,
				})
			}

			// TUIアプリケーションを実行する
			if err := displayFileTui(results); err != nil {
				fmt.Printf("Error running TUI: %v\n", err)
				os.Exit(1)
			}

		} else {
			// オプションがない場合、シート名の一覧を出力する
			for _, sheet := range f.GetSheetList() {
				fmt.Println(sheet)
			}
		}
	},
}

func init() {
	// コマンドとフラグを登録する
	rootCmd.AddCommand(catCmd)
	catCmd.Flags().BoolVarP(&all, "all", "a", false, "Get cell values of all sheets separated by commas and display them in a TUI pager with separate tabs for each sheet.")
	catCmd.Flags().BoolVarP(&catFormula, "formula", "f", false, "If the cell value is a formula, display the formula instead of the value.")
	catCmd.Flags().StringVarP(&sheetName, "name", "n", "", "Display the cell values of the specified sheet separated by commas.")
}
