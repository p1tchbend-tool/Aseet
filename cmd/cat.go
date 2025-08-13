package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var all bool

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

		if all {
			for _, sheet := range f.GetSheetList() {
				fmt.Println(sheet)
				rows, err := f.GetRows(sheet)
				if err != nil {
					fmt.Fprintf(os.Stderr, "シート %s の行取得中にエラーが発生しました: %v\n", sheet, err)
					continue
				}
				for _, row := range rows {
					fmt.Println(strings.Join(row, ","))
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
}
