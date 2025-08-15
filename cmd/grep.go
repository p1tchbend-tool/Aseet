package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var grepCmd = &cobra.Command{
	Use:   "grep [pattern] [file]",
	Short: "Excelファイルから指定した文字列を含む行を検索します。",
	Long:  `指定されたExcelファイルの全シートから、指定した文字列を含む行を検索して表示します。`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		pattern := args[0]
		filePath := args[1]

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

			for i, row := range rows {
				match := false
				for _, cell := range row {
					if strings.Contains(cell, pattern) {
						match = true
						break
					}
				}
				if match {
					fmt.Printf("%s:%d:%s\n", sheetName, i+1, strings.Join(row, ","))
				}
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(grepCmd)
}
