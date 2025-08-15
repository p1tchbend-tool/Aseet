package cmd

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var grepCmd = &cobra.Command{
	Use:   "grep [pattern] [file or directory]",
	Short: "Excelファイルまたはディレクトリから指定した文字列を含む行を検索します。",
	Long:  `指定されたExcelファイルまたはディレクトリ内の全シートから、指定した文字列を含む行を検索して表示します。`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		pattern := args[0]
		path := args[1]

		info, err := os.Stat(path)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error accessing path %s: %v\n", path, err)
			os.Exit(1)
		}

		var filesToProcess []string
		if info.IsDir() {
			err := filepath.Walk(path, func(p string, info os.FileInfo, err error) error {
				if err != nil {
					return err
				}
				if !info.IsDir() {
					ext := strings.ToLower(filepath.Ext(p))
					if ext == ".xls" || ext == ".xlsx" || ext == ".xlsm" || ext == ".ods" {
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
			filesToProcess = append(filesToProcess, path)
		}

		for _, filePath := range filesToProcess {
			f, err := excelize.OpenFile(filePath)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", filePath, err)
				continue
			}

			for _, sheetName := range f.GetSheetList() {
				rows, err := f.GetRows(sheetName)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error getting rows from sheet %s in file %s: %v\n", sheetName, filePath, err)
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
						fmt.Printf("%s:%s:%d:%s\n", filePath, sheetName, i+1, strings.Join(row, ","))
					}
				}
			}

			if err := f.Close(); err != nil {
				fmt.Fprintln(os.Stderr, err)
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(grepCmd)
}
