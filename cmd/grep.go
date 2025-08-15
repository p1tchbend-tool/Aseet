package cmd

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var grepRecursive bool
var grepIgnoreCase bool

var grepCmd = &cobra.Command{
	Use:   "grep [pattern] [file or directory]",
	Short: "Excelファイルまたはディレクトリから指定した文字列を含む行を検索します。",
	Long:  `指定されたExcelファイルまたはディレクトリ内の全シートから、指定した文字列を含む行を検索して表示します。`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		pattern := args[0]
		path := args[1]

		if grepIgnoreCase {
			pattern = strings.ToLower(pattern)
		}

		info, err := os.Stat(path)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error accessing path %s: %v\n", path, err)
			os.Exit(1)
		}

		var filesToProcess []string
		if info.IsDir() {
			if grepRecursive {
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
				entries, err := os.ReadDir(path)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error reading directory %s: %v\n", path, err)
					os.Exit(1)
				}
				for _, entry := range entries {
					if !entry.IsDir() {
						ext := strings.ToLower(filepath.Ext(entry.Name()))
						if ext == ".xls" || ext == ".xlsx" || ext == ".xlsm" || ext == ".ods" {
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

			for _, sheetName := range f.GetSheetList() {
				rows, err := f.GetRows(sheetName)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error getting rows from sheet %s in file %s: %v\n", sheetName, filePath, err)
					continue
				}

				for i, row := range rows {
					match := false
					for _, cell := range row {
						cellToSearch := cell
						if grepIgnoreCase {
							cellToSearch = strings.ToLower(cell)
						}
						if strings.Contains(cellToSearch, pattern) {
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
	grepCmd.Flags().BoolVarP(&grepRecursive, "recursive", "r", false, "サブディレクトリまで再帰的に検索します。")
	grepCmd.Flags().BoolVarP(&grepIgnoreCase, "ignore-case", "i", false, "検索時に大文字小文字を区別しません。")
}
