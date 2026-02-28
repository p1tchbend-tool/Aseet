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

var grepFormula bool
var grepIgnoreCase bool
var grepRecursive bool

// 対応するExcelファイルの拡張子かどうかを判定する
func isExcelFile(ext string) bool {
	return ext == ".xlsx" || ext == ".xlsm" || ext == ".xlam" || ext == ".xltm" || ext == ".xltx"
}

var grepCmd = &cobra.Command{
	Use:   "grep [pattern] [file or directory]",
	Short: "Search for lines containing the specified string from an Excel file or directory.",
	Long:  `Search and display lines containing the specified string from all sheets in the specified Excel file or directory.`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		pattern := args[0]
		path := args[1]

		var re *regexp.Regexp
		var err error
		// 大文字小文字を区別しないオプションが指定された場合
		if grepIgnoreCase {
			re, err = regexp.Compile("(?i)" + pattern)
		} else {
			re, err = regexp.Compile(pattern)
		}
		if err != nil {
			fmt.Printf("Error compiling regex: %v\n", err)
			os.Exit(1)
		}

		// 指定されたパスの情報を取得する
		info, err := os.Stat(path)
		if err != nil {
			fmt.Printf("Error accessing path %s\n", path)
			os.Exit(1)
		}

		var filesToProcess []string

		// パスがディレクトリの場合
		if info.IsDir() {
			// 再帰的に検索する場合
			if grepRecursive {
				_ = filepath.Walk(path, func(p string, info os.FileInfo, err error) error {
					if err != nil {
						// 探索中のエラー（アクセス権限など）は無視して続行する
						return nil
					}
					if !info.IsDir() {
						ext := strings.ToLower(filepath.Ext(p))
						if isExcelFile(ext) {
							filesToProcess = append(filesToProcess, p)
						}
					}
					return nil
				})
			} else {
				// ディレクトリ直下のみを検索する場合
				entries, err := os.ReadDir(path)
				if err == nil {
					for _, entry := range entries {
						if !entry.IsDir() {
							ext := strings.ToLower(filepath.Ext(entry.Name()))
							if isExcelFile(ext) {
								filesToProcess = append(filesToProcess, filepath.Join(path, entry.Name()))
							}
						}
					}
				}
			}
		} else {
			// パスがファイルの場合
			filesToProcess = append(filesToProcess, path)
		}

		// ファイルが見つからなかった場合
		if len(filesToProcess) == 0 {
			fmt.Println("File not found.")
			os.Exit(1)
		}

		// 収集したファイルを順に処理する
		for _, filePath := range filesToProcess {
			// Excelファイルを開く
			f, err := excelize.OpenFile(filePath)
			if err != nil {
				fmt.Printf("Error opening file %s\n", filePath)
				continue
			}

			// 全シートをループ処理する
			for _, sheetName := range f.GetSheetList() {
				// シートのすべての行を取得する
				rows, err := f.GetRows(sheetName)
				if err != nil {
					fmt.Printf("Error getting rows from sheet %s in file %s\n", sheetName, filePath)
					continue
				}

				// 各行をループ処理する
				for i, row := range rows {
					isMatched := false
					// 各セルをループ処理する
					for c, cell := range row {
						searchTarget := cell
						// 数式を検索対象にする場合
						if grepFormula {
							cellName, err := excelize.CoordinatesToCellName(c+1, i+1)
							if err == nil {
								formula, err := f.GetCellFormula(sheetName, cellName)
								if err == nil && formula != "" {
									searchTarget = formula
								}
							}
						}

						// 正規表現でマッチするか判定する
						if re.MatchString(searchTarget) {
							isMatched = true
							break
						}
					}
					// マッチした場合、ファイル名、シート名、行番号を出力する
					if isMatched {
						fmt.Printf("[Matched] %s: %s: Row %d\n", filePath, sheetName, i+1)
					}
				}
			}

			// ファイルを閉じる
			f.Close()
		}
	},
}

func init() {
	// コマンドとフラグを登録する
	rootCmd.AddCommand(grepCmd)
	grepCmd.Flags().BoolVarP(&grepFormula, "formula", "f", false, "If a cell contains a formula, search the formula instead.")
	grepCmd.Flags().BoolVarP(&grepIgnoreCase, "ignore-case", "i", false, "Ignore case distinctions during the search.")
	grepCmd.Flags().BoolVarP(&grepRecursive, "recursive", "r", false, "Search recursively through subdirectories.")
}
