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

var sdIgnoreCase bool
var sdSheetName string
var sdRecursive bool
var sdFormula bool
var sdHyperlink bool

var sdCmd = &cobra.Command{
	Use:   "sd [search] [replace] [file or directory]",
	Short: "Search and replace strings in an Excel file or directory.",
	Long:  `Search for a string and replace it with another string in all cells of all sheets in the specified Excel file or directory.`,
	Args:  cobra.ExactArgs(3),
	Run: func(cmd *cobra.Command, args []string) {
		search := args[0]
		replace := args[1]
		path := args[2]

		var re *regexp.Regexp
		var err error
		// 大文字小文字を区別しないオプションが指定された場合
		if sdIgnoreCase {
			re, err = regexp.Compile("(?i)" + search)
		} else {
			re, err = regexp.Compile(search)
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
			// 再帰的に処理する場合
			if sdRecursive {
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
				// ディレクトリ直下のみを処理する場合
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

			var sheetsToProcess []string
			// 特定のシート名が指定された場合
			if sdSheetName != "" {
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
					fmt.Printf("Error: Sheet '%s' not found in file %s\n", sdSheetName, filePath)
					f.Close()
					continue
				}
			} else {
				// 全シートを対象にする
				sheetsToProcess = f.GetSheetList()
			}

			// 対象シートをループ処理する
			for _, sheetName := range sheetsToProcess {
				// シートのすべての行を取得する
				rows, err := f.GetRows(sheetName)
				if err != nil {
					fmt.Printf("Error getting rows from sheet %s in file %s\n", sheetName, filePath)
					continue
				}

				// 各行をループ処理する
				for r, row := range rows {
					isRowModified := false

					// 各セルをループ処理する
					for c, cellValue := range row {
						// セルの座標からセル名（例: A1）を取得する
						cellName, err := excelize.CoordinatesToCellName(c+1, r+1)
						if err != nil {
							continue
						}

						if sdHyperlink {
							// ハイパーリンクのみを置換対象にする
							hasLink, target, err := f.GetCellHyperLink(sheetName, cellName)
							if err == nil && hasLink && target != "" {
								if re.MatchString(target) {
									isRowModified = true
									newTarget := re.ReplaceAllString(target, replace)
									
									// リンクタイプを判定（簡易的に ! が含まれていれば Location、それ以外は External とする）
									linkType := "External"
									if strings.Contains(newTarget, "!") {
										linkType = "Location"
									}
									
									if err := f.SetCellHyperLink(sheetName, cellName, newTarget, linkType); err != nil {
										fmt.Printf("Error setting cell hyperlink for %s on sheet %s\n", cellName, sheetName)
										continue
									}
								}
							}
						} else if sdFormula {
							// 数式のみを置換対象にする
							formula, err := f.GetCellFormula(sheetName, cellName)
							if err == nil && formula != "" {
								if re.MatchString(formula) {
									isRowModified = true
									newFormula := re.ReplaceAllString(formula, replace)
									if err := f.SetCellFormula(sheetName, cellName, newFormula); err != nil {
										fmt.Printf("Error setting cell formula for %s on sheet %s\n", cellName, sheetName)
										continue
									}
									// 数式を更新したため再計算フラグを立てる
									if f.WorkBook != nil && f.WorkBook.CalcPr != nil {
										f.WorkBook.CalcPr.FullCalcOnLoad = true
									}
								}
							}
						} else {
							// セルの値のみを置換対象にする
							if re.MatchString(cellValue) {
								isRowModified = true
								newCellValue := re.ReplaceAllString(cellValue, replace)
								if err := f.SetCellValue(sheetName, cellName, newCellValue); err != nil {
									fmt.Printf("Error setting cell value for %s on sheet %s\n", cellName, sheetName)
									continue
								}
							}
						}
					}

					// 行内で置換が発生した場合、結果を出力する
					if isRowModified {
						fmt.Printf("[Replaced] %s: %s: Row %d\n", filePath, sheetName, r+1)
					}
				}
			}

			// 変更をファイルに保存する
			if err := f.Save(); err != nil {
				fmt.Printf("Error saving file %s\n", filePath)
			}

			// ファイルを閉じる
			f.Close()
		}
	},
}

func init() {
	// コマンドとフラグを登録する
	rootCmd.AddCommand(sdCmd)
	sdCmd.Flags().BoolVarP(&sdIgnoreCase, "ignore-case", "i", false, "Ignore case distinctions during the search.")
	sdCmd.Flags().StringVarP(&sdSheetName, "name", "n", "", "Replace cell values in the specified sheet.")
	sdCmd.Flags().BoolVarP(&sdRecursive, "recursive", "r", false, "Process subdirectories recursively.")
	sdCmd.Flags().BoolVarP(&sdFormula, "formula", "f", false, "Replace only cell formulas.")
	sdCmd.Flags().BoolVarP(&sdHyperlink, "hyperlink", "l", false, "Replace only hyperlinks.")
}
