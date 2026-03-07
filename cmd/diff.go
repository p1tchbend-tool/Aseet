package cmd

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/pmezard/go-difflib/difflib"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// 差分表示用のカラーコード
const (
	colorDel    = "[#d55e00]"
	colorAdd    = "[#56b4e9]"
	colorChange = "[#f0e442]"
	colorReset  = "[-]"
)

var diffFormula bool
var diffOpen bool
var diffSheetName string

// 変更されたシートとそのセル座標を保持する構造体
type modifiedSheet struct {
	name  string
	cells []string
}

var diffCmd = &cobra.Command{
	Use:   "diff [file1] [file2]",
	Short: "Compare sheet names and cell contents of two Excel files",
	Long:  `Compare sheet names of two Excel files and output the differences in unified diff format. For sheets with the same name, compare the cell contents cell by cell.`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		file1 := args[0]
		file2 := args[1]

		// --open オプションが指定された場合の処理
		if diffOpen {
			cacheDir, err := os.UserCacheDir()
			if err != nil {
				fmt.Printf("Error getting cache directory: %v\n", err)
				os.Exit(1)
			}
			aseetDir := filepath.Join(cacheDir, "aseet")
			if err := os.MkdirAll(aseetDir, 0755); err != nil {
				fmt.Printf("Error creating cache directory: %v\n", err)
				os.Exit(1)
			}

			localPath := filepath.Join(aseetDir, "[OLD]"+filepath.Base(file1))
			remotePath := filepath.Join(aseetDir, "[NEW]"+filepath.Base(file2))

			if err := copyFile(file1, localPath); err != nil {
				fmt.Printf("Error copying file1: %v\n", err)
				os.Exit(1)
			}
			if err := copyFile(file2, remotePath); err != nil {
				fmt.Printf("Error copying file2: %v\n", err)
				os.Exit(1)
			}

			openFile(localPath)
			openFile(remotePath)
		}

		// 1つ目のExcelファイルを開く
		f1, err := excelize.OpenFile(file1)
		if err != nil {
			fmt.Printf("Error opening file %s\n", file1)
			os.Exit(1)
		}
		defer f1.Close()

		// 2つ目のExcelファイルを開く
		f2, err := excelize.OpenFile(file2)
		if err != nil {
			fmt.Printf("Error opening file %s\n", file2)
			os.Exit(1)
		}
		defer f2.Close()

		var results []sheetResult

		// 両方のファイルからシート名のリストを取得する
		sheets1 := f1.GetSheetList()
		sheets2 := f2.GetSheetList()

		// 差分を取る前にシート名を昇順にソートする
		sort.Strings(sheets1)
		sort.Strings(sheets2)

		var sheetListDiff string

		// シート名が指定されていない場合のみ、シート名一覧の差分を生成する
		if diffSheetName == "" {
			text1 := strings.Join(sheets1, "\n") + "\n"
			text2 := strings.Join(sheets2, "\n") + "\n"

			// シート名のリストを比較するためのUnified Diffを設定する
			diff := difflib.UnifiedDiff{
				A:        difflib.SplitLines(text1),
				B:        difflib.SplitLines(text2),
				FromFile: file1,
				ToFile:   file2,
				Context:  3,
			}

			// 差分文字列を生成する
			text, err := difflib.GetUnifiedDiffString(diff)
			if err != nil {
				fmt.Printf("Error generating diff: %v\n", err)
				os.Exit(1)
			}

			if text != "" {
				var diffLines []string
				lines := strings.Split(text, "\n")
				for _, line := range lines {
					// 空行やヘッダー行を除外する
					if line == "" || strings.HasPrefix(line, "---") || strings.HasPrefix(line, "+++") || strings.HasPrefix(line, "@@") {
						continue
					}

					// Unified Diffの出力を色付けする
					if strings.HasPrefix(line, "-") {
						diffLines = append(diffLines, colorDel+line+colorReset)
					} else if strings.HasPrefix(line, "+") {
						diffLines = append(diffLines, colorAdd+line+colorReset)
					} else {
						diffLines = append(diffLines, line)
					}
				}
				sheetListDiff = strings.Join(diffLines, "\n")
			}
		}

		// 全てのユニークなシート名を取得し、存在チェック用のマップを作成する
		sheetMap1 := make(map[string]bool)
		sheetMap2 := make(map[string]bool)
		var allSheets []string

		for _, s := range sheets1 {
			sheetMap1[s] = true
		}
		for _, s := range sheets2 {
			sheetMap2[s] = true
		}

		// 比較対象のシートを決定する
		if diffSheetName != "" {
			if !sheetMap1[diffSheetName] && !sheetMap2[diffSheetName] {
				fmt.Printf("Sheet %s does not exist in either file.\n", diffSheetName)
				os.Exit(1)
			}
			allSheets = []string{diffSheetName}
		} else {
			for _, s := range sheets1 {
				allSheets = append(allSheets, s)
			}
			for _, s := range sheets2 {
				if !sheetMap1[s] {
					allSheets = append(allSheets, s)
				}
			}
		}

		var modifiedSheets []modifiedSheet

		// 各シートについてセルの内容を比較または出力する
		for _, sheet := range allSheets {
			in1 := sheetMap1[sheet]
			in2 := sheetMap2[sheet]

			if in1 && in2 {
				// 両方のファイルに存在する場合、差分を計算する
				rows1, err := getSheetData(f1, sheet, diffFormula)
				if err != nil {
					fmt.Printf("Error reading sheet %s from %s\n", sheet, file1)
					continue
				}

				rows2, err := getSheetData(f2, sheet, diffFormula)
				if err != nil {
					fmt.Printf("Error reading sheet %s from %s\n", sheet, file2)
					continue
				}

				rowAlign := align(rows1, rows2)
				colAlign := align(transpose(rows1), transpose(rows2))

				hasSheetDiff := false
				var sheetOutput []string
				var changedCells []string

				for _, rPair := range rowAlign {
					r1, r2 := rPair[0], rPair[1]
					var diffCells []string

					for _, cPair := range colAlign {
						c1, c2 := cPair[0], cPair[1]
						val1, val2 := "", ""

						if r1 != -1 && c1 != -1 && r1 < len(rows1) && c1 < len(rows1[r1]) {
							val1 = rows1[r1][c1]
						}
						if r2 != -1 && c2 != -1 && r2 < len(rows2) && c2 < len(rows2[r2]) {
							val2 = rows2[r2][c2]
						}

						if val1 == val2 {
							diffCells = append(diffCells, escapeCSVField(val1))
						} else {
							hasSheetDiff = true

							// 変更されたセルの座標を取得
							rIdx := r2
							if rIdx == -1 {
								rIdx = r1
							}
							cIdx := c2
							if cIdx == -1 {
								cIdx = c1
							}
							if rIdx != -1 && cIdx != -1 {
								cellName, _ := excelize.CoordinatesToCellName(cIdx+1, rIdx+1)
								if cellName != "" {
									changedCells = append(changedCells, cellName)
								}
							}

							var cellDiff string
							if val1 != "" && val2 != "" {
								cellDiff = fmt.Sprintf("%s-%s%s %s+%s%s", colorDel, escapeCSVField(val1), colorReset, colorAdd, escapeCSVField(val2), colorReset)
							} else if val1 != "" {
								cellDiff = fmt.Sprintf("%s-%s%s", colorDel, escapeCSVField(val1), colorReset)
							} else if val2 != "" {
								cellDiff = fmt.Sprintf("%s+%s%s", colorAdd, escapeCSVField(val2), colorReset)
							}
							diffCells = append(diffCells, cellDiff)
						}
					}

					sheetOutput = append(sheetOutput, strings.Join(diffCells, ","))
				}

				if hasSheetDiff {
					modifiedSheets = append(modifiedSheets, modifiedSheet{
						name:  sheet,
						cells: changedCells,
					})
					results = append(results, sheetResult{
						title:   sheet,
						content: strings.Join(sheetOutput, "\n"),
					})
				} else {
					// 差分が全くない場合、catコマンドと同様にそのまま出力する
					content, err := getSheetContents(f1, sheet, diffFormula)
					if err != nil {
						fmt.Printf("Error reading sheet %s\n", sheet)
						continue
					}
					results = append(results, sheetResult{
						title:   sheet,
						content: content,
					})
				}
			} else if in1 {
				// 1つ目のファイルにのみ存在する場合、catコマンドと同様にそのまま出力する
				content, err := getSheetContents(f1, sheet, diffFormula)
				if err != nil {
					fmt.Printf("Error reading sheet %s\n", sheet)
					continue
				}
				results = append(results, sheetResult{
					title:   fmt.Sprintf("%s%s : %s%s", colorDel, filepath.Base(file1), sheet, colorReset),
					content: content,
				})
			} else if in2 {
				// 2つ目のファイルにのみ存在する場合、catコマンドと同様にそのまま出力する
				content, _ := getSheetContents(f2, sheet, diffFormula)
				if err != nil {
					fmt.Printf("Error reading sheet %s\n", sheet)
					continue
				}
				results = append(results, sheetResult{
					title:   fmt.Sprintf("%s%s : %s%s", colorAdd, filepath.Base(file2), sheet, colorReset),
					content: content,
				})
			}
		}

		// サマリーを作成して先頭に追加する
		var summaryBuilder strings.Builder
		if sheetListDiff != "" {
			summaryBuilder.WriteString("\n")
			summaryBuilder.WriteString("[Sheet Name Differences]")
			summaryBuilder.WriteString("\n\n")
			summaryBuilder.WriteString(sheetListDiff)
			summaryBuilder.WriteString("\n")
		}
		if len(modifiedSheets) > 0 {
			summaryBuilder.WriteString("\n")
			summaryBuilder.WriteString("[Modified Sheets (Cell Differences)]")
			summaryBuilder.WriteString("\n\n")

			// 変更されたシート名を昇順ソートする
			sort.Slice(modifiedSheets, func(i, j int) bool {
				return modifiedSheets[i].name < modifiedSheets[j].name
			})

			for _, ms := range modifiedSheets {
				summaryBuilder.WriteString(fmt.Sprintf("%s%s: %s%s\n", colorChange, ms.name, strings.Join(ms.cells, ", "), colorReset))
			}
		}

		summaryText := summaryBuilder.String()
		if summaryText != "" {
			results = append([]sheetResult{{
				title:   "Summary",
				content: summaryText,
			}}, results...)
		}

		// 差分が全くない場合の処理
		if len(results) == 0 {
			fmt.Println("No differences found.")
			return
		}

		// TUIアプリケーションを実行する
		if err := displayTui(results); err != nil {
			fmt.Printf("Error running TUI: %v\n", err)
			os.Exit(1)
		}
	},
}

func init() {
	// コマンドを登録する
	rootCmd.AddCommand(diffCmd)
	diffCmd.Flags().BoolVarP(&diffFormula, "formula", "f", false, "If the cell value is a formula, compare the formula instead of the value.")
	diffCmd.Flags().StringVarP(&diffSheetName, "name", "n", "", "Compare only the specified sheet.")
	diffCmd.Flags().BoolVarP(&diffOpen, "open", "o", false, "Copy the two files to the cache directory with [LOCAL] and [REMOTE] prefixes and open them.")
}
