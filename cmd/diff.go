package cmd

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/pmezard/go-difflib/difflib"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// 差分表示用のカラーコード
const (
	colorLightOrange = "[#ffaf00]"
	colorLightBlue   = "[#87d7ff]"
	colorReset       = "[-]"
)

var diffFormula bool
var diffOpen bool

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

			localPath := filepath.Join(aseetDir, "[LOCAL]"+filepath.Base(file1))
			remotePath := filepath.Join(aseetDir, "[REMOTE]"+filepath.Base(file2))

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

		var sheetListDiff string
		if text != "" {
			// Unified Diffの出力を色付けする
			lines := strings.Split(text, "\n")
			for i, line := range lines {
				if strings.HasPrefix(line, "---") || strings.HasPrefix(line, "+++") {
					// ファイルヘッダーはそのままにする
				} else if strings.HasPrefix(line, "-") {
					lines[i] = colorLightOrange + line + colorReset
				} else if strings.HasPrefix(line, "+") {
					lines[i] = colorLightBlue + line + colorReset
				}
			}
			sheetListDiff = strings.Join(lines, "\n")
		}

		// 全てのユニークなシート名を取得し、存在チェック用のマップを作成する
		sheetMap1 := make(map[string]bool)
		sheetMap2 := make(map[string]bool)
		var allSheets []string

		for _, s := range sheets1 {
			sheetMap1[s] = true
			allSheets = append(allSheets, s)
		}
		for _, s := range sheets2 {
			sheetMap2[s] = true
			if !sheetMap1[s] {
				allSheets = append(allSheets, s)
			}
		}

		var modifiedSheets []string

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
							var cellDiff string
							if val1 != "" && val2 != "" {
								cellDiff = fmt.Sprintf("%s-%s%s %s+%s%s", colorLightOrange, escapeCSVField(val1), colorReset, colorLightBlue, escapeCSVField(val2), colorReset)
							} else if val1 != "" {
								cellDiff = fmt.Sprintf("%s-%s%s", colorLightOrange, escapeCSVField(val1), colorReset)
							} else if val2 != "" {
								cellDiff = fmt.Sprintf("%s+%s%s", colorLightBlue, escapeCSVField(val2), colorReset)
							}
							diffCells = append(diffCells, cellDiff)
						}
					}

					sheetOutput = append(sheetOutput, strings.Join(diffCells, ","))
				}

				if hasSheetDiff {
					modifiedSheets = append(modifiedSheets, sheet)
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
					title:   fmt.Sprintf("%s%s : %s%s", colorLightOrange, filepath.Base(file1), sheet, colorReset),
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
					title:   fmt.Sprintf("%s%s : %s%s", colorLightBlue, filepath.Base(file2), sheet, colorReset),
					content: content,
				})
			}
		}

		// サマリーを作成して先頭に追加する
		var summaryBuilder strings.Builder
		if sheetListDiff != "" {
			summaryBuilder.WriteString("[Sheet Name Differences]\n")
			summaryBuilder.WriteString(sheetListDiff)
		}
		if len(modifiedSheets) > 0 {
			if summaryBuilder.Len() > 0 {
				summaryBuilder.WriteString("\n")
			}
			summaryBuilder.WriteString("[Modified Sheets (Cell Differences)]\n")
			for _, s := range modifiedSheets {
				summaryBuilder.WriteString(fmt.Sprintf("%s\n", s))
			}
		}

		summaryText := strings.TrimSpace(summaryBuilder.String())
		if summaryText != "" {
			results = append([]sheetResult{{
				title:   "[yellow]Summary[-]",
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
	diffCmd.Flags().BoolVarP(&diffOpen, "open", "o", false, "Copy the two files to the cache directory with [LOCAL] and [REMOTE] prefixes and open them.")
}

// 空でないセルの数をカウントする
func countNonEmpty(row []string) int {
	c := 0
	for _, v := range row {
		if v != "" {
			c++
		}
	}
	return c
}

// 2次元配列（マトリックス）を転置する
func transpose(matrix [][]string) [][]string {
	maxCol := 0
	for _, row := range matrix {
		if len(row) > maxCol {
			maxCol = len(row)
		}
	}
	res := make([][]string, maxCol)
	for i := 0; i < maxCol; i++ {
		res[i] = make([]string, len(matrix))
		for j, row := range matrix {
			if i < len(row) {
				res[i][j] = row[i]
			}
		}
	}
	return res
}

// 2つの行の不一致要素数を計算する
func calcMatchCost(row1, row2 []string) int {
	matchCost := 0
	maxL := len(row1)
	if len(row2) > maxL {
		maxL = len(row2)
	}
	for k := 0; k < maxL; k++ {
		v1, v2 := "", ""
		if k < len(row1) {
			v1 = row1[k]
		}
		if k < len(row2) {
			v2 = row2[k]
		}
		if v1 != v2 {
			matchCost++
		}
	}
	return matchCost
}

// 動的計画法（DP）を用いて2つの2次元配列のアライメント（差分）を計算する
func align(a, b [][]string) [][2]int {
	n, m := len(a), len(b)
	dp := make([][]int, n+1)
	for i := range dp {
		dp[i] = make([]int, m+1)
	}

	// 初期化：削除コスト
	for i := 1; i <= n; i++ {
		dp[i][0] = dp[i-1][0] + countNonEmpty(a[i-1])
	}
	// 初期化：挿入コスト
	for j := 1; j <= m; j++ {
		dp[0][j] = dp[0][j-1] + countNonEmpty(b[j-1])
	}

	// DPテーブルを埋める
	for i := 1; i <= n; i++ {
		for j := 1; j <= m; j++ {
			costDel := dp[i-1][j] + countNonEmpty(a[i-1])
			costIns := dp[i][j-1] + countNonEmpty(b[j-1])
			costMatch := dp[i-1][j-1] + calcMatchCost(a[i-1], b[j-1])

			minCost := costDel
			if costIns < minCost {
				minCost = costIns
			}
			if costMatch < minCost {
				minCost = costMatch
			}
			dp[i][j] = minCost
		}
	}

	// バックトラックして最適なパス（アライメント）を復元する
	var path [][2]int
	i, j := n, m
	for i > 0 || j > 0 {
		if i > 0 && j > 0 {
			if dp[i][j] == dp[i-1][j-1]+calcMatchCost(a[i-1], b[j-1]) {
				path = append([][2]int{{i - 1, j - 1}}, path...)
				i--
				j--
				continue
			}
		}
		if i > 0 && dp[i][j] == dp[i-1][j]+countNonEmpty(a[i-1]) {
			path = append([][2]int{{i - 1, -1}}, path...)
			i--
		} else {
			path = append([][2]int{{-1, j - 1}}, path...)
			j--
		}
	}
	return path
}
