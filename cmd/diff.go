package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/gdamore/tcell/v2"
	"github.com/pmezard/go-difflib/difflib"
	"github.com/rivo/tview"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// 差分表示用のカラーコード
const (
	colorLightOrange = "[#ffaf00]"
	colorLightBlue   = "[#87d7ff]"
	colorReset       = "[-]"
)

// シートごとの差分結果を保持する構造体
type diffResult struct {
	title   string
	content string
}

var diffCmd = &cobra.Command{
	Use:   "diff [file1] [file2]",
	Short: "Compare sheet names and cell contents of two Excel files",
	Long:  `Compare sheet names of two Excel files and output the differences in unified diff format. For sheets with the same name, compare the cell contents cell by cell.`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		file1 := args[0]
		file2 := args[1]

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

		var results []diffResult

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
			results = append(results, diffResult{
				title:   "Sheet List",
				content: strings.Join(lines, "\n"),
			})
		}

		// 全てのユニークなシート名を取得する
		allSheetsMap := make(map[string]bool)
		var allSheets []string
		for _, s := range sheets1 {
			if !allSheetsMap[s] {
				allSheetsMap[s] = true
				allSheets = append(allSheets, s)
			}
		}
		for _, s := range sheets2 {
			if !allSheetsMap[s] {
				allSheetsMap[s] = true
				allSheets = append(allSheets, s)
			}
		}

		// 各シートについてセルの内容を比較または出力する
		for _, sheet := range allSheets {
			in1 := false
			for _, s := range sheets1 {
				if s == sheet {
					in1 = true
					break
				}
			}
			in2 := false
			for _, s := range sheets2 {
				if s == sheet {
					in2 = true
					break
				}
			}

			if in1 && in2 {
				// 両方のファイルに存在する場合、差分を計算する
				rows1, err := f1.GetRows(sheet)
				if err != nil {
					fmt.Printf("Error reading sheet %s from %s\n", sheet, file1)
					continue
				}

				rows2, err := f2.GetRows(sheet)
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

					var rowNumStr string
					if r1 != -1 {
						rowNumStr = fmt.Sprintf("%d", r1+1)
					} else {
						rowNumStr = fmt.Sprintf("%d", r2+1)
					}

					sheetOutput = append(sheetOutput, fmt.Sprintf("Row %s: %s", rowNumStr, strings.Join(diffCells, ",")))
				}

				if hasSheetDiff {
					results = append(results, diffResult{
						title:   sheet,
						content: strings.Join(sheetOutput, "\n"),
					})
				}
			} else if in1 {
				// 1つ目のファイルにのみ存在する場合、catコマンドと同様にそのまま出力する
				rows1, err := f1.GetRows(sheet)
				if err != nil {
					fmt.Printf("Error reading sheet %s from %s\n", sheet, file1)
					continue
				}

				// シート内の最大列数を取得する
				maxCols := 0
				for _, row := range rows1 {
					if len(row) > maxCols {
						maxCols = len(row)
					}
				}

				var sheetOutput []string
				for _, row := range rows1 {
					var cells []string
					for c := 0; c < maxCols; c++ {
						val := ""
						if c < len(row) {
							val = row[c]
						}
						cells = append(cells, escapeCSVField(val))
					}
					sheetOutput = append(sheetOutput, strings.Join(cells, ","))
				}

				results = append(results, diffResult{
					title:   sheet,
					content: strings.Join(sheetOutput, "\n"),
				})
			} else if in2 {
				// 2つ目のファイルにのみ存在する場合、catコマンドと同様にそのまま出力する
				rows2, err := f2.GetRows(sheet)
				if err != nil {
					fmt.Printf("Error reading sheet %s from %s\n", sheet, file2)
					continue
				}

				// シート内の最大列数を取得する
				maxCols := 0
				for _, row := range rows2 {
					if len(row) > maxCols {
						maxCols = len(row)
					}
				}

				var sheetOutput []string
				for _, row := range rows2 {
					var cells []string
					for c := 0; c < maxCols; c++ {
						val := ""
						if c < len(row) {
							val = row[c]
						}
						cells = append(cells, escapeCSVField(val))
					}
					sheetOutput = append(sheetOutput, strings.Join(cells, ","))
				}

				results = append(results, diffResult{
					title:   sheet,
					content: strings.Join(sheetOutput, "\n"),
				})
			}
		}

		// 差分が全くない場合の処理
		if len(results) == 0 {
			fmt.Println("No differences found.")
			return
		}

		// TUIアプリケーションとページコンテナの構築
		app := tview.NewApplication()
		pages := tview.NewPages()

		// タブバーの作成と設定
		tabBar := tview.NewTextView().
			SetDynamicColors(true).
			SetRegions(true).
			SetWrap(false).
			SetHighlightedFunc(func(added, removed, remaining []string) {
				if len(added) > 0 {
					pages.SwitchToPage(added[0])
				}
			})
		tabBar.SetBackgroundColor(tcell.ColorDefault)

		var tabTitles []string
		// 各シートの差分結果をページとして追加する
		for i, res := range results {
			pageID := fmt.Sprintf("page_%d", i)
			tabTitles = append(tabTitles, fmt.Sprintf(`["%s"] %s [""]`, pageID, res.title))

			textView := tview.NewTextView().
				SetDynamicColors(true).
				SetText(res.content).
				SetScrollable(true).
				SetWrap(false)
			textView.SetBackgroundColor(tcell.ColorDefault)

			pages.AddPage(pageID, textView, true, i == 0)
		}

		// タブバーにタイトルを設定し、最初のタブをハイライトする
		tabBar.SetText(strings.Join(tabTitles, " | "))
		if len(results) > 0 {
			tabBar.Highlight(fmt.Sprintf("page_%d", 0))
		}

		// キー入力のハンドリング（タブ切り替えと終了）
		currentTab := 0
		app.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
			if event.Key() == tcell.KeyRight || event.Key() == tcell.KeyTab {
				currentTab = (currentTab + 1) % len(results)
				tabBar.Highlight(fmt.Sprintf("page_%d", currentTab))
				return nil
			} else if event.Key() == tcell.KeyLeft {
				currentTab = (currentTab - 1 + len(results)) % len(results)
				tabBar.Highlight(fmt.Sprintf("page_%d", currentTab))
				return nil
			} else if event.Key() == tcell.KeyEscape || event.Rune() == 'q' {
				app.Stop()
				return nil
			}
			return event
		})

		// レイアウトの組み立て（上にタブバー、下にページ内容）
		layout := tview.NewFlex().
			SetDirection(tview.FlexRow).
			AddItem(tabBar, 1, 1, false).
			AddItem(pages, 0, 1, true)

		// TUIアプリケーションを実行する
		if err := app.SetRoot(layout, true).EnableMouse(true).Run(); err != nil {
			fmt.Printf("Error running TUI: %v\n", err)
			os.Exit(1)
		}
	},
}

func init() {
	// コマンドを登録する
	rootCmd.AddCommand(diffCmd)
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

			matchCost := 0
			maxL := len(a[i-1])
			if len(b[j-1]) > maxL {
				maxL = len(b[j-1])
			}
			for k := 0; k < maxL; k++ {
				v1, v2 := "", ""
				if k < len(a[i-1]) {
					v1 = a[i-1][k]
				}
				if k < len(b[j-1]) {
					v2 = b[j-1][k]
				}
				if v1 != v2 {
					matchCost++
				}
			}
			costMatch := dp[i-1][j-1] + matchCost

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
			matchCost := 0
			maxL := len(a[i-1])
			if len(b[j-1]) > maxL {
				maxL = len(b[j-1])
			}
			for k := 0; k < maxL; k++ {
				v1, v2 := "", ""
				if k < len(a[i-1]) {
					v1 = a[i-1][k]
				}
				if k < len(b[j-1]) {
					v2 = b[j-1][k]
				}
				if v1 != v2 {
					matchCost++
				}
			}
			if dp[i][j] == dp[i-1][j-1]+matchCost {
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
