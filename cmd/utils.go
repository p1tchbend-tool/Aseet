package cmd

import (
	"fmt"
	"os"
	"os/exec"
	"runtime"
	"strings"

	"github.com/gdamore/tcell/v2"
	"github.com/rivo/tview"
	"github.com/xuri/excelize/v2"
)

// TUIで表示する各タブ（シート）のデータを保持する構造体
type sheetResult struct {
	title   string
	content string
}

// TUIアプリケーションを構築して表示する共通処理
func displayTui(results []sheetResult) error {
	app := tview.NewApplication()
	pages := tview.NewPages()

	var lastFocus tview.Primitive = pages

	tabBar := tview.NewTextView().
		SetDynamicColors(true).
		SetRegions(true).
		SetWrap(false).
		SetScrollable(true).
		SetHighlightedFunc(func(added, removed, remaining []string) {
			if len(added) > 0 {
				pages.SwitchToPage(added[0])
			}
		})
	tabBar.SetBackgroundColor(tcell.ColorDefault)

	// tabBarにフォーカスが当たったことを記録する
	tabBar.SetFocusFunc(func() {
		lastFocus = tabBar
	})

	tabBar.SetMouseCapture(func(action tview.MouseAction, event *tcell.EventMouse) (tview.MouseAction, *tcell.EventMouse) {
		if action == tview.MouseRightClick || action == tview.MouseMiddleClick {
			x, y := event.Position()
			newEvent := tcell.NewEventMouse(x, y, tcell.Button1, event.Modifiers())
			return tview.MouseLeftClick, newEvent
		}
		return action, event
	})

	var tabTitles []string
	for i, res := range results {
		pageID := fmt.Sprintf("page_%d", i)
		tabTitles = append(tabTitles, fmt.Sprintf(`["%s"] %s [""]`, pageID, res.title))

		textView := tview.NewTextView().
			SetDynamicColors(true).
			SetText(res.content).
			SetScrollable(true).
			SetWrap(false)
		textView.SetBackgroundColor(tcell.ColorDefault)

		// テキストビューにフォーカスが当たったことを記録する
		textView.SetFocusFunc(func() {
			lastFocus = pages
		})

		pages.AddPage(pageID, textView, true, i == 0)
	}

	tabBar.SetText(strings.Join(tabTitles, " | "))
	if len(results) > 0 {
		tabBar.Highlight(fmt.Sprintf("page_%d", 0))
	}

	// 操作方法を表示するヘルプバーを作成
	helpText := " [yellow]Tab[-]: Next tab | [yellow]Shift + Tab[-]: Previous tab | [yellow]h / j / k / l[-]: Scroll | [yellow]g[-]: Scroll to top | [yellow]Shift + g[-]: Scroll to bottom | [yellow]q[-]: Quit "
	helpBar := tview.NewTextView().
		SetDynamicColors(true).
		SetText(helpText).
		SetTextAlign(tview.AlignCenter)
	helpBar.SetBackgroundColor(tcell.ColorDefault)

	currentTab := 0
	app.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		// Tabキーで次のタブへ
		if event.Key() == tcell.KeyTab {
			currentTab = (currentTab + 1) % len(results)
			tabBar.Highlight(fmt.Sprintf("page_%d", currentTab))
			return nil
			// Shift+Tabキーで前のタブへ
		} else if event.Key() == tcell.KeyBacktab {
			currentTab = (currentTab - 1 + len(results)) % len(results)
			tabBar.Highlight(fmt.Sprintf("page_%d", currentTab))
			return nil
			// タブバーを左スクロール
		} else if event.Rune() == 'b' {
			app.SetFocus(tabBar)
			for range 100 {
				app.QueueEvent(tcell.NewEventKey(tcell.KeyLeft, 0, tcell.ModNone))
			}
			// タブバーを右スクロール
		} else if event.Rune() == 'f' {
			app.SetFocus(tabBar)
			for range 100 {
				app.QueueEvent(tcell.NewEventKey(tcell.KeyRight, 0, tcell.ModNone))
			}
			// Esq or qで終了
		} else if event.Key() == tcell.KeyEscape || event.Rune() == 'q' {
			app.Stop()
			return nil
		}
		return event
	})

	// 左スクロールボタン
	leftBtn := tview.NewButton("◀").
		SetStyle(tcell.StyleDefault.Background(tcell.ColorDefault)).
		SetActivatedStyle(tcell.StyleDefault.Background(tcell.ColorDefault)).
		SetSelectedFunc(func() {
			for range 100 {
				app.QueueEvent(tcell.NewEventKey(tcell.KeyLeft, 0, tcell.ModNone))
			}
			app.SetFocus(lastFocus) // 最後にフォーカスがあったコンポーネントに戻す
		})

	// 右スクロールボタン
	rightBtn := tview.NewButton("▶").
		SetStyle(tcell.StyleDefault.Background(tcell.ColorDefault)).
		SetActivatedStyle(tcell.StyleDefault.Background(tcell.ColorDefault)).
		SetSelectedFunc(func() {
			for range 100 {
				app.QueueEvent(tcell.NewEventKey(tcell.KeyRight, 0, tcell.ModNone))
			}
			app.SetFocus(lastFocus) // 最後にフォーカスがあったコンポーネントに戻す
		})

	mainContent := tview.NewFlex().
		SetDirection(tview.FlexRow).
		AddItem(tabBar, 1, 1, false).
		AddItem(pages, 0, 1, true).
		AddItem(helpBar, 1, 1, false)

	layout := tview.NewFlex().
		SetDirection(tview.FlexColumn).
		AddItem(leftBtn, 3, 0, false).
		AddItem(mainContent, 0, 1, true).
		AddItem(rightBtn, 3, 0, false)

	return app.SetRoot(layout, true).EnableMouse(true).Run()
}

// 対応するExcelファイルの拡張子かどうかを判定する
func isExcelFile(ext string) bool {
	return ext == ".xlsx" || ext == ".xlsm" || ext == ".xlam" || ext == ".xltm" || ext == ".xltx"
}

// セルの値にカンマ、改行、ダブルクォーテーションが含まれる場合はエスケープ処理を行う
func escapeCSVField(value string) string {
	needsQuotes := strings.Contains(value, "\"") || strings.Contains(value, "\n") || strings.Contains(value, ",")

	// 1. ダブルクォーテーションを2つにする
	value = strings.ReplaceAll(value, "\"", "\"\"")
	// 2. ラインフィールドが含まれる場合、\nに変換する
	value = strings.ReplaceAll(value, "\n", "\\n")

	// 3. ダブルクォーテーション・ラインフィールド・カンマが含まれていた場合は、フィールド全体をダブルクォーテーションで囲む
	if needsQuotes {
		return fmt.Sprintf("\"%s\"", value)
	}
	return value
}

// シートのデータを2次元配列として取得する。isFormulaがtrueの場合は数式を取得する
func getSheetData(f *excelize.File, sheetName string, isFormula bool) ([][]string, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	if !isFormula {
		return rows, nil
	}

	var result [][]string
	for r, row := range rows {
		var newRow []string
		for c, val := range row {
			cellName, err := excelize.CoordinatesToCellName(c+1, r+1)
			if err == nil {
				formula, err := f.GetCellFormula(sheetName, cellName)
				if err == nil && formula != "" {
					newRow = append(newRow, formula)
					continue
				}
			}
			newRow = append(newRow, val)
		}
		result = append(result, newRow)
	}
	return result, nil
}

// シートの内容を文字列として取得する
func getSheetContents(f *excelize.File, sheetName string, isFormula bool) (string, error) {
	// シートのデータを取得する
	rows, err := getSheetData(f, sheetName, isFormula)
	if err != nil {
		return "", err
	}

	// シート内の最大列数を取得する
	maxCols := 0
	for _, row := range rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}

	var sb strings.Builder
	// 各行をループ処理する
	for _, row := range rows {
		var outputCells []string
		// 最大列数に合わせて各セルをループ処理する
		for c := 0; c < maxCols; c++ {
			var value string
			if c < len(row) {
				value = row[c]
			}
			outputCells = append(outputCells, escapeCSVField(value))
		}
		// セルの値をカンマ区切りで結合し、改行を追加する
		sb.WriteString(strings.Join(outputCells, ",") + "\n")
	}
	return sb.String(), nil
}

// ファイルをコピーする
func copyFile(src, dst string) error {
	input, err := os.ReadFile(src)
	if err != nil {
		return err
	}
	return os.WriteFile(dst, input, 0644)
}

// OSの関連付けられたアプリケーションでファイルを開く
func openFile(path string) {
	var cmd *exec.Cmd
	switch runtime.GOOS {
	case "windows":
		cmd = exec.Command("cmd", "/c", "start", "", path)
	case "darwin":
		cmd = exec.Command("open", path)
	default: // linux, etc
		cmd = exec.Command("xdg-open", path)
	}
	err := cmd.Start()
	if err != nil {
		fmt.Printf("Error opening file %s: %v\n", path, err)
	}
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
