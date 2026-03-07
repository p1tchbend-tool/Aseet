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

		pages.AddPage(pageID, textView, true, i == 0)
	}

	tabBar.SetText(strings.Join(tabTitles, " | "))
	if len(results) > 0 {
		tabBar.Highlight(fmt.Sprintf("page_%d", 0))
	}

	// 操作方法を表示するヘルプバーを作成
	helpText := " [yellow]Tab[-]: Next tab | [yellow]Shift + Tab[-]: Prev tab | [yellow]Arrow keys[-]: Scroll | [yellow]Ctrl + c[-]: Quit "
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
		}
		return event
	})

	layout := tview.NewFlex().
		SetDirection(tview.FlexRow).
		AddItem(tabBar, 1, 1, false).
		AddItem(pages, 0, 1, true).
		AddItem(helpBar, 1, 1, false) // ヘルプバーを画面下部に追加

	return app.SetRoot(layout, true).EnableMouse(true).Run()
}

// 対応するExcelファイルの拡張子かどうかを判定する
func isExcelFile(ext string) bool {
	return ext == ".xlsx" || ext == ".xlsm" || ext == ".xlam" || ext == ".xltm" || ext == ".xltx"
}

// セルの値にカンマが含まれる場合はダブルクォーテーションで囲む
func escapeCSVField(value string) string {
	if strings.Contains(value, ",") {
		return fmt.Sprintf("\"%s\"", strings.ReplaceAll(value, "\"", "\"\""))
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
