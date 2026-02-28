package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/gdamore/tcell/v2"
	"github.com/rivo/tview"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var all bool
var catFormula bool
var sheetName string

// シートの内容を文字列として取得する
func getSheetContents(f *excelize.File, sheetName string) (string, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return "", err
	}

	var sb strings.Builder
	for r, row := range rows {
		var outputCells []string
		for c, originalValue := range row {
			cellName, _ := excelize.CoordinatesToCellName(c+1, r+1)
			formulaText, err := f.GetCellFormula(sheetName, cellName)
			if catFormula && err == nil && formulaText != "" {
				outputCells = append(outputCells, fmt.Sprintf("\"%s\"", strings.ReplaceAll(formulaText, "\"", "\"\"")))
			} else {
				value, err := f.GetCellValue(sheetName, cellName)
				if err != nil {
					// GetCellValueが失敗した場合は元の値にフォールバック
					outputCells = append(outputCells, fmt.Sprintf("\"%s\"", strings.ReplaceAll(originalValue, "\"", "\"\"")))
				} else {
					outputCells = append(outputCells, fmt.Sprintf("\"%s\"", strings.ReplaceAll(value, "\"", "\"\"")))
				}
			}
		}
		sb.WriteString(strings.Join(outputCells, ",") + "\n")
	}
	return sb.String(), nil
}

// シートの内容を標準出力に表示する
func printSheetContents(f *excelize.File, sheetName string) error {
	content, err := getSheetContents(f, sheetName)
	if err != nil {
		return err
	}
	fmt.Print(content)
	return nil
}

var catCmd = &cobra.Command{
	Use:   "cat [file]",
	Short: "Excelファイルのシート名を出力します。",
	Long:  `指定されたExcelファイルのシート名を1行ずつ出力します。`,
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		filePath := args[0]

		f, err := excelize.OpenFile(filePath)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", filePath, err)
			os.Exit(1)
		}
		defer f.Close()

		if sheetName != "" {
			if err := printSheetContents(f, sheetName); err != nil {
				fmt.Fprintf(os.Stderr, "Error reading sheet %s: %v\n", sheetName, err)
				os.Exit(1)
			}
		} else if all {
			type sheetResult struct {
				title   string
				content string
			}
			var results []sheetResult

			// 全シートの内容を取得
			for _, sheet := range f.GetSheetList() {
				content, err := getSheetContents(f, sheet)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error reading sheet %s: %v\n", sheet, err)
					continue
				}
				results = append(results, sheetResult{
					title:   sheet,
					content: content,
				})
			}

			if len(results) == 0 {
				fmt.Println("No sheets found.")
				return
			}

			// TUIの構築
			app := tview.NewApplication()
			pages := tview.NewPages()

			// タブバーの作成
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

			// キー入力のハンドリング
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

			// レイアウトの組み立て
			layout := tview.NewFlex().
				SetDirection(tview.FlexRow).
				AddItem(tabBar, 1, 1, false).
				AddItem(pages, 0, 1, true)

			if err := app.SetRoot(layout, true).EnableMouse(true).Run(); err != nil {
				fmt.Fprintf(os.Stderr, "Error running TUI: %v\n", err)
				os.Exit(1)
			}
		} else {
			for _, sheet := range f.GetSheetList() {
				fmt.Println(sheet)
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(catCmd)
	catCmd.Flags().BoolVarP(&all, "all", "a", false, "すべてのシートのセルの値をカンマ区切りで表示します。")
	catCmd.Flags().BoolVarP(&catFormula, "formula", "f", false, "セルの値が数式の場合は値でなく数式を表示します。")
	catCmd.Flags().StringVarP(&sheetName, "name", "n", "", "指定したシートのセルの値をカンマ区切りで表示します。")
}
