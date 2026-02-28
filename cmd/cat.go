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

var catCmd = &cobra.Command{
	Use:   "cat [file]",
	Short: "Output sheet names or cell values of an Excel file.",
	Long:  `When executed without options, outputs the sheet names of the specified Excel file line by line.`,
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		filePath := args[0]

		// Excelファイルを開く
		f, err := excelize.OpenFile(filePath)
		if err != nil {
			fmt.Printf("Error opening file %s\n", filePath)
			os.Exit(1)
		}
		defer f.Close()

		// 特定のシート名が指定された場合の処理
		if sheetName != "" {
			content, err := getSheetContents(f, sheetName, catFormula)
			if err != nil {
				fmt.Printf("Error reading sheet %s\n", sheetName)
				os.Exit(1)
			}
			fmt.Print(content)

		} else if all {
			// --allオプションが指定された場合、全シートをTUIで表示する
			type sheetResult struct {
				title   string
				content string
			}
			var results []sheetResult

			// 全シートの内容を取得してスライスに保存する
			for _, sheet := range f.GetSheetList() {
				content, err := getSheetContents(f, sheet, catFormula)
				if err != nil {
					fmt.Printf("Error reading sheet %s\n", sheet)
					continue
				}
				results = append(results, sheetResult{
					title:   sheet,
					content: content,
				})
			}

			// シートが見つからなかった場合の処理
			if len(results) == 0 {
				fmt.Println("No sheets found.")
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
			// 各シートの内容をページとして追加する
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
					// 次のタブへ移動
					currentTab = (currentTab + 1) % len(results)
					tabBar.Highlight(fmt.Sprintf("page_%d", currentTab))
					return nil
				} else if event.Key() == tcell.KeyLeft {
					// 前のタブへ移動
					currentTab = (currentTab - 1 + len(results)) % len(results)
					tabBar.Highlight(fmt.Sprintf("page_%d", currentTab))
					return nil
				} else if event.Key() == tcell.KeyEscape || event.Rune() == 'q' {
					// アプリケーションを終了
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

		} else {
			// オプションがない場合、シート名の一覧を出力する
			for _, sheet := range f.GetSheetList() {
				fmt.Println(sheet)
			}
		}
	},
}

func init() {
	// コマンドとフラグを登録する
	rootCmd.AddCommand(catCmd)
	catCmd.Flags().BoolVarP(&all, "all", "a", false, "Get cell values of all sheets separated by commas and display them in a TUI pager with separate tabs for each sheet.")
	catCmd.Flags().BoolVarP(&catFormula, "formula", "f", false, "If the cell value is a formula, display the formula instead of the value.")
	catCmd.Flags().StringVarP(&sheetName, "name", "n", "", "Display the cell values of the specified sheet separated by commas.")
}
