package internal

import (
	"fmt"
	"strings"

	"github.com/gdamore/tcell/v2"
	"github.com/rivo/tview"
	"github.com/xuri/excelize/v2"
)

// TUIで表示する各タブ（シート）のデータを保持する構造体
type SheetResult struct {
	Title   string     // タブに表示されるタイトル（シート名など）
	Content string     // タブ内に表示されるテキストコンテンツ（Summaryなど）
	Cells   [][]string // テーブル表示用の2次元配列データ
	IsTable bool       // trueの場合はTableとして、falseの場合はTextViewとして表示する
}

// ディレクトリ比較時の各ファイル（ブック）のデータを保持する構造体
type BookResult struct {
	FileName string        // リストに表示されるファイル名
	Sheets   []SheetResult // ファイルに含まれるシートのデータ
}

// シートのタブ画面（メインコンテンツ）を構築する共通モジュール
// app: tviewアプリケーションのインスタンス
// results: 表示するシートのデータリスト
func CreateSheetTabs(app *tview.Application, results []SheetResult) tview.Primitive {
	// ページビュー（タブのコンテンツ部分）を作成
	pages := tview.NewPages()
	var lastFocus tview.Primitive = pages

	// タブバー（上部のタブ一覧）を作成
	tabBar := tview.NewTextView().
		SetDynamicColors(true).
		SetRegions(true).
		SetWrap(false).
		SetScrollable(true).
		SetHighlightedFunc(func(added, removed, remaining []string) {
			// タブが選択されたら、対応するページに切り替える
			if len(added) > 0 {
				pages.SwitchToPage(added[0])
			}
		})
	tabBar.SetBackgroundColor(tcell.ColorDefault)

	// タブバーにフォーカスが当たった際の処理
	tabBar.SetFocusFunc(func() {
		lastFocus = tabBar
	})

	// マウス操作のキャプチャ（右クリックや中クリックを左クリックに変換してタブ選択を容易にする）
	tabBar.SetMouseCapture(func(action tview.MouseAction, event *tcell.EventMouse) (tview.MouseAction, *tcell.EventMouse) {
		if action == tview.MouseRightClick || action == tview.MouseMiddleClick {
			x, y := event.Position()
			newEvent := tcell.NewEventMouse(x, y, tcell.Button1, event.Modifiers())
			return tview.MouseLeftClick, newEvent
		}
		return action, event
	})

	var tabTitles []string
	// 各シートのデータをページとして追加し、タブのタイトルを生成する
	for i, res := range results {
		pageID := fmt.Sprintf("page_%d", i)
		// tviewのRegion機能を使ってタブのタイトルをフォーマット
		tabTitles = append(tabTitles, fmt.Sprintf(`["%s"] %s [""]`, pageID, res.Title))

		var page tview.Primitive

		if res.IsTable {
			// テーブルビューを作成
			table := tview.NewTable().
				SetBorders(true).
				SetBordersColor(tcell.GetColor("#e5e5e5")).
				SetFixed(1, 1) // ヘッダー行とヘッダー列を固定
			table.SetBackgroundColor(tcell.ColorDefault)

			// 最大列数を計算
			maxCols := 0
			for _, row := range res.Cells {
				if len(row) > maxCols {
					maxCols = len(row)
				}
			}
			// 常に1つ余分な列を追加
			maxCols++

			// 左上の空白セル
			table.SetCell(0, 0, tview.NewTableCell("").SetSelectable(false))

			// 列ヘッダー (A, B, C...) を追加
			for c := 0; c < maxCols; c++ {
				colName, _ := excelize.ColumnNumberToName(c + 1)
				table.SetCell(0, c+1, tview.NewTableCell(colName).
					SetSelectable(false).
					SetAlign(tview.AlignCenter).
					SetTextColor(tcell.GetColor("#f0e442")))
			}

			// 行データと行ヘッダー (1, 2, 3...) を追加
			// 常に1つ余分な行を追加
			rowCount := len(res.Cells) + 1
			for r := 0; r < rowCount; r++ {
				table.SetCell(r+1, 0, tview.NewTableCell(fmt.Sprintf("%d", r+1)).
					SetSelectable(false).
					SetAlign(tview.AlignRight).
					SetTextColor(tcell.GetColor("#f0e442")))

				var row []string
				if r < len(res.Cells) {
					row = res.Cells[r]
				}

				for c := 0; c < maxCols; c++ {
					val := ""
					if c < len(row) {
						val = row[c]
					}
					// 改行が含まれる場合は表示崩れを防ぐためエスケープ
					val = strings.ReplaceAll(val, "\n", "\\n")
					table.SetCell(r+1, c+1, tview.NewTableCell(val).SetSelectable(true))
				}
			}

			// テーブルにフォーカスが当たった際の処理
			table.SetFocusFunc(func() {
				lastFocus = pages
			})
			page = table
		} else {
			// テキストビューを作成（Summaryなど用）
			textView := tview.NewTextView().
				SetDynamicColors(true).
				SetText(res.Content).
				SetScrollable(true).
				SetWrap(false)
			textView.SetBackgroundColor(tcell.ColorDefault)

			// テキストビューにフォーカスが当たった際の処理
			textView.SetFocusFunc(func() {
				lastFocus = pages
			})
			page = textView
		}

		// 最初のページのみ表示状態にする
		pages.AddPage(pageID, page, true, i == 0)
	}

	// タブバーにタイトル一覧を設定
	tabBar.SetText(strings.Join(tabTitles, " | "))
	if len(results) > 0 {
		// 最初のタブをハイライト状態にする
		tabBar.Highlight(fmt.Sprintf("page_%d", 0))
	}

	// タブバーとページビューを縦に並べるレイアウトを作成
	layout := tview.NewFlex().
		SetDirection(tview.FlexRow).
		AddItem(tabBar, 1, 1, false).
		AddItem(pages, 0, 1, true)

	currentTab := 0
	// コンポーネント単位でキーバインドを設定
	layout.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		if event.Key() == tcell.KeyTab {
			// Tabキーで次のタブへ
			currentTab = (currentTab + 1) % len(results)
			tabBar.Highlight(fmt.Sprintf("page_%d", currentTab))
			return nil
		} else if event.Key() == tcell.KeyBacktab {
			// Shift+Tabキーで前のタブへ
			currentTab = (currentTab - 1 + len(results)) % len(results)
			tabBar.Highlight(fmt.Sprintf("page_%d", currentTab))
			return nil
		} else if event.Rune() == 'b' {
			// 'b'キーでタブバーを左にスクロール
			row, col := tabBar.GetScrollOffset()
			newCol := col - 1
			if newCol < 0 {
				newCol = 0
			}
			tabBar.ScrollTo(row, newCol)
			return nil
		} else if event.Rune() == 'f' {
			// 'f'キーでタブバーを右にスクロール
			row, col := tabBar.GetScrollOffset()
			tabBar.ScrollTo(row, col+1)
			return nil
		} else if event.Rune() == 'B' {
			// 'B'キーでタブバーを大きく左にスクロール
			row, col := tabBar.GetScrollOffset()
			newCol := col - 100
			if newCol < 0 {
				newCol = 0
			}
			tabBar.ScrollTo(row, newCol)
			return nil
		} else if event.Rune() == 'F' {
			// 'F'キーでタブバーを大きく右にスクロール
			row, col := tabBar.GetScrollOffset()
			tabBar.ScrollTo(row, col+100)
			return nil
		} else if event.Rune() == 'H' {
			// 'H'キーでコンテンツを大きく左にスクロール
			_, frontPage := pages.GetFrontPage()
			if tv, ok := frontPage.(*tview.TextView); ok {
				row, col := tv.GetScrollOffset()
				newCol := col - 100
				if newCol < 0 {
					newCol = 0
				}
				tv.ScrollTo(row, newCol)
			} else if tb, ok := frontPage.(*tview.Table); ok {
				row, col := tb.GetOffset()
				newCol := col - 10
				if newCol < 0 {
					newCol = 0
				}
				tb.SetOffset(row, newCol)
			}
			return nil
		} else if event.Rune() == 'J' {
			// 'J'キーでコンテンツを下にスクロール
			_, frontPage := pages.GetFrontPage()
			if tv, ok := frontPage.(*tview.TextView); ok {
				row, col := tv.GetScrollOffset()
				tv.ScrollTo(row+10, col)
			} else if tb, ok := frontPage.(*tview.Table); ok {
				row, col := tb.GetOffset()
				tb.SetOffset(row+10, col)
			}
			return nil
		} else if event.Rune() == 'K' {
			// 'K'キーでコンテンツを上にスクロール
			_, frontPage := pages.GetFrontPage()
			if tv, ok := frontPage.(*tview.TextView); ok {
				row, col := tv.GetScrollOffset()
				newRow := row - 10
				if newRow < 0 {
					newRow = 0
				}
				tv.ScrollTo(newRow, col)
			} else if tb, ok := frontPage.(*tview.Table); ok {
				row, col := tb.GetOffset()
				newRow := row - 10
				if newRow < 0 {
					newRow = 0
				}
				tb.SetOffset(newRow, col)
			}
			return nil
		} else if event.Rune() == 'L' {
			// 'L'キーでコンテンツを大きく右にスクロール
			_, frontPage := pages.GetFrontPage()
			if tv, ok := frontPage.(*tview.TextView); ok {
				row, col := tv.GetScrollOffset()
				tv.ScrollTo(row, col+100)
			} else if tb, ok := frontPage.(*tview.Table); ok {
				row, col := tb.GetOffset()
				tb.SetOffset(row, col+10)
			}
			return nil
		} else if event.Rune() == 'n' {
			// 'n'キーで次の差分へスクロール
			_, frontPage := pages.GetFrontPage()
			if tb, ok := frontPage.(*tview.Table); ok {
				row, col := tb.GetOffset()
				searchR := row + 1
				searchC := col + 1
				rowCount := tb.GetRowCount()
				colCount := tb.GetColumnCount()
				found := false
				for r := searchR; r < rowCount; r++ {
					startCol := 1
					if r == searchR {
						startCol = searchC + 1
					}
					for c := startCol; c < colCount; c++ {
						cell := tb.GetCell(r, c)
						if cell != nil && (strings.Contains(cell.Text, "[#d55e00]") || strings.Contains(cell.Text, "[#56b4e9]") || strings.Contains(cell.Text, "[#f0e442]")) {
							targetR := r - 1
							if targetR < 0 {
								targetR = 0
							}
							targetC := c - 1
							if targetC < 0 {
								targetC = 0
							}
							tb.SetOffset(targetR, targetC)
							found = true
							break
						}
					}
					if found {
						break
					}
				}
			} else if tv, ok := frontPage.(*tview.TextView); ok {
				text := tv.GetText(false)
				lines := strings.Split(text, "\n")
				row, _ := tv.GetScrollOffset()
				searchR := row + 1
				for r := searchR + 1; r < len(lines); r++ {
					if strings.Contains(lines[r], "[#d55e00]") || strings.Contains(lines[r], "[#56b4e9]") || strings.Contains(lines[r], "[#f0e442]") {
						targetR := r - 1
						if targetR < 0 {
							targetR = 0
						}
						tv.ScrollTo(targetR, 0)
						break
					}
				}
			}
			return nil
		} else if event.Rune() == 'N' {
			// 'N'キーで前の差分へスクロール
			_, frontPage := pages.GetFrontPage()
			if tb, ok := frontPage.(*tview.Table); ok {
				row, col := tb.GetOffset()
				searchR := row + 1
				searchC := col + 1
				colCount := tb.GetColumnCount()
				found := false
				for r := searchR; r >= 0; r-- {
					startCol := colCount - 1
					if r == searchR {
						startCol = searchC - 1
					}
					for c := startCol; c >= 1; c-- {
						cell := tb.GetCell(r, c)
						if cell != nil && (strings.Contains(cell.Text, "[#d55e00]") || strings.Contains(cell.Text, "[#56b4e9]") || strings.Contains(cell.Text, "[#f0e442]")) {
							targetR := r - 1
							if targetR < 0 {
								targetR = 0
							}
							targetC := c - 1
							if targetC < 0 {
								targetC = 0
							}
							tb.SetOffset(targetR, targetC)
							found = true
							break
						}
					}
					if found {
						break
					}
				}
			} else if tv, ok := frontPage.(*tview.TextView); ok {
				text := tv.GetText(false)
				lines := strings.Split(text, "\n")
				row, _ := tv.GetScrollOffset()
				searchR := row + 1
				for r := searchR - 1; r >= 0; r-- {
					if strings.Contains(lines[r], "[#d55e00]") || strings.Contains(lines[r], "[#56b4e9]") || strings.Contains(lines[r], "[#f0e442]") {
						targetR := r - 1
						if targetR < 0 {
							targetR = 0
						}
						tv.ScrollTo(targetR, 0)
						break
					}
				}
			}
			return nil
		}
		return event
	})

	// 左スクロール用のボタン（マウス操作用）
	leftBtn := tview.NewButton("◀").
		SetStyle(tcell.StyleDefault.Background(tcell.ColorDefault)).
		SetActivatedStyle(tcell.StyleDefault.Background(tcell.ColorDefault)).
		SetSelectedFunc(func() {
			for range 100 {
				app.QueueEvent(tcell.NewEventKey(tcell.KeyLeft, 0, tcell.ModNone))
			}
			app.SetFocus(lastFocus)
		})

	// 右スクロール用のボタン（マウス操作用）
	rightBtn := tview.NewButton("▶").
		SetStyle(tcell.StyleDefault.Background(tcell.ColorDefault)).
		SetActivatedStyle(tcell.StyleDefault.Background(tcell.ColorDefault)).
		SetSelectedFunc(func() {
			for range 100 {
				app.QueueEvent(tcell.NewEventKey(tcell.KeyRight, 0, tcell.ModNone))
			}
			app.SetFocus(lastFocus)
		})

	// ボタンとメインレイアウトを横に並べるラッパーを作成
	wrapper := tview.NewFlex().
		SetDirection(tview.FlexColumn).
		AddItem(leftBtn, 3, 0, false).
		AddItem(layout, 0, 1, true).
		AddItem(rightBtn, 3, 0, false)

	return wrapper
}

// ファイル比較用のTUIアプリケーションを構築して表示する
func DisplayFileTui(results []SheetResult) error {
	app := tview.NewApplication()
	// シートタブ画面を構築
	layout := CreateSheetTabs(app, results)

	// ヘルプテキスト（操作説明）の作成
	helpText1 := " [#f0e442]Tab[-]: Switch tab | [#f0e442]b / f[-]: Scroll tab | [#f0e442]h / j / k / l[-]: Scroll text | [#f0e442]n[-]: Next diff | [#f0e442]g[-]: Go to edge | [#f0e442]q[-]: Quit "
	helpBar1 := tview.NewTextView().
		SetDynamicColors(true).
		SetText(helpText1).
		SetTextAlign(tview.AlignCenter)
	helpBar1.SetBackgroundColor(tcell.ColorDefault)

	helpText2 := "Hold Shift to change the key behavior."
	helpBar2 := tview.NewTextView().
		SetDynamicColors(true).
		SetText(helpText2).
		SetTextAlign(tview.AlignCenter)
	helpBar2.SetBackgroundColor(tcell.ColorDefault)

	// 全体のレイアウトを構築（メイン画面 + ヘルプテキスト）
	rootLayout := tview.NewFlex().SetDirection(tview.FlexRow).
		AddItem(layout, 0, 1, true).
		AddItem(helpBar1, 1, 1, false).
		AddItem(helpBar2, 1, 1, false)

	// アプリケーション全体のキーバインド（終了操作）
	app.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		if event.Key() == tcell.KeyEscape || event.Rune() == 'q' {
			app.Stop()
			return nil
		}
		return event
	})

	// アプリケーションを実行
	return app.SetRoot(rootLayout, true).EnableMouse(true).Run()
}

// ディレクトリ比較用の2ペインTUIを表示する
func DisplayDirTui(books []BookResult) error {
	app := tview.NewApplication()

	// 左ペイン：ファイル一覧を表示するリスト
	list := tview.NewList().ShowSecondaryText(false)
	list.SetBorder(true).SetTitle(" File Differences ")
	list.SetBackgroundColor(tcell.ColorDefault)
	list.SetMainTextStyle(tcell.StyleDefault.Background(tcell.ColorDefault))
	list.SetSelectedStyle(tcell.StyleDefault.Reverse(true))

	// hjklキーでリストをスクロールできるようにする（Vimライクな操作）
	list.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		switch event.Rune() {
		case 'j':
			return tcell.NewEventKey(tcell.KeyDown, 0, tcell.ModNone)
		case 'k':
			return tcell.NewEventKey(tcell.KeyUp, 0, tcell.ModNone)
		case 'h':
			return tcell.NewEventKey(tcell.KeyLeft, 0, tcell.ModNone)
		case 'l':
			return tcell.NewEventKey(tcell.KeyRight, 0, tcell.ModNone)
		}
		return event
	})

	// 右ペイン：選択されたファイルのシート内容を表示するページビュー
	rightPages := tview.NewPages()

	// 各ブック（ファイル）のデータをリストと右ペインに追加する
	for i, book := range books {
		bookPageID := fmt.Sprintf("book_%d", i)
		// ブックごとのシートタブ画面を構築
		sheetTabs := CreateSheetTabs(app, book.Sheets)
		// 右ペインにページとして追加（最初のブックのみ表示）
		rightPages.AddPage(bookPageID, sheetTabs, true, i == 0)
		// リストにファイル名を追加
		list.AddItem(book.FileName, "", 0, nil)
	}

	// リストの選択が変更されたら、右ペインの表示を切り替える
	list.SetChangedFunc(func(index int, mainText string, secondaryText string, shortcut rune) {
		rightPages.SwitchToPage(fmt.Sprintf("book_%d", index))
	})

	// ヘルプテキスト（操作説明）の作成
	helpText1 := " [#f0e442]Space[-]: Switch pain | [#f0e442]Tab[-]: Switch file / sheet | [#f0e442]b / f[-]: Scroll tab | [#f0e442]h / j / k / l[-]: Scroll text | [#f0e442]n[-]: Next diff | [#f0e442]g[-]: Go to edge | [#f0e442]q[-]: Quit "
	helpBar1 := tview.NewTextView().
		SetDynamicColors(true).
		SetText(helpText1).
		SetTextAlign(tview.AlignCenter)
	helpBar1.SetBackgroundColor(tcell.ColorDefault)

	helpText2 := "Hold Shift to change the key behavior."
	helpBar2 := tview.NewTextView().
		SetDynamicColors(true).
		SetText(helpText2).
		SetTextAlign(tview.AlignCenter)
	helpBar2.SetBackgroundColor(tcell.ColorDefault)

	// メインレイアウト（左ペインと右ペインを横に並べる）
	mainLayout := tview.NewFlex().
		AddItem(list, 30, 1, true).
		AddItem(rightPages, 0, 3, false)

	// 全体のレイアウト（メインレイアウト + ヘルプテキスト）
	rootLayout := tview.NewFlex().SetDirection(tview.FlexRow).
		AddItem(mainLayout, 0, 1, true).
		AddItem(helpBar1, 1, 1, false).
		AddItem(helpBar2, 1, 1, false)

	// アプリケーション全体のキーバインド
	app.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		// Spaceキーで左ペイン（リスト）と右ペイン（シート内容）のフォーカスを切り替える
		if event.Rune() == ' ' {
			if app.GetFocus() == list {
				app.SetFocus(rightPages)
			} else {
				app.SetFocus(list)
			}
			return nil
		} else if event.Key() == tcell.KeyEscape || event.Rune() == 'q' {
			// Escキーまたは'q'キーで終了
			app.Stop()
			return nil
		}
		return event
	})

	// アプリケーションを実行
	return app.SetRoot(rootLayout, true).EnableMouse(true).Run()
}
