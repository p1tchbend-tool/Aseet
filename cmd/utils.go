package cmd

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"

	"github.com/gdamore/tcell/v2"
	"github.com/rivo/tview"
	"github.com/xuri/excelize/v2"
)

// TUIで表示する各タブ（シート）のデータを保持する構造体
type sheetResult struct {
	title   string // タブに表示されるタイトル（シート名など）
	content string // タブ内に表示されるコンテンツ（CSV形式の文字列など）
}

// ディレクトリ比較時の各ファイル（ブック）のデータを保持する構造体
type bookResult struct {
	fileName string        // リストに表示されるファイル名
	sheets   []sheetResult // ファイルに含まれるシートのデータ
}

// シートのタブ画面（メインコンテンツ）を構築する共通モジュール
// app: tviewアプリケーションのインスタンス
// results: 表示するシートのデータリスト
func createSheetTabs(app *tview.Application, results []sheetResult) tview.Primitive {
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
		tabTitles = append(tabTitles, fmt.Sprintf(`["%s"] %s [""]`, pageID, res.title))

		// シートの内容を表示するテキストビューを作成
		textView := tview.NewTextView().
			SetDynamicColors(true).
			SetText(res.content).
			SetScrollable(true).
			SetWrap(false)
		textView.SetBackgroundColor(tcell.ColorDefault)

		// テキストビューにフォーカスが当たった際の処理
		textView.SetFocusFunc(func() {
			lastFocus = pages
		})

		// 最初のページのみ表示状態にする
		pages.AddPage(pageID, textView, true, i == 0)
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
			// 'H'キーでテキストビューを大きく左にスクロール
			_, frontPage := pages.GetFrontPage()
			if tv, ok := frontPage.(*tview.TextView); ok {
				row, col := tv.GetScrollOffset()
				newCol := col - 100
				if newCol < 0 {
					newCol = 0
				}
				tv.ScrollTo(row, newCol)
			}
			return nil
		} else if event.Rune() == 'J' {
			// 'J'キーでテキストビューを下にスクロール
			_, frontPage := pages.GetFrontPage()
			if tv, ok := frontPage.(*tview.TextView); ok {
				row, col := tv.GetScrollOffset()
				tv.ScrollTo(row+10, col)
			}
			return nil
		} else if event.Rune() == 'K' {
			// 'K'キーでテキストビューを上にスクロール
			_, frontPage := pages.GetFrontPage()
			if tv, ok := frontPage.(*tview.TextView); ok {
				row, col := tv.GetScrollOffset()
				newRow := row - 10
				if newRow < 0 {
					newRow = 0
				}
				tv.ScrollTo(newRow, col)
			}
			return nil
		} else if event.Rune() == 'L' {
			// 'L'キーでテキストビューを大きく右にスクロール
			_, frontPage := pages.GetFrontPage()
			if tv, ok := frontPage.(*tview.TextView); ok {
				row, col := tv.GetScrollOffset()
				tv.ScrollTo(row, col+100)
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
func displayFileTui(results []sheetResult) error {
	app := tview.NewApplication()
	// シートタブ画面を構築
	layout := createSheetTabs(app, results)

	// ヘルプテキスト（操作説明）の作成
	helpText1 := " [#f0e442]Tab[-]: Switch tab | [#f0e442]b / f[-]: Scroll tab | [#f0e442]h / j / k / l[-]: Scroll text | [#f0e442]g[-]: Scroll text to edge | [#f0e442]q[-]: Quit "
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
func displayDirTui(books []bookResult) error {
	app := tview.NewApplication()

	// 左ペイン：ファイル一覧を表示するリスト
	list := tview.NewList().ShowSecondaryText(false)
	list.SetBorder(true).SetTitle(" File Differences ")
	list.SetBackgroundColor(tcell.ColorDefault)
	list.SetMainTextStyle(tcell.StyleDefault.Background(tcell.ColorDefault))

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
		sheetTabs := createSheetTabs(app, book.sheets)
		// 右ペインにページとして追加（最初のブックのみ表示）
		rightPages.AddPage(bookPageID, sheetTabs, true, i == 0)
		// リストにファイル名を追加
		list.AddItem(book.fileName, "", 0, nil)
	}

	// リストの選択が変更されたら、右ペインの表示を切り替える
	list.SetChangedFunc(func(index int, mainText string, secondaryText string, shortcut rune) {
		rightPages.SwitchToPage(fmt.Sprintf("book_%d", index))
	})

	// ヘルプテキスト（操作説明）の作成
	helpText1 := " [#f0e442]Space[-]: Switch pain | [#f0e442]Tab[-]: Switch file / sheet | [#f0e442]b / f[-]: Scroll tab | [#f0e442]h / j / k / l[-]: Scroll text | [#f0e442]g[-]: Scroll text to edge | [#f0e442]q[-]: Quit "
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

// 対応するExcelファイルの拡張子かどうかを判定する
func isExcelFile(ext string) bool {
	return ext == ".xlsx" || ext == ".xlsm" || ext == ".xlam" || ext == ".xltm" || ext == ".xltx"
}

// セルの値にカンマ、改行、ダブルクォーテーションが含まれる場合はCSV形式としてエスケープ処理を行う
func escapeCSVField(value string) string {
	needsQuotes := strings.Contains(value, "\"") || strings.Contains(value, "\n") || strings.Contains(value, ",")

	// 1. ダブルクォーテーションを2つにする（エスケープ）
	value = strings.ReplaceAll(value, "\"", "\"\"")
	// 2. 改行コードが含まれる場合、文字としての "\n" に変換する（表示崩れを防ぐため）
	value = strings.ReplaceAll(value, "\n", "\\n")

	// 3. ダブルクォーテーション・改行・カンマのいずれかが含まれていた場合は、フィールド全体をダブルクォーテーションで囲む
	if needsQuotes {
		return fmt.Sprintf("\"%s\"", value)
	}
	return value
}

// シートのデータを2次元配列として取得する。isFormulaがtrueの場合は値ではなく数式を取得する
func getSheetData(f *excelize.File, sheetName string, isFormula bool) ([][]string, error) {
	// シートの全行を取得
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	// 数式を取得しない場合はそのまま返す
	if !isFormula {
		return rows, nil
	}

	var result [][]string
	// 各セルについて数式を取得する処理
	for r, row := range rows {
		var newRow []string
		for c, val := range row {
			// 列番号と行番号からセル名（例: "A1"）を生成
			cellName, err := excelize.CoordinatesToCellName(c+1, r+1)
			if err == nil {
				// セルの数式を取得
				formula, err := f.GetCellFormula(sheetName, cellName)
				if err == nil && formula != "" {
					// 数式が存在する場合は数式を格納
					newRow = append(newRow, formula)
					continue
				}
			}
			// 数式がない場合やエラーの場合は元の値を格納
			newRow = append(newRow, val)
		}
		result = append(result, newRow)
	}
	return result, nil
}

// シートの内容をCSV形式の文字列として取得する
func getSheetContents(f *excelize.File, sheetName string, isFormula bool) (string, error) {
	// シートのデータを2次元配列で取得
	rows, err := getSheetData(f, sheetName, isFormula)
	if err != nil {
		return "", err
	}

	// シート内の最大列数を計算（行によって列数が異なる場合があるため）
	maxCols := 0
	for _, row := range rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}

	var sb strings.Builder
	// 各行をループ処理してCSV文字列を構築
	for _, row := range rows {
		var outputCells []string
		// 最大列数に合わせて各セルをループ処理（足りない列は空文字で埋める）
		for c := 0; c < maxCols; c++ {
			var value string
			if c < len(row) {
				value = row[c]
			}
			// セルの値をエスケープ処理して追加
			outputCells = append(outputCells, escapeCSVField(value))
		}
		// セルの値をカンマ区切りで結合し、改行を追加
		sb.WriteString(strings.Join(outputCells, ",") + "\n")
	}
	return sb.String(), nil
}

// ファイルをコピーする（一時ファイルの作成などに使用）
func copyFile(src, dst string) error {
	input, err := os.ReadFile(src)
	if err != nil {
		return err
	}
	return os.WriteFile(dst, input, 0644)
}

// OSの関連付けられた既定のアプリケーションでファイルを開く
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

// 行内の空でないセルの数をカウントする（アライメントのコスト計算に使用）
func countNonEmpty(row []string) int {
	c := 0
	for _, v := range row {
		if v != "" {
			c++
		}
	}
	return c
}

// 2次元配列（マトリックス）を転置する（行と列を入れ替える）
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

// 2つの行を比較し、不一致要素数（コスト）を計算する
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

// 動的計画法（DP）を用いて2つの2次元配列のアライメント（差分パス）を計算する
// 行の挿入・削除を考慮して、最も変更コストが少ない対応関係を見つける
func align(a, b [][]string) [][2]int {
	n, m := len(a), len(b)
	// DPテーブルの初期化
	dp := make([][]int, n+1)
	for i := range dp {
		dp[i] = make([]int, m+1)
	}

	// 初期化：aの要素をすべて削除する場合のコスト
	for i := 1; i <= n; i++ {
		dp[i][0] = dp[i-1][0] + countNonEmpty(a[i-1])
	}
	// 初期化：bの要素をすべて挿入する場合のコスト
	for j := 1; j <= m; j++ {
		dp[0][j] = dp[0][j-1] + countNonEmpty(b[j-1])
	}

	// DPテーブルを埋める（最小コストを計算）
	for i := 1; i <= n; i++ {
		for j := 1; j <= m; j++ {
			costDel := dp[i-1][j] + countNonEmpty(a[i-1])           // 削除コスト
			costIns := dp[i][j-1] + countNonEmpty(b[j-1])           // 挿入コスト
			costMatch := dp[i-1][j-1] + calcMatchCost(a[i-1], b[j-1]) // 置換（マッチ）コスト

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

	// バックトラックして最適なパス（アライメント結果）を復元する
	var path [][2]int
	i, j := n, m
	for i > 0 || j > 0 {
		if i > 0 && j > 0 {
			// マッチ（置換）のパスから来た場合
			if dp[i][j] == dp[i-1][j-1]+calcMatchCost(a[i-1], b[j-1]) {
				path = append([][2]int{{i - 1, j - 1}}, path...)
				i--
				j--
				continue
			}
		}
		// 削除のパスから来た場合
		if i > 0 && dp[i][j] == dp[i-1][j]+countNonEmpty(a[i-1]) {
			path = append([][2]int{{i - 1, -1}}, path...)
			i--
		} else {
			// 挿入のパスから来た場合
			path = append([][2]int{{-1, j - 1}}, path...)
			j--
		}
	}
	return path
}

// 指定されたディレクトリ内のExcelファイルを探索してパスの配列を返す
// recursiveがtrueの場合はサブディレクトリも再帰的に探索する
func findExcelFiles(dirPath string, recursive bool) []string {
	var files []string

	if recursive {
		// 再帰的にディレクトリを探索
		_ = filepath.Walk(dirPath, func(p string, info os.FileInfo, err error) error {
			if err != nil {
				// 探索中のエラー（アクセス権限など）は無視して続行する
				return nil
			}
			if !info.IsDir() {
				ext := strings.ToLower(filepath.Ext(p))
				if isExcelFile(ext) {
					files = append(files, p)
				}
			}
			return nil
		})
	} else {
		// 指定されたディレクトリ直下のみを探索
		entries, err := os.ReadDir(dirPath)
		if err == nil {
			for _, entry := range entries {
				if !entry.IsDir() {
					ext := strings.ToLower(filepath.Ext(entry.Name()))
					if isExcelFile(ext) {
						files = append(files, filepath.Join(dirPath, entry.Name()))
					}
				}
			}
		}
	}

	return files
}
