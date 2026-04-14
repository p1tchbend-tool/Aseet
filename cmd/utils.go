package cmd

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"

	"github.com/xuri/excelize/v2"
)

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
			costDel := dp[i-1][j] + countNonEmpty(a[i-1])             // 削除コスト
			costIns := dp[i][j-1] + countNonEmpty(b[j-1])             // 挿入コスト
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
