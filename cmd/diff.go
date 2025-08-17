package cmd

import (
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"sort"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// findHeaderRow は、指定されたシートの最初の100行をスキャンしてヘッダー行を見つけます。
// ヘッダー行の候補は、最初と最後の空でないセルの間に空のセルがない行です。
// 最もセル数が多い候補がヘッダー行として採用されます。
func findHeaderRow(f *excelize.File, sheetName string) ([]string, int, error) {
	// シートからすべての行を取得
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, 0, err
	}

	var headerRow []string
	headerRowNum := 0
	maxCells := -1

	// スキャンする行数を決定（最大100行）
	numRowsToCheck := 100
	if len(rows) < numRowsToCheck {
		numRowsToCheck = len(rows)
	}

	// 指定された行数までループしてヘッダー行候補を探す
	for i := 0; i < numRowsToCheck; i++ {
		row := rows[i]
		if len(row) == 0 {
			continue
		}

		firstCellIdx := -1
		lastCellIdx := -1

		// 行内の最初と最後の空でないセルのインデックスを見つける
		for j, cell := range row {
			if cell != "" {
				if firstCellIdx == -1 {
					firstCellIdx = j
				}
				lastCellIdx = j
			}
		}

		// 行が実質的に空である場合、次の行へ
		if firstCellIdx == -1 {
			continue
		}

		// 最初と最後の空でないセルの間に空のセルがあるかチェック
		isCandidate := true
		sliceToCheck := row[firstCellIdx : lastCellIdx+1]
		for _, cell := range sliceToCheck {
			if cell == "" {
				isCandidate = false
				break
			}
		}

		if !isCandidate {
			continue
		}

		// この行はヘッダー行の候補。これまでで最もセル数が多いかチェック
		cellCount := len(sliceToCheck)
		if cellCount > maxCells {
			maxCells = cellCount
			headerRow = row
			headerRowNum = i + 1 // 1-basedの行番号
		}
	}

	// 適切なヘッダー行が見つからなかった場合
	if headerRowNum == 0 {
		return nil, 0, nil
	}

	return headerRow, headerRowNum, nil
}

// handleSameFileName は2つのファイル名が同じ場合（git diffなどでの利用を想定）、
// 2つ目のファイルを一時ディレクトリにコピーして比較対象とします。
// 新しいファイルパスとエラーを返します。
func handleSameFileName(localPath, remotePath string) (string, error) {
	if filepath.Base(localPath) != filepath.Base(remotePath) {
		return remotePath, nil
	}

	// ユーザーキャッシュディレクトリを取得
	cacheDir, err := os.UserCacheDir()
	if err != nil {
		return "", fmt.Errorf("Error getting user cache dir: %v", err)
	}
	// aseet用の一時ディレクトリを作成
	tempDir := filepath.Join(cacheDir, "aseet", "temp")
	if err := os.MkdirAll(tempDir, 0755); err != nil {
		return "", fmt.Errorf("Error creating temp dir: %v", err)
	}

	// コピー先のファイルパスを生成
	baseName := filepath.Base(remotePath)
	newFileName := "[REMOTE]_" + baseName
	destPath := filepath.Join(tempDir, newFileName)

	// ファイルをコピー
	sourceFile, err := os.Open(remotePath)
	if err != nil {
		return "", fmt.Errorf("Error opening source file for copy: %v", err)
	}
	defer sourceFile.Close()

	destFile, err := os.Create(destPath)
	if err != nil {
		return "", fmt.Errorf("Error creating destination file for copy: %v", err)
	}
	defer destFile.Close()

	_, err = io.Copy(destFile, sourceFile)
	if err != nil {
		return "", fmt.Errorf("Error copying file: %v", err)
	}

	return destPath, nil
}

var openFiles bool

var diffCmd = &cobra.Command{
	Use:   "diff [file1] [file2]",
	Short: "Show the difference in sheet names and header row content between two excel files",
	Long:  `Show the difference in sheet names and header row content between two excel files. This command compares header columns by their content, accounting for additions and deletions. The header row is identified by scanning the first 100 rows. Empty cells in header rows are ignored. It also compares data rows cell by cell for columns with matching headers, prioritizing formulas over calculated values.`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		localPath := args[0]
		remotePath := args[1]

		remotePath, err := handleSameFileName(localPath, remotePath)
		if err != nil {
			fmt.Fprintf(os.Stderr, "%v\n", err)
			os.Exit(1)
		}

		// 1つ目のExcelファイルを開く
		f1, err := excelize.OpenFile(localPath)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", localPath, err)
			os.Exit(1)
		}
		// 処理終了時にファイルをクローズする
		defer func() {
			if err := f1.Close(); err != nil {
				fmt.Fprintf(os.Stderr, "Error closing file %s: %v\n", localPath, err)
			}
		}()

		// 2つ目のExcelファイルを開く
		f2, err := excelize.OpenFile(remotePath)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", remotePath, err)
			os.Exit(1)
		}
		// 処理終了時にファイルをクローズする
		defer func() {
			if err := f2.Close(); err != nil {
				fmt.Fprintf(os.Stderr, "Error closing file %s: %v\n", remotePath, err)
			}
		}()

		// 各ファイルのシートリストを取得
		sheets1 := f1.GetSheetList()
		sheets2 := f2.GetSheetList()

		// シート名の存在を高速にチェックするためのマップを作成
		sheetMap1 := make(map[string]bool)
		for _, s := range sheets1 {
			sheetMap1[s] = true
		}
		sheetMap2 := make(map[string]bool)
		for _, s := range sheets2 {
			sheetMap2[s] = true
		}

		// 両方のファイルに存在するすべてのシート名を重複なく集め、ソートする
		allSheetsMap := make(map[string]bool)
		for _, s := range sheets1 {
			allSheetsMap[s] = true
		}
		for _, s := range sheets2 {
			allSheetsMap[s] = true
		}
		var allSheets []string
		for s := range allSheetsMap {
			allSheets = append(allSheets, s)
		}
		sort.Strings(allSheets)

		// すべてのシートをループして比較
		for _, sheet := range allSheets {
			_, existsIn1 := sheetMap1[sheet]
			_, existsIn2 := sheetMap2[sheet]

			// シートが片方のファイルにしか存在しない場合の処理
			if !existsIn1 {
				fmt.Printf("Sheet '%s' only in %s\n\n", sheet, remotePath)
				continue
			}
			if !existsIn2 {
				fmt.Printf("Sheet '%s' only in %s\n\n", sheet, localPath)
				continue
			}

			// 両方のファイルに存在するシートの比較ロジック
			// 各シートからヘッダー行を特定
			row1, rowNum1, err1 := findHeaderRow(f1, sheet)
			if err1 != nil {
				fmt.Fprintf(os.Stderr, "Error reading sheet %s from %s: %v\n", sheet, localPath, err1)
				continue
			}

			row2, rowNum2, err2 := findHeaderRow(f2, sheet)
			if err2 != nil {
				fmt.Fprintf(os.Stderr, "Error reading sheet %s from %s: %v\n", sheet, remotePath, err2)
				continue
			}

			// ヘッダー行を比較
			// まず、ヘッダー行から空のセルを除外したスライスを作成
			var r1NonEmpty []string
			for _, cell := range row1 {
				if cell != "" {
					r1NonEmpty = append(r1NonEmpty, cell)
				}
			}
			var r2NonEmpty []string
			for _, cell := range row2 {
				if cell != "" {
					r2NonEmpty = append(r2NonEmpty, cell)
				}
			}

			// 各ヘッダー項目（列名）の出現回数をマップに記録（重複列対応）
			map1 := make(map[string]int)
			for _, s := range r1NonEmpty {
				map1[s]++
			}
			map2 := make(map[string]int)
			for _, s := range r2NonEmpty {
				map2[s]++
			}

			// 各ファイルにのみ存在するヘッダー項目を特定
			var onlyInFile1, onlyInFile2 []string
			for val, count1 := range map1 {
				count2 := map2[val]
				if count1 > count2 {
					for i := 0; i < count1-count2; i++ {
						onlyInFile1 = append(onlyInFile1, val)
					}
				}
			}
			for val, count2 := range map2 {
				count1 := map1[val]
				if count2 > count1 {
					for i := 0; i < count2-count1; i++ {
						onlyInFile2 = append(onlyInFile2, val)
					}
				}
			}

			// ヘッダーに差分があれば結果を出力
			if len(onlyInFile1) > 0 || len(onlyInFile2) > 0 {
				fmt.Printf("Sheet '%s': Header row content mismatch. Comparing %s (Row %d) and %s (Row %d):\n", sheet, localPath, rowNum1, remotePath, rowNum2)
				if len(onlyInFile1) > 0 {
					fmt.Printf("  Columns only in %s:\n", localPath)
					for _, s := range onlyInFile1 {
						fmt.Printf("    - %s\n", s)
					}
				}
				if len(onlyInFile2) > 0 {
					fmt.Printf("  Columns only in %s:\n", remotePath)
					for _, s := range onlyInFile2 {
						fmt.Printf("    - %s\n", s)
					}
				}
				fmt.Println()
			}

			// ヘッダーが同一の場合、データ行を比較
			// シートからすべての行を取得
			allRows1, err := f1.GetRows(sheet)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error reading all rows from sheet %s in %s: %v\n", sheet, localPath, err)
				continue
			}
			allRows2, err := f2.GetRows(sheet)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error reading all rows from sheet %s in %s: %v\n", sheet, remotePath, err)
				continue
			}

			// ヘッダー名から列インデックスへのマップを作成
			header1Indices := make(map[string]int)
			for i, h := range row1 {
				if h != "" {
					if _, exists := header1Indices[h]; !exists {
						header1Indices[h] = i
					}
				}
			}
			header2Indices := make(map[string]int)
			for i, h := range row2 {
				if h != "" {
					if _, exists := header2Indices[h]; !exists {
						header2Indices[h] = i
					}
				}
			}

			// 両方のファイルに存在する共通のヘッダー名をスライスに格納
			var commonHeaderSlice []string
			for h := range header1Indices {
				if _, ok := header2Indices[h]; ok {
					commonHeaderSlice = append(commonHeaderSlice, h)
				}
			}
			sort.Strings(commonHeaderSlice)

			// 比較する最大行数を決定
			maxRows := len(allRows1)
			if len(allRows2) > maxRows {
				maxRows = len(allRows2)
			}

			rowContentDiff := false
			// 1行ずつデータを比較
			for i := 0; i < maxRows; i++ {
				physicalRowNum := i + 1

				// ヘッダー行自体は比較対象から除外
				if (rowNum1 > 0 && physicalRowNum == rowNum1) || (rowNum2 > 0 && physicalRowNum == rowNum2) {
					continue
				}

				rowHasDiff := false
				var row1Vals, row2Vals []string

				// 共通のヘッダー列についてセルを比較
				for _, hName := range commonHeaderSlice {
					idx1 := header1Indices[hName]
					idx2 := header2Indices[hName]

					// 比較対象のセルを特定
					cellName1, _ := excelize.CoordinatesToCellName(idx1+1, physicalRowNum)
					// まず数式を取得し、なければ値を取得する
					val1, _ := f1.GetCellFormula(sheet, cellName1)
					if val1 == "" {
						val1, _ = f1.GetCellValue(sheet, cellName1)
					}

					cellName2, _ := excelize.CoordinatesToCellName(idx2+1, physicalRowNum)
					val2, _ := f2.GetCellFormula(sheet, cellName2)
					if val2 == "" {
						val2, _ = f2.GetCellValue(sheet, cellName2)
					}

					// セルの値を比較
					if val1 != val2 {
						rowHasDiff = true
					}
					row1Vals = append(row1Vals, val1)
					row2Vals = append(row2Vals, val2)
				}

				// 行に差分があれば結果を出力
				if rowHasDiff {
					if !rowContentDiff {
						fmt.Printf("Sheet '%s': Found differences in row content:\n", sheet)
						rowContentDiff = true
					}
					row1Str := strings.Join(row1Vals, ", ")
					row2Str := strings.Join(row2Vals, ", ")
					fmt.Printf("  - Row %d: [%s] vs [%s]\n", physicalRowNum, row1Str, row2Str)
				}
			}
			if rowContentDiff {
				fmt.Println()
			}
		}

		// --open フラグが指定されている場合、比較した2つのファイルをアプリケーションで開く
		if openFiles {
			exec.Command("cmd", "/C", "start", localPath).Start()
			exec.Command("cmd", "/C", "start", remotePath).Start()
		}
	},
}

func init() {
	// diffコマンドをルートコマンドに追加
	rootCmd.AddCommand(diffCmd)
	// --open, -o フラグを定義
	diffCmd.Flags().BoolVarP(&openFiles, "open", "o", false, "最後に2つのファイルを関連付けられたアプリケーションで開きます。")
}
