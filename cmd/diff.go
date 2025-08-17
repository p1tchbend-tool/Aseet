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
		return "", fmt.Errorf("error getting user cache dir: %v", err)
	}
	// aseet用の一時ディレクトリを作成
	tempDir := filepath.Join(cacheDir, "aseet", "temp")
	if err := os.MkdirAll(tempDir, 0755); err != nil {
		return "", fmt.Errorf("error creating temp dir: %v", err)
	}

	// コピー先のファイルパスを生成
	baseName := filepath.Base(remotePath)
	newFileName := "[REMOTE]_" + baseName
	destPath := filepath.Join(tempDir, newFileName)

	// ファイルをコピー
	sourceFile, err := os.Open(remotePath)
	if err != nil {
		return "", fmt.Errorf("error opening source file for copy: %v", err)
	}
	defer sourceFile.Close()

	destFile, err := os.Create(destPath)
	if err != nil {
		return "", fmt.Errorf("error creating destination file for copy: %v", err)
	}
	defer destFile.Close()

	_, err = io.Copy(destFile, sourceFile)
	if err != nil {
		return "", fmt.Errorf("error copying file: %v", err)
	}

	return destPath, nil
}

func contains(slice []string, item string) bool {
	for _, a := range slice {
		if a == item {
			return true
		}
	}
	return false
}

var openFiles bool
var diffFormula bool

var diffCmd = &cobra.Command{
	Use:   "diff [file1] [file2]",
	Short: "Show the difference in sheet names and header row content between two excel files",
	Long:  `Show the difference in sheet names and header row content between two excel files. This command compares header columns by their content, accounting for additions and deletions. The header row is identified by scanning the first 100 rows. Empty cells in header rows are ignored. It also compares data rows cell by cell for columns with matching headers. With the -f flag, it prioritizes formulas over calculated values.`,
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

		// 両方のファイルに存在するすべてのシート名を重複なく集め、ソートする
		var allSheets []string
		allSheets = append(allSheets, sheets1...)
		for _, s := range sheets2 {
			if !contains(sheets1, s) {
				allSheets = append(allSheets, s)
			}
		}
		sort.Strings(allSheets)

		// すべてのシートをループして比較
		for _, sheet := range allSheets {
			existsIn1 := contains(sheets1, sheet)
			existsIn2 := contains(sheets2, sheet)

			// シートが片方のファイルにしか存在しない場合の処理
			if !existsIn1 {
				fmt.Println("================================================================================")
				fmt.Println(sheet)
				fmt.Println("================================================================================")
				fmt.Printf("Sheet '%s' only in %s\n\n", sheet, remotePath)
				continue
			}
			if !existsIn2 {
				fmt.Println("================================================================================")
				fmt.Println(sheet)
				fmt.Println("================================================================================")
				fmt.Printf("Sheet '%s' only in %s\n\n", sheet, localPath)
				continue
			}

			// 両方のファイルに存在するシートの比較ロジック
			// 各シートからヘッダー行を特定
			row1, _, err1 := findHeaderRow(f1, sheet)
			if err1 != nil {
				fmt.Fprintf(os.Stderr, "Error reading sheet %s from %s: %v\n", sheet, localPath, err1)
				continue
			}

			row2, _, err2 := findHeaderRow(f2, sheet)
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
			unmatchedColumnMap := make(map[int]bool)
			for val, count1 := range map1 {
				count2 := map2[val]
				if count1 > count2 {
					for i := 0; i < count1-count2; i++ {
						onlyInFile1 = append(onlyInFile1, val)
					}
					for i, h := range row1 {
						if h == val {
							unmatchedColumnMap[i+1] = true
						}
					}
				}
			}
			for val, count2 := range map2 {
				count1 := map1[val]
				if count2 > count1 {
					for i := 0; i < count2-count1; i++ {
						onlyInFile2 = append(onlyInFile2, val)
					}
					for i, h := range row2 {
						if h == val {
							unmatchedColumnMap[i+1] = true
						}
					}
				}
			}

			isShownSheetName := false

			// ヘッダーに差分があれば結果を出力
			if len(onlyInFile1) > 0 || len(onlyInFile2) > 0 {
				if len(onlyInFile1) > 0 {
					fmt.Println("================================================================================")
					fmt.Println(sheet)
					fmt.Println("================================================================================")
					for _, s := range onlyInFile1 {
						fmt.Printf("Columns '%s' only in %s\n", s, localPath)
					}
					isShownSheetName = true
				}

				if len(onlyInFile2) > 0 {
					if !isShownSheetName {
						fmt.Println("================================================================================")
						fmt.Println(sheet)
						fmt.Println("================================================================================")
					}
					for _, s := range onlyInFile2 {
						fmt.Printf("Columns '%s' only in %s\n", s, remotePath)
					}
					isShownSheetName = true
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

			// 比較する最大行数を決定
			maxRows := len(allRows1)
			if len(allRows2) > maxRows {
				maxRows = len(allRows2)
			}

			isRowContentDiff := false
			// 1行ずつデータを比較
			for i := 0; i < maxRows; i++ {
				physicalRowNum := i + 1
				isRowHasDiff := false

				var row1, row2 []string
				if i < len(allRows1) {
					row1 = allRows1[i]
				}
				if i < len(allRows2) {
					row2 = allRows2[i]
				}

				maxCols := len(row1)
				if len(row2) > maxCols {
					maxCols = len(row2)
				}

				var row1Vals, row2Vals []string
				for j := 0; j < maxCols; j++ {
					physicalColNum := j + 1

					var val1, val2 string
					// Get value from file 1
					cellName1, _ := excelize.CoordinatesToCellName(physicalColNum, physicalRowNum)
					if diffFormula {
						val1, _ = f1.GetCellFormula(sheet, cellName1)
						if val1 == "" {
							val1, _ = f1.GetCellValue(sheet, cellName1)
						}
					} else {
						val1, _ = f1.GetCellValue(sheet, cellName1)
					}
					row1Vals = append(row1Vals, val1)

					// Get value from file 2
					cellName2, _ := excelize.CoordinatesToCellName(physicalColNum, physicalRowNum)
					if diffFormula {
						val2, _ = f2.GetCellFormula(sheet, cellName2)
						if val2 == "" {
							val2, _ = f2.GetCellValue(sheet, cellName2)
						}
					} else {
						val2, _ = f2.GetCellValue(sheet, cellName2)
					}
					row2Vals = append(row2Vals, val2)

					// Compare if not in unmatched columns
					if _, exists := unmatchedColumnMap[physicalColNum]; !exists {
						if val1 != val2 {
							isRowHasDiff = true
						}
					}
				}

				// 行に差分があれば結果を出力
				if isRowHasDiff {
					if !isShownSheetName {
						fmt.Println("================================================================================")
						fmt.Println(sheet)
						fmt.Println("================================================================================")
						isShownSheetName = true
					}

					if !isRowContentDiff {
						fmt.Printf("Found differences in row content: [%s] vs [%s]\n", localPath, remotePath)
						isRowContentDiff = true
					}

					var formattedRow1Vals []string
					for _, v := range row1Vals {
						escapedV := strings.ReplaceAll(v, "\"", "\"\"")
						formattedRow1Vals = append(formattedRow1Vals, fmt.Sprintf("\"%s\"", escapedV))
					}
					row1Str := strings.Join(formattedRow1Vals, ",")

					var formattedRow2Vals []string
					for _, v := range row2Vals {
						escapedV := strings.ReplaceAll(v, "\"", "\"\"")
						formattedRow2Vals = append(formattedRow2Vals, fmt.Sprintf("\"%s\"", escapedV))
					}
					row2Str := strings.Join(formattedRow2Vals, ",")
					fmt.Printf("  - Row %d: [%s] vs [%s]\n", physicalRowNum, row1Str, row2Str)
				}
			}

			if isRowContentDiff {
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
	diffCmd.Flags().BoolVarP(&diffFormula, "formula", "f", false, "セルに数式がある場合は数式を比較対象にします。")
}
