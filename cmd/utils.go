package cmd

import (
	"fmt"
	"os"
	"os/exec"
	"runtime"
	"strings"

	"github.com/xuri/excelize/v2"
)

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
