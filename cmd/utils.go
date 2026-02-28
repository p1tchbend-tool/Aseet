package cmd

import (
	"fmt"
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

// シートの内容を文字列として取得する
func getSheetContents(f *excelize.File, sheetName string, isFormula bool) (string, error) {
	// シートのすべての行を取得する
	rows, err := f.GetRows(sheetName)
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
	for r, row := range rows {
		var outputCells []string
		// 最大列数に合わせて各セルをループ処理する
		for c := 0; c < maxCols; c++ {
			var value string
			if c < len(row) {
				value = row[c]
			}

			// セルの座標からセル名（例: A1）を取得する
			cellName, _ := excelize.CoordinatesToCellName(c+1, r+1)

			// 数式を取得する
			formula, err := f.GetCellFormula(sheetName, cellName)

			// 数式フラグが有効かつ数式が存在する場合
			if isFormula && err == nil && formula != "" {
				outputCells = append(outputCells, escapeCSVField(formula))
			} else {
				// 値を取得する
				outputCells = append(outputCells, escapeCSVField(value))
			}
		}
		// セルの値をカンマ区切りで結合し、改行を追加する
		sb.WriteString(strings.Join(outputCells, ",") + "\n")
	}
	return sb.String(), nil
}
