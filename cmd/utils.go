package cmd

import (
	"fmt"
	"strings"
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

// シートの内容をカンマ区切りの文字列としてフォーマットする
func formatSheetContents(rows [][]string) string {
	maxCols := 0
	for _, row := range rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}

	var output []string
	for _, row := range rows {
		var cells []string
		for c := 0; c < maxCols; c++ {
			val := ""
			if c < len(row) {
				val = row[c]
			}
			cells = append(cells, escapeCSVField(val))
		}
		output = append(output, strings.Join(cells, ","))
	}
	return strings.Join(output, "\n")
}
