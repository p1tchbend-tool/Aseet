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
