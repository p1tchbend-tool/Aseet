package cmd

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var touchCmd = &cobra.Command{
	Use:   "touch [file]",
	Short: "Update the access and modification times of a file or create an empty Excel file",
	Long:  `Update the access and modification times of a file to the current time. If the file does not exist, create an empty .xlsx file. If the extension is not .xlsx, it will be appended automatically.`,
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		filePath := args[0]

		// 拡張子が.xlsxでない場合は補完する
		if strings.ToLower(filepath.Ext(filePath)) != ".xlsx" {
			filePath += ".xlsx"
		}

		// ファイルが存在するか確認する
		if _, err := os.Stat(filePath); os.IsNotExist(err) {
			// ファイルが存在しない場合は空のxlsxファイルを作成する
			f := excelize.NewFile()
			defer f.Close()

			if err := f.SaveAs(filePath); err != nil {
				fmt.Printf("Error creating file %s: %v\n", filePath, err)
				os.Exit(1)
			}
			fmt.Printf("Created empty Excel file: %s\n", filePath)
		} else {
			// ファイルが存在する場合はタイムスタンプを更新する
			currentTime := time.Now()
			if err := os.Chtimes(filePath, currentTime, currentTime); err != nil {
				fmt.Printf("Error updating timestamp for %s: %v\n", filePath, err)
				os.Exit(1)
			}
			fmt.Printf("Updated timestamp for: %s\n", filePath)
		}
	},
}

func init() {
	// コマンドを登録する
	rootCmd.AddCommand(touchCmd)
}
