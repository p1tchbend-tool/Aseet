package cmd

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/spf13/cobra"
)

var clearCmd = &cobra.Command{
	Use:   "clear",
	Short: "一時ファイルを削除します。",
	Long:  `aseet/temp フォルダにあるすべての一時ファイルを削除します。`,
	Args:  cobra.NoArgs,
	Run: func(cmd *cobra.Command, args []string) {
		cacheDir, _ := os.UserCacheDir()
		tempDir := filepath.Join(cacheDir, "aseet")

		if _, err := os.Stat(tempDir); os.IsNotExist(err) {
			fmt.Println("一時ファイルはありません。")
			return
		}

		files, _ := os.ReadDir(tempDir)
		if len(files) == 0 {
			fmt.Println("一時ファイルはありません。")
			return
		}

		for _, file := range files {
			filePath := filepath.Join(tempDir, file.Name())
			if err := os.RemoveAll(filePath); err != nil {
				fmt.Printf("ファイル '%s' の削除中にエラーが発生しました\n", filePath)
			} else {
				fmt.Printf("ファイル '%s' を削除しました\n", filePath)
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(clearCmd)
}
