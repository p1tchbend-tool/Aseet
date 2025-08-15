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
	Long:  `asheet/temp フォルダにあるすべての一時ファイルを削除します。`,
	Args:  cobra.NoArgs,
	Run: func(cmd *cobra.Command, args []string) {
		cacheDir, err := os.UserCacheDir()
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error getting user cache dir: %v\n", err)
			os.Exit(1)
		}
		tempDir := filepath.Join(cacheDir, "asheet", "temp")

		if _, err := os.Stat(tempDir); os.IsNotExist(err) {
			fmt.Println("一時ファイルはありません。")
			return
		}

		files, err := os.ReadDir(tempDir)
		if err != nil {
			fmt.Fprintf(os.Stderr, "一時ディレクトリの読み込み中にエラーが発生しました: %v\n", err)
			os.Exit(1)
		}

		if len(files) == 0 {
			fmt.Println("一時ファイルはありません。")
			return
		}

		for _, file := range files {
			filePath := filepath.Join(tempDir, file.Name())
			if err := os.RemoveAll(filePath); err != nil {
				fmt.Fprintf(os.Stderr, "ファイル '%s' の削除中にエラーが発生しました: %v\n", filePath, err)
			}
		}

		fmt.Println("一時ファイルを削除しました。")
	},
}

func init() {
	rootCmd.AddCommand(clearCmd)
}
