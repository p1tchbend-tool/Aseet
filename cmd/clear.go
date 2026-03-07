package cmd

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/spf13/cobra"
)

var clearCmd = &cobra.Command{
	Use:   "clear",
	Short: "Delete temporary files",
	Long:  `Delete all temporary files in the aseet folder.`,
	Args:  cobra.NoArgs,
	Run: func(cmd *cobra.Command, args []string) {
		// ユーザーのキャッシュディレクトリを取得
		cacheDir, _ := os.UserCacheDir()
		// aseet用の一時フォルダのパスを構築
		tempDir := filepath.Join(cacheDir, "aseet")

		// 一時フォルダが存在しない場合は終了
		if _, err := os.Stat(tempDir); os.IsNotExist(err) {
			fmt.Println("No temporary files found.")
			return
		}

		// 一時フォルダ内のファイル一覧を取得
		files, _ := os.ReadDir(tempDir)
		// フォルダが空の場合は終了
		if len(files) == 0 {
			fmt.Println("No temporary files found.")
			return
		}

		// フォルダ内の各ファイル・ディレクトリを削除
		for _, file := range files {
			filePath := filepath.Join(tempDir, file.Name())
			if err := os.RemoveAll(filePath); err != nil {
				// 削除に失敗した場合
				fmt.Printf("Error occurred while deleting file '%s'\n", filePath)
			} else {
				// 削除に成功した場合
				fmt.Printf("Deleted file '%s'\n", filePath)
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(clearCmd)
}
