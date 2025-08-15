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

		err = os.RemoveAll(tempDir)
		if err != nil {
			fmt.Fprintf(os.Stderr, "一時ファイルの削除中にエラーが発生しました: %v\n", err)
			os.Exit(1)
		}

		fmt.Println("一時ファイルを削除しました。")
	},
}

func init() {
	rootCmd.AddCommand(clearCmd)
}
