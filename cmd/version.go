package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

// Version はアプリケーションのバージョン情報を保持する
var Version = "0.6.0"

var versionCmd = &cobra.Command{
	Use:   "version",
	Short: "Print the version number of aseet",
	Long:  `Print the version number of aseet.`,
	Run: func(cmd *cobra.Command, args []string) {
		// バージョン情報を出力する
		fmt.Printf("aseet version %s\n", Version)
	},
}

func init() {
	// ルートコマンドにversionコマンドを追加する
	rootCmd.AddCommand(versionCmd)
}
