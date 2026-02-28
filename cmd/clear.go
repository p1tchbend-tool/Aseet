package cmd

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/spf13/cobra"
)

var clearCmd = &cobra.Command{
	Use:   "clear",
	Short: "Delete temporary files.",
	Long:  `Delete all temporary files in the aseet folder.`,
	Args:  cobra.NoArgs,
	Run: func(cmd *cobra.Command, args []string) {
		cacheDir, _ := os.UserCacheDir()
		tempDir := filepath.Join(cacheDir, "aseet")

		if _, err := os.Stat(tempDir); os.IsNotExist(err) {
			fmt.Println("No temporary files found.")
			return
		}

		files, _ := os.ReadDir(tempDir)
		if len(files) == 0 {
			fmt.Println("No temporary files found.")
			return
		}

		for _, file := range files {
			filePath := filepath.Join(tempDir, file.Name())
			if err := os.RemoveAll(filePath); err != nil {
				fmt.Printf("Error occurred while deleting file '%s'\n", filePath)
			} else {
				fmt.Printf("Deleted file '%s'\n", filePath)
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(clearCmd)
}
