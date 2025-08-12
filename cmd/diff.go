package cmd

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var diffCmd = &cobra.Command{
	Use:   "diff [file1] [file2]",
	Short: "Show sheet names of two excel files",
	Long:  `Show sheet names of two excel files.`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		for _, filePath := range args {
			f, err := excelize.OpenFile(filePath)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", filePath, err)
				continue
			}

			fmt.Printf("Sheets in %s:\n", filePath)
			for _, name := range f.GetSheetList() {
				fmt.Printf("- %s\n", name)
			}
			fmt.Println()

			if err := f.Close(); err != nil {
				fmt.Fprintf(os.Stderr, "Error closing file %s: %v\n", filePath, err)
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(diffCmd)
}
