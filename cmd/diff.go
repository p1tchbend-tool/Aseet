package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/pmezard/go-difflib/difflib"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var diffCmd = &cobra.Command{
	Use:   "diff [file1] [file2]",
	Short: "Compare sheet names of two Excel files",
	Long:  `Compare sheet names of two Excel files and output the differences in unified diff format.`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		file1 := args[0]
		file2 := args[1]

		f1, err := excelize.OpenFile(file1)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", file1, err)
			os.Exit(1)
		}
		defer f1.Close()

		f2, err := excelize.OpenFile(file2)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", file2, err)
			os.Exit(1)
		}
		defer f2.Close()

		sheets1 := f1.GetSheetList()
		sheets2 := f2.GetSheetList()

		text1 := strings.Join(sheets1, "\n") + "\n"
		text2 := strings.Join(sheets2, "\n") + "\n"

		diff := difflib.UnifiedDiff{
			A:        difflib.SplitLines(text1),
			B:        difflib.SplitLines(text2),
			FromFile: file1,
			ToFile:   file2,
			Context:  3,
		}

		text, err := difflib.GetUnifiedDiffString(diff)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error generating diff: %v\n", err)
			os.Exit(1)
		}

		if text != "" {
			fmt.Print(text)
		}
	},
}

func init() {
	rootCmd.AddCommand(diffCmd)
}
