package cmd

import (
	"bytes"
	"encoding/csv"
	"fmt"
	"os"
	"strings"

	"github.com/pmezard/go-difflib/difflib"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var diffCmd = &cobra.Command{
	Use:   "diff [file1] [file2]",
	Short: "Compare sheet names and cell contents of two Excel files",
	Long:  `Compare sheet names of two Excel files and output the differences in unified diff format. For sheets with the same name, compare the cell contents and output the differences in CSV format.`,
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

		// Compare cell contents for common sheets
		sheetMap1 := make(map[string]bool)
		for _, s := range sheets1 {
			sheetMap1[s] = true
		}

		for _, sheet := range sheets2 {
			if sheetMap1[sheet] {
				csv1, err := getSheetCSV(f1, sheet)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error reading sheet %s from %s: %v\n", sheet, file1, err)
					continue
				}

				csv2, err := getSheetCSV(f2, sheet)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error reading sheet %s from %s: %v\n", sheet, file2, err)
					continue
				}

				sheetDiff := difflib.UnifiedDiff{
					A:        difflib.SplitLines(csv1),
					B:        difflib.SplitLines(csv2),
					FromFile: fmt.Sprintf("%s/%s", file1, sheet),
					ToFile:   fmt.Sprintf("%s/%s", file2, sheet),
					Context:  3,
				}

				sheetText, err := difflib.GetUnifiedDiffString(sheetDiff)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error generating diff for sheet %s: %v\n", sheet, err)
					continue
				}

				if sheetText != "" {
					fmt.Print(sheetText)
				}
			}
		}
	},
}

func getSheetCSV(f *excelize.File, sheetName string) (string, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return "", err
	}

	var buf bytes.Buffer
	w := csv.NewWriter(&buf)
	err = w.WriteAll(rows)
	if err != nil {
		return "", err
	}

	return buf.String(), nil
}

func init() {
	rootCmd.AddCommand(diffCmd)
}
