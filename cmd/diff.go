package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/pmezard/go-difflib/difflib"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

const (
	colorRed   = "\033[31m"
	colorGreen = "\033[32m"
	colorReset = "\033[0m"
)

var diffCmd = &cobra.Command{
	Use:   "diff [file1] [file2]",
	Short: "Compare sheet names and cell contents of two Excel files",
	Long:  `Compare sheet names of two Excel files and output the differences in unified diff format. For sheets with the same name, compare the cell contents cell by cell.`,
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
				rows1, err := f1.GetRows(sheet)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error reading sheet %s from %s: %v\n", sheet, file1, err)
					continue
				}

				rows2, err := f2.GetRows(sheet)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error reading sheet %s from %s: %v\n", sheet, file2, err)
					continue
				}

				maxRows := len(rows1)
				if len(rows2) > maxRows {
					maxRows = len(rows2)
				}

				for r := 0; r < maxRows; r++ {
					var row1, row2 []string
					if r < len(rows1) {
						row1 = rows1[r]
					}
					if r < len(rows2) {
						row2 = rows2[r]
					}

					maxCols := len(row1)
					if len(row2) > maxCols {
						maxCols = len(row2)
					}

					for c := 0; c < maxCols; c++ {
						val1, val2 := "", ""
						if c < len(row1) {
							val1 = row1[c]
						}
						if c < len(row2) {
							val2 = row2[c]
						}

						if val1 != val2 {
							cellName, err := excelize.CoordinatesToCellName(c+1, r+1)
							if err != nil {
								cellName = fmt.Sprintf("R%dC%d", r+1, c+1)
							}

							if val1 != "" {
								fmt.Printf("%s- %s!%s: %s%s\n", colorRed, sheet, cellName, val1, colorReset)
							}
							if val2 != "" {
								fmt.Printf("%s+ %s!%s: %s%s\n", colorGreen, sheet, cellName, val2, colorReset)
							}
						}
					}
				}
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(diffCmd)
}
