package main

import (
	"bufio"
	"fmt"
	"os"
	"time"

	//"github.com/pkg/profile"

	"github.com/LIJUCHACKO/ods2csv"
)

// writeLines writes the lines to the given file.
func writeLines(lines []string, path string) error {
	file, err := os.Create(path)
	if err != nil {
		return err
	}
	defer file.Close()

	w := bufio.NewWriter(file)
	for _, line := range lines {
		fmt.Fprintln(w, line+"\r")
	}
	return w.Flush()
}

func main() {
	// CPU profiling by default
	//defer profile.Start().Stop()
	start := time.Now()
	Filecontent, eerr := ods.ReadODSFile(os.Args[1])
	fmt.Printf("\nexecution time-")
	fmt.Println(time.Since(start))
	if eerr != nil {
		fmt.Printf("Read : %s\n", eerr)
		var yes string
		fmt.Scan(&yes)
		os.Exit(0)
	}
	for _, sheet := range Filecontent.Sheets {
		outputcontent := []string{}
		fmt.Printf("Sheet %s No of rows-%d\n  ", sheet.Name+".csv ", len(sheet.Rows))
		for _, row := range sheet.Rows {
			rowString := ""

			for _, cell := range row.Cells {

				rowString = rowString + cell.Text + "," //cell.Type or cell.Formula or cell.Value
			}

			outputcontent = append(outputcontent, rowString)
		}
		fmt.Printf("writing   %s", sheet.Name+".csv\n")
		if err := writeLines(outputcontent, sheet.Name+".csv"); err != nil {
			fmt.Printf("writing: %s", err)
			var yes string
			fmt.Scan(&yes)
			return
		}
	}

}
