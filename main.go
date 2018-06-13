package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/extrame/xls"
	"github.com/tealeg/xlsx"
)

func main() {
	//var csvRow []string
	//sourceFile := "Book1.xlsx"
	sourceFile := flag.String("ds", "./", "Path to an Excel file or directory")
	targetDirectory := flag.String("dt", "./", "Path to where CSV will be saved")
	flag.Parse()

	if _, err := os.Stat(*sourceFile); os.IsNotExist(err) {
		fmt.Println("Source not found!")
		log.Fatal(err)
	}
	if _, err := os.Stat(*targetDirectory); os.IsNotExist(err) {
		fmt.Println("Destination directory not found!")
		log.Fatal(err)
	}

	if _, err := os.Stat(*sourceFile); err == nil {
		fileExt := filepath.Ext(*sourceFile)
		if fileExt == "" {
			//Get files in the directory
			files, _ := ioutil.ReadDir(*sourceFile)

			for _, file := range files {
				//fmt.Println(file.Name())
				if strings.ToLower(filepath.Ext(file.Name())) == ".xlsx" && file.Name()[:1] != "~" {
					fmt.Println("Processing " + file.Name())
					parseXlsx(*sourceFile+file.Name(), *targetDirectory)
				}
			}

		} else if fileExt == ".xlsx" {
			parseXlsx(*sourceFile, *targetDirectory)
		} else if fileExt == ".xls" {
			fmt.Println("Xls format is not supported yet")
		} else {
			fmt.Println(fileExt + " extension is not supported")
		}

	}

	//parseXls(*sourceFile, *targetDirectory)

}

func parseXlsx(sourceFile string, targetDirectory string) {
	sourceFileName := sourceFile[:len(sourceFile)-5]
	//fmt.Println(sourceFile)
	xlFile, _ := xlsx.OpenFile(sourceFile)

	for _, sheet := range xlFile.Sheets {
		targetFile, _ := os.Create(targetDirectory + sourceFileName + "_" + sheet.Name + ".csv")
		defer targetFile.Close()
		writer := csv.NewWriter(targetFile)
		defer writer.Flush()
		for _, row := range sheet.Rows {
			csvRow := []string{}
			for _, cell := range row.Cells {
				text := cell.Value
				csvRow = append(csvRow, text)
				//fmt.Printf("%s\n", text)
			}
			writer.Write(csvRow)
		}
	}
}

func parseXls(sourceFile string, targetDirectory string) {
	xlFile, _ := xls.Open(sourceFile, "utf-8")

	for i := 0; i < xlFile.NumSheets(); i++ {
		//fmt.Println(xlFile.GetSheet(i))
		sheet := xlFile.GetSheet(i)
		//fmt.Println(sheet.MaxRow)
		for j := 0; j <= int(sheet.MaxRow); j++ {
			row := sheet.Row(j)
			for k := 0; k < row.LastCol(); k++ {
				cell := row.Col(k)
				fmt.Println(cell)
			}
		}

	}

	//fmt.Println("test:", xlFile2.NumSheets())
}
