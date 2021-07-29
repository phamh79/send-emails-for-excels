package main

import (
	"fmt"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/joho/godotenv"
)

func main() {
	// load env
	if err := godotenv.Load("setting.env"); err != nil {
		fmt.Println("error loadding setting.env", err)
		return
	}
	sourceFile := os.Getenv("SOURCE_FILE")
	templateFile := os.Getenv("TEMPLATE_FILE")
	dataSheet := os.Getenv("DATA_SHEET")
	senderColumnStr := os.Getenv("SENDER_COLUMN")
	senderColumn, err := strconv.Atoi(senderColumnStr)
	if err != nil {
		fmt.Println(err)
		return
	}
	startRowStr := os.Getenv("START_ROW")
	startRow, err := strconv.Atoi(startRowStr)
	if err != nil {
		fmt.Println(err)
		return
	}
	// load source file
	dataFile, err := excelize.OpenFile(sourceFile)
	if err != nil {
		fmt.Println(err)
		return
	}
	rows, err := dataFile.GetRows(dataSheet)
	if err != nil {
		fmt.Println("error reading source file", err)
		return
	}

	// writing data
	currentEmail := ""
	var saveFile *excelize.File
	cursor := startRow
	for i := startRow; i < len(rows); i++ {
		row := rows[i]
		email := row[senderColumn]
		if email != currentEmail {
			fmt.Println("writing data for", email)
			if saveFile != nil {
				err := saveFile.SaveAs(fmt.Sprintf("%v.%v", currentEmail, "xlsx"))
				if err != nil {
					fmt.Println(err)
					return
				}
			}
			saveFile, err = excelize.OpenFile(templateFile)
			if err != nil {
				fmt.Println(err)
				return
			}
			cursor = startRow
		}
		currentEmail = email
		cursor++
		if err := saveFile.SetSheetRow(dataSheet, fmt.Sprintf("%v%d", "A", cursor), &row); err != nil {
			fmt.Println(err)
			return
		}
	}
	saveFile.SaveAs(fmt.Sprintf("%v.%v", currentEmail, "xlsx"))
	if err != nil {
		fmt.Println(err)
		return
	}
}
