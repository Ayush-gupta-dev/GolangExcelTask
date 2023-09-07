package main

import (
	"encoding/csv"
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

type Student struct {
	Name   string
	BITSID string
	Email  string
	Branch string
}

// get email and branch from ID
func getEmailAndBranch(bitsID string) (email string, branch string) {
	// mapping ID digit:
	mapping := map[string]string{
		"A1": "chemical",
		"A5": "B.Pharma",
		"A7": "CS",
		"A3": "EE",
		"A4": "mechanical",
		"A8": "EnI",
		"A2": "civil",
		"AB": "Manufacturing",
		"B1": "Msc Bio",
		"B2": "Chem",
		"B3": "Eco",
		"B4": "Mathematics",
		"B5": "Physics",
		"D2": "General Studies",
		"AA": "ECE",
	}
	trimmedBranch := bitsID[4:]
	digits := trimmedBranch[:2]
	//fmt.Println(digits)
	email = bitsID + "@bitspilani.ac.in"
	if value, ok := mapping[digits]; ok {
		return email, value
	} else {
		return email, "unknown"
	}
}

func main() {

	// new formatted data
	var students []Student

	//reading excel file
	excelFile, err := excelize.OpenFile("BITS_inputExcel.xlsx")
	if err != nil {
		fmt.Println(err)
	}
	rows, err := excelFile.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
	}
	// fmt.Println("excel info is")
	for _, row := range rows[1:] {
		//getting only ids
		bitsId := row[1]
		//fmt.Println(bitsId)
		//getting email and branch from id
		email, branch := getEmailAndBranch(bitsId)
		// appending our dataset of students
		student := Student{
			Name:   row[0],
			BITSID: bitsId,
			Email:  email,
			Branch: branch,
		}
		students = append(students, student)
	}
	//fmt.Println(students)

	// creating new excel file:
	// outputFile := excelize.NewFile()
	// Set headers in the output sheet
	// outputFile.SetCellValue("Sheet1", "A1", "Name")
	// outputFile.SetCellValue("Sheet1", "B1", "BITS ID")
	// outputFile.SetCellValue("Sheet1", "C1", "Email")
	// outputFile.SetCellValue("Sheet1", "D1", "Branch")
	// outputFile.SaveAs("outputData.xlsx")

	// creating file in CSV
	CSVfile := "output.csv"
	file, err := os.Create(CSVfile)
	if err != nil {
		fmt.Println(err)
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	//headers
	headers := []string{"Name", "BITS-ID", "Email", "Branch"}
	writer.Write(headers)
	// add students data to csv file
	for _, student := range students {
		data := []string{student.Name, student.BITSID, student.Email, student.Branch}
		err := writer.Write(data)
		if err != nil {
			fmt.Println(err)
		}
	}
	writer.Flush()
	fmt.Println("successfully created csv file")

}
