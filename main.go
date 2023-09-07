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
	// i currently dont know the logic how to calculate for dual, and what is AB, so just appending with @pilani.com, for now
	email = bitsID + "@pilani-dummy.com"
	return email, "Computer Science"
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
	fmt.Println("successfully craeted csv file")

}
