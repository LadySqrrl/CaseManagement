//the main body of the case management program
package main

import (
	"errors"
	"fmt"
	"os"

	"github.com/tealeg/xlsx/v3"
)

//make sure it opens
func check(e error) {
	if e != nil {
		panic(e)
	}
}

func cellVisitor(c *xlsx.Cell) error {
	value, err := c.FormattedValue()
	if err != nil {
		fmt.Println(err.Error())
	} else {
		fmt.Println("Cell value:", value)
	}
	return err
}

func rowVisitor(r *xlsx.Row) error {
	return r.ForEachCell(cellVisitor)
}

func main() {

	filename := "Test_File.xlsx"
	wb, err := xlsx.OpenFile(filename)
	if err != nil {
		panic(err)
	}
	sh, ok := wb.Sheet["Sheet1"]
	if !ok {
		panic(errors.New("Sheet not found"))
	}
	/*fmt.Println("Max row is", sh.MaxRow)
	sh.ForEachRow(rowVisitor)*/

	lastName, err := sh.Cell(1, 0)
	if err != nil {
		panic(err)
	}

	fmt.Println(lastName.String())

	b := []byte(lastName.String())

	outFile, err := os.Create("output.txt")
	check(err)

	defer outFile.Close()

	outFile.Write(b)

}
