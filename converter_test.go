package excel

import (
	"fmt"
	"github.com/mattn/go-ole"
	"os"
	"testing"
)

func inputFile() string {
	wd, err := os.Getwd()
	if err != nil {
		panic(err)
	}
	return fmt.Sprintf("%s\\%s", wd, "example.xls")
}

func outputFile() string {
	wd, err := os.Getwd()
	if err != nil {
		panic(err)
	}
	return fmt.Sprintf("%s\\%s", wd, "example_out.xls")
}

func TestExcel(t *testing.T) {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()
	e, err := NewExcel()
	if err != nil {
		t.Error(err.Error())
		return
	}
	err = e.Quit()
	if err != nil {
		t.Error(err.Error())
		return
	}
}

func TestWorkbooks(t *testing.T) {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()
	e, err := NewExcel()
	if err != nil {
		t.Error(err.Error())
		return
	}
	defer e.Quit()
	workbooks, err := e.Workbooks()
	if err != nil {
		t.Error(err.Error())
		return
	}
	err = workbooks.Close()
	if err != nil {
		t.Error(err.Error())
		return
	}
}

func TestWorkbook(t *testing.T) {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()
	e, err := NewExcel()
	if err != nil {
		t.Error(err.Error())
		return
	}
	defer e.Quit()
	workbooks, err := e.Workbooks()
	if err != nil {
		t.Error(err.Error())
		return
	}
	defer workbooks.Close()
	workbook, err := workbooks.Workbook(inputFile())
	if err != nil {
		t.Error(err.Error())
		return
	}
	err = workbook.Saved(true)
	if err != nil {
		t.Error(err.Error())
		return
	}
	version, err := e.Version()
	if err != nil {
		t.Error(err.Error())
		return
	}
	_ = os.Remove(outputFile())
	err = workbook.SaveAs(outputFile(), version)
	if err != nil {
		t.Error(err.Error())
		return
	}
}

func TestCell(t *testing.T) {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()
	e, err := NewExcel()
	if err != nil {
		t.Error(err.Error())
		return
	}
	defer e.Quit()
	workbooks, err := e.Workbooks()
	if err != nil {
		t.Error(err.Error())
		return
	}
	defer workbooks.Close()
	workbook, err := workbooks.Workbook(inputFile())
	if err != nil {
		t.Error(err.Error())
		return
	}
	worksheet, err := workbook.Worksheet(1)
	if err != nil {
		t.Error(err.Error())
		return
	}
	cell, err := worksheet.Cell(1, 1)
	if err != nil {
		t.Error(err.Error())
		return
	}
	err = cell.Set(12346)
	if err != nil {
		t.Error(err.Error())
		return
	}
	err = workbook.Saved(true)
	if err != nil {
		t.Error(err.Error())
		return
	}
	version, err := e.Version()
	if err != nil {
		t.Error(err.Error())
		return
	}
	_ = os.Remove(outputFile())
	err = workbook.SaveAs(outputFile(), version)
	if err != nil {
		t.Error(err.Error())
		return
	}
}
