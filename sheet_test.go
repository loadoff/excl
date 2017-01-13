package excl

import (
	"bufio"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"strings"
	"testing"

	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
)

func TestNewSheet(t *testing.T) {
	sheet := NewSheet("hello", 0)
	if sheet.xml.Name != "hello" {
		t.Error(`sheet name should be "hello".`)
	}
	if sheet.xml.SheetID != "1" {
		t.Error(`sheet id should be "1" but [`, sheet.xml.SheetID, "]")
	}
}

func TestOpen(t *testing.T) {
	os.MkdirAll("temp/xl/worksheets", 0755)
	defer os.RemoveAll("temp/xl")
	sheet := NewSheet("hello", 0)
	err := sheet.Open("")
	if err == nil {
		t.Error("sheet should not be opened. sheet does not exsist.")
	}
	f, _ := os.Create("temp/xl/worksheets/sheet1.xml")
	f.Close()
	if err = sheet.Open("temp1"); err == nil {
		t.Error("sheet should not be opened. file path does not exist.")
	}
	if err = sheet.Open("temp"); err == nil {
		t.Error("sheet should not be opened. xml file is currupt.")
	}
	f, _ = os.Create("temp/xl/worksheets/sheet1.xml")
	f.WriteString("<hoge></hoge>")
	f.Close()
	if err = sheet.Open("temp"); err == nil {
		t.Error("sheet should not be opened. worksheet tag does not exist.")
	}
	f, _ = os.Create("temp/xl/worksheets/sheet1.xml")
	f.WriteString("<worksheet></worksheet>")
	f.Close()
	if err = sheet.Open("temp"); err == nil {
		t.Error("sheet should not be opened. sheetData tag does not exist.")
	}
	f, _ = os.Create("temp/xl/worksheets/sheet1.xml")
	str := "<worksheet><sheetData><row></row></sheetData><hoge></hoge></worksheet>"
	f.WriteString(str)
	f.Close()
	if err = sheet.Open("temp"); err != nil {
		t.Error("sheet should be opened. [", err.Error(), "]")
	} else {
		sheet.Close()
		f, _ = os.Open("temp/xl/worksheets/sheet1.xml")
		b, _ := ioutil.ReadAll(f)
		f.Close()
		if string(b) != str {
			t.Error("new sheet file should be same as before string. [", string(b), "]")
		}
	}
}
func TestClose(t *testing.T) {
	var err error
	sheet := NewSheet("sheet", 0)
	if err = sheet.Close(); err != nil {
		t.Error("sheet should be closed because sheet is not opened.")
	}
	sheet.opened = true
	if err = sheet.Close(); err == nil {
		t.Error("sheet should not be closed because tempFile does not exist.")
	}
}

func TestGetRow(t *testing.T) {
	//var err error
	sheet := &Sheet{}
	row := &Row{rowID: 2}
	sheet.Rows = append(sheet.Rows, row)
	row2 := sheet.GetRow(2)
	if row != row2 {
		t.Error("row should be same.")
	}
	if row3 := sheet.GetRow(3); row3.rowID != 3 {
		t.Error("rowID should be 3 but [", row3.rowID, "]")
	}
	if row1 := sheet.GetRow(1); row1.rowID != 1 {
		t.Error("rowID should be 1 but [", row1.rowID, "]")
	}
}

func TestShowGridlines(t *testing.T) {
	os.MkdirAll("temp/xl/worksheets", 0755)
	defer os.RemoveAll("temp/xl")
	sheet := NewSheet("hoge", 0)
	sheet.ShowGridlines(true)
	if sheet.sheetView != nil {
		t.Error("sheetView should be nil.")
	}
	sheet.Create("temp")
	sheet.ShowGridlines(true)
	if v, err := sheet.sheetView.getAttr("showGridLines"); err != nil {
		t.Error("showGridLines should be exist.")
	} else if v != "1" {
		t.Error("value should be 1 but ", v)
	}
	sheet.Close()

	if b, err := ioutil.ReadFile("temp/xl/worksheets/sheet1.xml"); err != nil {
		t.Error("sheet1.xml should be readable.", err.Error())
	} else if string(b) != `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><sheetViews><sheetView workbookViewId="0" showGridLines="1"></sheetView></sheetViews><sheetData></sheetData></worksheet>` {
		t.Error("[" + string(b) + "]")
	}
	os.Remove("temp/xl/worksheets/sheet1.xml")

	sheet = NewSheet("hoge", 1)
	sheet.Create("temp")
	sheet.ShowGridlines(false)
	if v, err := sheet.sheetView.getAttr("showGridLines"); err != nil {
		t.Error("showGridLines should be exist.")
	} else if v != "0" {
		t.Error("value should be 0 but ", v)
	}
	sheet.Close()
	b, err := ioutil.ReadFile("temp/xl/worksheets/sheet2.xml")
	if err != nil {
		t.Error("sheet2.xml should be readable.", err.Error())
	} else if string(b) != `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><sheetViews><sheetView workbookViewId="0" showGridLines="0"></sheetView></sheetViews><sheetData></sheetData></worksheet>` {
		t.Error(string(b))
	}
	os.Remove("temp/xl/worksheets/sheet2.xml")

}

// UTF-8 から ShiftJIS
func utf8ToSjis(str string) (string, error) {
	ret, err := ioutil.ReadAll(transform.NewReader(strings.NewReader(str), japanese.ShiftJIS.NewEncoder()))
	if err != nil {
		return "", err
	}
	return string(ret), err
}

// ShiftJIS から UTF-8
func sjisToUtf8(str string) (string, error) {
	ret, err := ioutil.ReadAll(transform.NewReader(strings.NewReader(str), japanese.ShiftJIS.NewDecoder()))
	if err != nil {
		return "", err
	}
	return string(ret), err
}

func BenchmarkCreateRowsSJIS(b *testing.B) {
	f, _ := os.Create("temp/__sharedStrings.xml")
	f2, _ := os.Create("temp/__sheet1.xml")
	sjis, _ := utf8ToSjis("あいうえお")
	defer f.Close()
	defer f2.Close()
	sharedStrings := &SharedStrings{tempFile: f}
	sheet := &Sheet{sharedStrings: sharedStrings, tempFile: f2}
	for j := 0; j < 10; j++ {
		rows := sheet.CreateRows(10000*j+1, 10000*(j+1))
		for i := 10000 * j; i < 10000*(j+1); i++ {
			cells := rows[i].CreateCells(1, 20)
			for _, cell := range cells {
				cell.SetString(sjis)
			}
		}
		sheet.OutputThroughRowNo(10000 * (j + 1))
	}
	f2.Close()
	f.Seek(0, os.SEEK_SET)
	reader := bufio.NewReader(japanese.ShiftJIS.NewDecoder().Reader(f))
	f2, err := os.Create("temp/__sharedStrings2.xml")
	if err != nil {
		fmt.Println(err.Error())
	} else {
		io.Copy(f2, reader)
	}
	f.Close()
	f2.Close()
}

func BenchmarkCreateRowsUTF8(b *testing.B) {
	f, _ := os.Create("temp/__sharedStrings_utf8.xml")
	f2, _ := os.Create("temp/__sheet_utf8.xml")
	utf8 := "あいうえお"
	defer f.Close()
	defer f2.Close()
	sharedStrings := &SharedStrings{tempFile: f}
	sheet := &Sheet{sharedStrings: sharedStrings, tempFile: f2}
	for j := 0; j < 10; j++ {
		rows := sheet.CreateRows(10000*j+1, 10000*(j+1))
		for i := 10000 * j; i < 10000*(j+1); i++ {
			cells := rows[i].CreateCells(1, 20)
			for _, cell := range cells {
				cell.SetString(utf8)
			}
		}
		sheet.OutputThroughRowNo(10000 * (j + 1))
	}
	f2.Close()
	f.Close()
}

func BenchmarkCreateRowsNumber(b *testing.B) {
	f2, _ := os.Create("temp/__sheet_number.xml")
	defer f2.Close()
	sheet := &Sheet{tempFile: f2}
	for j := 0; j < 10; j++ {
		rows := sheet.CreateRows(10000*j+1, 10000*(j+1))
		for i := 10000 * j; i < 10000*(j+1); i++ {
			cells := rows[i].CreateCells(1, 20)
			for _, cell := range cells {
				cell.SetNumber("12345678901234567890")
			}
		}
		sheet.OutputThroughRowNo(10000 * (j + 1))
	}
	f2.Close()
}

func BenchmarkConvert(b *testing.B) {
	sjis, _ := utf8ToSjis("あいうえお")
	f, _ := os.Create("temp/test1.txt")
	for i := 1; i < 100000*20; i++ {
		f.WriteString(sjis)
	}
	f.Seek(0, os.SEEK_SET)
	f2, _ := os.Create("temp/test2.txt")
	r := transform.NewReader(f, japanese.ShiftJIS.NewDecoder())
	io.Copy(f2, r)
	f.Close()
	f2.Close()
}
