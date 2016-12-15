package excl

import (
	"archive/zip"
	"os"
	"path/filepath"
	"testing"
)

func createCurruputXLSX(from string, to string, delfile string) {
	os.Mkdir("temp/output", 0755)
	defer os.RemoveAll("temp/output")
	unzip(from, "temp/output")
	if delfile != "" {
		os.Remove(filepath.Join("temp/output", delfile))
	}
	createZip(to, getFiles("temp/output"), "temp/output")
}

func TestNewWorkbook(t *testing.T) {
	workbook, err := NewWorkbook("path", "tempPath", "tempPath/text.xlsx")
	if err == nil {
		t.Error(`tempPath should not be exists.`)
		workbook.Close()
	}
	os.Mkdir("cancreate", 0700)
	workbook, err = NewWorkbook("path", "cancreate", "tempPath/text.xlsx")
	if err == nil {
		t.Error(`error should be happen.`)
		workbook.Close()
	}
	os.RemoveAll("cancreate")
}

func TestFileDirExist(t *testing.T) {
	if ok := isDirExist("no/exist/dir"); ok {
		t.Error("directory should not be exist.")
	}
	if ok := isDirExist("temp"); !ok {
		t.Error("directory should be exist.")
	}
	if ok := isDirExist("temp/test.xlsx"); ok {
		t.Error("temp/test.xlsx is not directory.")
	}
	if ok := isFileExist("no/file/exist"); ok {
		t.Error("file should not be exist.")
	}
	if ok := isFileExist("temp"); ok {
		t.Error("temp/out should be directory.")
	}
	if ok := isFileExist("temp/test.xlsx"); !ok {
		t.Error("temp/test.xlsx should be file.")
	}
}

func TestOpenWorkbook(t *testing.T) {
	os.MkdirAll("temp/out", 0755)
	defer os.RemoveAll("temp/out")
	workbook, _ := NewWorkbook("temp/test.xlsx", "temp/out", "temp/out/test.xlsx")
	if err := workbook.Open(); err != nil {
		t.Error(err.Error())
		t.Error("workbook should be opened.")
	}
	if ok := isFileExist(filepath.Join(workbook.TempPath, "[Content_Types].xml")); !ok {
		t.Error("[Content_Types].xml should be exists.")
	}
	workbook.Close()
	if ok := isFileExist(filepath.Join(workbook.TempPath, "[Content_Types].xml")); ok {
		t.Error("[Content_Types].xml should be deleted.")
	}
	if ok := isFileExist("temp/out/test.xlsx"); !ok {
		t.Error("test.xml should be created.")
	}
	// error patern
	_, err := NewWorkbook("no/path/excel.xlsx", "temp/out", "temp/out/test.xlsx")
	if err == nil {
		t.Error("error should be happen.")
	}
	_, err = NewWorkbook("temp/test.xlsx", "nopath", "temp/out/text.xlsx")
	if err == nil {
		t.Error("error should be happen because path does not exist.")
	}
	workbook, _ = NewWorkbook("temp/test.xlsx", "temp/out", "temp/out/text.xlsx")
	workbook.Path = "no/path/excel.xlsx"
	if err := workbook.Open(); err == nil {
		t.Error("workbook should not be opened.")
	}
	workbook.Close()
	workbook, _ = NewWorkbook("temp/test.xlsx", "temp/out", "temp/out/text.xlsx")
	os.Mkdir(workbook.TempPath, 0755)
	if err := workbook.Open(); err == nil {
		t.Error("workbook should not be opened.")
	}
	workbook.Close()

	f, _ := os.Create("temp/currupt.xlsx")
	z := zip.NewWriter(f)
	z.Close()
	f.Close()
	defer os.Remove("temp/currupt.xlsx")
	workbook, _ = NewWorkbook("temp/currupt.xlsx", "temp/out", "temp/out/text.xlsx")
	if err := workbook.Open(); err == nil {
		t.Error("workbook should not be opened because excel file must be currupt.")
	}
	workbook.Close()

	createZip("temp/empty.xlsx", nil, "")
	defer os.Remove("temp/empty.xlsx")
	workbook, _ = NewWorkbook("temp/empty.xlsx", "temp/out", "temp/out/text.xlsx")
	if err := workbook.Open(); err == nil {
		t.Error("workbook should not be opened beacause excel file is not zip file.")
	}
	workbook.Close()

	createCurruputXLSX("temp/test.xlsx", "temp/no_content_types.xlsx", "[Content_Types].xml")
	defer os.Remove("temp/no_content_types.xlsx")
	workbook, _ = NewWorkbook("temp/no_content_types.xlsx", "temp/out", "temp/out/test.xlsx")
	if err := workbook.Open(); err == nil {
		t.Error("workbook should not be opened beacause excel file does not include [Content_Types].xml.")
	}
	workbook.Close()

	createCurruputXLSX("temp/test.xlsx", "temp/no_workbook_xml.xlsx", "xl/workbook.xml")
	defer os.Remove("temp/no_workbook_xml.xlsx")
	workbook, _ = NewWorkbook("temp/no_workbook_xml.xlsx", "temp/out", "temp/out/test.xlsx")
	if err := workbook.Open(); err == nil {
		t.Error("workbook should not be opened beacause excel file does not include workbook.xml.")
	}
	workbook.Close()

}

func TestOpenWorkbookXML(t *testing.T) {
	workbook := Workbook{TempPath: ""}
	workbook.TempPath = ""
	err := workbook.openWorkbook()
	if err == nil {
		t.Error("workbook.xml should not be opened because workbook.xml does not exist.")
	}
	os.MkdirAll("temp/workbook/xl", 0755)
	defer os.RemoveAll("temp/workbook")
	f1, _ := os.Create("temp/workbook/xl/workbook.xml")
	f1.Close()
	err = workbook.openWorkbook()
	if err == nil {
		t.Error("workbook.xml should not be parsed.")
	}
	f1, _ = os.Create("temp/workbook/xl/workbook.xml")
	f1.WriteString("<workbook><sheets><sheet></sheet><sheet></sheet></sheets></workbook>")
	f1.Close()

	workbook.TempPath = "temp/workbook"
	err = workbook.openWorkbook()
	if err != nil {
		t.Error("workbook.xml should be opened. error[", err.Error(), "]")
	} else if len(workbook.sheets) != 2 {
		t.Error("sheet count should be 2 but [", len(workbook.sheets), "]")
	}
}

func TestOpenSheet(t *testing.T) {
	os.Mkdir("temp/out", 0755)
	defer os.RemoveAll("temp/out")
	workbook, _ := NewWorkbook("temp/test.xlsx", "temp/out", "temp/out/test.xlsx")
	workbook.Open()
	sheet, _ := workbook.OpenSheet("Sheet1")
	if sheet == nil {
		t.Error("Sheet1 should be exist.")
	} else if sheet != workbook.sheets[0] {
		t.Error("Sheet1 should be same as workbook sheet1.")
	}
	if _, err := workbook.OpenSheet("Sheet2"); err != nil {
		t.Error("Sheet2 should be created. [", err.Error(), "]")
	}
	workbook.Close()
}

func TestSetInfo(t *testing.T) {
	os.Mkdir("temp/out", 0755)
	defer os.RemoveAll("temp/out")
	workbook, _ := NewWorkbook("temp/test.xlsx", "temp/out", "temp/out/text.xlsx")
	workbook.TempPath = ""
	err := workbook.setInfo()
	if err == nil {
		t.Error("[Content_Types].xml should not be opened.")
	}
	workbook.TempPath = "temp"
	err = workbook.setInfo()
	if err == nil {
		t.Error("workbook.xml should not be opened.")
	}
	f, _ := os.Create(filepath.Join("temp", "xl", "workbook.xml"))
	f.WriteString(`<workbook><sheets><sheet name="Sheet1"></sheet></sheets></workbook>`)
	f.Close()
	defer os.Remove(filepath.Join("temp", "xl", "workbook.xml"))
	f, _ = os.Create(filepath.Join("temp", "xl", "sharedStrings.xml"))
	f.Close()
	defer os.Remove(filepath.Join("temp", "xl", "sharedStrings.xml"))
	err = workbook.setInfo()
	if err == nil {
		t.Error("sharedStrings.xml should not be opened.")
	}
}
