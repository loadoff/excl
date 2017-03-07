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

func TestCreateWorkbook(t *testing.T) {
	wb, err := Create()
	if err != nil {
		t.Error("error should not be happen but ", err.Error())
	}
	wb.OpenSheet("hello")
	wb.Save("temp/new.xlsx")
	if !isFileExist("temp/new.xlsx") {
		t.Error("new.xlsx should be created.")
	}
	os.Remove("temp/new.xlsx")
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
	var err error
	var workbook *Workbook

	os.MkdirAll("temp/out", 0755)
	defer os.RemoveAll("temp/out")
	if workbook, err = Open("temp/test.xlsx"); err != nil {
		t.Error(err.Error())
		t.Error("workbook should be opened.")
	}
	if ok := isFileExist(filepath.Join(workbook.TempPath, "[Content_Types].xml")); !ok {
		t.Error("[Content_Types].xml should be exists.")
	}

	if err = workbook.Save("temp/out/test.xlsx"); err != nil {
		t.Error("Close should be succeed.", err.Error())
	}
	if ok := isFileExist(filepath.Join(workbook.TempPath, "[Content_Types].xml")); ok {
		t.Error("[Content_Types].xml should be deleted.")
	}
	if ok := isFileExist("temp/out/test.xlsx"); !ok {
		t.Error("test.xml should be created.")
	}
	// error patern

	if workbook, err = Open("no/path/excel.xlsx"); err == nil {
		t.Error("workbook should not be opened.")
		workbook.Close()
	}

	f, _ := os.Create("temp/currupt.xlsx")
	z := zip.NewWriter(f)
	z.Close()
	f.Close()
	defer os.Remove("temp/currupt.xlsx")
	if workbook, err = Open("temp/currupt.xlsx"); err == nil {
		t.Error("workbook should not be opened because excel file must be currupt.")
	}
	workbook.Close()

	createZip("temp/empty.xlsx", nil, "")
	defer os.Remove("temp/empty.xlsx")
	if workbook, err = Open("temp/empty.xlsx"); err == nil {
		t.Error("workbook should not be opened beacause excel file is not zip file.")
		workbook.Close()
	}

	createCurruputXLSX("temp/test.xlsx", "temp/no_content_types.xlsx", "[Content_Types].xml")
	defer os.Remove("temp/no_content_types.xlsx")
	if workbook, err = Open("temp/no_content_types.xlsx"); err == nil {
		t.Error("workbook should not be opened beacause excel file does not include [Content_Types].xml.")
		workbook.Close()
	}

	createCurruputXLSX("temp/test.xlsx", "temp/no_workbook_xml.xlsx", "xl/workbook.xml")
	defer os.Remove("temp/no_workbook_xml.xlsx")
	if workbook, err = Open("temp/no_workbook_xml.xlsx"); err == nil {
		t.Error("workbook should not be opened beacause excel file does not include workbook.xml.")
	}
	workbook.Close()

}

func TestOpenWorkbookXML(t *testing.T) {
	var err error
	workbook := &Workbook{TempPath: ""}
	workbook.TempPath = ""

	if err = workbook.openWorkbook(); err == nil {
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
	workbook, _ := Open("temp/test.xlsx")
	sheet, _ := workbook.OpenSheet("Sheet1")
	if sheet == nil {
		t.Error("Sheet1 should be exist.")
	}
	if _, err := workbook.OpenSheet("Sheet2"); err != nil {
		t.Error("Sheet2 should be created. [", err.Error(), "]")
	}

	sheet, _ = workbook.OpenSheet("ペンギンﾍﾟﾝｷﾞﾝAaＡａ0０")
	sheet.Close()
	tempSheet, _ := workbook.OpenSheet("ﾍﾟﾝｷﾞﾝペンギンaaaa00")
	if sheet != tempSheet {
		t.Error("ペンギンﾍﾟﾝｷﾞﾝAaＡａ0０ sheet should be same as ﾍﾟﾝｷﾞﾝペンギンaaaa00 sheet.")
	}
	workbook.Close()
}

func TestSetInfo(t *testing.T) {
	os.Mkdir("temp/out", 0755)
	defer os.RemoveAll("temp/out")
	workbook := &Workbook{TempPath: "temp/out"}
	err := workbook.setInfo()
	if err == nil {
		t.Error("[Content_Types].xml should not be opened.")
	}
	createContentTypes("temp/out")
	err = workbook.setInfo()
	if err == nil {
		t.Error("workbook.xml should not be opened.")
	}
	createWorkbook("temp/out")
	f, _ := os.Create(filepath.Join("temp/out/xl/sharedStrings.xml"))
	f.Close()
	err = workbook.setInfo()
	if err == nil {
		t.Error("sharedStrings.xml should not be opened.")
	}
	os.Remove(filepath.Join("temp/out/xl/sharedStrings.xml"))

}

func TestCalcPr(t *testing.T) {
	workbook := &Workbook{calcPr: &Tag{}}
	workbook.SetForceFormulaRecalculation(true)
	if v, _ := workbook.calcPr.getAttr("fullCalcOnLoad"); v != "1" {
		t.Error("fullCalcOnLoad attribute should be 1 but", v)
	}
	workbook.SetForceFormulaRecalculation(false)
	if _, err := workbook.calcPr.getAttr("fullCalcOnLoad"); err == nil {
		t.Error("fullCalcOnLoad attribute should not be found.")
	}

}
