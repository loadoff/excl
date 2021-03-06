package excl

import (
	"os"
	"path/filepath"
	"testing"
)

func TestOpenWorkbookRels(t *testing.T) {
	os.MkdirAll("./temp/xl/_rels", 0755)
	defer os.RemoveAll("./temp/xl")
	path := filepath.Join("temp", "xl", "_rels", "workbook.xml.rels")
	_, err := OpenWorkbookRels("nopath")
	if err == nil {
		t.Error("workbook.xml.rels should not be opened.")
	}
	f, _ := os.Create(path)
	f.Close()
	_, err = OpenWorkbookRels("temp")
	if err == nil {
		t.Error("workbook.xml.rels should not be opened because syntax error.")
	}
	f, _ = os.Create(path)
	f.WriteString("<Relationships><Relationship></Relationship></Relationships>")
	f.Close()
	wbr, err := OpenWorkbookRels("temp")
	if err != nil {
		t.Error("workbook.xml.rels should be opened.")
	}
	if wbr == nil {
		t.Error("WorkbookRels should be created.")
	}
}

func TestAddSheetWorkbookRels(t *testing.T) {
	wbr := &WorkbookRels{rels: &Relationships{}}
	rid1 := wbr.addSheet("sheet1")
	if len(wbr.rels.Rels) != 1 {
		t.Error("Rels count should be 1 but [", len(wbr.rels.Rels), "]")
	}
	rid2 := wbr.addSharedStrings()
	wbr.addSheet("sheet2")
	wbr.addSharedStrings()
	if wbr.rels.Rels[0].ID != rid1 {
		t.Error("id should be [", rid1, "] but [", wbr.rels.Rels[0].ID, "]", wbr.rels.Rels[0].Target)
	}
	if wbr.rels.Rels[1].ID != rid2 {
		t.Error("id should be [", rid2, "] but [", wbr.rels.Rels[1].ID, "]")
	}
	wbr.addSheet("sheet3")
	err := wbr.Close()
	if err == nil {
		t.Error("close error should be happen.")
	}
	wbr.path = filepath.Join("temp", "workbook.xml.rels")
	defer os.Remove(filepath.Join("temp", "workbook.xml.rels"))
	err = wbr.Close()
	if err != nil {
		t.Error("workbook.xml.rels should be closed.", err.Error())
	}
}

func TestCloseWorkbookRels(t *testing.T) {
	var wbr *WorkbookRels
	var err error
	if err = wbr.Close(); err != nil {
		t.Error("error should not be happen.", err.Error())
	}

	wbr = &WorkbookRels{}
	wbr.rels = &Relationships{}
	if err = wbr.Close(); err == nil {
		t.Error("error should be happen.")
	}

	wbr.path = "temp/rels.xml"
	if err = wbr.Close(); err != nil {
		t.Error("error should not be happen.", err.Error())
	}
	os.Remove("temp/rels.xml")
}
