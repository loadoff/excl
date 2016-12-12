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
	wbr.addSheet("sheet1")
	if len(wbr.rels.Rels) != 1 {
		t.Error("Rels count should be 1 but [", len(wbr.rels.Rels), "]")
	}
	wbr.addSharedStrings()
	wbr.addSheet("sheet2")
	wbr.addSharedStrings()
	wbr.setRID()
	if wbr.rels.Rels[0].ID != "rId2" {
		t.Error("id should be rId2 but [", wbr.rels.Rels[0].ID, "]", wbr.rels.Rels[0].Target)
	}
	if wbr.rels.Rels[1].ID != "rId3" {
		t.Error("id should be rId3 but [", wbr.rels.Rels[1].ID, "]")
	}
	wbr.addSheet("sheet3")
	wbr.setRID()
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
