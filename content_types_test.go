package excl

import (
	"os"
	"testing"
)

func TestOpenContentTypes(t *testing.T) {
	_, err := OpenContentTypes("./no/path/exist")
	if err == nil {
		t.Error(`xml file should not open.`)
	}
	f, _ := os.Create("./[Content_Types].xml")
	f.Close()
	defer os.Remove("./[Content_Types].xml")
	_, err = OpenContentTypes("./")
	if err == nil {
		t.Error("[Content_Types].xml is not xml file.")
	}

	f, _ = os.Create("./temp/[Content_Types].xml")
	f.WriteString("<Types></Types>")
	f.Close()
	defer os.Remove("./temp/[Content_Types].xml")
	_, err = OpenContentTypes("./temp")
	if err != nil {
		t.Error("[Content_Types].xml should be opened. [", err.Error(), "]")
	}
}

func TestSheetCount(t *testing.T) {
	f, _ := os.Create("./temp/[Content_Types].xml")
	f.WriteString(`<Types><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`)
	f.Close()
	defer os.Remove("./temp/[Content_Types].xml")
	types, _ := OpenContentTypes("./temp")
	count := types.sheetCount()
	if count != 1 {
		t.Error("sheet count should be 1 not [", count, "]")
	}
}

func TestAddSheet(t *testing.T) {
	f, _ := os.Create("./temp/[Content_Types].xml")
	f.WriteString(`<Types><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`)
	f.Close()
	defer os.Remove("./temp/[Content_Types].xml")
	types, _ := OpenContentTypes("./temp")
	name := types.addSheet()
	if name != "sheet2.xml" {
		t.Error(`sheet name should be "sheet2.xml" but [`, name, "]")
	}
	count := types.sheetCount()
	if count != 2 {
		t.Error("sheet count should be 2 not [", count, "]")
	}
	if types.types.Overrides[len(types.types.Overrides)-1].PartName != "/xl/worksheets/sheet2.xml" {
		t.Error(`part name should be "/xl/worksheets/sheet2.xml" but [`, types.types.Overrides[count-1].PartName, "]")
	}
}

func TestHasSharedString(t *testing.T) {
	f, _ := os.Create("./temp/[Content_Types].xml")
	f.WriteString(`<Types><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`)
	f.Close()
	defer os.Remove("./temp/[Content_Types].xml")

	types, _ := OpenContentTypes("./temp")
	if !types.hasSharedString() {
		t.Error("sharedString.xml file should be exists.")
	}

	f, _ = os.Create("./temp/[Content_Types].xml")
	f.WriteString(`<Types></Types>`)
	f.Close()

	types, _ = OpenContentTypes("./temp")
	if types.hasSharedString() {
		t.Error("sharedString.xml file should not be exists.")
	}
}

func TestAddSharedString(t *testing.T) {
	f, _ := os.Create("./temp/[Content_Types].xml")
	f.WriteString(`<Types></Types>`)
	f.Close()
	defer os.Remove("./temp/[Content_Types].xml")

	types, _ := OpenContentTypes("./temp")
	if types.hasSharedString() {
		t.Error("sharedString.xml file should not be exists.")
	}
	types.addSharedString()
	if !types.hasSharedString() {
		t.Error("sharedString.xml file should be exists.")
	}
}
