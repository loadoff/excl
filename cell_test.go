package excl

import (
	"bytes"
	"encoding/xml"
	"os"
	"testing"
)

func TestNewCell(t *testing.T) {
	tag := &Tag{}
	cell := NewCell(tag, nil, nil)
	if cell != nil {
		t.Error("cell should be nil because colNo does not exist.")
	}
	attr := xml.Attr{
		Name:  xml.Name{Local: "r"},
		Value: "",
	}
	tag.Attr = append(tag.Attr, attr)
	cell = NewCell(tag, nil, nil)
	if cell != nil {
		t.Error("cell should be nil because colNo is not correct.")
	}
	attr.Value = "A1"
	tag.Attr = []xml.Attr{attr}
	cell = NewCell(tag, nil, nil)
	if cell == nil {
		t.Error("cell should be created.")
	} else if cell.colNo != 1 {
		t.Error("colNo should be 1 but [", cell.colNo, "]")
	}
	tag.setAttr("s", "2")
	if cell = NewCell(tag, nil, nil); cell == nil {
		t.Error("cell should be created.")
	}

}

func TestSetNumber(t *testing.T) {
	tag := &Tag{}
	attr := xml.Attr{
		Name:  xml.Name{Local: "r"},
		Value: "A1",
	}
	tag.Attr = []xml.Attr{attr}
	cell := &Cell{cell: tag, colNo: 1}
	cell.SetNumber("123")
	val := cell.cell.Children[0].(*Tag)
	if val.Name.Local != "v" {
		t.Error("tag should be v but [", val.Name.Local, "]")
	} else {
		data := val.Children[0].(xml.CharData)
		if string(data) != "123" {
			t.Error("value should be 123 but [", data, "]")
		}
	}
	typeAttr := xml.Attr{
		Name:  xml.Name{Local: "t"},
		Value: "s",
	}
	tag.Attr = []xml.Attr{attr, typeAttr}
	cell = &Cell{cell: tag, colNo: 1}
	cell.SetNumber("456")
	for _, attr := range cell.cell.Attr {
		if attr.Name.Local == "t" {
			t.Error("t attribute should be deleted.")
		}
	}
}

func TestSetString(t *testing.T) {
	f, _ := os.Create("temp/sharedStrings.xml")
	sharedStrings := &SharedStrings{count: 0, tempFile: f, buffer: &bytes.Buffer{}}
	tag := &Tag{}
	tag.setAttr("r", "AB12")
	cell := &Cell{cell: tag, colNo: 1}
	cell.sharedStrings = sharedStrings
	cell.SetString("こんにちは")
	cTag := cell.cell.Children[0].(*Tag)
	if cTag.Name.Local != "v" {
		t.Error("tag name should be [v] but [", cTag.Name.Local, "]")
	} else if string(cTag.Children[0].(xml.CharData)) == "こんにちは" {
		t.Error("tag value should be こんにちは but [", cTag.Children[0].(xml.CharData), "]")
	} else if cell.cell.Attr[0].Value == "s" {
		t.Error("tag attribute value should be s but [", cTag.Attr[0].Value, "]")
	}
	f.Close()
	os.Remove("temp/sharedStrings.xml")
}

func TestSetCellNumFmt(t *testing.T) {
	cell := &Cell{}
	cell.styles = &Styles{}
	cell.styleIndex = 10

	if cell.SetNumFmt("format"); cell.style.NumFmtID != 0 {
		t.Error("numFmtId should be 0 but ", cell.style.NumFmtID)
	}

	if cell.SetNumFmt("format"); cell.style.NumFmtID != 1 {
		t.Error("numFmtId should be 1 but ", cell.style.NumFmtID)
	}
}

func TestCellSetFont(t *testing.T) {
	cell := &Cell{}
	cell.styles = &Styles{fonts: &Tag{}}
	cell.styleIndex = 10

	if cell.SetFont(Font{}); cell.style.FontID != 0 {
		t.Error("fontID should be 0 but ", cell.style.FontID)
	}

	if cell.SetFont(Font{}); cell.style.FontID != 1 {
		t.Error("fontID should be 1 but ", cell.style.FontID)
	}
}

func TestCellSetBackgroundColor(t *testing.T) {
	cell := &Cell{}
	cell.styles = &Styles{fills: &Tag{}}
	cell.styleIndex = 10

	if cell.SetBackgroundColor("FFFFFF"); cell.style.FillID != 0 {
		t.Error("fillID should be 0 but ", cell.style.FillID)
	}

	if cell.SetBackgroundColor("000000"); cell.style.FillID != 1 {
		t.Error("fillID should be 1 but ", cell.style.FillID)
	}
}

func TestCellSetBorder(t *testing.T) {
	cell := &Cell{}
	cell.styles = &Styles{borders: &Tag{}}
	cell.styleIndex = 10

	if cell.SetBorder(Border{}); cell.style.BorderID != 0 {
		t.Error("BorderID should be 0 but ", cell.style.BorderID)
	}

	if cell.SetBorder(Border{}); cell.style.BorderID != 1 {
		t.Error("BorderID should be 1 but ", cell.style.BorderID)
	}
}

func TestCellSetStyle(t *testing.T) {
	cell := &Cell{}
	cell.styles = &Styles{}
	cell.styleIndex = 10
	style := &Style{}
	cell.SetStyle(nil)
	if cell.style != nil {
		t.Error("style should be nil.")
	}
	cell.SetStyle(style)
	if cell.style.NumFmtID != 0 {
		t.Error("NumFmtID should be 0 but", cell.style.NumFmtID)
	}
	if cell.style.FontID != 0 {
		t.Error("FontID should be 0 but", cell.style.FontID)
	}
	if cell.style.FillID != 0 {
		t.Error("FillID should be 0 but", cell.style.FillID)
	}
	if cell.style.BorderID != 0 {
		t.Error("BorderID should be 0 but", cell.style.BorderID)
	}
	if cell.style.Horizontal != "" {
		t.Error("Horizontal should be empty but", cell.style.Horizontal)
	}
	if cell.style.Vertical != "" {
		t.Error("Vertical should be empty but", cell.style.Vertical)
	}

	style.NumFmtID = 1
	style.FontID = 2
	style.FillID = 3
	style.BorderID = 4
	style.Horizontal = "center"
	style.Vertical = "top"
	cell.SetStyle(style)
	if cell.style.NumFmtID != style.NumFmtID {
		t.Error("NumFmtID should be 1 but", cell.style.NumFmtID)
	}

	if cell.style.FontID != style.FontID {
		t.Error("FontID should be 2 but", cell.style.FontID)
	}

	if cell.style.FillID != style.FillID {
		t.Error("FillID should be 3 but", cell.style.FillID)
	}

	if cell.style.BorderID != style.BorderID {
		t.Error("BorderID should be 3 but", cell.style.BorderID)
	}

	if cell.style.Horizontal != style.Horizontal {
		t.Error("Horizontal should be center but", cell.style.Horizontal)
	}

	if cell.style.Vertical != style.Vertical {
		t.Error("Vertical should be top but", cell.style.Vertical)
	}
}

func TestResetStyleIndex(t *testing.T) {
	var cell *Cell
	if cell.resetStyleIndex(); cell != nil {
		t.Error("cell should be nil")
	}
	cell = &Cell{}
	cell.style = &Style{}
	cell.cell = &Tag{}
	cell.styles = &Styles{cellXfs: &Tag{}}
	cell.resetStyleIndex()
	if _, err := cell.cell.getAttr("s"); err == nil {
		t.Error("cell attribute should not be found.")
	}

	cell.changed = true
	cell.resetStyleIndex()
	if val, _ := cell.cell.getAttr("s"); val != "0" {
		t.Error("style index should be 0 but", val)
	}
}
