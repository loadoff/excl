package excl

import (
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
	sharedStrings := &SharedStrings{count: 0, tempFile: f}
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
