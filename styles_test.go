package excl

import (
	"encoding/xml"
	"os"
	"path/filepath"
	"testing"
)

func TestOpenStypes(t *testing.T) {
	defer os.Remove(filepath.Join("temp", "xl", "styles.xml"))

	_, err := OpenStyles("nopath")
	if err == nil {
		t.Error("styles.xml should not be opened.")
	}
	f, _ := os.Create(filepath.Join("temp", "xl", "styles.xml"))
	f.Close()
	_, err = OpenStyles("temp")
	if err == nil {
		t.Error("styles.xml should not be opened because syntax error.")
	}
	f, _ = os.Create(filepath.Join("temp", "xl", "styles.xml"))
	f.WriteString("<hoge></hoge>")
	f.Close()
	_, err = OpenStyles("temp")
	if err == nil {
		t.Error("styles.xml should not be opened because worksheet tag does not exist.")
	}
}

func TestSetData(t *testing.T) {
	styles := &Styles{}
	err := styles.setData()
	if err == nil {
		t.Error("error should be occurred because worksheet tag is not exist.")
	}
	styles.styles = &Tag{}
	err = styles.setData()
	if err == nil {
		t.Error("error should be occurred because worksheet tag is not exist.")
	}
	styles.styles = &Tag{Name: xml.Name{Local: "styleSheet"}}
	err = styles.setData()
	if err != nil {
		t.Error("error should not be occurred.", err.Error())
	}
	if styles.numFmtNumber != defaultMaxNumfmt {
		t.Error("styles.numFmtNumber should be ", defaultMaxNumfmt, " but ", styles.numFmtNumber)
	}
	fonts := &Tag{Name: xml.Name{Local: "fonts"}}
	fills := &Tag{Name: xml.Name{Local: "fills"}}
	borders := &Tag{Name: xml.Name{Local: "borders"}}
	cellStyleXfs := &Tag{Name: xml.Name{Local: "cellStyleXfs"}}
	cellXfs := &Tag{Name: xml.Name{Local: "cellXfs"}}
	cellStyles := &Tag{Name: xml.Name{Local: "cellStyles"}}
	dxfs := &Tag{Name: xml.Name{Local: "dxfs"}}
	extLst := &Tag{Name: xml.Name{Local: "extLst"}}
	tag := styles.styles
	tag.Children = append(tag.Children, fonts)
	tag.Children = append(tag.Children, fills)
	tag.Children = append(tag.Children, borders)
	tag.Children = append(tag.Children, cellStyleXfs)
	tag.Children = append(tag.Children, cellXfs)
	tag.Children = append(tag.Children, cellStyles)
	tag.Children = append(tag.Children, dxfs)
	tag.Children = append(tag.Children, extLst)
	err = styles.setData()
	if err != nil {
		t.Error("error should not be occurred.", err.Error())
	}
}

func TestSetNumFmt(t *testing.T) {
	//	styles := &Styles{}
	//	&Tag{}
}
