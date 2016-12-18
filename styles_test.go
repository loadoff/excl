package excl

import (
	"bytes"
	"encoding/xml"
	"os"
	"path/filepath"
	"strings"
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

func TestSetStyleList(t *testing.T) {
	styles := &Styles{}
	r := strings.NewReader(`<cellXfs></cellXfs>`)
	tag := &Tag{}
	xml.NewDecoder(r).Decode(tag)
	styles.cellXfs = tag
	styles.setStyleList()
	if len(styles.styleList) != 0 {
		t.Error("styleList should be 0 but ", styles.styleList)
	}
	r = strings.NewReader(`<cellXfs><xf numFmtId="1" fontId="2" fillId="3" borderId="4" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"></xf></cellXfs>`)
	tag = &Tag{}
	xml.NewDecoder(r).Decode(tag)
	styles.cellXfs = tag
	styles.setStyleList()
	if len(styles.styleList) != 1 {
		t.Error("styleList should be 1 but ", len(styles.styleList))
	} else if styles.styleList[0].numFmtID != 1 {
		t.Error("numFmtID should be 1 but ", styles.styleList[0].numFmtID)
	} else if styles.styleList[0].fontID != 2 {
		t.Error("fontID should be 2 but ", styles.styleList[0].fontID)
	} else if styles.styleList[0].fillID != 3 {
		t.Error("fillID should be 3 but ", styles.styleList[0].fillID)
	} else if styles.styleList[0].borderID != 4 {
		t.Error("borderID should be 4 but ", styles.styleList[0].borderID)
	} else if styles.styleList[0].applyNumberFormat != 1 {
		t.Error("applyNumberFormat should be 1 but ", styles.styleList[0].applyNumberFormat)
	} else if styles.styleList[0].applyFont != 1 {
		t.Error("applyFont should be 1 but ", styles.styleList[0].applyFont)
	} else if styles.styleList[0].applyFill != 1 {
		t.Error("applyFill should be 1 but ", styles.styleList[0].applyFill)
	} else if styles.styleList[0].applyBorder != 1 {
		t.Error("applyBorder should be 1 but ", styles.styleList[0].applyBorder)
	} else if styles.styleList[0].applyAlignment != 1 {
		t.Error("applyAlignment should be 1 but ", styles.styleList[0].applyAlignment)
	} else if styles.styleList[0].applyProtection != 1 {
		t.Error("applyProtection should be 1 but ", styles.styleList[0].applyProtection)
	}
}

func TestSetFont(t *testing.T) {
	r := strings.NewReader(`<fonts></fonts>`)
	tag := &Tag{}
	xml.NewDecoder(r).Decode(tag)
	styles := &Styles{fonts: tag}
	font := Font{Size: 12, Color: "FFFF00FF"}
	index := styles.SetFont(font)
	if index != 0 {
		t.Error("index should be 0 but ", index)
	}
	b := new(bytes.Buffer)
	xml.NewEncoder(b).Encode(styles.fonts)
	if b.String() != `<fonts><font><sz val="12"></sz><color rgb="FFFF00FF"></color></font></fonts>` {
		t.Error("xml is corrupt [", b.String(), "]")
	}
	font = Font{Size: 13}
	index = styles.SetFont(font)
	if index != 1 {
		t.Error("index should be 1 but ", index)
	}
	b = new(bytes.Buffer)
	xml.NewEncoder(b).Encode(styles.fonts.Children[index].(*Tag))
	if b.String() != `<font><sz val="13"></sz></font>` {
		t.Error("xml is currupt [", b.String(), "]")
	}
	font = Font{Color: "FF00FFFF"}
	index = styles.SetFont(font)
	if index != 2 {
		t.Error("index should be 2 but ", index)
	}
	b = new(bytes.Buffer)
	xml.NewEncoder(b).Encode(styles.fonts.Children[index].(*Tag))
	if b.String() != `<font><color rgb="FF00FFFF"></color></font>` {
		t.Error("xml is currupt [", b.String(), "]")
	}
}

func TestSetBackgroundColor(t *testing.T) {
}

func TestSetNumFmt(t *testing.T) {
	//	styles := &Styles{}
	//	&Tag{}
}
