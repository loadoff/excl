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
	r = strings.NewReader(`<cellXfs><xf numFmtId="1" fontId="2" fillId="3" borderId="4" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"><alignment horizontal="left" vertical="top" wrapText="1"></alignment></xf></cellXfs>`)
	tag = &Tag{}
	xml.NewDecoder(r).Decode(tag)
	styles.cellXfs = tag
	styles.setStyleList()
	if len(styles.styleList) != 1 {
		t.Error("styleList should be 1 but ", len(styles.styleList))
	} else if styles.styleList[0].NumFmtID != 1 {
		t.Error("numFmtID should be 1 but ", styles.styleList[0].NumFmtID)
	} else if styles.styleList[0].FontID != 2 {
		t.Error("fontID should be 2 but ", styles.styleList[0].FontID)
	} else if styles.styleList[0].FillID != 3 {
		t.Error("fillID should be 3 but ", styles.styleList[0].FillID)
	} else if styles.styleList[0].BorderID != 4 {
		t.Error("borderID should be 4 but ", styles.styleList[0].BorderID)
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
	} else if styles.styleList[0].Horizontal != "left" {
		t.Error("Horizontal should be left but ", styles.styleList[0].Horizontal)
	} else if styles.styleList[0].Vertical != "top" {
		t.Error("Vertical should be top but ", styles.styleList[0].Vertical)
	} else if styles.styleList[0].Wrap != 1 {
		t.Error("Wrap should be 1 but ", styles.styleList[0].Wrap)
	}
}

func TestSetNumFmt(t *testing.T) {
	r := strings.NewReader(`<numFmts></numFmts>`)
	tag := &Tag{}
	xml.NewDecoder(r).Decode(tag)
	styles := &Styles{numFmts: tag}
	styles.setNumFmtNumber()
	index := styles.SetNumFmt("#,##0.0")
	if index != 200 {
		t.Error("index should be 200 but ", index)
	}
	b := new(bytes.Buffer)
	xml.NewEncoder(b).Encode(styles.numFmts)
	if b.String() != `<numFmts><numFmt numFmtId="200" formatCode="#,##0.0"></numFmt></numFmts>` {
		t.Error("xml is corrupt [", b.String(), "]")
	}
	index = styles.SetNumFmt(`"`)
	if index != 201 {
		t.Error("index should be 201 but ", index)
	}
	b = new(bytes.Buffer)
	xml.NewEncoder(b).Encode(styles.numFmts.Children[1])
	if b.String() != `<numFmt numFmtId="201" formatCode="&#34;"></numFmt>` {
		t.Error("xml is corrupt [", b.String(), "]")
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
	r := strings.NewReader(`<fills></fills>`)
	tag := &Tag{}
	xml.NewDecoder(r).Decode(tag)
	styles := &Styles{fills: tag}
	index := styles.SetBackgroundColor("FF00FF00")
	if index != 0 {
		t.Error("index should be 0 but", index)
	}
	b := new(bytes.Buffer)
	xml.NewEncoder(b).Encode(styles.fills)
	if b.String() != `<fills><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"></fgColor></patternFill></fill></fills>` {
		t.Error("xml is corrupt [", b.String(), "]")
	}
}

func TestSetBorder(t *testing.T) {
	r := strings.NewReader(`<borders></borders>`)
	tag := &Tag{}
	xml.NewDecoder(r).Decode(tag)
	styles := &Styles{borders: tag}
	border := Border{}
	index := styles.SetBorder(border)
	if index != 0 {
		t.Error("index should be 0 but", index)
	}
	b := new(bytes.Buffer)
	xml.NewEncoder(b).Encode(styles.borders)
	if b.String() != `<borders><border><left></left><right></right><top></top><bottom></bottom></border></borders>` {
		t.Error("xml is corrupt [", b.String(), "]")
	}
	border.Left = &BorderSetting{Style: "thin"}
	index = styles.SetBorder(border)
	if index != 1 {
		t.Error("index should be 1 but", index)
	}
	b = new(bytes.Buffer)
	xml.NewEncoder(b).Encode(styles.borders.Children[index])
	if b.String() != `<border><left style="thin"></left><right></right><top></top><bottom></bottom></border>` {
		t.Error("xml is corrupt [", b.String(), "]")
	}
	border.Left.Color = "FFFFFFFF"
	b = new(bytes.Buffer)
	index = styles.SetBorder(border)
	if index != 2 {
		t.Error("index should be 2 but", index)
	}
	xml.NewEncoder(b).Encode(styles.borders.Children[index])
	if b.String() != `<border><left style="thin"><color rgb="FFFFFFFF"></color></left><right></right><top></top><bottom></bottom></border>` {
		t.Error("xml is corrupt [", b.String(), "]")
	}
	border.Right = &BorderSetting{Style: "hair", Color: "FF000000"}
	border.Top = &BorderSetting{Style: "dashDotDot", Color: "FF111111"}
	border.Bottom = &BorderSetting{Style: "dotted", Color: "FF222222"}
	index = styles.SetBorder(border)
	if index != 3 {
		t.Error("index should be 3 but", index)
	}
	b = new(bytes.Buffer)
	xml.NewEncoder(b).Encode(styles.borders.Children[index])
	if b.String() != `<border><left style="thin"><color rgb="FFFFFFFF"></color></left><right style="hair"><color rgb="FF000000"></color></right><top style="dashDotDot"><color rgb="FF111111"></color></top><bottom style="dotted"><color rgb="FF222222"></color></bottom></border>` {
		t.Error("xml is corrupt [", b.String(), "]")
	}
}

func TestSetStyle(t *testing.T) {
	r := strings.NewReader(`<cellXfs></cellXfs>`)
	tag := &Tag{}
	xml.NewDecoder(r).Decode(tag)
	styles := &Styles{cellXfs: tag}
	style := &Style{}
	index := styles.SetStyle(style)
	if index != 0 {
		t.Error("index should be 0 but", index)
	}
	b := new(bytes.Buffer)
	xml.NewEncoder(b).Encode(styles.cellXfs)
	if b.String() != `<cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"></xf></cellXfs>` {
		t.Error("xml is corrupt [", b.String(), "]")
	}
	style.NumFmtID = 1
	style.FontID = 2
	style.FillID = 3
	style.BorderID = 4
	style.XfID = 5
	index = styles.SetStyle(style)
	if index != 1 {
		t.Error("index should be 1 but", index)
	}
	b = new(bytes.Buffer)
	xml.NewEncoder(b).Encode(styles.cellXfs.Children[index].(*Tag))
	if b.String() != `<xf numFmtId="1" fontId="2" fillId="3" borderId="4" xfId="5" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"></xf>` {
		t.Error("xml is corrupt [", b.String(), "]")
	}
	style.Horizontal = "left"
	style.Vertical = "top"
	index = styles.SetStyle(style)
	if index != 2 {
		t.Error("index should be 2 but", index)
	}
	b = new(bytes.Buffer)
	xml.NewEncoder(b).Encode(styles.cellXfs.Children[index].(*Tag))
	if b.String() != `<xf numFmtId="1" fontId="2" fillId="3" borderId="4" xfId="5" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top"></alignment></xf>` {
		t.Error("xml is corrupt [", b.String(), "]")
	}

}
