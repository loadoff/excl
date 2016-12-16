package excl

import (
	"encoding/xml"
	"errors"
	"os"
	"path/filepath"
	"strconv"
)

const defaultMaxNumfmt = 200

// Styles スタイルの情報を持った構造体
type Styles struct {
	path         string
	styles       *Tag
	numFmts      *Tag
	fonts        *Tag
	fills        *Tag
	borders      *Tag
	cellStyleXfs *Tag
	cellXfs      *Tag
	numFmtNumber int
}

// Style セルの書式情報
type Style struct {
	numFmt            int
	font              int
	fill              int
	border            int
	xfID              int
	applyNumberFormat int
	applyFont         int
	applyFill         int
	applyBorder       int
	applyAlignment    int
	applyProtection   int
}

// OpenStyles styles.xmlファイルを開く
func OpenStyles(dir string) (*Styles, error) {
	var f *os.File
	var err error
	path := filepath.Join(dir, "xl", "styles.xml")
	f, err = os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tag := &Tag{}
	err = xml.NewDecoder(f).Decode(tag)
	if err != nil {
		return nil, err
	}
	styles := &Styles{styles: tag, path: path}
	err = styles.setData()
	if err != nil {
		return nil, err
	}
	return styles, nil
}

// Close styles.xmlファイルを閉じる
func (styles *Styles) Close() error {
	f, err := os.Create(styles.path)
	if err != nil {
		return err
	}
	defer f.Close()
	f.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
	err = xml.NewEncoder(f).Encode(styles)
	if err != nil {
		return err
	}
	return nil
}

func (styles *Styles) setData() error {
	tag := styles.styles
	if tag == nil {
		return errors.New("Tag (Styles.styles) is nil.")
	}
	if tag.Name.Local != "styleSheet" {
		return errors.New("The styles.xml file is currupt.")
	}
	for _, child := range tag.Children {
		switch tag := child.(type) {
		case *Tag:
			if tag.Name.Local == "numFmts" {
				styles.numFmts = tag
			} else if tag.Name.Local == "fonts" {
				styles.fonts = tag
			} else if tag.Name.Local == "fills" {
				styles.fills = tag
			} else if tag.Name.Local == "borders" {
				styles.borders = tag
			} else if tag.Name.Local == "cellStyleXfs" {
				styles.cellStyleXfs = tag
			} else if tag.Name.Local == "cellXfs" {
				styles.cellXfs = tag
			}
		}
	}
	styles.numFmtNumber = defaultMaxNumfmt
	if styles.numFmts != nil {
		styles.setNumFmtNumber()
	}
	return nil
}

// setNumFmtNumber フォーマットID
func (styles *Styles) setNumFmtNumber() {
	max := defaultMaxNumfmt
	for _, child := range styles.numFmts.Children {
		switch tag := child.(type) {
		case *Tag:
			for _, attr := range tag.Attr {
				if attr.Name.Local == "numFmtId" {
					i, _ := strconv.Atoi(attr.Value)
					if max <= i {
						max = i + 1
					}
				}
			}
		}
	}
	styles.numFmtNumber = max
}

// SetNumFmt 数値フォーマットをセットする
func (styles *Styles) SetNumFmt(format string) int {
	if styles.numFmts == nil {
		styles.numFmts = &Tag{Name: xml.Name{Local: "numFmts"}}
	}
	styles.numFmtNumber++
	tag := &Tag{Name: xml.Name{Local: "numFmt"}}
	tag.setAttr("numFmtId", strconv.Itoa(styles.numFmtNumber))
	tag.setAttr("formatCode", format)
	styles.numFmts.Children = append(styles.numFmts.Children, tag)
	return styles.numFmtNumber
}

func (styles *Styles) setCellXfs(numFmtID int, fontID int, fillID int, borderID int, xfID int) int {
	attr := []xml.Attr{
		xml.Attr{
			Name:  xml.Name{Local: "numFmtId"},
			Value: strconv.Itoa(numFmtID),
		},
		xml.Attr{
			Name:  xml.Name{Local: "fontId"},
			Value: strconv.Itoa(fontID),
		},
		xml.Attr{
			Name:  xml.Name{Local: "fillId"},
			Value: strconv.Itoa(fillID),
		},
		xml.Attr{
			Name:  xml.Name{Local: "borderId"},
			Value: strconv.Itoa(borderID),
		},
		xml.Attr{
			Name:  xml.Name{Local: "xfId"},
			Value: strconv.Itoa(xfID),
		},
	}
	if numFmtID != 0 {
		attr = append(attr, xml.Attr{
			Name:  xml.Name{Local: "applyNumberFormat"},
			Value: "1",
		})
	}
	if fontID != 0 {
		attr = append(attr, xml.Attr{
			Name:  xml.Name{Local: "applyFont"},
			Value: "1",
		})
	}
	if fillID != 0 {
		attr = append(attr, xml.Attr{
			Name:  xml.Name{Local: "applyFill"},
			Value: "1",
		})
	}
	if borderID != 0 {
		attr = append(attr, xml.Attr{
			Name:  xml.Name{Local: "applyBorder"},
			Value: "1",
		})
	}
	tag := &Tag{
		Name: xml.Name{Local: "xf"},
		Attr: attr,
	}
	styles.cellXfs.Children = append(styles.cellXfs.Children, tag)
	return len(styles.cellXfs.Children) - 1
}

func (styles *Styles) getCellXfs(index int) *Tag {
	if len(styles.cellXfs.Children) < index {
		return styles.cellXfs.Children[index].(*Tag)
	}
	return nil
}

// MarshalXML stylesからXMLを作り直す
func (styles *Styles) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = styles.styles.Name
	start.Attr = styles.styles.Attr
	e.EncodeToken(start)
	if styles.numFmts != nil {
		e.Encode(styles.numFmts)
	}
	if styles.fonts != nil {
		e.Encode(styles.fonts)
	}
	if styles.fills != nil {
		e.Encode(styles.fills)
	}
	if styles.borders != nil {
		e.Encode(styles.borders)
	}
	if styles.cellStyleXfs != nil {
		e.Encode(styles.cellStyleXfs)
	}
	if styles.cellXfs != nil {
		e.Encode(styles.cellXfs)
	}
	outputsList := []string{"numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs"}
	for _, v := range styles.styles.Children {
		switch v.(type) {
		case *Tag:
			child := v.(*Tag)
			if !IsExistString(outputsList, child.Name.Local) {
				if err := e.Encode(child); err != nil {
					return err
				}
			}
		}
	}
	e.EncodeToken(start.End())
	return nil
}

// IsExistString 配列内に文字列が存在するかを確認する
func IsExistString(strs []string, str string) bool {
	for _, s := range strs {
		if s == str {
			return true
		}
	}
	return false
}
