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
	styleList    []*Style
	numFmtNumber int
	fontCount    int
	fillCount    int
	borderCount  int
}

// Style セルの書式情報
type Style struct {
	xf                *Tag
	NumFmtID          int
	FontID            int
	FillID            int
	BorderID          int
	XfID              int
	applyNumberFormat int
	applyFont         int
	applyFill         int
	applyBorder       int
	applyAlignment    int
	applyProtection   int
	Horizontal        string
	Vertical          string
}

// Font フォントの設定
type Font struct {
	Size      int
	Color     string
	Name      string
	Bold      bool
	Italic    bool
	Underline bool
}

// BorderSetting 罫線の設定
type BorderSetting struct {
	Style string
	Color string
}

// Border 罫線の設定
type Border struct {
	Left   *BorderSetting
	Right  *BorderSetting
	Top    *BorderSetting
	Bottom *BorderSetting
}

// createStyles styles.xmlを作成する
func createStyles(dir string) error {
	os.Mkdir(filepath.Join(dir, "xl"), 0755)
	path := filepath.Join(dir, "xl", "styles.xml")
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	f.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
	f.WriteString(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main">`)
	f.WriteString(`<fonts><font/></fonts>`)
	f.WriteString(`<fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>`)
	f.WriteString(`<borders><border/></borders>`)
	f.WriteString(`<cellStyleXfs><xf/></cellStyleXfs>`)
	f.WriteString(`<cellXfs><xf/></cellXfs>`)
	f.WriteString(`</styleSheet>`)
	f.Close()
	return nil
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
	if styles == nil {
		return nil
	}
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
	styles.numFmtNumber = defaultMaxNumfmt
	for _, child := range tag.Children {
		switch tag := child.(type) {
		case *Tag:
			switch tag.Name.Local {
			case "numFmts":
				styles.numFmts = tag
				styles.setNumFmtNumber()
			case "fonts":
				styles.fonts = tag
				styles.setFontCount()
			case "fills":
				styles.fills = tag
				styles.setFillCount()
			case "borders":
				styles.borders = tag
				styles.setBorderCount()
			case "cellStyleXfs":
				styles.cellStyleXfs = tag
			case "cellXfs":
				styles.cellXfs = tag
				styles.setStyleList()
			}
		}
	}
	return nil
}

func (styles *Styles) setStyleList() {
	for _, child := range styles.cellXfs.Children {
		switch child.(type) {
		case *Tag:
			t := child.(*Tag)
			if t.Name.Local == "xf" {
				style := &Style{xf: t}
				for _, attr := range t.Attr {
					index, _ := strconv.Atoi(attr.Value)
					switch attr.Name.Local {
					case "numFmtId":
						style.NumFmtID = index
					case "fontId":
						style.FontID = index
					case "fillId":
						style.FillID = index
					case "borderId":
						style.BorderID = index
					case "applyNumberFormat":
						style.applyNumberFormat = index
					case "applyFont":
						style.applyFont = index
					case "applyFill":
						style.applyFill = index
					case "applyBorder":
						style.applyBorder = index
					case "applyAlignment":
						style.applyAlignment = index
					case "applyProtection":
						style.applyProtection = index
					}
				}
				// alignment
				if style.applyAlignment == 1 {
					for _, xfChild := range t.Children {
						switch xfChild.(type) {
						case *Tag:
							cTag := xfChild.(*Tag)
							if cTag.Name.Local == "alignment" {
								for _, attr := range cTag.Attr {
									if attr.Name.Local == "horizontal" {
										style.Horizontal = attr.Value
									} else if attr.Name.Local == "vertical" {
										style.Vertical = attr.Value
									}
								}
							}
							break
						}
					}
				}
				styles.styleList = append(styles.styleList, style)
			}
		}
	}
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

func (styles *Styles) setFontCount() {
	for _, tag := range styles.fonts.Children {
		switch tag.(type) {
		case *Tag:
			if tag.(*Tag).Name.Local == "font" {
				styles.fontCount++
			}
		}
	}
}

func (styles *Styles) setFillCount() {
	for _, tag := range styles.fills.Children {
		switch tag.(type) {
		case *Tag:
			if tag.(*Tag).Name.Local == "fill" {
				styles.fillCount++
			}
		}
	}
}

func (styles *Styles) setBorderCount() {
	for _, tag := range styles.borders.Children {
		switch tag.(type) {
		case *Tag:
			if tag.(*Tag).Name.Local == "border" {
				styles.borderCount++
			}
		}
	}
}

// SetNumFmt 数値フォーマットをセットする
func (styles *Styles) SetNumFmt(format string) int {
	if styles.numFmts == nil {
		styles.numFmts = &Tag{Name: xml.Name{Local: "numFmts"}}
	}
	tag := &Tag{Name: xml.Name{Local: "numFmt"}}
	tag.setAttr("numFmtId", strconv.Itoa(styles.numFmtNumber))
	tag.setAttr("formatCode", format)
	styles.numFmts.Children = append(styles.numFmts.Children, tag)
	styles.numFmtNumber++
	return styles.numFmtNumber - 1
}

// SetFont フォント情報を追加する
func (styles *Styles) SetFont(font Font) int {
	tag := &Tag{Name: xml.Name{Local: "font"}}
	var t *Tag
	if font.Size > 0 {
		t = &Tag{Name: xml.Name{Local: "sz"}}
		t.setAttr("val", strconv.Itoa(font.Size))
		tag.Children = append(tag.Children, t)
	}
	if font.Name != "" {
		t = &Tag{Name: xml.Name{Local: "name"}}
		t.setAttr("val", font.Name)
		tag.Children = append(tag.Children, t)
	}
	if font.Color != "" {
		t = &Tag{Name: xml.Name{Local: "color"}}
		t.setAttr("rgb", font.Color)
		tag.Children = append(tag.Children, t)
	}
	if font.Bold {
		t = &Tag{Name: xml.Name{Local: "b"}}
		tag.Children = append(tag.Children, t)
	}
	if font.Italic {
		t = &Tag{Name: xml.Name{Local: "i"}}
		tag.Children = append(tag.Children, t)
	}
	if font.Underline {
		t = &Tag{Name: xml.Name{Local: "u"}}
		tag.Children = append(tag.Children, t)
	}
	styles.fonts.Children = append(styles.fonts.Children, tag)
	styles.fontCount++
	return styles.fontCount - 1
}

// SetBackgroundColor 背景色を追加する
func (styles *Styles) SetBackgroundColor(color string) int {
	tag := &Tag{Name: xml.Name{Local: "fill"}}
	patternFill := &Tag{Name: xml.Name{Local: "patternFill"}}
	patternFill.setAttr("patternType", "solid")
	fgColor := &Tag{Name: xml.Name{Local: "fgColor"}}
	fgColor.setAttr("rgb", color)
	patternFill.Children = []interface{}{fgColor}
	tag.Children = []interface{}{patternFill}
	styles.fills.Children = append(styles.fills.Children, tag)
	styles.fillCount++
	return styles.fillCount - 1
}

// SetBorder 罫線を設定する
func (styles *Styles) SetBorder(border Border) int {
	var color *Tag
	tag := &Tag{Name: xml.Name{Local: "border"}}
	left := &Tag{Name: xml.Name{Local: "left"}}
	right := &Tag{Name: xml.Name{Local: "right"}}
	top := &Tag{Name: xml.Name{Local: "top"}}
	bottom := &Tag{Name: xml.Name{Local: "bottom"}}

	if border.Left != nil {
		left.setAttr("style", border.Left.Style)
		if border.Left.Color != "" {
			color = &Tag{Name: xml.Name{Local: "color"}}
			color.setAttr("rgb", border.Left.Color)
			left.Children = []interface{}{color}
		}
	}
	if border.Right != nil {
		right.setAttr("style", border.Right.Style)
		if border.Right.Color != "" {
			color = &Tag{Name: xml.Name{Local: "color"}}
			color.setAttr("rgb", border.Right.Color)
			right.Children = []interface{}{color}
		}
	}

	if border.Top != nil {
		top.setAttr("style", border.Top.Style)
		if border.Top.Color != "" {
			color = &Tag{Name: xml.Name{Local: "color"}}
			color.setAttr("rgb", border.Top.Color)
			top.Children = []interface{}{color}
		}
	}
	if border.Bottom != nil {
		bottom.setAttr("style", border.Bottom.Style)
		if border.Bottom.Color != "" {
			color = &Tag{Name: xml.Name{Local: "color"}}
			color.setAttr("rgb", border.Bottom.Color)
			bottom.Children = []interface{}{color}
		}
	}

	tag.Children = append(tag.Children, left)
	tag.Children = append(tag.Children, right)
	tag.Children = append(tag.Children, top)
	tag.Children = append(tag.Children, bottom)
	styles.borders.Children = append(styles.borders.Children, tag)
	styles.borderCount++
	return styles.borderCount - 1
}

// SetStyle セルの書式を設定
func (styles *Styles) SetStyle(style *Style) int {
	// すでに同じ書式が存在する場合はその書式を使用する
	for index, s := range styles.styleList {
		if s.NumFmtID == style.NumFmtID &&
			s.FontID == style.FontID &&
			s.FillID == style.FillID &&
			s.BorderID == style.BorderID &&
			s.XfID == style.XfID &&
			s.Horizontal == style.Horizontal &&
			s.Vertical == style.Vertical {
			return index
		}
	}
	return styles.SetCellXfs(style)
}

// GetStyle Style構造体を取得する
func (styles *Styles) GetStyle(index int) *Style {
	if len(styles.styleList) < index {
		return nil
	}
	return styles.styleList[index]
}

// SetCellXfs cellXfsにタグを追加する
func (styles *Styles) SetCellXfs(style *Style) int {
	s := &Style{
		NumFmtID:   style.NumFmtID,
		FontID:     style.FontID,
		FillID:     style.FillID,
		XfID:       style.XfID,
		Horizontal: style.Horizontal,
		Vertical:   style.Vertical,
	}
	attr := []xml.Attr{
		xml.Attr{
			Name:  xml.Name{Local: "numFmtId"},
			Value: strconv.Itoa(style.NumFmtID),
		},
		xml.Attr{
			Name:  xml.Name{Local: "fontId"},
			Value: strconv.Itoa(style.FontID),
		},
		xml.Attr{
			Name:  xml.Name{Local: "fillId"},
			Value: strconv.Itoa(style.FillID),
		},
		xml.Attr{
			Name:  xml.Name{Local: "borderId"},
			Value: strconv.Itoa(style.BorderID),
		},
		xml.Attr{
			Name:  xml.Name{Local: "xfId"},
			Value: strconv.Itoa(style.XfID),
		},
	}
	if style.NumFmtID != 0 {
		attr = append(attr, xml.Attr{
			Name:  xml.Name{Local: "applyNumberFormat"},
			Value: "1",
		})
		s.applyNumberFormat = 1
	}
	if style.FontID != 0 {
		attr = append(attr, xml.Attr{
			Name:  xml.Name{Local: "applyFont"},
			Value: "1",
		})
		s.applyFont = 1
	}
	if style.FillID != 0 {
		attr = append(attr, xml.Attr{
			Name:  xml.Name{Local: "applyFill"},
			Value: "1",
		})
		s.applyFill = 1
	}
	if style.BorderID != 0 {
		attr = append(attr, xml.Attr{
			Name:  xml.Name{Local: "applyBorder"},
			Value: "1",
		})
		s.applyBorder = 1
	}
	tag := &Tag{
		Name: xml.Name{Local: "xf"},
		Attr: attr,
	}
	if style.Horizontal != "" || style.Vertical != "" {
		alignment := &Tag{Name: xml.Name{Local: "alignment"}}
		if style.Horizontal != "" {
			alignment.setAttr("horizontal", style.Horizontal)
		}
		if style.Vertical != "" {
			alignment.setAttr("vertical", style.Vertical)
		}
		tag.Children = []interface{}{alignment}
		tag.Attr = append(tag.Attr, xml.Attr{
			Name:  xml.Name{Local: "applyAlignment"},
			Value: "1",
		})
		s.applyAlignment = 1
	}
	s.xf = tag
	styles.cellXfs.Children = append(styles.cellXfs.Children, tag)
	styles.styleList = append(styles.styleList, s)
	return len(styles.styleList) - 1
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
