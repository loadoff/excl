package excl

import (
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

// Sheet シート操作用の構造体
type Sheet struct {
	xml           *SheetXML
	opened        bool
	Rows          []*Row
	Styles        *Styles
	sheetData     *Tag
	tempFile      *os.File
	afterString   string
	sharedStrings *SharedStrings
	sheetPath     string
	tempSheetPath string
	colsStyles    []ColsStyle
	maxRow        int
}

// SheetXML シートの情報を補完する
type SheetXML struct {
	XMLName xml.Name `xml:"sheet"`
	Name    string   `xml:"name,attr"`
	SheetID string   `xml:"sheetId,attr"`
	RID     string   `xml:"id,attr"`
}

// ColsStyle 列のスタイル情報
type ColsStyle struct {
	min   int
	max   int
	style string
}

// NewSheet シートを作成する
func NewSheet(name string, index int) *Sheet {
	return &Sheet{xml: &SheetXML{
		XMLName: xml.Name{Space: "", Local: "sheet"},
		Name:    name,
		SheetID: fmt.Sprintf("%d", index+1),
		RID:     fmt.Sprintf("rId%d", index+1),
	}}
}

// Create シートを新規に作成する
func (sheet *Sheet) Create(dir string) error {
	f, err := os.Create(filepath.Join(dir, "xl", "worksheets", "sheet"+sheet.xml.SheetID+".xml"))
	if err != nil {
		return err
	}
	defer f.Close()
	f.WriteString(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">`)
	f.WriteString("<sheetData></sheetData>")
	f.WriteString("</worksheet>")
	f.Close()
	sheet.Open(dir)
	return nil
}

// Open シートを開く
func (sheet *Sheet) Open(dir string) error {
	var err error
	sheet.sheetPath = filepath.Join(dir, "xl", "worksheets", "sheet"+sheet.xml.SheetID+".xml")
	sheet.tempSheetPath = filepath.Join(dir, "xl", "worksheets", "__sheet"+sheet.xml.SheetID+".xml")
	f, err := os.Open(sheet.sheetPath)
	if err != nil {
		return err
	}
	defer f.Close()
	tag := &Tag{}
	if err = xml.NewDecoder(f).Decode(tag); err != nil {
		return err
	}
	if err = sheet.setData(tag); err != nil {
		return err
	}
	sheet.setSeparatePoint()

	var b bytes.Buffer
	xml.NewEncoder(&b).Encode(tag)
	strs := strings.Split(b.String(), "<separate_tag></separate_tag>")
	if sheet.tempFile, err = os.Create(sheet.tempSheetPath); err != nil {
		return err
	}
	sheet.tempFile.WriteString(strs[0])
	sheet.afterString = strs[1]
	sheet.opened = true
	return nil
}

// Close シートを閉じる
func (sheet *Sheet) Close() error {
	var err error
	if sheet.opened == false {
		return nil
	}
	sheet.OutputAll()
	if _, err = sheet.tempFile.WriteString(sheet.afterString); err != nil {
		return err
	}
	sheet.tempFile.Close()
	os.Remove(sheet.sheetPath)
	if err := os.Rename(sheet.tempSheetPath, sheet.sheetPath); err != nil {
		return err
	}
	sheet.opened = false
	return nil
}

func (sheet *Sheet) setData(tag *Tag) error {
	if tag.Name.Local != "worksheet" {
		return errors.New("The file [" + sheet.sheetPath + "] is currupt.")
	}
	for _, child := range tag.Children {
		switch tag := child.(type) {
		case *Tag:
			if tag.Name.Local == "sheetData" {
				for _, data := range tag.Children {
					switch row := data.(type) {
					case *Tag:
						if row.Name.Local == "row" {
							newRow := NewRow(row, sheet.sharedStrings, sheet.Styles)
							if newRow == nil {
								return errors.New("The file [" + sheet.sheetPath + "] is currupt.")
							}
							newRow.colsStyles = sheet.colsStyles
							sheet.Rows = append(sheet.Rows, newRow)
							sheet.maxRow = newRow.rowID
						}
					}
				}
				sheet.sheetData = tag
				break
			} else if tag.Name.Local == "cols" {
				sheet.colsStyles = getColsStyles(tag)
			}
		}
	}
	if sheet.sheetData == nil {
		return errors.New("The file[sheet" + sheet.xml.SheetID + ".xml] is currupt. No sheetData tag found.")
	}
	return nil
}

func (sheet *Sheet) setSeparatePoint() {
	sheet.sheetData.Children = []interface{}{separateTag()}
}

// CreateRows 行を作成する
func (sheet *Sheet) CreateRows(from int, to int) []*Row {
	if sheet.maxRow < to {
		sheet.maxRow = to
	}
	rows := make([]*Row, sheet.maxRow)
	for _, row := range sheet.Rows {
		if row == nil {
			continue
		}
		rows[row.rowID-1] = row
	}
	sheet.Rows = rows
	for i := from - 1; i < to; i++ {
		if rows[i] != nil {
			continue
		}
		attr := []xml.Attr{
			xml.Attr{
				Name:  xml.Name{Local: "r"},
				Value: strconv.Itoa(from + i),
			},
		}
		tag := &Tag{
			Name: xml.Name{Local: "row"},
			Attr: attr,
		}
		rows[i] = &Row{rowID: i + 1, row: tag, sharedStrings: sheet.sharedStrings}
	}
	return sheet.Rows
}

// GetRow 行を取得する(インデックス情報は1から開始)
func (sheet *Sheet) GetRow(rowNo int) *Row {
	for _, row := range sheet.Rows {
		if row.rowID == rowNo {
			return row
		}
		if row.rowID > rowNo {
			break
		}
	}
	// 行番号が存在しない行を作成する
	attr := []xml.Attr{
		xml.Attr{
			Name:  xml.Name{Local: "r"},
			Value: strconv.Itoa(int(rowNo)),
		},
	}
	tag := &Tag{
		Name: xml.Name{Local: "row"},
		Attr: attr,
	}
	row := NewRow(tag, sheet.sharedStrings, sheet.Styles)
	row.colsStyles = sheet.colsStyles
	added := false
	rows := make([]*Row, len(sheet.Rows)+1)
	for i := 0; i < len(sheet.Rows); i++ {
		if sheet.Rows[i].rowID < rowNo {
			rows[i] = sheet.Rows[i]
		} else if sheet.Rows[i].rowID > rowNo {
			if !added {
				rows[i] = row
				added = true
			}
			rows[i+1] = sheet.Rows[i]
		}
	}
	if !added {
		rows[len(sheet.Rows)] = row
	}
	sheet.Rows = rows
	return row
}

// OutputAll すべて出力する
func (sheet *Sheet) OutputAll() {
	for _, row := range sheet.Rows {
		if row != nil {
			row.resetStyleIndex()
			xml.NewEncoder(sheet.tempFile).Encode(row)
		}
	}
	sheet.Rows = nil
}

// OutputThroughRowNo rowNoまですべて出力する
func (sheet *Sheet) OutputThroughRowNo(rowNo int) {
	var i int
	for i = 0; i < len(sheet.Rows); i++ {
		if sheet.Rows[i] == nil {
			continue
		}
		if rowNo < sheet.Rows[i].rowID {
			break
		}
		sheet.Rows[i].resetStyleIndex()
		xml.NewEncoder(sheet.tempFile).Encode(sheet.Rows[i])
	}
	sheet.Rows = sheet.Rows[i:]
}

// getColsStyles 列に設定されているスタイルを取得する
func getColsStyles(tag *Tag) []ColsStyle {
	var styles []ColsStyle
	for _, child := range tag.Children {
		switch t := child.(type) {
		case *Tag:
			if t.Name.Local != "col" {
				continue
			}
			var style ColsStyle
			for _, attr := range t.Attr {
				if attr.Name.Local == "min" {
					style.min, _ = strconv.Atoi(attr.Value)
				} else if attr.Name.Local == "max" {
					style.max, _ = strconv.Atoi(attr.Value)
				} else if attr.Name.Local == "style" {
					style.style = attr.Value
				}
			}
			styles = append(styles, style)
		}
	}
	return styles
}

func getStyleNo(styles []ColsStyle, colNo int) string {
	for _, style := range styles {
		if style.min <= 0 || style.max <= 0 {
			continue
		}
		if style.min <= colNo && style.max >= colNo {
			return style.style
		}
	}
	return ""
}
