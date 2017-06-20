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

// Sheet struct for control sheet data
type Sheet struct {
	xml           *SheetXML
	opened        bool
	Rows          []*Row
	Styles        *Styles
	worksheet     *Tag
	sheetView     *Tag
	cols          *Tag
	sheetData     *Tag
	tempFile      *os.File
	afterString   string
	sharedStrings *SharedStrings
	sheetPath     string
	tempSheetPath string
	colInfos      colInfos
	maxRow        int
	sheetIndex    int
}

// SheetXML sheet.xml information
type SheetXML struct {
	XMLName xml.Name `xml:"sheet"`
	Name    string   `xml:"name,attr"`
	SheetID string   `xml:"sheetId,attr"`
	RID     string   `xml:"id,attr"`
}

// colInfo 列のスタイル情報
type colInfo struct {
	min         int
	max         int
	style       string
	width       float64
	customWidth bool
}

type colInfos []colInfo

// NewSheet create new sheet information.
func NewSheet(name string, index int, maxSheetID int) *Sheet {
	return &Sheet{
		xml: &SheetXML{
			XMLName: xml.Name{Space: "", Local: "sheet"},
			Name:    name,
			SheetID: fmt.Sprintf("%d", maxSheetID+1),
			RID:     fmt.Sprintf("rId%d", index+1),
		},
		sheetIndex: index + 1,
	}
}

// Create シートを新規に作成する
func (sheet *Sheet) Create(dir string) error {
	f, err := os.Create(filepath.Join(dir, "xl", "worksheets", fmt.Sprintf("sheet%d.xml", sheet.sheetIndex)))
	if err != nil {
		return err
	}
	defer f.Close()
	f.WriteString(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">`)
	f.WriteString(`<sheetViews><sheetView workbookViewId="0"></sheetView></sheetViews>`)
	f.WriteString("<sheetData></sheetData>")
	f.WriteString("</worksheet>")
	f.Close()
	sheet.Open(dir)
	return nil
}

// Open open sheet.xml in directory
func (sheet *Sheet) Open(dir string) error {
	var err error
	sheet.sheetPath = filepath.Join(dir, "xl", "worksheets", fmt.Sprintf("sheet%d.xml", sheet.sheetIndex))
	sheet.tempSheetPath = filepath.Join(dir, "xl", "worksheets", fmt.Sprintf("__sheet%d.xml", sheet.sheetIndex))
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
	sheet.worksheet = tag
	sheet.setSeparatePoint()
	if sheet.tempFile, err = os.Create(sheet.tempSheetPath); err != nil {
		return err
	}
	sheet.opened = true
	return nil
}

// Close シートを閉じる
func (sheet *Sheet) Close() error {
	var err error
	if sheet == nil || sheet.opened == false {
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
	sheet.worksheet = nil
	sheet.sheetView = nil
	sheet.sheetData = nil
	sheet.tempFile = nil
	return nil
}

func (sheet *Sheet) setData(sheetTag *Tag) error {
	if sheetTag.Name.Local != "worksheet" {
		return errors.New("The file [" + sheet.sheetPath + "] is currupt.")
	}
	for _, child := range sheetTag.Children {
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
							newRow.colInfos = sheet.colInfos
							sheet.Rows = append(sheet.Rows, newRow)
							sheet.maxRow = newRow.rowID
						}
					}
				}
				sheet.sheetData = tag
				break
			} else if tag.Name.Local == "cols" {
				sheet.cols = tag
				sheet.colInfos = getColInfos(tag)
			} else if tag.Name.Local == "sheetViews" {
				for _, view := range tag.Children {
					if v, ok := view.(*Tag); ok {
						if v.Name.Local == "sheetView" {
							sheet.sheetView = v
						}
					}
				}
			}
		}
	}
	if sheet.sheetData == nil {
		return errors.New("The file[sheet" + sheet.xml.SheetID + ".xml] is currupt. No sheetData tag found.")
	}
	return nil
}

func (sheet *Sheet) setSeparatePoint() {
	for i := 0; i < len(sheet.worksheet.Children); i++ {
		if sheet.cols != nil && sheet.worksheet.Children[i] == sheet.cols {
			sheet.worksheet.Children[i] = separateTag()
			break
		} else if sheet.worksheet.Children[i] == sheet.sheetData {
			sheet.cols = &Tag{Name: xml.Name{Local: "cols"}}
			sheet.worksheet.Children = append(sheet.worksheet.Children[:i], append([]interface{}{separateTag()},
				sheet.worksheet.Children[i:]...)...)
			break
		}
	}
	sheet.sheetData.Children = []interface{}{separateTag()}
}

// CreateRows create multiple rows
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
	row.colInfos = sheet.colInfos
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

// ShowGridlines グリッド線の表示非表示
func (sheet *Sheet) ShowGridlines(show bool) {
	if sheet.sheetView != nil {
		if show {
			sheet.sheetView.setAttr("showGridLines", "1")
		} else {
			sheet.sheetView.setAttr("showGridLines", "0")
		}
	}
}

func (sheet *Sheet) outputFirst() {
	var b bytes.Buffer
	xml.NewEncoder(&b).Encode(sheet.worksheet)
	strs := strings.Split(b.String(), "<separate_tag></separate_tag>")
	sheet.tempFile.WriteString(strs[0])
	if len(sheet.colInfos) != 0 {
		xml.NewEncoder(sheet.tempFile).Encode(sheet.colInfos)
	}
	sheet.tempFile.WriteString(strs[1])
	sheet.afterString = strs[2]
	sheet.worksheet = nil
}

// OutputAll すべて出力する
func (sheet *Sheet) OutputAll() {
	if sheet.worksheet != nil {
		sheet.outputFirst()
	}
	for _, row := range sheet.Rows {
		if row != nil {
			row.resetStyleIndex()
			xml.NewEncoder(sheet.tempFile).Encode(row)
		}
	}
	sheet.Rows = nil
}

// OutputThroughRowNo output through to rowno
func (sheet *Sheet) OutputThroughRowNo(rowNo int) {
	var i int
	if sheet.worksheet != nil {
		sheet.outputFirst()
	}
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
func getColInfos(tag *Tag) []colInfo {
	var infos []colInfo
	for _, child := range tag.Children {
		switch t := child.(type) {
		case *Tag:
			if t.Name.Local != "col" {
				continue
			}
			info := colInfo{width: -1}
			for _, attr := range t.Attr {
				if attr.Name.Local == "min" {
					info.min, _ = strconv.Atoi(attr.Value)
				} else if attr.Name.Local == "max" {
					info.max, _ = strconv.Atoi(attr.Value)
				} else if attr.Name.Local == "style" {
					info.style = attr.Value
				} else if attr.Name.Local == "width" {
					info.width, _ = strconv.ParseFloat(attr.Value, 64)
				} else if attr.Name.Local == "customWidth" {
					info.customWidth = false
					if attr.Value == "1" {
						info.customWidth = true
					}
				}
			}
			infos = append(infos, info)
		}
	}
	return infos
}

func getStyleNo(styles []colInfo, colNo int) string {
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

// SetColWidth set column width
func (sheet *Sheet) SetColWidth(width float64, colNo int) {
	for i, info := range sheet.colInfos {
		if info.min == colNo && colNo == info.max {
			// if min and max are same as colNo just replace width and customWidth value
			sheet.colInfos[i].width = width
			sheet.colInfos[i].customWidth = true
			return
		} else if info.min == colNo {
			// insert info before index
			info.width = width
			info.max = colNo
			info.customWidth = true
			sheet.colInfos[i].min++
			if i > 0 {
				sheet.colInfos = append(sheet.colInfos[:i-1], append([]colInfo{info}, sheet.colInfos[i-1:]...)...)
			} else {
				sheet.colInfos = append([]colInfo{info}, sheet.colInfos...)
			}
			return
		} else if info.max == colNo {
			// insert info after index
			sheet.colInfos[i].max--
			info.width = width
			info.min = colNo
			info.customWidth = true
			sheet.colInfos = append(sheet.colInfos[:i+1], append([]colInfo{info}, sheet.colInfos[i+1:]...)...)
			return
		} else if info.min < colNo && colNo < info.max {
			// devide three deferent informations
			beforeInfo := info.clone()
			afterInfo := info.clone()
			beforeInfo.max = colNo - 1
			afterInfo.min = colNo + 1
			info.width = width
			info.min = colNo
			info.max = colNo
			info.customWidth = true
			sheet.colInfos = append(sheet.colInfos[:i], append([]colInfo{beforeInfo, info, afterInfo}, sheet.colInfos[i+1:]...)...)
			return
		} else if info.min > colNo {
			// insert colInfo after index
			info.width = width
			info.min = colNo
			info.max = colNo
			info.customWidth = true
			sheet.colInfos = append(sheet.colInfos[:i], append([]colInfo{info}, sheet.colInfos[i:]...)...)
			return
		}
	}
	info := colInfo{min: colNo, max: colNo, width: width, customWidth: true}
	sheet.colInfos = append(sheet.colInfos, info)
}

func (info colInfo) clone() colInfo {
	return colInfo{min: info.min, max: info.max, style: info.style, width: info.width, customWidth: info.customWidth}
}

func (infos colInfos) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "cols"}
	e.EncodeToken(start)
	for _, info := range infos {
		if err := e.Encode(info); err != nil {
			return err
		}
	}
	e.EncodeToken(start.End())
	return nil
}

func (info colInfo) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "col"}
	start.Attr = []xml.Attr{
		xml.Attr{Name: xml.Name{Local: "min"}, Value: strconv.Itoa(info.min)},
		xml.Attr{Name: xml.Name{Local: "max"}, Value: strconv.Itoa(info.max)},
	}
	if info.style != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "style"}, Value: info.style})
	}
	if info.customWidth || info.width != -1 {
		if info.customWidth {
			start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "width"}, Value: fmt.Sprint(info.width)})
		} else {
			start.Attr = append(start.Attr, []xml.Attr{
				xml.Attr{Name: xml.Name{Local: "width"}, Value: fmt.Sprint(info.width)},
				xml.Attr{Name: xml.Name{Local: "customWidth"}, Value: "1"},
			}...)
		}
	}
	e.EncodeToken(start)
	e.EncodeToken(start.End())
	return nil
}
