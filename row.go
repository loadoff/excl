package excl

import (
	"encoding/xml"
	"fmt"
	"math"
	"sort"
	"strconv"
	"time"
)

// Row 行の構造体
type Row struct {
	rowID         int
	row           *Tag
	cells         []*Cell
	sharedStrings *SharedStrings
	colInfos      []colInfo
	style         string
	minColNo      int
	maxColNo      int
	styles        *Styles
}

// NewRow は新しく行を追加する際に使用する
func NewRow(tag *Tag, sharedStrings *SharedStrings, styles *Styles) *Row {
	row := &Row{row: tag, sharedStrings: sharedStrings, styles: styles}
	for _, attr := range tag.Attr {
		if attr.Name.Local == "r" {
			row.rowID, _ = strconv.Atoi(attr.Value)
		} else if attr.Name.Local == "s" {
			row.style = attr.Value
		}
	}
	for _, child := range tag.Children {
		switch col := child.(type) {
		case *Tag:
			if col.Name.Local == "c" {
				cell := NewCell(col, sharedStrings, styles)
				if cell == nil {
					return nil
				}
				cell.styles = row.styles
				row.cells = append(row.cells, cell)
				row.maxColNo = cell.colNo
				if row.minColNo == 0 {
					row.minColNo = cell.colNo
				}
			}
		}
	}
	return row
}

// CreateCells セル一覧を用意する
func (row *Row) CreateCells(from int, to int) []*Cell {
	if row.maxColNo < to {
		row.maxColNo = to
	}
	cells := make([]*Cell, row.maxColNo)
	for _, cell := range row.cells {
		if cell == nil || cell.colNo == 0 {
			continue
		}
		cells[cell.colNo-1] = cell
	}
	for i := from; i <= to; i++ {
		if cells[i-1] != nil {
			continue
		}
		attr := []xml.Attr{
			xml.Attr{
				Name:  xml.Name{Local: "r"},
				Value: fmt.Sprintf("%s%d", ColStringPosition(i), row.rowID),
			},
		}
		tag := &Tag{
			Name: xml.Name{Local: "c"},
			Attr: attr,
		}
		style := 0
		if row.style != "" {
			style, _ = strconv.Atoi(row.style)
		} else {
			for _, colInfo := range row.colInfos {
				if colInfo.style != "" && colInfo.min <= i && i <= colInfo.max {
					style, _ = strconv.Atoi(colInfo.style)
					break
				}
			}
		}
		cells[i-1] = &Cell{cell: tag, colNo: i, sharedStrings: row.sharedStrings, styleIndex: style, styles: row.styles}
	}
	row.cells = cells
	return row.cells
}

// GetCell セル番号のセルを取得する
func (row *Row) GetCell(colNo int) *Cell {

	for i := len(row.cells) - 1; i >= 0; i-- {
		//	for _, cell := range row.cells {
		cell := row.cells[i]
		if cell.colNo == colNo {
			return cell
		}
		if cell.colNo < colNo {
			break
		}
	}

	// 存在しない場合はセルを追加する
	tag := &Tag{Name: xml.Name{Local: "c"}}
	tag.setAttr("r", fmt.Sprintf("%s%d", ColStringPosition(int(colNo)), row.rowID))
	if row.style != "" {
		tag.setAttr("s", row.style)
	} else if style := getStyleNo(row.colInfos, int(colNo)); style != "" {
		tag.setAttr("s", style)
	}

	cell := NewCell(tag, row.sharedStrings, row.styles)
	row.cells = append(row.cells, cell)
	return cell
}

// SetString set string at a row
func (row *Row) SetString(val string, colNo int) *Cell {
	cell := row.GetCell(colNo).SetString(val)
	return cell
}

// SetNumber set number at a row
func (row *Row) SetNumber(val interface{}, colNo int) *Cell {
	cell := row.GetCell(colNo).SetNumber(val)
	return cell
}

// SetFormula set a formula at a row
func (row *Row) SetFormula(val string, colNo int) *Cell {
	cell := row.GetCell(colNo).SetFormula(val)
	return cell
}

// SetDate set a date at a row
func (row *Row) SetDate(val time.Time, colNo int) *Cell {
	cell := row.GetCell(colNo).SetDate(val)
	return cell
}

// SetHeight set row height
func (row *Row) SetHeight(height float64) {
	row.row.setAttr("customHeight", "1")
	row.row.setAttr("ht", strconv.FormatFloat(height, 'f', 4, 64))
}

// ColStringPosition obtain AtoZ column string from column no
func ColStringPosition(num int) string {
	atoz := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	if num <= 26 {
		return atoz[num-1]
	}
	return ColStringPosition((num-1)/26) + atoz[(num-1)%26]
}

// ColNumPosition obtain column no from AtoZ column string
func ColNumPosition(col string) int {
	var num int
	for i := len(col) - 1; i >= 0; i-- {
		p := math.Pow(26, float64(len(col)-i-1))
		num += int(p) * int(col[i]-0x40)
	}
	return num
}

// MarshalXML Create xml tags
func (row *Row) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	sort.Slice(row.cells, func(i, j int) bool {
		return row.cells[i].colNo < row.cells[j].colNo
	})
	start.Name = row.row.Name
	start.Attr = row.row.Attr
	e.EncodeToken(start)
	if err := e.Encode(row.cells); err != nil {
		return err
	}
	e.EncodeToken(start.End())
	return nil
}

func (row *Row) resetStyleIndex() {
	for _, cell := range row.cells {
		cell.resetStyleIndex()
	}
}
