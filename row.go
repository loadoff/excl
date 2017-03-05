package excl

import (
	"encoding/xml"
	"fmt"
	"math"
	"sort"
	"strconv"
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
		cells[i-1] = &Cell{cell: tag, colNo: i, sharedStrings: row.sharedStrings, styleIndex: style}
	}
	row.cells = cells
	return row.cells
}

// ByCell セルに対するソート用
type ByCell []*Cell

func (c ByCell) Len() int           { return len(c) }
func (c ByCell) Swap(i, j int)      { c[i], c[j] = c[j], c[i] }
func (c ByCell) Less(i, j int) bool { return c[i].colNo < c[j].colNo }

// GetCell セル番号のセルを取得する
func (row *Row) GetCell(colNo int) *Cell {

	for _, cell := range row.cells {
		if cell.colNo == colNo {
			return cell
		}
		if cell.colNo > colNo {
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
	sort.Sort(ByCell(row.cells))
	return cell
}

// SetString 文字を出力する
func (row *Row) SetString(val string, colNo int) *Cell {
	cell := row.GetCell(colNo).SetString(val)
	return cell
}

// SetNumber 数値を出力する
func (row *Row) SetNumber(val string, colNo int) *Cell {
	cell := row.GetCell(colNo).SetNumber(val)
	return cell
}

// ColStringPosition カラム番号からA-Z文字列を取得する
func ColStringPosition(num int) string {
	atoz := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	if num <= 26 {
		return atoz[num-1]
	}
	return ColStringPosition(num/26) + atoz[(num-1)%26]
}

// ColNumPosition カラム番号をA-Z文字列から取得する
func ColNumPosition(col string) int {
	var num int
	for i := len(col) - 1; i >= 0; i-- {
		p := math.Pow(26, float64(len(col)-i-1))
		num += int(p) * int(col[i]-0x40)
	}
	return num
}

// MarshalXML タグを作成しなおす
func (row *Row) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = row.row.Name
	start.Attr = row.row.Attr
	e.EncodeToken(start)
	for _, c := range row.cells {
		if err := e.Encode(c.cell); err != nil {
			return err
		}
	}
	e.EncodeToken(start.End())
	return nil
}

func (row *Row) resetStyleIndex() {
	for _, cell := range row.cells {
		cell.resetStyleIndex()
	}
}
