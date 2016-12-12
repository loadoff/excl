package excl

import (
	"encoding/xml"
	"regexp"
	"strconv"
)

// Cell はセル一つ一つに対する構造体
type Cell struct {
	cell          *Tag
	colNo         int
	R             string
	sharedStrings *SharedStrings
	style         string
}

// NewCell は新しくcellを作成する
func NewCell(tag *Tag, sharedStrings *SharedStrings) *Cell {
	cell := &Cell{cell: tag, sharedStrings: sharedStrings, colNo: -1}
	r := regexp.MustCompile("^([A-Z]+)[0-9]+$")
	for _, attr := range tag.Attr {
		if attr.Name.Local == "r" {
			strs := r.FindStringSubmatch(attr.Value)
			if len(strs) != 2 {
				return nil
			}
			cell.colNo = int(ColNumPosition(strs[1]))
		} else if attr.Name.Local == "s" {
			cell.style = attr.Value
		}
	}
	if cell.colNo == -1 {
		return nil
	}
	return cell
}

// setValue セルに文字列を追加する
func (cell *Cell) setValue(val string) *Cell {
	tag := &Tag{
		Name: xml.Name{Local: "v"},
		Children: []interface{}{
			xml.CharData(val),
		},
	}
	cell.cell.Children = []interface{}{tag}
	return cell
}

// SetString 文字列を追加する
func (cell *Cell) SetString(val string) *Cell {
	v := cell.sharedStrings.AddString(val)
	cell.setValue(strconv.Itoa(v))
	cell.cell.setAttr("t", "s")
	return cell
}

// SetNumber 数値を追加する
func (cell *Cell) SetNumber(val string) *Cell {
	cell.setValue(val)
	cell.cell.deleteAttr("t")
	return cell
}
