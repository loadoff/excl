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
	styleIndex    int
	styles        *Styles
	style         *Style
	changed       bool
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
			cell.styleIndex, _ = strconv.Atoi(attr.Value)
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

// GetStyle Style構造体を取得する
func (cell *Cell) GetStyle() *Style {
	if cell.style == nil {
		style := cell.styles.GetStyle(cell.styleIndex)
		if style == nil {
			style = &Style{}
		}
		cell.style = &Style{
			NumFmtID:   style.NumFmtID,
			FontID:     style.FontID,
			FillID:     style.FillID,
			BorderID:   style.BorderID,
			XfID:       style.XfID,
			Horizontal: style.Horizontal,
			Vertical:   style.Vertical,
		}
	}
	return cell.style
}

// SetNumFmt 数値フォーマット
func (cell *Cell) SetNumFmt(fmt string) {
	if cell.style == nil {
		cell.GetStyle()
	}
	cell.style.NumFmtID = cell.styles.SetNumFmt(fmt)
	cell.changed = true
}

// SetFont フォント情報をセットする
func (cell *Cell) SetFont(font Font) {
	if cell.style == nil {
		cell.GetStyle()
	}
	cell.style.FontID = cell.styles.SetFont(font)
	cell.changed = true
}

// SetBackgroundColor 背景色をセットする
func (cell *Cell) SetBackgroundColor(color string) {
	if cell.style == nil {
		cell.GetStyle()
	}
	cell.style.FillID = cell.styles.SetBackgroundColor(color)
	cell.changed = true
}

// SetBorder 罫線情報をセットする
func (cell *Cell) SetBorder(border Border) {
	if cell.style == nil {
		cell.GetStyle()
	}
	cell.style.BorderID = cell.styles.SetBorder(border)
	cell.changed = true
}

func (cell *Cell) resetStyleIndex() {
	if cell.changed {
		index := cell.styles.SetStyle(cell.style)
		cell.cell.setAttr("s", strconv.Itoa(index))
	}
}
