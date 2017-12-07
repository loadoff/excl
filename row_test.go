package excl

import (
	"bytes"
	"encoding/xml"
	"os"
	"testing"
	"time"
)

func TestNewRow(t *testing.T) {
	tag := &Tag{}
	tag.setAttr("r", "10")
	tag.setAttr("s", "5")
	row := NewRow(tag, nil, nil)
	if row.rowID != 10 {
		t.Error("row no should be 10 but", row.rowID)
	}
	if row.style != "5" {
		t.Error("row style index should be 5 but", row.style)
	}

	cellTag := &Tag{Name: xml.Name{Local: "c"}}
	cellTag.setAttr("r", "J10")
	tag.Children = append(tag.Children, cellTag)
	row = NewRow(tag, nil, nil)
	if row == nil {
		t.Error("row should not be nil.")
	}
	if row.maxColNo != 10 {
		t.Error("row maxColNo should be 10 but", row.maxColNo)
	}
	if row.minColNo != 10 {
		t.Error("row minColNo should be 10 but", row.minColNo)
	}

	tag.Children = append(tag.Children, &Tag{Name: xml.Name{Local: "c"}})
	row = NewRow(tag, nil, nil)
	if row != nil {
		t.Error("row should be nil.")
	}
}

func TestCreateCells(t *testing.T) {
	tag := &Tag{}
	tag.setAttr("r", "10")
	row := NewRow(tag, nil, nil)
	cells := row.CreateCells(2, 3)
	if len(cells) != 3 {
		t.Error("3 cells should be created")
	}
	colInfo := colInfo{min: 4, max: 5, style: "2"}
	row.colInfos = append(row.colInfos, colInfo)
	cells = row.CreateCells(3, 10)
	if len(cells) != 10 {
		t.Error("10 cells should be created")
	}
	if cells[3].styleIndex != 2 || cells[4].styleIndex != 2 || cells[5].styleIndex != 0 {
		t.Error("cell style index should be set.", cells[3].styleIndex, cells[4].styleIndex, cells[5].styleIndex)
	}
	row.style = "5"
	cells = row.CreateCells(12, 13)
	if cells[3].styleIndex != 2 || cells[4].styleIndex != 2 || cells[5].styleIndex != 0 || cells[11].styleIndex != 5 || cells[12].styleIndex != 5 {
		t.Error("row style should be set")
	}
}

func TestGetCell(t *testing.T) {
	row := &Row{row: &Tag{}}
	tag := &Tag{}
	row.row.Children = append(row.row.Children, tag)
	cell := &Cell{cell: tag, colNo: 3}
	row.cells = append(row.cells, cell)
	c := row.GetCell(3)
	if c != cell {
		t.Error("cell should be get.")
	}
	c = row.GetCell(4)
	if c.colNo != 4 {
		t.Error("colNo should be 4 but", c.colNo)
	}
	row.colInfos = append(row.colInfos, colInfo{min: 1, max: 1, style: "3"})
	c = row.GetCell(1)
	if c.colNo != 1 {
		t.Error("colNo should be 1 but", c.colNo)
	}

	if c.styleIndex != 3 {
		t.Error("cell style should be 3 but", c.style)
	}

	row.style = "4"
	c = row.GetCell(2)
	if c.colNo != 2 {
		t.Error("colNo should be 2 but", c.colNo)
	}
	if c.styleIndex != 4 {
		t.Error("cell style should be 4 but", c.style)
	}
}

func TestSetRowString(t *testing.T) {
	f, _ := os.Create("temp/test.xml")
	defer func() {
		f.Close()
		os.Remove("temp/test.xml")
	}()
	ss := &SharedStrings{tempFile: f, buffer: &bytes.Buffer{}}
	tag := &Tag{}
	tag.setAttr("r", "10")
	row := NewRow(tag, ss, nil)
	c := row.SetString("hello world", 10)
	if val, _ := c.cell.getAttr("t"); val != "s" {
		t.Error("cell attribute should be s but", val)
	}

	tag = c.cell.Children[0].(*Tag)
	if tag.Name.Local != "v" {
		t.Error("cell tag should be v but", c.cell.Children[0].(*Tag).Name.Local)
	}
	if string(tag.Children[0].(xml.CharData)) != "0" {
		t.Error("cell value should be 0 but", string(tag.Children[0].(xml.CharData)))
	}
}

func TestSetRowNumber(t *testing.T) {
	tag := &Tag{}
	tag.setAttr("r", "10")
	row := NewRow(tag, nil, nil)
	c := row.SetNumber("20", 10)
	if val, _ := c.cell.getAttr("t"); val == "s" {
		t.Error("cell attribute should not be s.")
	}

	tag = c.cell.Children[0].(*Tag)
	if tag.Name.Local != "v" {
		t.Error("cell tag should be v but", c.cell.Children[0].(*Tag).Name.Local)
	}
	if string(tag.Children[0].(xml.CharData)) != "20" {
		t.Error("cell value should be 20 but", string(tag.Children[0].(xml.CharData)))
	}

	c = row.SetNumber(20.1, 11)
	tag = c.cell.Children[0].(*Tag)
	if string(tag.Children[0].(xml.CharData)) != "20.1" {
		t.Error("cell value should be 20 but", string(tag.Children[0].(xml.CharData)))
	}
}

func TestSetRowDate(t *testing.T) {
	tag := &Tag{}
	tag.setAttr("r", "10")
	row := NewRow(tag, nil, &Styles{})
	now := time.Now()
	c := row.SetDate(now, 10)
	if val, _ := c.cell.getAttr("t"); val != "d" {
		t.Error("cell attribute should be d but", val)
	}
	tag = c.cell.Children[0].(*Tag)
	if string(tag.Children[0].(xml.CharData)) != now.Format("2006-01-02T15:04:05.999999999") {
		t.Error("cell value should be", now.Format("2006-01-02T15:04:05.999999999"), "but", string(tag.Children[0].(xml.CharData)))
	}
}

func TestSetRowFormula(t *testing.T) {
	tag := &Tag{}
	tag.setAttr("r", "10")
	row := NewRow(tag, nil, &Styles{})
	c := row.SetFormula("SUM(A1:A2)", 10)
	if val, _ := c.cell.getAttr("t"); val != "" {
		t.Error("cell attribute should be empty but", val)
	}
	tag = c.cell.Children[0].(*Tag)
	if tag.Name.Local != "f" {
		t.Error("tag name should be f but", tag.Name.Local)
	}
	if string(tag.Children[0].(xml.CharData)) != "SUM(A1:A2)" {
		t.Error("cell value should be SUM(A1:A2) but", string(tag.Children[0].(xml.CharData)))
	}
}
func TestColStringPosition(t *testing.T) {
	if ColStringPosition(26) != "Z" {
		t.Error("col id should be Z but", ColStringPosition(26))
	}
	if ColStringPosition(27) != "AA" {
		t.Error("col id should be AA but", ColStringPosition(27))
	}
	if ColStringPosition(52) != "AZ" {
		t.Error("col id should be AZ but", ColStringPosition(52))
	}
}

func TestColNumPosition(t *testing.T) {
	if ColNumPosition("Z") != 26 {
		t.Error("col no should be 26 but", ColNumPosition("Z"))
	}
	if ColNumPosition("AA") != 27 {
		t.Error("col no should be 27 but", ColNumPosition("AA"))
	}
}

func TestCreateRowXML(t *testing.T) {
	stdout := new(bytes.Buffer)
	tag := &Tag{Name: xml.Name{Local: "row"}}
	row := &Row{row: tag}
	xml.NewEncoder(stdout).Encode(row)
	if string(stdout.Bytes()) != "<row></row>" {
		t.Error("xml should be empty but", string(stdout.Bytes()))
	}

	cellTag := &Tag{Name: xml.Name{Local: "c"}}
	row.cells = append(row.cells, &Cell{cell: cellTag})
	xml.NewEncoder(stdout).Encode(row)
	if string(stdout.Bytes()) != "<row></row><row><c></c></row>" {
		t.Error(`xml should be "<row></row><row><c></c></row>" but`, string(stdout.Bytes()))
	}
}

func TestRowResetStyleIndex(t *testing.T) {
	tag := &Tag{Name: xml.Name{Local: "row"}}
	row := &Row{row: tag}
	cellTag := &Tag{Name: xml.Name{Local: "c"}}
	row.cells = append(row.cells, &Cell{cell: cellTag})
	row.resetStyleIndex()
}

func TestSetRowHeight(t *testing.T) {
	tag := &Tag{Name: xml.Name{Local: "row"}}
	row := &Row{row: tag}
	row.SetHeight(1.23)
	if val, err := row.row.getAttr("customHeight"); err != nil {
		t.Error(`tag's customHeight attribute should be set but`, err)
	} else if val != "1" {
		t.Error(`tag's customHeight attribute should be 1 but`, val)
	}
	if val, err := row.row.getAttr("ht"); err != nil {
		t.Error(`tag's ht attribute should be set but`, err)
	} else if val != "1.2300" {
		t.Error(`tag's ht attribute should be 1.2300 but`, val)
	}
}

func BenchmarkNewRow(b *testing.B) {
	row := &Row{rowID: 10}
	row.CreateCells(1, b.N)
}
