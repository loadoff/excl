package excl

import (
	"encoding/xml"
	"errors"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
)

// WorkbookRels workbook.xml.relの情報をもつ構造体
type WorkbookRels struct {
	rels *Relationships
	path string
}

// Relationships Relationshipsタグの情報
type Relationships struct {
	XMLName xml.Name       `xml:"Relationships"`
	Xmlns   string         `xml:"xmlns,attr"`
	Rels    []relationship `xml:"Relationship"`
}

type relationship struct {
	XMLName xml.Name `xml:"Relationship"`
	ID      string   `xml:"Id,attr"`
	Type    string   `xml:"Type,attr"`
	Target  string   `xml:"Target,attr"`
}

// OpenWorkbookRels workbook.xml.relsファイルを開く
func OpenWorkbookRels(dir string) (*WorkbookRels, error) {
	path := filepath.Join(dir, "xl", "_rels", "workbook.xml.rels")
	if !isFileExist(path) {
		return nil, errors.New("The workbook.xml.rels is not exists.")
	}
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	data, err := ioutil.ReadAll(f)
	if err != nil {
		return nil, err
	}
	rels := &Relationships{}
	err = xml.Unmarshal(data, rels)
	if err != nil {
		return nil, err
	}
	return &WorkbookRels{rels: rels, path: path}, nil
}

// Close workbook.xml.relsファイルを閉じる
func (wbr *WorkbookRels) Close() error {
	wbr.setRID()
	f, err := os.Create(wbr.path)
	if err != nil {
		return err
	}
	defer f.Close()
	data, err := xml.Marshal(wbr.rels)
	if err != nil {
		return err
	}
	f.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
	f.Write(data)
	return nil
}

// addSharedStrings sharedStrings.xmlファイルの関連情報を追加する
func (wbr *WorkbookRels) addSharedStrings() {
	for _, rel := range wbr.rels.Rels {
		if rel.Target == "sharedStrings.xml" {
			return
		}
	}
	rel := relationship{
		XMLName: xml.Name{Local: "Relationship"},
		Target:  "sharedStrings.xml",
		Type:    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
	}
	wbr.rels.Rels = append(wbr.rels.Rels, rel)
}

// addSheet workbook.xml.relsにシート情報を追加する
func (wbr *WorkbookRels) addSheet(name string) {
	rel := relationship{
		XMLName: xml.Name{Local: "Relationship"},
		Target:  "worksheets/" + name,
		Type:    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
	}
	wbr.rels.Rels = append(wbr.rels.Rels, rel)
}

func (wbr *WorkbookRels) setRID() {
	rep := regexp.MustCompile(`worksheets\/sheet([0-9]+)\.xml`)
	id := 1
	for index, rel := range wbr.rels.Rels {
		if !rep.MatchString(rel.Target) {
			continue
		}
		i, _ := strconv.Atoi(rep.ReplaceAllString(rel.Target, "$1"))
		wbr.rels.Rels[index].ID = fmt.Sprintf("rId%d", i)
		if id < i {
			id = i
		}
	}
	for index, rel := range wbr.rels.Rels {
		if rep.MatchString(rel.Target) {
			continue
		}
		id++
		wbr.rels.Rels[index].ID = fmt.Sprintf("rId%d", id)
	}
}
