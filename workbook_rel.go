package excl

import (
	"encoding/xml"
	"errors"
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"
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

// createWorkbookRels workbook.xml.relsファイルを作成する
func createWorkbookRels(dir string) error {
	os.Mkdir(filepath.Join(dir, "xl", "_rels"), 0755)
	path := filepath.Join(dir, "xl", "_rels", "workbook.xml.rels")
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	f.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
	f.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	f.WriteString(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`)
	f.WriteString(`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`)
	f.WriteString(`</Relationships>`)
	f.Close()
	return nil
}

// OpenWorkbookRels open workbook.xml.rels
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

// Close close workbook.xml.rels
func (wbr *WorkbookRels) Close() error {
	if wbr == nil {
		return nil
	}
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

// addSharedStrings add sharedStrings.xml information
func (wbr *WorkbookRels) addSharedStrings() string {
	for _, rel := range wbr.rels.Rels {
		if rel.Target == "sharedStrings.xml" {
			return rel.ID
		}
	}
	rel := relationship{
		XMLName: xml.Name{Local: "Relationship"},
		Target:  "sharedStrings.xml",
		Type:    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
		ID:      strings.Replace(time.Now().Format("rId060102030405.000"), ".", "", 1) + random(),
	}
	wbr.rels.Rels = append(wbr.rels.Rels, rel)
	return rel.ID
}

// addSheet add sheet information to workbook.xml.rels
func (wbr *WorkbookRels) addSheet(name string) string {
	rel := relationship{
		XMLName: xml.Name{Local: "Relationship"},
		Target:  "worksheets/" + name,
		Type:    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
		ID:      strings.Replace(time.Now().Format("rId060102030405.000"), ".", "", 1) + random(),
	}
	wbr.rels.Rels = append(wbr.rels.Rels, rel)
	return rel.ID
}

func (wbr *WorkbookRels) getTarget(rid string) string {
	if wbr == nil {
		return ""
	}
	for _, rel := range wbr.rels.Rels {
		if rel.ID == rid {
			return rel.Target
		}
	}
	return ""
}

func (wbr *WorkbookRels) getSheetIndex(rid string) int {
	if wbr == nil {
		return -1
	}
	re := regexp.MustCompile(`\Aworksheets\/sheet([0-9]+)\.xml\z`)
	for _, rel := range wbr.rels.Rels {
		if rel.ID == rid {
			vals := re.FindStringSubmatch(rel.Target)
			if len(vals) != 2 {
				return -1
			}
			index, _ := strconv.Atoi(vals[1])
			return index
		}
	}
	return -1
}

func (wbr *WorkbookRels) getSheetMaxIndex() int {
	maxIndex := 0
	re := regexp.MustCompile(`\Aworksheets\/sheet([0-9]+)\.xml\z`)
	for _, rel := range wbr.rels.Rels {
		val := re.FindStringSubmatch(rel.Target)
		if len(val) == 2 {
			index, _ := strconv.Atoi(val[1])
			if index > maxIndex {
				maxIndex = index
			}
		}
	}
	return maxIndex
}
