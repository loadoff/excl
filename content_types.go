package excl

import (
	"encoding/xml"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
)

type ContentTypes struct {
	path  string
	types *ContentTypesXML
}

// ContentTypesXML [Content_Types].xmlファイルを読み込む
type ContentTypesXML struct {
	XMLName   xml.Name          `xml:"Types"`
	Xmlns     string            `xml:"xmlns,attr"`
	Defaults  []contentDefault  `xml:"Default"`
	Overrides []contentOverride `xml:"Override"`
}

type contentOverride struct {
	XMLName     xml.Name `xml:"Override"`
	PartName    string   `xml:"PartName,attr"`
	ContentType string   `xml:"ContentType,attr"`
}

type contentDefault struct {
	XMLName     xml.Name `xml:"Default"`
	Extension   string   `xml:"Extension,attr"`
	ContentType string   `xml:"ContentType,attr"`
}

// OpenContentTypes [Content_Types].xmlファイルを開き構造体に読み込む
func OpenContentTypes(dir string) (*ContentTypes, error) {
	path := filepath.Join(dir, "[Content_Types].xml")
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	data, err := ioutil.ReadAll(f)
	if err != nil {
		return nil, err
	}
	v := ContentTypesXML{}
	err = xml.Unmarshal(data, &v)
	if err != nil {
		return nil, err
	}
	f.Close()
	types := &ContentTypes{path, &v}
	return types, nil
}

// Close [Content_Types].xmlファイルを閉じる
func (types *ContentTypes) Close() error {
	f, err := os.Create(types.path)
	if err != nil {
		return err
	}
	defer f.Close()
	d, err := xml.Marshal(types.types)
	if err != nil {
		return err
	}
	f.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
	f.Write(d)
	return nil
}

// sheetCount シートの数を返す
func (types *ContentTypes) sheetCount() int {
	var count int
	for _, override := range types.types.Overrides {
		if override.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" {
			count++
		}
	}
	return count
}

// addSheet シートを追加する
func (types *ContentTypes) addSheet() string {
	count := types.sheetCount()
	name := fmt.Sprintf("sheet%d.xml", count+1)

	override := contentOverride{
		XMLName:     xml.Name{Space: "", Local: "Override"},
		PartName:    "/xl/worksheets/" + name,
		ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"}
	types.types.Overrides = append(types.types.Overrides, override)
	return name
}

// hasSharedString sharedString.xmlファイルが存在するか確認する
func (types *ContentTypes) hasSharedString() bool {
	for _, override := range types.types.Overrides {
		if override.PartName == "/xl/sharedStrings.xml" {
			return true
		}
	}
	return false
}

// addSharedString sharedString.xmlファイルを追加する
func (types *ContentTypes) addSharedString() {
	if types.hasSharedString() {
		return
	}
	override := contentOverride{
		XMLName:     xml.Name{Space: "", Local: "Override"},
		PartName:    "/xl/sharedStrings.xml",
		ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"}
	types.types.Overrides = append(types.types.Overrides, override)
}
