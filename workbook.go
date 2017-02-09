package excl

import (
	"archive/zip"
	"crypto/rand"
	"encoding/binary"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"golang.org/x/text/unicode/norm"
)

// Workbook はワークブック内の情報を格納する
type Workbook struct {
	Path          string
	TempPath      string
	types         *ContentTypes
	opened        bool
	sheets        []*Sheet
	SharedStrings *SharedStrings
	workbookRels  *WorkbookRels
	Styles        *Styles
	outputPath    string
	workbookTag   *Tag
	sheetsTag     *Tag
}

// WorkbookXML workbook.xmlに記載されている<workbook>タグの中身
type WorkbookXML struct {
	XMLName xml.Name  `xml:"workbook"`
	Sheets  sheetsXML `xml:"sheets"`
}

// sheetsXML workbook.xmlに記載されている<sheets>タグの中身
type sheetsXML struct {
	XMLName   xml.Name   `xml:"sheets"`
	Sheetlist []SheetXML `xml:"sheet"`
}

// CreateWorkbook 新しくワークブックを作成する
func CreateWorkbook(expand, outputPath string) (*Workbook, error) {
	if !isDirExist(expand) {
		return nil, errors.New("Directory[" + expand + "] does not exist.")
	}
	dir := filepath.Join(expand, strings.Replace(time.Now().Format("TEMP_20060102030405.000"), ".", "", 1)+random())
	workbook := &Workbook{Path: "", TempPath: dir, outputPath: outputPath}
	defer func() {
		if !workbook.opened {
			workbook.Close()
		}
	}()
	if err := os.Mkdir(workbook.TempPath, 0755); err != nil {
		return nil, err
	}
	if err := createContentTypes(dir); err != nil {
		return nil, err
	}
	if err := createRels(dir); err != nil {
		return nil, err
	}
	if err := createWorkbook(dir); err != nil {
		return nil, err
	}
	if err := createWorkbookRels(dir); err != nil {
		return nil, err
	}
	if err := createStyles(dir); err != nil {
		return nil, err
	}
	if err := createTheme1(dir); err != nil {
		return nil, err
	}
	if err := os.Mkdir(filepath.Join(dir, "xl", "worksheets"), 0755); err != nil {
		return nil, err
	}
	if err := workbook.setInfo(); err != nil {
		return nil, err
	}
	workbook.opened = true
	return workbook, nil
}

// NewWorkbook は新しいワークブック構造体を作成する
// path にはワークブックのパス
// expand には展開するパス
func NewWorkbook(path, expand, outputPath string) (*Workbook, error) {
	if !isFileExist(path) {
		return nil, errors.New("Excel file does not exist.")
	}
	if !isDirExist(expand) {
		return nil, errors.New("Directory[" + expand + "] does not exist.")
	}
	dir := filepath.Join(expand, strings.Replace(time.Now().Format("TEMP_20060102030405.000"), ".", "", 1)+random())
	workbook := &Workbook{Path: path, TempPath: dir, outputPath: outputPath}
	return workbook, nil
}

// Open ワークブックを開く
func (workbook *Workbook) Open() error {
	if ok := isFileExist(workbook.Path); !ok {
		return errors.New("File[" + workbook.Path + "] does not exist.")
	}
	err := os.Mkdir(workbook.TempPath, 0755)
	if err != nil {
		return err
	}
	if err := unzip(workbook.Path, workbook.TempPath); err != nil {
		return err
	}
	defer func() {
		if !workbook.opened {
			workbook.Close()
		}
	}()
	if !isFileExist(filepath.Join(workbook.TempPath, "[Content_Types].xml")) {
		return errors.New("This excel file is corrupt.")
	}
	if !isFileExist(filepath.Join(workbook.TempPath, "xl", "workbook.xml")) {
		return errors.New("This excel file is corrupt.")
	}
	if !isFileExist(filepath.Join(workbook.TempPath, "xl", "_rels", "workbook.xml.rels")) {
		return errors.New("This excel file is corrupt.")
	}
	if !isFileExist(filepath.Join(workbook.TempPath, "xl", "styles.xml")) {
		return errors.New("This excel file is corrupt.")
	}
	err = workbook.setInfo()
	if err != nil {
		return err
	}
	workbook.opened = true
	return nil
}

// Close 操作中のブックを閉じる
func (workbook *Workbook) Close() error {
	if workbook == nil || !workbook.opened {
		return nil
	}
	var err, sheetErr, ssErr, relsErr, stylesErr, typesErr error
	var f *os.File
	defer os.RemoveAll(workbook.TempPath)
	for _, sheet := range workbook.sheets {
		tempErr := sheet.Close()
		if sheetErr == nil && tempErr != nil {
			sheetErr = tempErr
		}
	}
	ssErr = workbook.SharedStrings.Close()
	relsErr = workbook.workbookRels.Close()
	stylesErr = workbook.Styles.Close()
	typesErr = workbook.types.Close()
	workbook.opened = false
	if sheetErr != nil {
		return sheetErr
	} else if ssErr != nil {
		return ssErr
	} else if relsErr != nil {
		return relsErr
	} else if stylesErr != nil {
		return stylesErr
	} else if typesErr != nil {
		return typesErr
	}
	f, err = os.Create(filepath.Join(workbook.TempPath, "xl", "workbook.xml"))
	if err != nil {
		return err
	}
	defer f.Close()
	f.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
	if err = xml.NewEncoder(f).Encode(workbook.workbookTag); err != nil {
		return err
	}
	f.Close()
	createZip(workbook.outputPath, getFiles(workbook.TempPath), workbook.TempPath)
	return nil
}

// OpenSheet 指定されたシートを開く
// 存在しない場合は作成する
func (workbook *Workbook) OpenSheet(name string) (*Sheet, error) {
	compName := strings.ToLower(string(norm.NFKC.Bytes([]byte(name))))
	for _, sheet := range workbook.sheets {
		sheetName := strings.ToLower(string(norm.NFKC.Bytes([]byte(sheet.xml.Name))))
		if sheetName != compName {
			continue
		}
		err := sheet.Open(workbook.TempPath)
		if err != nil {
			return nil, err
		}
		return sheet, nil
	}
	count := workbook.types.sheetCount()
	sheetName := workbook.types.addSheet()
	workbook.workbookRels.addSheet(sheetName)
	workbook.sheetsTag.Children = append(workbook.sheetsTag.Children, createSheetTag(name, count+1))
	sheet := NewSheet(name, count)
	sheet.sharedStrings = workbook.SharedStrings
	sheet.Styles = workbook.Styles
	if err := sheet.Create(workbook.TempPath); err != nil {
		return nil, err
	}
	workbook.sheets = append(workbook.sheets, sheet)
	return sheet, nil
}

func createSheetTag(name string, id int) *Tag {
	tag := &Tag{Name: xml.Name{Local: "sheet"}}
	tag.setAttr("name", name)
	tag.setAttr("sheetId", strconv.Itoa(id))
	tag.setAttr("r:id", fmt.Sprintf("rId%d", id))
	return tag
}

// setInfo xlsx情報を読み込みセットする
func (workbook *Workbook) setInfo() error {
	var err error
	workbook.types, err = OpenContentTypes(workbook.TempPath)
	if err != nil {
		return err
	}
	workbook.Styles, err = OpenStyles(workbook.TempPath)
	if err != nil {
		return err
	}
	workbook.workbookRels, err = OpenWorkbookRels(workbook.TempPath)
	if err != nil {
		return err
	}
	workbook.SharedStrings, err = OpenSharedStrings(workbook.TempPath)
	if err != nil {
		return err
	}
	workbook.workbookRels.addSharedStrings()
	workbook.types.addSharedString()
	err = workbook.openWorkbook()
	if err != nil {
		return err
	}

	return nil
}

// createWorkbook workbook.xmlファイルを作成する
func createWorkbook(dir string) error {
	os.Mkdir(filepath.Join(dir, "xl"), 0755)
	path := filepath.Join(dir, "xl", "workbook.xml")
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	f.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
	f.WriteString(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">`)
	f.WriteString(`<sheets/>`)
	f.WriteString(`</workbook>`)
	f.Close()

	return nil
}

// createRels .relsファイルを作成する
func createRels(dir string) error {
	os.Mkdir(filepath.Join(dir, "_rels"), 0755)
	path := filepath.Join(dir, "_rels", ".rels")
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	f.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
	f.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	f.WriteString(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>`)
	f.WriteString(`</Relationships>`)
	f.Close()
	return nil
}

// openWorkbook xmlファイルを開きシート情報を取得する
func (workbook *Workbook) openWorkbook() error {
	workbookPath := filepath.Join(workbook.TempPath, "xl", "workbook.xml")
	f, err := os.Open(workbookPath)
	if err != nil {
		return err
	}
	defer f.Close()
	data, _ := ioutil.ReadAll(f)
	f.Close()
	val := WorkbookXML{}
	err = xml.Unmarshal(data, &val)
	if err != nil {
		return err
	}
	for i := range val.Sheets.Sheetlist {
		sheet := &val.Sheets.Sheetlist[i]
		workbook.sheets = append(workbook.sheets,
			&Sheet{
				xml:           sheet,
				Styles:        workbook.Styles,
				sharedStrings: workbook.SharedStrings,
			})
	}
	tag := &Tag{}
	f, _ = os.Open(workbookPath)
	defer f.Close()
	if err = xml.NewDecoder(f).Decode(tag); err != nil {
		return err
	}
	f.Close()
	workbook.workbookTag = tag
	workbook.setSheetsTag()
	return nil
}

func (workbook *Workbook) setSheetsTag() {
	for _, child := range workbook.workbookTag.Children {
		switch t := child.(type) {
		case *Tag:
			if t.Name.Local == "sheets" {
				workbook.sheetsTag = t
				return
			}
		}
	}
}

// getFiles dir以下に存在するファイルの一覧を取得する
func getFiles(dir string) []string {
	fileList := []string{}
	files, _ := ioutil.ReadDir(dir)
	for _, f := range files {
		if f.IsDir() == true {
			list := getFiles(path.Join(dir, f.Name()))
			fileList = append(fileList, list...)
			continue
		}
		fileList = append(fileList, path.Join(dir, f.Name()))
	}
	return fileList
}

// unzip はzipを解凍する
func unzip(src, dest string) error {
	r, err := zip.OpenReader(src)
	if err != nil {
		return err
	}
	defer r.Close()
	os.MkdirAll(dest, 0755)

	for _, f := range r.File {
		rc, err := f.Open()
		if err != nil {
			return err
		}
		defer rc.Close()

		path := filepath.Join(dest, f.Name)
		if f.FileInfo().IsDir() {
			os.MkdirAll(path, 0755)
		} else {
			os.MkdirAll(filepath.Dir(path), 0755)
			to, err := os.OpenFile(
				path, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0755)
			if err != nil {
				return err
			}
			defer to.Close()
			_, err = io.Copy(to, rc)
			if err != nil {
				return err
			}
			to.Close()
		}
	}
	return nil
}

func createZip(zipPath string, fileList []string, replace string) {
	var zipfile *os.File
	var err error
	if zipfile, err = os.Create(zipPath); err != nil {
		log.Fatalln(err)
	}
	defer zipfile.Close()
	w := zip.NewWriter(zipfile)
	for _, file := range fileList {
		body, err := ioutil.ReadFile(file)
		if err != nil {
			fmt.Println(err.Error())
			continue
		}
		f, err := w.Create(strings.Replace(file, replace+"/", "", 1))
		if err != nil {
			fmt.Println(err.Error())
			continue
		}
		_, err = f.Write(body)
		if err != nil {
			fmt.Println(err.Error())
		}
	}
	w.Close()
}

// isFileExist ファイルの存在確認
func isFileExist(filename string) bool {
	stat, err := os.Stat(filename)
	if err != nil {
		return false
	}
	return !stat.IsDir()
}

// isDirExist ディレクトリの存在確認
func isDirExist(dirname string) bool {
	stat, err := os.Stat(dirname)
	if err != nil {
		return false
	}
	return stat.IsDir()
}

func random() string {
	var n uint64
	binary.Read(rand.Reader, binary.LittleEndian, &n)
	return strconv.FormatUint(n, 36)
}
