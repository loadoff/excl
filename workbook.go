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
	styles        *Styles
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

// NewWorkbook は新しいワークブック構造体を作成する
// path にはワークブックのパス
// tempPath には展開するパス
func NewWorkbook(path, tempPath, outputPath string) (*Workbook, error) {
	if !isFileExist(path) {
		return nil, errors.New("Excel file does not exist.")
	}
	if !isDirExist(tempPath) {
		return nil, errors.New("Directory[" + tempPath + "] does not exist.")
	}
	dir := filepath.Join(tempPath, strings.Replace(time.Now().Format("TEMP_20060102030405.000"), ".", "", 1)+random())
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
	var err, tempErr error
	defer os.RemoveAll(workbook.TempPath)
	if workbook.SharedStrings != nil {
		workbook.SharedStrings.Close()
		workbook.SharedStrings = nil
	}
	if workbook.workbookRels != nil {
		tempErr = workbook.workbookRels.Close()
		workbook.workbookRels = nil
	}
	err = tempErr
	if workbook.styles != nil {
		tempErr = workbook.styles.Close()
		workbook.styles = nil
	}
	if err == nil && tempErr != nil {
		err = tempErr
	}
	if workbook.types != nil {
		tempErr = workbook.types.Close()
		workbook.types = nil
	}
	if err == nil && tempErr != nil {
		err = tempErr
	}

	for _, sheet := range workbook.sheets {
		if sheet.opened {
			tempErr = sheet.Close()
			if err == nil && tempErr != nil {
				err = tempErr
			}
		}
	}

	if err == nil {
		f, err := os.Create(filepath.Join(workbook.TempPath, "xl", "workbook.xml"))
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
	}
	workbook.opened = false
	return err
}

// OpenSheet 指定されたシートを開く
// 存在しない場合は作成する
func (workbook *Workbook) OpenSheet(name string) (*Sheet, error) {
	for _, sheet := range workbook.sheets {
		if sheet.xml.Name != name {
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
	sheet := NewSheet(sheetName, count)
	sheet.sharedStrings = workbook.SharedStrings
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
	workbook.styles, err = OpenStyles(workbook.TempPath)
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

// openXML xmlファイルを開きシート情報を取得する
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
	for _, sheet := range val.Sheets.Sheetlist {
		workbook.sheets = append(workbook.sheets,
			&Sheet{
				xml:           &sheet,
				Styles:        workbook.styles,
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
			os.MkdirAll(path, f.Mode())
		} else {
			os.MkdirAll(filepath.Dir(path), f.Mode())
			to, err := os.OpenFile(
				path, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, f.Mode())
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
