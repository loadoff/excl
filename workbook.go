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
	TempPath      string
	types         *ContentTypes
	opened        bool
	maxSheetID    int
	sheets        []*Sheet
	SharedStrings *SharedStrings
	workbookRels  *WorkbookRels
	Styles        *Styles
	workbookTag   *Tag
	sheetsTag     *Tag
	calcPr        *Tag
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

// Create 新しくワークブックを作成する
func Create() (*Workbook, error) {
	dir, err := ioutil.TempDir("", "excl_"+strings.Replace(time.Now().Format("20060102030405.000"), ".", "", 1))
	if err != nil {
		return nil, err
	}
	workbook := &Workbook{TempPath: dir}
	defer func() {
		if !workbook.opened {
			workbook.Close()
		}
	}()
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

// Open Excelファイルを開く
func Open(path string) (*Workbook, error) {
	if !isFileExist(path) {
		return nil, errors.New("Excel file does not exist.")
	}
	dir, err := ioutil.TempDir("", "excl"+strings.Replace(time.Now().Format("20060102030405"), ".", "", 1))
	if err != nil {
		return nil, err
	}
	workbook := &Workbook{TempPath: dir}
	if err := unzip(path, dir); err != nil {
		return nil, err
	}

	defer func() {
		if !workbook.opened {
			workbook.Close()
		}
	}()

	if !isFileExist(filepath.Join(dir, "[Content_Types].xml")) {
		return nil, errors.New("this excel file is corrupt.")
	}
	if !isFileExist(filepath.Join(dir, "xl", "workbook.xml")) {
		return nil, errors.New("This excel file is corrupt.")
	}
	if !isFileExist(filepath.Join(dir, "xl", "_rels", "workbook.xml.rels")) {
		return nil, errors.New("This excel file is corrupt.")
	}
	if !isFileExist(filepath.Join(dir, "xl", "styles.xml")) {
		return nil, errors.New("This excel file is corrupt.")
	}
	err = workbook.setInfo()
	if err != nil {
		return nil, err
	}
	workbook.opened = true
	return workbook, nil
}

// Save 操作中のブックを保存し閉じる
func (workbook *Workbook) Save(path string) error {
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
	if path != "" {
		createZip(path, getFiles(workbook.TempPath), workbook.TempPath)
	}
	return nil
}

// Close 操作中のブックを閉じる(保存はしない)
func (workbook *Workbook) Close() error {
	return workbook.Save("")
}

// OpenSheet Open specified sheet
// if there is no specified sheet then create new sheet
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
	count := len(workbook.sheets)
	sheetName := workbook.types.addSheet(count)
	workbook.workbookRels.addSheet(sheetName)
	workbook.sheetsTag.Children = append(workbook.sheetsTag.Children, createSheetTag(name, count+1, workbook.maxSheetID+1))
	sheet := NewSheet(name, count, workbook.maxSheetID)
	sheet.sharedStrings = workbook.SharedStrings
	sheet.Styles = workbook.Styles
	if err := sheet.Create(workbook.TempPath); err != nil {
		return nil, err
	}
	workbook.sheets = append(workbook.sheets, sheet)
	workbook.maxSheetID++
	return sheet, nil
}

// SetForceFormulaRecalculation set fullCalcOnLoad attribute to calcPr tag.
// When this excel file is opened, all calculation fomula will be recalculated.
func (workbook *Workbook) SetForceFormulaRecalculation(flg bool) {
	if workbook.calcPr != nil {
		if flg {
			workbook.calcPr.setAttr("fullCalcOnLoad", "1")
		} else {
			workbook.calcPr.deleteAttr("fullCalcOnLoad")
		}
	}
}

// RenameSheet rename sheet name from old name to new name.
func (workbook *Workbook) RenameSheet(old string, new string) {
	for i, sheet := range workbook.sheets {
		if sheet.xml.Name != old {
			continue
		}
		sheet.xml.Name = new
		switch t := workbook.sheetsTag.Children[i].(type) {
		case *Tag:
			t.setAttr("name", new)
		}
	}
}

// HideSheet hide sheet
func (workbook *Workbook) HideSheet(name string) {
	for i, sheet := range workbook.sheets {
		if sheet.xml.Name != name {
			continue
		}
		switch t := workbook.sheetsTag.Children[i].(type) {
		case *Tag:
			t.setAttr("state", "hidden")
		}
		break
	}
}

// ShowSheet show sheet
func (workbook *Workbook) ShowSheet(name string) {
	for i, sheet := range workbook.sheets {
		if sheet.xml.Name != name {
			continue
		}
		switch t := workbook.sheetsTag.Children[i].(type) {
		case *Tag:
			t.deleteAttr("state")
		}
		break
	}
}

func createSheetTag(name string, id int, sheetID int) *Tag {
	tag := &Tag{Name: xml.Name{Local: "sheet"}}
	tag.setAttr("name", name)
	tag.setAttr("sheetId", strconv.Itoa(sheetID))
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

// openWorkbook open workbook.xml and set workbook information
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
		index, _ := strconv.Atoi(strings.Replace(sheet.RID, "rId", "", 1))
		workbook.sheets = append(workbook.sheets,
			&Sheet{
				xml:           sheet,
				Styles:        workbook.Styles,
				sharedStrings: workbook.SharedStrings,
				sheetIndex:    index,
			})
		sheetID, _ := strconv.Atoi(sheet.SheetID)
		if workbook.maxSheetID < sheetID {
			workbook.maxSheetID = sheetID
		}
	}
	tag := &Tag{}
	f, _ = os.Open(workbookPath)
	defer f.Close()
	if err = xml.NewDecoder(f).Decode(tag); err != nil {
		return err
	}
	f.Close()
	workbook.workbookTag = tag
	workbook.setWorkbookInfo()
	return nil
}

func (workbook *Workbook) setWorkbookInfo() {
	for _, child := range workbook.workbookTag.Children {
		switch t := child.(type) {
		case *Tag:
			if t.Name.Local == "sheets" {
				workbook.sheetsTag = t
			} else if t.Name.Local == "calcPr" {
				workbook.calcPr = t
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

// unzip unzip excel file
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
		read, err := os.Open(file)
		defer read.Close()
		if err != nil {
			fmt.Println(err)
			continue
		}
		f, err := w.Create(strings.Replace(file, replace+"/", "", 1))
		if err != nil {
			fmt.Println(err)
			continue
		}
		if _, err = io.Copy(f, read); err != nil {
			fmt.Println(err)
			continue
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
