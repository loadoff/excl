package excl

import (
	"bytes"
	"encoding/xml"
	"errors"
	"io"
	"os"
	"path/filepath"
	"strings"
)

// SharedStrings 構造体
type SharedStrings struct {
	file        *os.File
	tempFile    *os.File
	count       int
	dir         string
	afterString string
	buffer      *bytes.Buffer
}

// OpenSharedStrings 新しいSharedString構造体を作成する
func OpenSharedStrings(dir string) (*SharedStrings, error) {
	var f *os.File
	var err error
	path := filepath.Join(dir, "xl", "sharedStrings.xml")
	if !isFileExist(path) {
		f, err = os.Create(path)
		if err != nil {
			return nil, err
		}
		f.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
		f.WriteString(`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></sst>`)
		f.Seek(0, os.SEEK_SET)
	} else {
		f, err = os.Open(path)
		if err != nil {
			return nil, err
		}
	}
	defer f.Close()
	tag := &Tag{}
	err = xml.NewDecoder(f).Decode(tag)
	if err != nil {
		return nil, err
	}
	f.Close()
	ss := &SharedStrings{dir: filepath.Join(dir, "xl"), buffer: &bytes.Buffer{}}
	ss.setStringCount(tag)
	if ss.count == -1 {
		return nil, errors.New("The sharedStrings.xml file is currupt.")
	}
	ss.setSeparatePoint(tag)
	var b bytes.Buffer
	xml.NewEncoder(&b).Encode(tag)
	strs := strings.Split(b.String(), "<separate_tag></separate_tag>")
	if len(strs) != 2 {
		return nil, errors.New("The sharedStrings.xml file is currupt.")
	}
	ss.file, err = os.Create(path)
	if err != nil {
		return nil, err
	}
	ss.tempFile, err = os.Create(filepath.Join(dir, "xl", "__sharedStrings.xml"))
	if err != nil {
		ss.file.Close()
		return nil, err
	}

	ss.file.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
	ss.file.WriteString(strs[0])
	ss.afterString = strs[1]
	return ss, nil
}

// Close sharedStrings情報をクローズする
func (ss *SharedStrings) Close() error {
	if ss == nil {
		return nil
	}
	defer ss.tempFile.Close()
	defer ss.file.Close()
	var err error
	io.Copy(ss.tempFile, ss.buffer)
	ss.tempFile.Seek(0, os.SEEK_SET)
	if _, err = io.Copy(ss.file, ss.tempFile); err != nil {
		return err
	}
	if _, err = ss.file.WriteString(ss.afterString); err != nil {
		return err
	}
	if err = ss.tempFile.Close(); err != nil {
		return err
	}
	if err = ss.file.Close(); err != nil {
		return err
	}
	os.Remove(filepath.Join(ss.dir, "__sharedStrings.xml"))
	ss.tempFile = nil
	ss.file = nil
	return nil
}

var (
	escQuot = []byte("&#34;") // shorter than "&quot;"
	escApos = []byte("&#39;") // shorter than "&apos;"
	escAmp  = []byte("&amp;")
	escLt   = []byte("&lt;")
	escGt   = []byte("&gt;")
	escTab  = []byte("&#x9;")
	escNl   = []byte("&#xA;")
	escCr   = []byte("&#xD;")
)

func escapeText(w io.Writer, s []byte) error {
	var esc []byte
	output := make([]byte, len(s)*5)
	j := 0
	for i := 0; i < len(s); i++ {
		switch s[i] {
		case '"':
			esc = escQuot
		case '\'':
			esc = escApos
		case '&':
			esc = escAmp
		case '<':
			esc = escLt
		case '>':
			esc = escGt
		case '\t':
			esc = escTab
		case '\n':
			esc = escNl
		case '\r':
			esc = escCr
		default:
			output[j] = s[i]
			j++
			continue
		}
		for _, e := range esc {
			output[j] = e
			j++
		}
	}
	if _, err := w.Write(output[:j]); err != nil {
		return err
	}
	return nil
}

// AddString 文字列データを追加する
// 戻り値はインデックス情報(0スタート)
func (ss *SharedStrings) AddString(text string) int {
	if len(text) != 0 && (text[0] == ' ' || text[len(text)-1] == ' ') {
		ss.buffer.WriteString(`<si><t xml:space="preserve">`)
	} else {
		ss.buffer.WriteString("<si><t>")
	}
	escapeText(ss.buffer, []byte(text))
	ss.buffer.WriteString("</t></si>")
	if ss.buffer.Len() > 1024 {
		io.Copy(ss.tempFile, ss.buffer)
		ss.buffer = &bytes.Buffer{}
	}
	ss.count++
	return ss.count - 1
}

// setStringCount 文字列のカウントをセットする
func (ss *SharedStrings) setStringCount(tag *Tag) {
	if tag.Name.Local != "sst" {
		ss.count = -1
		return
	}
	// countとuniqueCountを削除する
	var attrs []xml.Attr
	for _, attr := range tag.Attr {
		if attr.Name.Local == "count" || attr.Name.Local == "uniqueCount" {
			continue
		}
		attrs = append(attrs, attr)
	}
	tag.Attr = attrs
	ss.count = 0
	for _, t := range tag.Children {
		switch child := t.(type) {
		case *Tag:
			if child.Name.Local == "si" {
				ss.count++
			}
		}
	}
}

// setSeparatePoint sharedStrings.xmlのセパレートポイントをセットする
func (ss *SharedStrings) setSeparatePoint(tag *Tag) {
	if ss.count == 0 {
		tag.Children = append(tag.Children, separateTag())
		return
	}
	count := 0
	for i := 0; i < len(tag.Children); i++ {
		switch child := tag.Children[i].(type) {
		case *Tag:
			if child.Name.Local == "si" {
				count++
				if count == ss.count {
					children := make([]interface{}, len(tag.Children)+1)
					copy(children, tag.Children[:i+1])
					children[i+1] = separateTag()
					copy(children[i+2:], tag.Children[i+1:])
					tag.Children = children
					return
				}
			}
		}
	}
}
