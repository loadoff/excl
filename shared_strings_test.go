package excl

import (
	"encoding/xml"
	"io/ioutil"
	"os"
	"path/filepath"
	"testing"
)

func TestOpenSharedStrings(t *testing.T) {
	os.Mkdir("temp/xl", 0755)
	defer os.RemoveAll("temp/xl")
	_, err := OpenSharedStrings("nopath")
	if err == nil {
		t.Error("sharedStrings.xml file should not be opened.")
	}
	ss, err := OpenSharedStrings("temp")
	if ss == nil {
		t.Error("structure should be created but [", err.Error(), "]")
	} else {
		if !isFileExist(filepath.Join("temp", "xl", "sharedStrings.xml")) {
			t.Error("sharedStrings.xml file should be created.")
		}
		if ss.count != 0 {
			t.Error("count should be 0 but [", ss.count, "]")
		}
		if !isFileExist(filepath.Join("temp", "xl", "__sharedStrings.xml")) {
			t.Error("__sharedStrings.xml should be opened.")
		}
		ss.Close()
		if isFileExist(filepath.Join("temp", "xl", "__sharedStrings.xml")) {
			t.Error("__sharedStrings.xml should be removed.")
		}
	}

	f, _ := os.Create(filepath.Join("temp", "xl", "sharedStrings.xml"))
	f.Close()
	_, err = OpenSharedStrings("temp")
	if err == nil {
		t.Error("sharedStrings.xml should not be parsed.")
	}

	f, _ = os.Create(filepath.Join("temp", "xl", "sharedStrings.xml"))
	f.WriteString("<currupt></currupt>")
	f.Close()
	_, err = OpenSharedStrings("temp")
	if err == nil {
		t.Error("sharedStrings.xml file should be currupt.")
	}

	f, _ = os.Create(filepath.Join("temp", "xl", "sharedStrings.xml"))
	f.WriteString("<sst><si><t></t></si><si><t></t></si></sst>")
	f.Close()
	ss, err = OpenSharedStrings("temp")
	if err != nil {
		t.Error("sharedStrings.xml should be opend.[", err.Error(), "]")
	}
	if ss.count != 2 {
		t.Error("strings count should be 2 but [", ss.count, "]")
	}
	ss.Close()
	f, _ = os.Create(filepath.Join("temp", "xl", "sharedStrings.xml"))
	f.WriteString("<sst><separate_tag></separate_tag></sst>")
	f.Close()
	ss, err = OpenSharedStrings("temp")
	if err == nil {
		t.Error("sharedString.xml should not be opened because the file is currupt.")
	}
}

func TestAddString(t *testing.T) {
	os.Mkdir("temp/xl", 0755)
	defer os.RemoveAll("temp/xl")
	f, _ := os.Create(filepath.Join("temp", "xl", "sharedStrings.xml"))
	f.WriteString("<sst><si></si></sst>")
	f.Close()
	ss, _ := OpenSharedStrings("temp")
	if index := ss.AddString("hello world"); index != 1 {
		t.Error("index should be 1 but [", index, "]")
	}
	ss.Close()
	f, _ = os.Open(filepath.Join("temp", "xl", "sharedStrings.xml"))
	b, _ := ioutil.ReadAll(f)
	f.Close()
	if string(b) != "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst><si></si><si><t>hello world</t></si></sst>" {
		t.Error(string(b))
	}
}

func BenchmarkAddString(b *testing.B) {
	s := "あいうえお"
	f, _ := os.Create("temp/sharedStrings.xml")
	//defer f.Close()
	//sharedStrings := &SharedStrings{tempFile: f}
	for i := 1; i < 100000; i++ {
		for j := 0; j < 20; j++ {
			//sharedStrings.AddString("あいうえお")
			//var b bytes.Buffer
			f.WriteString("<si><t>")
			xml.EscapeText(f, []byte(s))
			f.WriteString("</t></si>")

		}
	}
}
