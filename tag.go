package excl

import (
	"encoding/xml"
	"errors"
	"io"
)

// Tag タグの情報をすべて保管する
type Tag struct {
	Name      xml.Name
	Attr      []xml.Attr
	Children  []interface{}
	XmlnsList []xml.Attr
}

// MarshalXML タグからXMLを作成しなおす
func (t *Tag) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = t.Name
	start.Attr = t.Attr
	e.EncodeToken(start)
	for _, v := range t.Children {
		switch v.(type) {
		case *Tag:
			child := v.(*Tag)
			if err := e.Encode(child); err != nil {
				return err
			}
		case xml.CharData:
			e.EncodeToken(v.(xml.CharData))
		case xml.Comment:
			e.EncodeToken(v.(xml.Comment))
		}
	}
	e.EncodeToken(start.End())
	return nil
}

// UnmarshalXML タグにXMLを読み込む
func (t *Tag) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	t.Name = start.Name
	t.Attr = start.Attr
	for _, at := range t.XmlnsList {
		if t.Name.Space != at.Value {
			continue
		}
		t.Name.Space = ""
		t.Name.Local = at.Name.Local + ":" + t.Name.Local
		break
	}
	if t.Name.Space != "" {
		t.Name.Space = ""
	}
	for index, attr := range start.Attr {
		if attr.Name.Space == "xmlns" {
			t.XmlnsList = append(t.XmlnsList, attr)
			start.Attr[index].Name.Local = start.Attr[index].Name.Space + ":" + start.Attr[index].Name.Local
			start.Attr[index].Name.Space = ""
			continue
		}
		for _, at := range t.XmlnsList {
			if attr.Name.Space != at.Value {
				continue
			}
			start.Attr[index].Name.Space = ""
			start.Attr[index].Name.Local = at.Name.Local + ":" + start.Attr[index].Name.Local
			break
		}
	}
	for {
		token, err := d.Token()
		if err != nil {
			if err == io.EOF {
				return nil
			}
			return err
		}
		switch token.(type) {
		case xml.StartElement:
			tok := token.(xml.StartElement)
			data := &Tag{XmlnsList: t.XmlnsList}
			if err := d.DecodeElement(&data, &tok); err != nil {
				return err
			}
			t.Children = append(t.Children, data)
		case xml.CharData:
			t.Children = append(t.Children, token.(xml.CharData).Copy())
		case xml.Comment:
			t.Children = append(t.Children, token.(xml.Comment).Copy())
		}
	}
}

func separateTag() *Tag {
	return &Tag{Name: xml.Name{Local: "separate_tag"}}
}

// setAttr 要素をセットする。要素がある場合は上書きする
// ない場合は追加する
func (t *Tag) setAttr(name string, val string) xml.Attr {
	for index, attr := range t.Attr {
		if attr.Name.Local == name {
			t.Attr[index].Value = val
			return t.Attr[index]
		}
	}
	attr := xml.Attr{
		Name:  xml.Name{Local: name},
		Value: val,
	}
	t.Attr = append(t.Attr, attr)
	return attr
}

// deleteAttr 要素を削除する
func (t *Tag) deleteAttr(name string) {
	for i := 0; i < len(t.Attr); i++ {
		attr := t.Attr[i]
		if attr.Name.Local == name {
			t.Attr = append(t.Attr[:i], t.Attr[i+1:]...)
			break
		}
	}
}

func (t *Tag) getAttr(name string) (string, error) {
	for _, attr := range t.Attr {
		if attr.Name.Local == name {
			return attr.Value, nil
		}
	}
	return "", errors.New("No attr found.")
}
