package excl

import "testing"

func TestSetAttr(t *testing.T) {
	tag := &Tag{}
	attr := tag.setAttr("s", "1")
	if attr.Name.Local != "s" {
		t.Error("attr name should be s but", attr.Name.Local)
	} else if attr.Value != "1" {
		t.Error("attr value should be 1 but", attr.Value)
	}

	tag.setAttr("s", "2")
	if tag.Attr[0].Value != "2" {
		t.Error("attr value should be 2 but", attr.Value)
	}
}

func TestDeleteAttr(t *testing.T) {
	tag := &Tag{}
	tag.setAttr("a", "1")
	tag.setAttr("b", "2")
	tag.setAttr("c", "3")
	tag.deleteAttr("b")
	if tag.Attr[0].Name.Local != "a" {
		t.Error("attr name should be a but", tag.Attr[0].Name.Local)
	} else if tag.Attr[0].Value != "1" {
		t.Error("attr value should be 1 but", tag.Attr[0].Value)
	} else if tag.Attr[1].Name.Local != "c" {
		t.Error("attr name should be c but", tag.Attr[1].Name.Local)
	} else if tag.Attr[1].Value != "3" {
		t.Error("attr value should be 3 but", tag.Attr[1].Value)
	} else if len(tag.Attr) != 2 {
		t.Error("attr count should be 2 but", len(tag.Attr))
	}
}
