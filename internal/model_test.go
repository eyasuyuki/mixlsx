package internal

import "testing"

const JSON = `{"worksheets": [
	{"sheet": "Sheet1",
	 "replacements": [
		{"title": "title",
		"value": "abc",
		"range": "B1",
		"color": "red"
		},
		{"title": "hogehoge",
		"value": "def",
		"range": "C1"
		}
	]}
]}`

func TestCreateBook(t *testing.T) {
	b := &Book{}
	b = CreateBook([]byte(JSON), b)
	if b == nil {
		t.Errorf("result is nil\n")
	}
	if len(b.Worksheets) < 1 {
		t.Errorf("worksheet length is 0\n")
	}
	w := b.Worksheets[0]
	if w.Sheet != "Sheet1" {
		t.Errorf("sheet is %v\n", w.Sheet)
	}
	//
	r := w.Replacements[0]
	if r.Title != "title" {
		t.Errorf("title is %v\n", r.Title)
	}
	if r.Value != "abc" {
		t.Errorf("value is %v\n", r.Value)
	}
	if r.Range != "B1" {
		t.Errorf("range is %v\n", r.Range)
	}
	if r.Color != "red" {
		t.Errorf("color is %v\n", r.Color)
	}
	//
	r = w.Replacements[1]
	if r.Title != "hogehoge" {
		t.Errorf("title is %v\n", r.Title)
	}
	if r.Value != "def" {
		t.Errorf("value is %v\n", r.Value)
	}
	if r.Range != "C1" {
		t.Errorf("range is %v\n", r.Range)
	}
	if r.Color != "" {
		t.Errorf("color is %v\n", r.Color)
	}
}
