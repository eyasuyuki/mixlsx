package internal

import (
	"encoding/json"
	"os"
)

type Replacement struct {
	Title string `json:"title"`
	Value string `json:"value"`
	Range string `json:"range"`
	Color string `json:"color"`
}

type Worksheet struct {
	Sheet        string        `json:"sheet"`
	Replacements []Replacement `json:"replacements"`
}

type Book struct {
	Path       string
	Worksheets []Worksheet `json:"worksheets"`
}

func ReadBook(p string) *Book {
	b := &Book{}
	b.Path = p
	bytes, err := os.ReadFile(b.Path)
	if err != nil {
		panic(err)
	}
	b = CreateBook(bytes, b)
	return b
}

func CreateBook(bytes []byte, b *Book) *Book {
	if err := json.Unmarshal(bytes, b); err != nil {
		panic(err)
	}
	return b
}
