package main

import (
	"fmt"
	"github.com/yangge2333/go-docx"
	"os"
)

func Fumiama() {
	readFile, err := os.Open("testdata/test.docx")
	if err != nil {
		panic(err)
	}
	fileinfo, err := readFile.Stat()
	if err != nil {
		panic(err)
	}
	size := fileinfo.Size()
	doc, err := docx.Parse(readFile, size)
	if err != nil {
		panic(err)
	}

	items := doc.Document.Body.Items
	for _, it := range items {
		switch it.(type) {
		case *docx.Paragraph:
			// printable
			para := it.(*docx.Paragraph)
			fmt.Println(para)
		case *docx.Table:
			table := it.(*docx.Table)
			fmt.Println(table)
		default:
			fmt.Println(it)
		}

	}
}

func main() {
	Fumiama()
}
