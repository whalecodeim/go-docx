package main

import (
	"fmt"
	"github.com/yangge2333/go-docx"
	"io"
	"os"
)

func Fumiama() {
	d := docx.NewA4()
	p := d.AddParagraph()
	p.AddLink("link", "b")
	table := d.AddTable(1, 2)
	para := table.TableRows[0].TableCells[1].AddParagraph()
	para.AddText("hello")

	f, err := os.Create("testdata/test_out.docx")
	if err != nil {
		fmt.Println(err.Error())
	}
	_, err = d.WriteTo(f)
	if err != nil {
		fmt.Println(err.Error())
	}
}

func Open() {
	file, err := os.Open("testdata/test.docx")
	if err != nil {
		return
	}
	defer file.Close()
	fileInfo, err := file.Stat()
	if err != nil {
		return
	}
	size := fileInfo.Size()
	reader := io.ReaderAt(file)
	doc, err := docx.Parse(reader, size)
	fmt.Println(doc)
}

func main() {
	Fumiama()
}
