package main

import (
	"fmt"
	"github.com/yangge2333/go-docx"
	"os"
)

func Fumiama() {
	d := docx.NewA4()
	p := d.AddParagraph()
	p.AddLink("link", "b")
	f, err := os.Create("testdata/test_out.docx")
	if err != nil {
		fmt.Println(err.Error())
	}
	_, err = d.WriteTo(f)
	if err != nil {
		fmt.Println(err.Error())
	}
}

func main() {
	Fumiama()
}
