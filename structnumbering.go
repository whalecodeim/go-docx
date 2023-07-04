package docx

import (
	"encoding/xml"
	"io"
)

type Numbering struct {
	XMLName xml.Name `xml:"w:numbering"`
	XMLW    string   `xml:"xmlns:w,attr"`             // cannot be unmarshalled in
	XMLR    string   `xml:"xmlns:r,attr,omitempty"`   // cannot be unmarshalled in
	XMLWP   string   `xml:"xmlns:wp,attr,omitempty"`  // cannot be unmarshalled in
	XMLWPS  string   `xml:"xmlns:wps,attr,omitempty"` // cannot be unmarshalled in
	XMLWPC  string   `xml:"xmlns:wpc,attr,omitempty"` // cannot be unmarshalled in
	XMLWPG  string   `xml:"xmlns:wpg,attr,omitempty"` // cannot be unmarshalled in

	AbstractNum []*AbstractNum

	file *Docx
}

type AbstractNum struct {
	XMLName        xml.Name `xml:"w:abstractNum,omitempty"`
	AbstractNumId  string   `xml:"w:abstractNumId"`
	Nsid           *Nsid
	MultiLevelType *MultiLevelType
	Tmpl           *Tmpl
}

type Nsid struct {
	XMLName xml.Name `xml:"w:nsid"`
	Val     string   `xml:"w:val"`
}

type MultiLevelType struct {
	XMLName xml.Name `xml:"w:multiLevelType"`
	Val     string   `xml:"w:val"`
}

type Tmpl struct {
	XMLName xml.Name `xml:"w:tmpl"`
	Val     string   `xml:"w:val"`
}

type Lvl struct {
	XMLName   xml.Name `xml:"w:lvl"`
	Ilvl      string   `xml:"w:ilvl"`
	Tentative string   `xml:"w:tentative"`
	Start     *Start
	NumFmt    *NumFmt
	LvlText   *LvlText
	LvlJc     *LvlJc
	PPr       *ParagraphProperties
	RPr       *RunProperties
}

type Start struct {
	XMLName xml.Name `xml:"w:start"`
	Val     string   `xml:"w:val"`
}

type NumFmt struct {
	XMLName xml.Name `xml:"w:numFmt"`
	Val     string   `xml:"w:val"`
}

type LvlText struct {
	XMLName xml.Name `xml:"w:lvlText"`
	Val     string   `xml:"w:val"`
}

type LvlJc struct {
	XMLName xml.Name `xml:"w:lvlJc"`
	Val     string   `xml:"w:val"`
}

func (p *Numbering) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
		}
	}
	return nil
}
