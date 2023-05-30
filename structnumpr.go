package docx

import (
	"encoding/xml"
	"io"
	"strconv"
)

// NumProperties <w:numPr>
type NumProperties struct {
	XMLName xml.Name `xml:"w:numPr,omitempty"`
	Ilvl    *Ilvl
	NumId   *NumId
}

// Ilvl ...
type Ilvl struct {
	XMLName xml.Name `xml:"w:ilvl,omitempty"`
	Val     int      `xml:"w:val,attr"`
}

// NumId ...
type NumId struct {
	XMLName xml.Name `xml:"w:numId,omitempty"`
	Val     int      `xml:"w:val,attr"`
}

func (p *NumProperties) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "ilvl":
				var value Ilvl
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = strconv.Atoi(v)
				if err != nil {
					return err
				}
				p.Ilvl = &value
			case "numId":
				var value NumId
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = strconv.Atoi(v)
				if err != nil {
					return err
				}
				p.NumId = &value
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
