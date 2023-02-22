package docxlib

import (
	"encoding/xml"
	"hash/crc64"
	"io"
	"os"
	"testing"
)

const decoded_doc_1 = `<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex wp14"><w:body><w:p w14:paraId="77CA082D" w14:textId="4AF3264D" w:rsidR="00D66E3F" w:rsidRDefault="003A3F42"><w:pPr><w:rPr><w:color w:val="808080"/></w:rPr></w:pPr><w:proofErr w:type="spellStart"/><w:r><w:t>test</w:t></w:r><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t>test</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t xml:space="preserve"> font </w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t>size</w:t></w:r><w:r><w:rPr><w:color w:val="808080"/></w:rPr><w:t>test</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:rPr><w:color w:val="808080"/></w:rPr><w:t xml:space="preserve"> color</w:t></w:r></w:p><w:p w14:paraId="6D114165" w14:textId="04580C29" w:rsidR="003A3F42" w:rsidRDefault="003A3F42" w:rsidP="003A3F42"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>New style 1</w:t></w:r></w:p><w:p w14:paraId="40D72B3B" w14:textId="76101901" w:rsidR="003A3F42" w:rsidRDefault="003A3F42" w:rsidP="003A3F42"><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>New style 2</w:t></w:r></w:p><w:p w14:paraId="1CA8A9B3" w14:textId="77777777" w:rsidR="00D66E3F" w:rsidRDefault="003A3F42"><w:r><w:rPr><w:color w:val="FF0000"/><w:sz w:val="44"/></w:rPr><w:t>test font size and color</w:t></w:r></w:p><w:p w14:paraId="0D82FB8B" w14:textId="77777777" w:rsidR="00D66E3F" w:rsidRDefault="003A3F42"><w:hyperlink r:id="rId4"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>google</w:t></w:r></w:hyperlink></w:p><w:sectPr w:rsidR="00D66E3F"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>`
const decoded_doc_2 = `<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:body><w:sdt><w:sdtPr><w:id w:val="-1247033294"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtEndPr/><w:sdtContent><w:p w14:paraId="308E3D65" w14:textId="77777777" w:rsidR="00FA66BB" w:rsidRPr="001D59EF" w:rsidRDefault="00FA66BB" w:rsidP="00A96827"><w:pPr><w:pStyle w:val="TOC1"/><w:jc w:val="center"/></w:pPr><w:r><w:t>Table of Contents</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p w14:paraId="1764C163" w14:textId="77777777" w:rsidR="009E307C" w:rsidRDefault="00A96827"><w:pPr><w:pStyle w:val="TOC1"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:b w:val="0"/><w:color w:val="auto"/></w:rPr></w:pPr><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:instrText xml:space="preserve"> TOC \h \z \t "Heading 1,2,S6,1,S0,1,S1,1,S2,1,S3,1,S4,1,S5,1" </w:instrText></w:r><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="_Toc420414504" w:history="1"><w:r w:rsidR="009E307C" w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Holy Grail [xref:bRJduW6hNR]</w:t></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr><w:tab/></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF _Toc420414504 \h </w:instrText></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr><w:t>2</w:t></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p><w:p w14:paraId="0F5BA552" w14:textId="77777777" w:rsidR="009E307C" w:rsidRDefault="009E307C"><w:pPr><w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="left" w:pos="3654"/></w:tabs><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/></w:rPr></w:pPr><w:hyperlink w:anchor="_Toc420414505" w:history="1"><w:r w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>1.</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/></w:rPr><w:tab/></w:r><w:r w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>What is your name? [xref:TH7u7QDqhD]</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF _Toc420414505 \h </w:instrText></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>2</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p><w:p w14:paraId="49E1F0AA" w14:textId="77777777" w:rsidR="009E307C" w:rsidRDefault="009E307C"><w:pPr><w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="left" w:pos="3654"/></w:tabs><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/></w:rPr></w:pPr><w:hyperlink w:anchor="_Toc420414506" w:history="1"><w:r w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>2.</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/></w:rPr><w:tab/></w:r><w:r w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>What is your quest? [xref:bC62HkFATC]</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF _Toc420414506 \h </w:instrText></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>2</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p><w:p w14:paraId="7BDA743C" w14:textId="77777777" w:rsidR="009E307C" w:rsidRDefault="009E307C"><w:pPr><w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="left" w:pos="3654"/></w:tabs><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/></w:rPr></w:pPr><w:hyperlink w:anchor="_Toc420414507" w:history="1"><w:r w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>3.</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/></w:rPr><w:tab/></w:r><w:r w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>What is your favourite colour? [xref:I3TphuHX6N]</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF _Toc420414507 \h </w:instrText></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>2</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p><w:p w14:paraId="4A0A0E88" w14:textId="77777777" w:rsidR="006C0D12" w:rsidRPr="009C59B6" w:rsidRDefault="00A96827" w:rsidP="009B657F"><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p w14:paraId="7EDC60AD" w14:textId="77777777" w:rsidR="0004272B" w:rsidRDefault="0004272B"><w:pPr><w:jc w:val="left"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:br w:type="page"/></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p><w:p w14:paraId="6775D4EA" w14:textId="4B1B4185" w:rsidR="00EF5BF6" w:rsidRDefault="00DE7E6E" w:rsidP="00EF5BF6"><w:pPr><w:pStyle w:val="S0"/></w:pPr><w:bookmarkStart w:id="1" w:name="_Toc388285991"/><w:bookmarkStart w:id="2" w:name="_Toc388366779"/><w:bookmarkStart w:id="3" w:name="_Toc388428327"/><w:bookmarkStart w:id="4" w:name="_Toc388451002"/><w:bookmarkStart w:id="5" w:name="_Toc420414504"/><w:r><w:lastRenderedPageBreak/><w:t>Holy Grail</w:t></w:r><w:bookmarkEnd w:id="1"/><w:bookmarkEnd w:id="2"/><w:bookmarkEnd w:id="3"/><w:bookmarkEnd w:id="4"/><w:r w:rsidR="009E307C"><w:t xml:space="preserve"> [</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="bookmark"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>xref</w:t></w:r><w:proofErr w:type="gramStart"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>:bRJduW6hNR</w:t></w:r><w:proofErr w:type="gramEnd"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:r w:rsidR="009E307C"><w:t>]</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p><w:p w14:paraId="2E760FD1" w14:textId="10909973" w:rsidR="00DE7E6E" w:rsidRDefault="00DE7E6E" w:rsidP="00DE7E6E"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="6" w:name="_Toc389482870"/><w:bookmarkStart w:id="7" w:name="_Toc420414505"/><w:r><w:t>What is your name?</w:t></w:r><w:bookmarkEnd w:id="6"/><w:r w:rsidR="009E307C"><w:t xml:space="preserve"> [</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="bookmark"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:proofErr w:type="gramStart"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>xref:</w:t></w:r><w:proofErr w:type="gramEnd"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>TH7u7QDqhD</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:r w:rsidR="009E307C"><w:t>]</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p><w:p w14:paraId="0249939B" w14:textId="77777777" w:rsidR="003946B5" w:rsidRPr="003946B5" w:rsidRDefault="00DE7E6E" w:rsidP="003946B5"><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r w:rsidRPr="0029440C"><w:t xml:space="preserve">My name is Sir </w:t></w:r><w:proofErr w:type="spellStart"/><w:r w:rsidRPr="0029440C"><w:t>Launcelot</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r w:rsidRPr="0029440C"><w:t xml:space="preserve"> of Camelot.</w:t></w:r></w:p><w:p w14:paraId="5BB04A25" w14:textId="04E09ADD" w:rsidR="006F5AAA" w:rsidRPr="006F5AAA" w:rsidRDefault="00DE7E6E" w:rsidP="00DE7E6E"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="8" w:name="_Toc389482871"/><w:bookmarkStart w:id="9" w:name="_Toc420414506"/><w:r><w:t>What is your quest?</w:t></w:r><w:bookmarkEnd w:id="8"/><w:r w:rsidR="009E307C"><w:t xml:space="preserve"> [</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="bookmark"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:proofErr w:type="gramStart"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>xref:</w:t></w:r><w:proofErr w:type="gramEnd"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>bC62HkFATC</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:r w:rsidR="009E307C"><w:t>]</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p><w:p w14:paraId="15194710" w14:textId="77777777" w:rsidR="002B0891" w:rsidRDefault="00DE7E6E" w:rsidP="002B0891"><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t xml:space="preserve">To seek the Holy </w:t></w:r><w:proofErr w:type="gramStart"/><w:r><w:t>Grail</w:t></w:r><w:r w:rsidRPr="00225D92"><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>[</w:t></w:r><w:proofErr w:type="gramEnd"/><w:r w:rsidRPr="00225D92"><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>or a grail shaped beacon]</w:t></w:r><w:r><w:t>.</w:t></w:r><w:r w:rsidR="00585075"><w:t xml:space="preserve"> </w:t></w:r></w:p><w:p w14:paraId="05C7DE39" w14:textId="2A77E45D" w:rsidR="00585075" w:rsidRDefault="00DE7E6E" w:rsidP="006F5AAA"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="10" w:name="_Toc389482872"/><w:bookmarkStart w:id="11" w:name="_Toc420414507"/><w:r><w:t>What is your favourite colour?</w:t></w:r><w:bookmarkEnd w:id="10"/><w:r w:rsidR="009E307C"><w:t xml:space="preserve"> [</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="bookmark"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r><w:bookmarkStart w:id="12" w:name="bookmark"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:proofErr w:type="gramStart"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>xref:</w:t></w:r><w:proofErr w:type="gramEnd"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>I3TphuHX6N</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="12"/><w:r w:rsidR="009E307C"><w:t>]</w:t></w:r><w:bookmarkEnd w:id="11"/></w:p><w:p w14:paraId="5FA4E707" w14:textId="77777777" w:rsidR="00DE7E6E" w:rsidRDefault="00DE7E6E" w:rsidP="00DE7E6E"><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t>Blue.</w:t></w:r></w:p><w:p w14:paraId="543FEBD5" w14:textId="77777777" w:rsidR="006F5AAA" w:rsidRPr="006F5AAA" w:rsidRDefault="00DE7E6E" w:rsidP="00DE7E6E"><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t>How many paragraphs here then?</w:t></w:r></w:p><w:sectPr w:rsidR="006F5AAA" w:rsidRPr="006F5AAA" w:rsidSect="002B3068"><w:footerReference w:type="default" r:id="rId9"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="709" w:footer="709" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>`

func TestUnmarshalPlainStructure(t *testing.T) {
	testCases := []struct {
		content       string
		numParagraphs int
	}{
		{decoded_doc_1, 5},
		{decoded_doc_2, 14},
	}
	for _, tc := range testCases {
		doc := Document{
			XMLW:    XMLNS_W,
			XMLR:    XMLNS_R,
			XMLWP:   XMLNS_WP,
			XMLName: xml.Name{Space: XMLNS_W, Local: "document"}}
		err := xml.Unmarshal(StringToBytes(tc.content), &doc)
		if err != nil {
			t.Fatal(err)
		}
		if len(doc.Body.Paragraphs) != tc.numParagraphs {
			t.Fatalf("We expected %d paragraphs, we got %d", tc.numParagraphs, len(doc.Body.Paragraphs))
		}
		for i, p := range doc.Body.Paragraphs {
			if len(p.Children) == 0 {
				t.Fatalf("We were not able to parse paragraph %d", i)
			}
			for _, child := range p.Children {
				if child == nil {
					t.Fatalf("There are Paragraph children with all fields nil")
				}
				if o, ok := child.(*Hyperlink); ok && o.ID == "" {
					t.Fatalf("We have a link without ID")
				}
			}
		}
	}
}

func TestInlineDrawingStructure(t *testing.T) {
	w := NewA4()
	// add new paragraph
	para1 := w.AddParagraph()
	// add text
	para1.AddText("直接粘贴 inline").AddTab()
	r, err := para1.AddAnchorDrawingFrom("testdata/fumiama.JPG")
	if err != nil {
		t.Fatal(err)
	}
	r.Drawing.Anchor.Graphic.GraphicData.Pic.BlipFill.Blip.AlphaModFix = &AAlphaModFix{Amount: 50000}
	r.Drawing.Anchor.Graphic.GraphicData.Pic.NonVisualPicProperties.CNvPicPr.Locks = &APicLocks{NoChangeAspect: 1}
	r.Drawing.Anchor.Graphic.GraphicData.Pic.SpPr.Xfrm.Rot = 50000
	para2 := w.AddParagraph().Justification("center")
	para2.AddInlineDrawingFrom("testdata/fumiama.JPG")
	para2.AddTab().AddTab().AppendTab().AppendTab()
	para2.AddInlineDrawingFrom("testdata/fumiama2x.webp")

	para3 := w.AddParagraph()
	para3.AddInlineDrawingFrom("testdata/fumiamayoko.png")

	f, err := os.Create("TestMarshalInlineDrawingStructure.xml")
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	_, err = marshaller{data: w.Document}.WriteTo(f)
	if err != nil {
		t.Fatal(err)
	}
	_, err = f.Seek(0, io.SeekStart)
	if err != nil {
		t.Fatal(err)
	}
	w = NewA4()
	err = xml.NewDecoder(f).Decode(&w.Document)
	if err != nil {
		t.Fatal(err)
	}
	f1, err := os.Create("TestUnmarshalInlineDrawingStructure.xml")
	if err != nil {
		t.Fatal(err)
	}
	defer f1.Close()
	_, err = marshaller{data: w.Document}.WriteTo(f1)
	if err != nil {
		t.Fatal(err)
	}
	_, err = f.Seek(0, io.SeekStart)
	if err != nil {
		t.Fatal(err)
	}
	_, err = f1.Seek(0, io.SeekStart)
	if err != nil {
		t.Fatal(err)
	}
	h := crc64.New(crc64.MakeTable(crc64.ECMA))
	_, err = io.Copy(h, f)
	if err != nil {
		t.Fatal(err)
	}
	md51 := h.Sum64()
	h.Reset()
	_, err = io.Copy(h, f1)
	if err != nil {
		t.Fatal(err)
	}
	md52 := h.Sum64()
	if md51 != md52 {
		t.Fail()
	}
}
