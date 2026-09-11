pub fn document(alternate: bool) -> Vec<u8> {
    let paragraphs = (1..=12)
        .map(|index| format!(r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>Paragraph {index}: a floating header must leave the body in its authored position.</w:t></w:r></w:p>"#))
        .collect::<String>();
    let mut parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_owned()),
        ("_rels/.rels".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml".to_owned(), format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{paragraphs}
<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>
<w:pgSz w:w="16840" w:h="11900" w:orient="landscape"/>
<w:pgMar w:top="1700" w:right="1700" w:bottom="1700" w:left="1700" w:header="850" w:footer="850"/>
</w:sectPr></w:body></w:document>"#)),
        ("word/_rels/document.xml.rels".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#.to_owned()),
        ("word/header1.xml".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>9620250</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>539750</wp:posOffset></wp:positionV>
<wp:extent cx="476250" cy="5539105"/><wp:wrapNone/><wp:docPr id="1" name="Floating header decoration"/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>
<wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="476250" cy="5539105"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="CDEBFA"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>
<wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>
</wp:anchor></w:drawing></w:r></w:p></w:hdr>"#.to_owned()),
    ];
    if alternate {
        for (name, value) in &mut parts {
            *name = name.replace("document.xml", "document2.xml");
            if matches!(name.as_str(), "[Content_Types].xml" | "_rels/.rels") {
                *value = value.replace("word/document.xml", "word/document2.xml");
            }
            if name == "word/document2.xml" {
                *value = value.replace("<w:document ", r#"<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" "#);
                for index in 1..=12 {
                    *value =
                        value.replacen("<w:p>", &format!(r#"<w:p w14:paraId="{index:08X}">"#), 1);
                }
            }
        }
    }
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name, value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}
