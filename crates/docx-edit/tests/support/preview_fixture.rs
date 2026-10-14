use std::collections::BTreeSet;

use docx_edit::EngineSession;

pub fn document(body: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.png"/></Relationships>"#.to_owned()),
        ("word/numbering.xml", r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body>{body}</w:body></w:document>"#)),
    ];
    let mut parts: Vec<_> = parts
        .into_iter()
        .map(|(name, value)| (name.to_owned(), value.into_bytes()))
        .collect();
    parts.push((
        "word/media/image.png".to_owned(),
        vec![
            137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1,
            8, 6, 0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 11, 73, 68, 65, 84, 120, 156, 99, 96, 0, 2,
            0, 0, 5, 0, 1, 169, 245, 232, 39, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
        ],
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

pub fn run(text: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

pub fn revision(kind: &str, id: &str, content: &str) -> String {
    format!(
        r#"<w:{kind} w:id="{id}" w:author="Ann" w:date="2026-09-29T12:00:00Z">{content}</w:{kind}>"#
    )
}

pub fn paragraph(id: u32, content: &str) -> String {
    format!(r#"<w:p w14:paraId="{id:08X}">{content}</w:p>"#)
}

pub fn table(content: &str) -> String {
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="2400" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>{content}</w:tc></w:tr></w:tbl>"#
    )
}

pub fn field(instruction: &str) -> String {
    format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
    )
}

pub fn plain() -> Vec<u8> {
    document(&format!(
        "{}{}{}{}",
        paragraph(1, &run("Before")),
        paragraph(
            2,
            &format!(
                "{}{}{}{}",
                run("A "),
                revision("ins", "1", &run("insert")),
                revision("del", "2", "<w:r><w:delText>delete</w:delText></w:r>"),
                revision("ins", "3", &run("again"))
            )
        ),
        paragraph(
            3,
            &revision("del", "4", "<w:r><w:delText>gone</w:delText></w:r>")
        ),
        paragraph(4, &run("After")),
    ))
}

pub fn nested() -> Vec<u8> {
    let inner = table(&paragraph(12, &revision("ins", "12", &run("nested"))));
    let outer = table(&format!(
        "{}{}{}",
        paragraph(
            10,
            &revision("del", "10", "<w:r><w:delText>cell</w:delText></w:r>")
        ),
        inner,
        paragraph(13, &run("cell end"))
    ));
    let block = format!(
        r#"<w:sdt><w:sdtPr><w:id w:val="20"/><w:tag w:val="block"/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>"#,
        paragraph(20, &revision("ins", "20", &run("block control")))
    );
    let inline = format!(
        r#"<w:sdt><w:sdtPr><w:id w:val="30"/><w:tag w:val="inline"/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>"#,
        revision("ins", "30", &run("inline control"))
    );
    document(&format!(
        "{}{}{}{}{}",
        paragraph(1, &run("Before")),
        outer,
        block,
        paragraph(30, &format!("{}{}", run("Inline "), inline)),
        paragraph(31, &run("After"))
    ))
}

pub fn breaks() -> Vec<u8> {
    let numbering = r#"<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>"#;
    document(&format!(
        "{}{}{}{}{}",
        paragraph(1, &run("Before")),
        paragraph(
            2,
            &format!(
                "{numbering}{}<w:r><w:br w:type=\"page\"/></w:r>",
                revision("ins", "1", &run("numbered"))
            )
        ),
        paragraph(
            3,
            &format!(
                "{numbering}{}{}",
                revision("ins", "2", &run("second")),
                revision("ins", "3", "<w:r><w:br w:type=\"page\"/></w:r>")
            )
        ),
        paragraph(
            4,
            &format!(
                "{}{}",
                revision("ins", "4", &run("column")),
                revision("ins", "5", "<w:r><w:br w:type=\"column\"/></w:r>")
            )
        ),
        paragraph(5, &run("After"))
    ))
}

pub fn drawings() -> Vec<u8> {
    let image = r#"<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="image"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="image"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#;
    let shape = r#"<w:r><w:drawing><wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:wrapNone/><wp:docPr id="2" name="shape"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#;
    document(&format!(
        "{}{}{}{}",
        paragraph(1, &run("Before")),
        paragraph(2, &revision("ins", "1", image)),
        paragraph(
            3,
            &format!("{}{}{}", run("A"), revision("ins", "2", shape), run("B"))
        ),
        paragraph(4, &run("After"))
    ))
}

pub fn fields() -> Vec<u8> {
    document(&format!(
        "{}{}{}",
        paragraph(
            1,
            &format!("{}{}", revision("ins", "1", &run("Before ")), field("PAGE"))
        ),
        paragraph(
            2,
            &format!(
                "{}{}",
                revision("del", "2", "<w:r><w:delText>Cached</w:delText></w:r>"),
                field("NUMPAGES")
            )
        ),
        paragraph(3, &run("After"))
    ))
}

pub fn sequence() -> Vec<u8> {
    document(&format!(
        "{}{}{}{}{}",
        paragraph(1, &revision("ins", "1", &run("Caption"))),
        paragraph(2, &field("SEQ Figure")),
        paragraph(3, &field("SEQ Figure")),
        paragraph(
            4,
            &format!(
                "{}{}",
                field("SEQ Figure"),
                revision("del", "3", &run(" old"))
            )
        ),
        paragraph(5, &field("SEQ Figure"))
    ))
}

pub fn ids(engine: &EngineSession) -> Vec<String> {
    engine
        .doc()
        .list_revisions()
        .unwrap()
        .into_iter()
        .map(|revision| revision.change.revision_id)
        .collect::<BTreeSet<_>>()
        .into_iter()
        .collect()
}

pub struct Random(u64);

impl Random {
    pub fn new(seed: u64) -> Self {
        Self(seed + 1)
    }

    pub fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }
}

pub fn corpus() -> [(&'static str, &'static [u8]); 3] {
    [
        (
            "comprehensive",
            include_bytes!(
                "../../../betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx"
            ),
        ),
        (
            "pages",
            include_bytes!("../fixtures/page-fragments/pages.docx"),
        ),
        (
            "principal",
            include_bytes!("../fixtures/structured-export/principal.docx"),
        ),
    ]
}

pub fn decide(env: &mut docx_edit::bridge::RenderEnv, ids: &[String], random: &mut Random) {
    use docx_edit::bridge::RevisionPreview::{Accepted, Rejected};
    let before = env.revision_preview.clone();
    for id in ids {
        if !random.next().is_multiple_of(3) {
            continue;
        }
        match random.next() % 3 {
            0 => {
                env.revision_preview.remove(id);
            }
            1 => {
                env.revision_preview.insert(id.clone(), Accepted);
            }
            _ => {
                env.revision_preview.insert(id.clone(), Rejected);
            }
        }
    }
    if before == env.revision_preview {
        let id = &ids[random.next() as usize % ids.len()];
        if env.revision_preview.get(id) == Some(&Accepted) {
            env.revision_preview.insert(id.clone(), Rejected);
        } else {
            env.revision_preview.insert(id.clone(), Accepted);
        }
    }
}

pub fn hidden_fields() -> Vec<u8> {
    document(&format!(
        "{}{}{}{}",
        paragraph(
            1,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> 123 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>"#
        ),
        paragraph(2, &run("Cached result")),
        paragraph(3, "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"),
        paragraph(4, &field("123"))
    ))
}

pub fn stamp_fields(engine: &EngineSession) {
    use docx_edit::{EditCtx, RawOp, SegmentContent};
    use yrs::Any;
    let mut index = 0;
    let mut ordinal = 0;
    let mut ops = Vec::new();
    for segment in engine.doc().story_segments("body").unwrap() {
        match segment.content {
            SegmentContent::Text(text) => {
                index += text.encode_utf16().count() as u32;
            }
            SegmentContent::OtherEmbed { kind, payload } if kind == "field" => {
                ordinal += 1;
                let deleted = payload.get("instruction").is_some_and(|value| {
                    matches!(value, Any::String(instruction) if instruction.trim() == "NUMPAGES")
                });
                let stamp = Any::Map(std::sync::Arc::new(std::collections::HashMap::from([(
                    "id".to_owned(),
                    Any::from(format!("field-{ordinal}")),
                )])));
                ops.push(RawOp::Format {
                    index,
                    len: 1,
                    attrs: [(if deleted { "del" } else { "ins" }.into(), stamp)].into(),
                });
                index += 1;
            }
            _ => {
                index += 1;
            }
        }
    }
    assert!(ordinal > 0);
    engine
        .doc()
        .apply_raw_ops("body", ops, &EditCtx::local("", ""))
        .unwrap();
}
