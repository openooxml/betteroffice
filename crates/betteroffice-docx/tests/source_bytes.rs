mod common;

use betteroffice_docx::Document;
use common::save_unedited;

const DECLARATION: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
const NAMESPACES: &str = r#"xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15""#;
const SECTION: &str = r#"<w:sectPr w:rsidR="00A1B2C3"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>"#;

fn main_part(body: &str) -> String {
    format!("{DECLARATION}<w:document {NAMESPACES}><w:body>{body}{SECTION}</w:body></w:document>")
}

fn docx(document: &str) -> Vec<u8> {
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), document.as_bytes().to_vec()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_main_part(package: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(package).unwrap();
    let (_, bytes) = parts
        .iter()
        .find(|(name, _)| name == "word/document.xml")
        .expect("word/document.xml");
    String::from_utf8(bytes.clone()).unwrap()
}

fn paragraph(id: &str, content: &str) -> String {
    format!(
        r#"<w:p w14:paraId="{id}" w14:textId="77777777" w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3"><w:pPr><w:spacing w:after="120"/></w:pPr>{content}</w:p>"#
    )
}

fn run(text: &str) -> String {
    format!(r#"<w:r w:rsidRPr="00D4E5F6"><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

fn paragraph_slices(xml: &str) -> Vec<&str> {
    xml.match_indices("<w:p")
        .filter(|(at, _)| matches!(xml.as_bytes()[at + 4], b' ' | b'>'))
        .map(|(at, _)| {
            let end = at + xml[at..].find("</w:p>").expect("closed paragraph") + "</w:p>".len();
            &xml[at..end]
        })
        .collect()
}

fn written_ids(xml: &str) -> Vec<String> {
    xml.match_indices("<w:p")
        .filter(|(at, _)| matches!(xml.as_bytes()[at + 4], b' ' | b'>' | b'/'))
        .map(|(at, _)| {
            let tag = &xml[at..at + xml[at..].find('>').unwrap()];
            tag.split(r#"w14:paraId=""#)
                .nth(1)
                .and_then(|rest| rest.split('"').next())
                .unwrap_or_default()
                .to_owned()
        })
        .collect()
}

fn without_paragraph_id_values(xml: &str) -> String {
    const ATTRIBUTE: &str = r#"w14:paraId=""#;
    let mut output = String::with_capacity(xml.len());
    let mut rest = xml;
    while let Some(at) = rest.find(ATTRIBUTE) {
        let value = at + ATTRIBUTE.len();
        output.push_str(&rest[..value]);
        let end = rest[value..]
            .find('"')
            .map_or(rest.len(), |end| value + end);
        rest = &rest[end..];
    }
    output.push_str(rest);
    output
}

#[test]
fn unedited_save_keeps_main_part_bytes() {
    let source = main_part(
        &[
            paragraph("0000000A", &run("First paragraph.")),
            paragraph(
                "0000000B",
                &format!(
                    r#"{}<w:proofErr w:type="spellStart"/>{}<w:proofErr w:type="spellEnd"/>"#,
                    run("Second "),
                    run("paragrph")
                ),
            ),
            paragraph("0000000C", &run("Third paragraph.")),
        ]
        .concat(),
    );
    let saved = saved_main_part(&save_unedited(&docx(&source)));
    assert_eq!(saved, source, "source bytes differ");
}

#[test]
fn reassigned_paragraph_ids_are_the_only_main_part_change() {
    let source = main_part(
        &[
            paragraph("1A2B3C4D", &run("first")),
            paragraph("1A2B3C4D", &run("second")),
            paragraph("0B000003", &run("third")),
        ]
        .concat(),
    );
    let mut document = Document::open(&docx(&source)).unwrap();
    let fresh: Vec<String> = document
        .paragraphs()
        .into_iter()
        .map(|paragraph| paragraph.para_id.clone().unwrap_or_default())
        .collect();
    assert_ne!(fresh[1], "1A2B3C4D");
    document.model_mut();
    let saved = saved_main_part(&document.save().unwrap());
    assert_eq!(
        without_paragraph_id_values(&saved),
        without_paragraph_id_values(&source),
        "source bytes differ beyond w14:paraId values"
    );
    assert_eq!(written_ids(&saved), fresh);
}

#[test]
fn replacing_one_paragraph_keeps_other_paragraph_bytes() {
    let source = main_part(
        &[
            paragraph("0000000A", &run("Alpha.")),
            paragraph(
                "0000000B",
                &format!(
                    r#"{}<w:proofErr w:type="spellStart"/>{}<w:proofErr w:type="spellEnd"/>"#,
                    run("Bravo "),
                    run("chrlie")
                ),
            ),
            paragraph("0000000C", &run("Old text")),
            paragraph(
                "0000000D",
                &format!(
                    r#"<w:bookmarkStart w:id="0" w:name="mark"/>{}<w:bookmarkEnd w:id="0"/>"#,
                    run("Delta.")
                ),
            ),
            paragraph("0000000E", &run("Echo.")),
        ]
        .concat(),
    );
    let mut document = Document::open(&docx(&source)).unwrap();
    document
        .replace_paragraph_text("0000000C", "New text")
        .unwrap();
    let saved = saved_main_part(&document.save().unwrap());
    let (before, after) = (paragraph_slices(&source), paragraph_slices(&saved));
    assert_eq!(after.len(), before.len());
    assert!(after[2].contains("New text"), "{}", after[2]);
    for index in [0, 1, 3, 4] {
        assert_eq!(
            after[index], before[index],
            "source bytes differ in paragraph {index}"
        );
    }
}

#[test]
fn unedited_save_keeps_table_of_contents_field_marks() {
    let entry = |anchor: &str, title: &str, page: &str| {
        format!(
            r#"<w:hyperlink w:anchor="{anchor}" w:history="1"><w:r><w:rPr><w:noProof/></w:rPr><w:t>{title}</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF {anchor} \h </w:instrText></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>{page}</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>"#
        )
    };
    let toc_properties = r#"<w:pPr><w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs><w:rPr><w:noProof/></w:rPr></w:pPr>"#;
    let toc = [
        format!(
            r#"<w:p w14:paraId="0000001A" w14:textId="77777777">{toc_properties}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>{}</w:p>"#,
            entry("_Toc100000001", "Introduction", "1")
        ),
        format!(
            r#"<w:p w14:paraId="0000001B" w14:textId="77777777">{toc_properties}{}</w:p>"#,
            entry("_Toc100000002", "Scope", "2")
        ),
        r#"<w:p w14:paraId="0000001C" w14:textId="77777777"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#.to_owned(),
    ]
    .concat();
    let heading = |id: &str, bookmark: u32, anchor: &str, title: &str| {
        format!(
            r#"<w:p w14:paraId="{id}" w14:textId="77777777"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="{bookmark}" w:name="{anchor}"/><w:r><w:t>{title}</w:t></w:r><w:bookmarkEnd w:id="{bookmark}"/></w:p>"#
        )
    };
    let source = main_part(
        &[
            toc.clone(),
            heading("0000001D", 0, "_Toc100000001", "Introduction"),
            heading("0000001E", 1, "_Toc100000002", "Scope"),
        ]
        .concat(),
    );
    let saved = saved_main_part(&save_unedited(&docx(&source)));
    assert!(
        saved.contains(&toc),
        "source bytes differ in the table of contents field"
    );
    assert_eq!(saved, source, "source bytes differ");
}

#[test]
fn unedited_save_keeps_proofing_marks() {
    let marked = r#"<w:p w14:paraId="0000002A" w14:textId="77777777"><w:r><w:t xml:space="preserve">A </w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t>speling</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:proofErr w:type="gramStart"/><w:r><w:t>grammer</w:t></w:r><w:proofErr w:type="gramEnd"/></w:p>"#;
    let source = main_part(
        &[
            marked,
            r#"<w:p w14:paraId="0000002B" w14:textId="77777777"><w:r><w:t>plain</w:t></w:r></w:p>"#,
        ]
        .concat(),
    );
    let saved = saved_main_part(&save_unedited(&docx(&source)));
    assert!(
        saved.contains(marked),
        "source bytes differ in the proofing marks"
    );
    assert_eq!(saved, source, "source bytes differ");
}

#[test]
fn unedited_save_keeps_run_revision_ids() {
    let marked = r#"<w:p w14:paraId="0000003A" w14:textId="77777777"><w:r w:rsidR="00C1D2E3" w:rsidRPr="00F4A5B6"><w:t xml:space="preserve">Revised </w:t></w:r><w:r w:rsidR="00C1D2E3"><w:t>run</w:t></w:r></w:p>"#;
    let source = main_part(
        &[
            marked,
            r#"<w:p w14:paraId="0000003B" w14:textId="77777777"><w:r><w:t>plain</w:t></w:r></w:p>"#,
        ]
        .concat(),
    );
    let saved = saved_main_part(&save_unedited(&docx(&source)));
    assert!(
        saved.contains(marked),
        "source bytes differ in the run revision ids"
    );
    assert_eq!(saved, source, "source bytes differ");
}
