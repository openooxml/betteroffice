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
fn model_mut_borrow_writes_the_reassigned_paragraph_ids() {
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
    assert_eq!(written_ids(&saved), fresh);
    let reopened = Document::open(&docx(&saved)).unwrap();
    let texts: Vec<String> = reopened
        .paragraphs()
        .into_iter()
        .map(betteroffice_docx::get_paragraph_text)
        .collect();
    assert_eq!(texts, ["first", "second", "third"]);
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

fn paragraph_span(xml: &str, id: &str) -> (usize, usize) {
    let at = xml
        .find(&format!(r#"w14:paraId="{id}""#))
        .expect("edited paragraph");
    let start = xml[..at].rfind("<w:p ").expect("paragraph start");
    let end = at + xml[at..].find("</w:p>").expect("closed paragraph") + "</w:p>".len();
    (start, end)
}

fn assert_bytes_kept_around(source: &str, saved: &str, id: &str, text: &str) {
    let (start, end) = paragraph_span(source, id);
    let (saved_start, saved_end) = paragraph_span(saved, id);
    assert!(saved[saved_start..saved_end].contains(text), "{saved}");
    assert_eq!(
        &saved[..saved_start],
        &source[..start],
        "source bytes differ before the edited paragraph"
    );
    assert_eq!(
        &saved[saved_end..],
        &source[end..],
        "source bytes differ after the edited paragraph"
    );
}

fn edit_and_save(source: &str, id: &str, text: &str) -> String {
    let mut document = Document::open(&docx(source)).unwrap();
    document.replace_paragraph_text(id, text).unwrap();
    saved_main_part(&document.save().unwrap())
}

fn cell(properties: &str, content: &str) -> String {
    format!(
        r#"<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2400"/>{properties}</w:tcPr>{content}</w:tc>"#
    )
}

fn table(columns: usize, rows: &[String]) -> String {
    let grid = r#"<w:gridCol w:w="2400"/>"#.repeat(columns);
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{}</w:tbl>"#,
        rows.iter()
            .map(|row| format!("<w:tr>{row}</w:tr>"))
            .collect::<String>()
    )
}

#[test]
fn editing_before_a_vertically_merged_cell_keeps_the_body_bytes() {
    let rows = [
        [
            cell(
                r#"<w:vMerge w:val="restart"/>"#,
                &paragraph("0000004A", &run("A")),
            ),
            cell("", &paragraph("0000004B", &run("B"))),
        ]
        .concat(),
        [
            cell("<w:vMerge/>", "<w:p/>"),
            cell("", &paragraph("0000004C", &run("D"))),
        ]
        .concat(),
    ];
    let source = main_part(
        &[
            paragraph("0000004E", &run("Old text")),
            table(2, &rows),
            paragraph("0000004D", &run("After.")),
        ]
        .concat(),
    );
    let saved = edit_and_save(&source, "0000004E", "New text");
    assert_bytes_kept_around(&source, &saved, "0000004E", "New text");
}

struct Generator(u64);

impl Generator {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }

    fn below(&mut self, bound: usize) -> usize {
        (self.next() % bound as u64) as usize
    }
}

fn next_id(used: &mut u32) -> String {
    *used += 1;
    format!("{:08X}", 0x1000_0000 + *used)
}

struct GeneratedCell {
    start: usize,
    span: usize,
    merge: Option<&'static str>,
}

fn merged_table(generator: &mut Generator, used: &mut u32, depth: usize) -> String {
    let columns = 2 + generator.below(3);
    let row_count = 2 + generator.below(3);
    let mut layout: Vec<Vec<GeneratedCell>> = Vec::new();
    for row in 0..row_count {
        let mut cells = Vec::new();
        let mut start = 0;
        while start < columns {
            let span = if start + 1 < columns && generator.below(3) == 0 {
                2
            } else {
                1
            };
            let above = row
                .checked_sub(1)
                .and_then(|previous| {
                    layout[previous]
                        .iter()
                        .position(|cell| cell.start == start && cell.span == span)
                })
                .filter(|_| generator.below(2) == 0);
            if let Some(index) = above {
                let previous = &mut layout[row - 1][index];
                if previous.merge.is_none() {
                    previous.merge = Some("restart");
                }
            }
            cells.push(GeneratedCell {
                start,
                span,
                merge: above.map(|_| "continue"),
            });
            start += span;
        }
        layout.push(cells);
    }
    let mut rows = Vec::new();
    for cells in &layout {
        let mut row = String::new();
        for generated in cells {
            let mut properties = String::new();
            if generated.span > 1 {
                properties.push_str(&format!(r#"<w:gridSpan w:val="{}"/>"#, generated.span));
            }
            let content = match generated.merge {
                Some("continue") => {
                    properties.push_str("<w:vMerge/>");
                    if generator.below(2) == 0 {
                        "<w:p/>".to_owned()
                    } else {
                        paragraph(&next_id(used), &run("continued"))
                    }
                }
                merge => {
                    if merge == Some("restart") {
                        properties.push_str(r#"<w:vMerge w:val="restart"/>"#);
                    }
                    let text = paragraph(&next_id(used), &run("cell"));
                    if depth == 0 && generator.below(4) == 0 {
                        format!("{}{text}", merged_table(generator, used, depth + 1))
                    } else {
                        text
                    }
                }
            };
            row.push_str(&cell(&properties, &content));
        }
        rows.push(row);
    }
    table(columns, &rows)
}

fn generated_body(generator: &mut Generator, used: &mut u32, edited: &str) -> String {
    let mut blocks = vec![paragraph(edited, &run("Opening."))];
    for control in 0..2 + generator.below(2) {
        let table = merged_table(generator, used, 0);
        if control == 0 {
            blocks.push(format!(
                r#"<w:sdt><w:sdtPr><w:id w:val="{}"/></w:sdtPr><w:sdtContent>{table}</w:sdtContent></w:sdt>"#,
                100 + *used
            ));
        } else {
            blocks.push(table);
        }
        blocks.push(paragraph(&next_id(used), &run("Between.")));
    }
    blocks.concat()
}

#[test]
fn editing_beside_generated_merged_tables_keeps_the_body_bytes() {
    let mut generator = Generator(0x9E37_79B9_7F4A_7C15);
    for case in 0..64 {
        let mut used = 0;
        let source = main_part(&generated_body(&mut generator, &mut used, "0FFFFFFF"));
        let saved = edit_and_save(&source, "0FFFFFFF", "New text");
        let (start, end) = paragraph_span(&source, "0FFFFFFF");
        let (saved_start, saved_end) = paragraph_span(&saved, "0FFFFFFF");
        assert!(
            saved[saved_start..saved_end].contains("New text"),
            "case {case}"
        );
        assert_eq!(
            (&saved[..saved_start], &saved[saved_end..]),
            (&source[..start], &source[end..]),
            "source bytes differ outside the edited paragraph in case {case}"
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
