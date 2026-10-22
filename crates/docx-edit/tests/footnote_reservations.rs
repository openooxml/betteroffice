#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use serde_json::Value;

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#
);
const NARROW_PAGE: &str = r#"<w:pgSz w:w="8400" w:h="6000"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/>"#;
const WIDE_PAGE: &str = r#"<w:pgSz w:w="16838" w:h="11906"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/>"#;

/// Paragraphs of uneven length; every `every`th ends in a reference to a note of uneven
/// length, so reference lines keep landing near page bottoms.
struct NoteHeavy {
    paragraphs: usize,
    every: usize,
    min_words: usize,
    word_span: usize,
    page: &'static str,
}

impl NoteHeavy {
    fn notes(&self) -> usize {
        self.paragraphs.div_ceil(self.every)
    }

    fn docx(&self) -> Vec<u8> {
        let mut body = String::new();
        let mut note = 0;
        for index in 0..self.paragraphs {
            let reference = if index % self.every == 0 {
                note += 1;
                format!(
                    r#"<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{note}"/></w:r>"#
                )
            } else {
                String::new()
            };
            body.push_str(&format!(
                r#"<w:p w14:paraId="{:08X}"><w:r><w:t xml:space="preserve">{} </w:t></w:r>{reference}</w:p>"#,
                0x1000_0000 + index,
                words(self.min_words + (index * 37) % self.word_span, index),
            ));
        }
        body.push_str(&format!("<w:sectPr>{}</w:sectPr>", self.page));
        let mut notes = format!(
            concat!(
                r#"<w:footnotes {}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p>"#,
                r#"<w:p/></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r>"#,
                r#"<w:continuationSeparator/></w:r></w:p><w:p/></w:footnote>"#
            ),
            NS
        );
        for index in 0..self.notes() {
            notes.push_str(&format!(
                concat!(
                    r#"<w:footnote w:id="{}"><w:p w14:paraId="{:08X}"><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>"#,
                    r#"<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>"#,
                    r#"<w:r><w:t xml:space="preserve"> {}</w:t></w:r></w:p></w:footnote>"#
                ),
                index + 1,
                0x2000_0000 + index,
                words(4 + (index * 53) % 29, index + 7)
            ));
        }
        notes.push_str("</w:footnotes>");
        let mut parts = fixture::page_parts(false);
        for (name, bytes) in &mut parts {
            match name.as_str() {
                "word/document.xml" => {
                    *bytes = format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#)
                        .into_bytes();
                }
                "word/footnotes.xml" => *bytes = notes.clone().into_bytes(),
                "word/styles.xml" => *bytes = styles().into_bytes(),
                _ => {}
            }
        }
        ooxml_opc::rezip_parts(&parts).unwrap()
    }
}

fn words(count: usize, seed: usize) -> String {
    (0..count)
        .map(|index| format!("w{}x{}", seed % 97, index))
        .collect::<Vec<_>>()
        .join(" ")
}

fn styles() -> String {
    format!(
        concat!(
            r#"<w:styles {}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>"#,
            r#"<w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="100" "#,
            r#"w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>"#,
            r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>"#,
            r#"<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/>"#,
            r#"<w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="18"/></w:rPr></w:style>"#,
            r#"<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>"#,
            r#"<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>"#
        ),
        NS
    )
}

struct PageNotes {
    reserved: f64,
    area: f64,
    separator: f64,
    first_pm: Option<(f64, f64)>,
    notes: Vec<(f64, f64)>,
}

fn page_notes(page: &Value) -> PageNotes {
    let areas = page["noteAreas"].as_array().cloned().unwrap_or_default();
    PageNotes {
        reserved: page["footnoteReservedHeight"].as_f64().unwrap_or(0.0),
        area: areas
            .iter()
            .map(|area| area["height"].as_f64().unwrap())
            .sum(),
        separator: areas
            .iter()
            .map(|area| area["separator"]["height"].as_f64().unwrap_or(0.0))
            .sum(),
        first_pm: page["fragments"]
            .as_array()
            .and_then(|fragments| fragments.first())
            .and_then(|fragment| {
                Some((fragment["pmStart"].as_f64()?, fragment["pmEnd"].as_f64()?))
            }),
        notes: areas
            .iter()
            .flat_map(|area| area["notes"].as_array().cloned().unwrap_or_default())
            .map(|note| {
                (
                    note["anchorDocStart"].as_f64().unwrap_or(-1.0),
                    note["height"].as_f64().unwrap(),
                )
            })
            .collect(),
    }
}

fn laid_out_notes(document: &NoteHeavy, client_id: u64) -> Vec<PageNotes> {
    let (engine, request) = fixture::laid_out(&document.docx(), client_id);
    let layout: Value =
        serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap()).unwrap();
    let pages: Vec<PageNotes> = layout["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(page_notes)
        .collect();
    let placed: usize = pages.iter().map(|page| page.notes.len()).sum();
    assert_eq!(placed, document.notes(), "every note is placed once");
    pages
}

#[test]
fn note_space_is_reserved_only_for_the_page_notes_or_the_reference_line_it_pushes_on() {
    let pages = laid_out_notes(
        &NoteHeavy {
            paragraphs: 120,
            every: 2,
            min_words: 30,
            word_span: 110,
            page: WIDE_PAGE,
        },
        31,
    );
    for (index, page) in pages.iter().enumerate() {
        let surplus = page.reserved - page.area;
        if surplus <= 0.5 {
            continue;
        }
        let next = &pages[index + 1];
        let (start, end) = next.first_pm.unwrap();
        let pushed: f64 = next
            .notes
            .iter()
            .filter(|(anchor, _)| (start..end).contains(anchor))
            .map(|(_, height)| height)
            .sum();
        assert!(
            pushed > 0.0 && surplus <= pushed + next.separator + 0.5,
            "page {} reserves {:.1} for {:.1} of notes; the next page opens with {:.1} of notes",
            index + 1,
            page.reserved,
            page.area,
            pushed
        );
    }
}

#[test]
fn unsettled_notes_do_not_inflate_the_page_count() {
    let pages = laid_out_notes(
        &NoteHeavy {
            paragraphs: 60,
            every: 3,
            min_words: 20,
            word_span: 60,
            page: NARROW_PAGE,
        },
        37,
    );
    assert!(pages.len() <= 41, "{} pages", pages.len());
}

/// Endnotes at the document end that together outgrow a page.
fn long_endnotes_docx() -> Vec<u8> {
    let mut body = String::new();
    for index in 0..24 {
        body.push_str(&format!(
            concat!(
                r#"<w:p w14:paraId="{:08X}"><w:r><w:t xml:space="preserve">{} </w:t></w:r>"#,
                r#"<w:r><w:endnoteReference w:id="{}"/></w:r></w:p>"#
            ),
            0x3000_0000 + index,
            words(25 + (index * 31) % 50, index),
            index + 1
        ));
    }
    body.push_str(&format!("<w:sectPr>{NARROW_PAGE}</w:sectPr>"));
    let mut notes = format!(
        concat!(
            r#"<w:endnotes {}><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p>"#,
            r#"</w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r>"#,
            r#"<w:continuationSeparator/></w:r></w:p></w:endnote>"#
        ),
        NS
    );
    for index in 0..24 {
        notes.push_str(&format!(
            r#"<w:endnote w:id="{}"><w:p w14:paraId="{:08X}"><w:r><w:t xml:space="preserve">{}</w:t></w:r></w:p></w:endnote>"#,
            index + 1,
            0x4000_0000 + index,
            words(30 + (index * 17) % 40, index + 3)
        ));
    }
    notes.push_str("</w:endnotes>");
    let mut parts = fixture::page_parts(false);
    for (name, bytes) in &mut parts {
        match name.as_str() {
            "word/document.xml" => {
                *bytes = format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#)
                    .into_bytes();
            }
            "word/endnotes.xml" => *bytes = notes.clone().into_bytes(),
            "word/styles.xml" => *bytes = styles().into_bytes(),
            _ => {}
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn endnotes_taller_than_a_page_do_not_push_out_blank_pages() {
    let (engine, request) = fixture::laid_out(&long_endnotes_docx(), 41);
    let layout: Value =
        serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap()).unwrap();
    let pages = layout["layout"]["pages"].as_array().unwrap();
    let blank = pages
        .iter()
        .filter(|page| page["fragments"].as_array().is_none_or(Vec::is_empty))
        .count();
    assert_eq!(blank, 0);
    assert!(pages.len() <= 17, "{} pages", pages.len());
}
