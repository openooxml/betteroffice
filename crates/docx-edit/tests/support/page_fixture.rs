//! The synthetic multi-page package the page-fragment tests lay out: small pages, a paragraph
//! split across pages with an emoji, tabs, atoms and a literal U+FFFC, a table with a repeated
//! header row, a vertical merge and a row split across pages with a footnote reference in its
//! first part, four sections numbered in Roman, restarted decimal, continued letters and an
//! unsupported format behind an odd-page parity filler, first-page and running headers, a footer,
//! footnotes and an endnote.

use std::path::PathBuf;

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

pub const FONT: &[u8] =
    include_bytes!("../../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
/// A second font with other metrics.
pub const OTHER_FONT: &[u8] =
    include_bytes!("../../../docx-raster/tests/assets/Carlito-Regular.ttf");

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#
);
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml";

pub fn fixture_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/page-fragments")
}

pub fn p(id: &str, content: &str) -> String {
    format!(r#"<w:p w14:paraId="{id}">{content}</w:p>"#)
}

pub fn r(text: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

fn words(count: usize, seed: &str) -> String {
    (0..count)
        .map(|index| format!("{seed}{index}"))
        .collect::<Vec<_>>()
        .join(" ")
}

fn cell(properties: &str, content: &str) -> String {
    format!(
        r#"<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/>{properties}</w:tcPr>{content}</w:tc>"#
    )
}

fn section(extra: &str) -> String {
    format!(
        r#"<w:sectPr>{extra}<w:pgSz w:w="7200" w:h="5760"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/></w:sectPr>"#
    )
}

fn body(revisions: bool) -> String {
    let split = format!(
        "{} {}\t{} \u{FFFC} {} {} {}",
        words(60, "alpha"),
        words(40, "beta"),
        words(10, "gamma"),
        words(19, "delta"),
        "\u{1F600}".repeat(40),
        (19..60)
            .map(|index| format!("delta{index}"))
            .collect::<Vec<_>>()
            .join(" ")
    );
    let atoms = format!(
        r#"{}<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>{}<w:r><w:br/></w:r>{}<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>{}<w:sdt><w:sdtPr><w:tag w:val="name"/><w:id w:val="7"/><w:text/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>{}"#,
        r("Page field "),
        r(" then a break"),
        r("after the break and a note"),
        r(" and a control "),
        r("Ada Lovelace"),
        r(" done.")
    );
    let tall = (0..24)
        .map(|index| {
            p(
                &format!("300000{index:02X}"),
                &r(&format!("Tall cell line {index}")),
            )
        })
        .collect::<String>();
    let note_in_split = p(
        "30000100",
        &format!(
            r#"{}<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r>"#,
            r("Early reference")
        ),
    );
    let rows = (0..6)
        .map(|index| {
            format!(
                "<w:tr>{}{}</w:tr>",
                cell(
                    "",
                    &p(
                        &format!("310000{index:02X}"),
                        &r(&format!("Row {index} left"))
                    )
                ),
                cell(
                    "",
                    &p(
                        &format!("320000{index:02X}"),
                        &r(&format!("Row {index} right"))
                    )
                )
            )
        })
        .collect::<String>();
    let table = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr>{}{}</w:tr><w:tr>{}{}</w:tr><w:tr>{}{}</w:tr>{rows}<w:tr>{}{}</w:tr></w:tbl>"#,
        cell("", &p("30000A01", &r("Header left"))),
        cell("", &p("30000A02", &r("Header right"))),
        cell(
            r#"<w:vMerge w:val="restart"/>"#,
            &p("30000A03", &r("Merged"))
        ),
        cell("", &p("30000A04", &r("Beside merge"))),
        cell(r#"<w:vMerge/>"#, &p("30000A05", &r(""))),
        cell("", &p("30000A06", &r("Below beside"))),
        cell("", &format!("{note_in_split}{tall}")),
        cell("", &p("30000A07", &r("Short neighbour"))),
    );
    let revised = if revisions {
        format!(
            r#"{}<w:ins w:id="41" w:author="Ann" w:date="2026-01-02T00:00:00Z">{}</w:ins><w:del w:id="42" w:author="Bob" w:date="2026-01-03T00:00:00Z"><w:r><w:delText xml:space="preserve">removed words </w:delText></w:r></w:del>{}"#,
            r("Kept "),
            r("inserted "),
            r("tail")
        )
    } else {
        r("Kept tail")
    };
    [
        format!(
            r#"<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>{}</w:p>"#,
            r("Page map")
        ),
        p("00000002", &r(&split)),
        p("00000003", &atoms),
        table,
        format!(
            r#"<w:p w14:paraId="00000004"><w:pPr>{}</w:pPr>{}</w:p>"#,
            section(
                r#"<w:headerReference w:type="default" r:id="rIdHeader2"/><w:headerReference w:type="first" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/><w:pgNumType w:fmt="lowerRoman" w:start="1"/><w:titlePg/>"#
            ),
            r("End of the Roman section")
        ),
        p("00000005", &r(&words(90, "epsilon"))),
        format!(
            r#"<w:p w14:paraId="00000006"><w:pPr>{}</w:pPr>{}<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="3"/></w:r></w:p>"#,
            section(
                r#"<w:type w:val="nextPage"/><w:pgNumType w:start="1"/><w:titlePg/>"#
            ),
            r("Restarted decimal section with a note")
        ),
        p("00000007", &r(&words(30, "zeta"))),
        format!(
            r#"<w:p w14:paraId="00000008"><w:pPr>{}</w:pPr>{}</w:p>"#,
            section(r#"<w:type w:val="oddPage"/><w:pgNumType w:fmt="upperLetter"/>"#),
            r("Continued letters")
        ),
        p("00000009", &revised),
        p(
            "0000000A",
            &format!(
                r#"{}<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="1"/></w:r>"#,
                r("Closing words with an endnote")
            ),
        ),
        section(
            r#"<w:type w:val="nextPage"/><w:pgNumType w:fmt="chineseCounting"/>"#,
        ),
    ]
    .concat()
}

fn styles() -> String {
    format!(
        r#"<w:styles {NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style></w:styles>"#
    )
}

/// The fixture's parts, in archive order.
pub fn page_parts(revisions: bool) -> Vec<(String, Vec<u8>)> {
    let rel = |id: &str, kind: &str, target: &str| {
        format!(r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{target}"/>"#)
    };
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="{OFFICE}.styles+xml"/><Override PartName="/word/header1.xml" ContentType="{OFFICE}.header+xml"/><Override PartName="/word/header2.xml" ContentType="{OFFICE}.header+xml"/><Override PartName="/word/footer1.xml" ContentType="{OFFICE}.footer+xml"/><Override PartName="/word/footnotes.xml" ContentType="{OFFICE}.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="{OFFICE}.endnotes+xml"/></Types>"#
    );
    let root_rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>"#,
        rel("rId1", "officeDocument", "word/document.xml")
    );
    let document_rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}{}{}{}{}{}</Relationships>"#,
        rel("rIdStyles", "styles", "styles.xml"),
        rel("rIdHeader1", "header", "header1.xml"),
        rel("rIdHeader2", "header", "header2.xml"),
        rel("rIdFooter1", "footer", "footer1.xml"),
        rel("rIdFootnotes", "footnotes", "footnotes.xml"),
        rel("rIdEndnotes", "endnotes", "endnotes.xml"),
    );
    let document = format!(
        r#"<w:document {NS}><w:body>{}</w:body></w:document>"#,
        body(revisions)
    );
    let story = |root: &str, content: &str| format!(r#"<w:{root} {NS}>{content}</w:{root}>"#);
    let notes = |root: &str, note: &str, bodies: &[(u32, &str, &str)]| {
        let mut xml = format!(
            r#"<w:{root} {NS}><w:{note} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:{note}><w:{note} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{note}>"#
        );
        for (id, para_id, text) in bodies {
            xml.push_str(&format!(
                r#"<w:{note} w:id="{id}">{}</w:{note}>"#,
                p(para_id, &r(text))
            ));
        }
        xml.push_str(&format!("</w:{root}>"));
        xml
    };
    [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root_rels),
        ("word/_rels/document.xml.rels", document_rels),
        ("word/document.xml", document),
        ("word/styles.xml", styles()),
        (
            "word/header1.xml",
            story("hdr", &p("20000001", &r("First page header"))),
        ),
        (
            "word/header2.xml",
            story(
                "hdr",
                &p(
                    "20000002",
                    &format!(
                        r#"{}<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
                        r("Running header ")
                    ),
                ),
            ),
        ),
        (
            "word/footer1.xml",
            story("ftr", &p("20000003", &r("Footer text"))),
        ),
        (
            "word/footnotes.xml",
            notes(
                "footnotes",
                "footnote",
                &[
                    (1, "40000001", "First footnote"),
                    (2, "40000002", "Footnote from a split row"),
                    (3, "40000003", "Third footnote"),
                    (4, "40000005", "Orphan footnote"),
                ],
            ),
        ),
        (
            "word/endnotes.xml",
            notes("endnotes", "endnote", &[(1, "40000004", "The endnote")]),
        ),
    ]
    .into_iter()
    .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
    .collect()
}

/// The committed fixture, rewritten from [`page_parts`] when `PAGE_FRAGMENTS_UPDATE=1`.
pub fn pages_docx() -> Vec<u8> {
    let path = fixture_dir().join("pages.docx");
    if std::env::var("PAGE_FRAGMENTS_UPDATE").as_deref() == Ok("1") {
        let bytes = ooxml_opc::rezip_parts(&page_parts(true)).unwrap();
        std::fs::create_dir_all(fixture_dir()).unwrap();
        std::fs::write(&path, &bytes).unwrap();
        return bytes;
    }
    std::fs::read(&path).unwrap()
}

/// A package holding `body` with the fixture's styles.
pub fn with_body(body: &str) -> Vec<u8> {
    let mut parts = page_parts(false);
    for (name, bytes) in &mut parts {
        if name == "word/document.xml" {
            *bytes =
                format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#).into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// `parts` with `settings` as the package's settings part.
fn with_settings(mut parts: Vec<(String, Vec<u8>)>, settings: &str) -> Vec<(String, Vec<u8>)> {
    for (name, bytes) in &mut parts {
        let mut xml = String::from_utf8(std::mem::take(bytes)).unwrap();
        match name.as_str() {
            "word/_rels/document.xml.rels" => {
                xml = xml.replacen(
                    "</Relationships>",
                    &format!(
                        r#"<Relationship Id="rIdSettings" Type="{REL}/settings" Target="settings.xml"/></Relationships>"#
                    ),
                    1,
                );
            }
            "[Content_Types].xml" => {
                xml = xml.replacen(
                    "</Types>",
                    &format!(
                        r#"<Override PartName="/word/settings.xml" ContentType="{OFFICE}.settings+xml"/></Types>"#
                    ),
                    1,
                );
            }
            _ => {}
        }
        *bytes = xml.into_bytes();
    }
    parts.push((
        "word/settings.xml".to_owned(),
        format!(r#"<w:settings {NS}>{settings}</w:settings>"#).into_bytes(),
    ));
    parts
}

/// `parts` with `from` replaced by `to` in the part `name`.
fn replaced(
    mut parts: Vec<(String, Vec<u8>)>,
    name: &str,
    from: &str,
    to: &str,
) -> Vec<(String, Vec<u8>)> {
    for (part, bytes) in &mut parts {
        if part == name {
            let xml = String::from_utf8(std::mem::take(bytes)).unwrap();
            assert!(xml.contains(from), "{name} holds {from}");
            *bytes = xml.replacen(from, to, 1).into_bytes();
        }
    }
    parts
}

/// The fixture without tracked changes, with odd and even headers: the first section's even
/// pages show `rIdHeader1`.
pub fn even_headers_docx() -> Vec<u8> {
    let parts = replaced(
        page_parts(false),
        "word/document.xml",
        r#"<w:headerReference w:type="first" r:id="rIdHeader1"/>"#,
        r#"<w:headerReference w:type="even" r:id="rIdHeader1"/>"#,
    );
    ooxml_opc::rezip_parts(&with_settings(parts, "<w:evenAndOddHeaders/>")).unwrap()
}

/// The fixture without tracked changes whose first section shows `rIdHeader2` both as its
/// first-page and as its running header.
pub fn shared_header_docx() -> Vec<u8> {
    let parts = replaced(
        page_parts(false),
        "word/document.xml",
        r#"<w:headerReference w:type="first" r:id="rIdHeader1"/>"#,
        r#"<w:headerReference w:type="first" r:id="rIdHeader2"/>"#,
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The fixture without tracked changes with endnotes placed at the end of each section.
pub fn section_endnotes_docx() -> Vec<u8> {
    ooxml_opc::rezip_parts(&with_settings(
        page_parts(false),
        r#"<w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr>"#,
    ))
    .unwrap()
}

/// The fixture's parts with a three-page body whose first page shows `rIdHeader1` and whose
/// others show `rIdHeader2`, which holds `header`.
fn headed_pages(header: &str) -> Vec<(String, Vec<u8>)> {
    let pages = format!(
        "{}{}",
        (0..3)
            .map(|index| p(&format!("0000000{index}"), &r(&words(120, "filler"))))
            .collect::<String>(),
        section(
            r#"<w:headerReference w:type="default" r:id="rIdHeader2"/><w:headerReference w:type="first" r:id="rIdHeader1"/><w:titlePg/>"#
        )
    );
    let mut parts = replaced(
        page_parts(false),
        "word/document.xml",
        &format!("<w:body>{}</w:body>", body(false)),
        &format!("<w:body>{pages}</w:body>"),
    );
    for (name, bytes) in &mut parts {
        if name == "word/header2.xml" {
            *bytes = format!(r#"<w:hdr {NS}>{header}</w:hdr>"#).into_bytes();
        }
    }
    parts
}

/// [`headed_pages`] whose two header relationships name one part holding a paragraph, a table
/// and a content control, so the export merges the two stories into one.
pub fn aliased_headers_docx() -> Vec<u8> {
    let header = format!(
        r#"{}<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr>{}</w:tr></w:tbl><w:sdt><w:sdtPr><w:tag w:val="banner"/><w:id w:val="5"/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>"#,
        p("22000001", &r("Shared header")),
        cell("", &p("22000002", &r("Header cell"))),
        p("22000003", &r("Header control")),
    );
    let parts = replaced(
        headed_pages(&header),
        "word/_rels/document.xml.rels",
        r#"Target="header1.xml""#,
        r#"Target="header2.xml""#,
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// [`headed_pages`] whose running header is a table row of exact height holding more lines than
/// it shows.
pub fn clipped_header_docx() -> Vec<u8> {
    let lines: String = (0..4)
        .map(|index| {
            p(
                &format!("2300000{index}"),
                &r(&format!("Clipped line {index}")),
            )
        })
        .collect();
    let header = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="300" w:hRule="exact"/></w:trPr>{}</w:tr></w:tbl>{}"#,
        cell("", &lines),
        p("23000010", &r("")),
    );
    ooxml_opc::rezip_parts(&headed_pages(&header)).unwrap()
}

/// [`headed_pages`] with its running header holding `header`.
pub fn headed_docx(header: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&headed_pages(header)).unwrap()
}

/// The fixture without tracked changes whose third section, lettered on odd pages, has been
/// merged into the final one.
pub fn without_third_section_docx() -> Vec<u8> {
    let third = format!(
        "<w:pPr>{}</w:pPr>",
        section(r#"<w:type w:val="oddPage"/><w:pgNumType w:fmt="upperLetter"/>"#)
    );
    ooxml_opc::rezip_parts(&replaced(
        page_parts(false),
        "word/document.xml",
        &third,
        "<w:pPr></w:pPr>",
    ))
    .unwrap()
}

/// A package holding `content` whose styles give text no font family.
pub fn without_default_fonts(content: &str) -> Vec<u8> {
    let parts = replaced(
        page_parts(false),
        "word/styles.xml",
        r#"<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>"#,
        "",
    );
    ooxml_opc::rezip_parts(&replaced(
        parts,
        "word/document.xml",
        &format!("<w:body>{}</w:body>", body(false)),
        &format!("<w:body>{content}</w:body>"),
    ))
    .unwrap()
}

/// A package holding `body` and a first footnote whose content is `note`.
pub fn with_body_and_note(body: &str, note: &str) -> Vec<u8> {
    let mut parts = page_parts(false);
    for (name, bytes) in &mut parts {
        if name == "word/document.xml" {
            *bytes =
                format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#).into_bytes();
        }
        if name == "word/footnotes.xml" {
            let xml = String::from_utf8(std::mem::take(bytes)).unwrap();
            let start = xml.find(r#"<w:footnote w:id="1">"#).unwrap();
            let end = start + xml[start..].find("</w:footnote>").unwrap();
            *bytes = format!(
                r#"{}<w:footnote w:id="1">{note}{}"#,
                &xml[..start],
                &xml[end..]
            )
            .into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The fixture without its tracked changes.
pub fn unrevised_docx() -> Vec<u8> {
    ooxml_opc::rezip_parts(&page_parts(false)).unwrap()
}

/// The region layout request a host builds for `bytes`: every section with its properties,
/// every normal note, the body story, and every font requirement measured with `font`.
pub fn region_request(engine: &EngineSession, bytes: &[u8], font: u32) -> Value {
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let mut sections: Vec<Value> = package
        .document
        .sections
        .clone()
        .unwrap_or_default()
        .into_iter()
        .map(|section| json!({"sectionId": section.id, "properties": section.properties}))
        .collect();
    let last = sections
        .last()
        .map(|section| section["properties"].clone())
        .unwrap_or_else(|| json!(package.document.final_section_properties));
    sections.push(json!({"properties": last}));
    let mut contents = Vec::new();
    for (notes, kind) in [
        (&package.footnotes, "footnote"),
        (&package.endnotes, "endnote"),
    ] {
        for note in notes.iter().flatten() {
            if note.note_type.is_empty() || note.note_type == "normal" {
                contents.push(json!({"id": note.id as i64, "noteKind": kind, "height": 0}));
            }
        }
    }
    let mut request = json!({
        "bodyStory": "body",
        "renderEnv": {},
        "options": {"pageGap": 24},
        "regions": {"sections": sections, "settings": package.settings},
        "notes": {"contents": contents},
    });
    let requirements: Vec<Value> = serde_json::from_str(
        &engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    let chains: serde_json::Map<String, Value> = requirements
        .iter()
        .map(|requirement| {
            (
                requirement["key"].as_str().unwrap().to_owned(),
                json!([font]),
            )
        })
        .collect();
    request["measurement"] = json!({
        "fontChains": chains,
        "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
        "compat": {"noLeading": false, "doNotExpandShiftReturn": false},
        "authoritativeShaping": true,
    });
    request
}

/// A session holding `bytes`, laid out with the pinned font.
pub fn laid_out(bytes: &[u8], client_id: u64) -> (EngineSession, String) {
    laid_out_with(bytes, client_id, FONT)
}

/// A session holding `bytes`, laid out with `font` alone.
pub fn laid_out_with(bytes: &[u8], client_id: u64, font: &[u8]) -> (EngineSession, String) {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(font).unwrap();
    let engine = EngineSession::new(client_id);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let request = region_request(&engine, bytes, font).to_string();
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    (engine, request)
}
