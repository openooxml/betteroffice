//! A TOC entry takes its line height from its entry text at the TOC style
//! size. Word renders the leader tab and the PAGEREF page number at that
//! size even when the tab run, the field result run or the paragraph mark
//! carries a larger one, and it does not apply a Hyperlink character style
//! size to entry text inside the TOC field result. Measured against the
//! Word reference PDFs of the corpus TOC documents (Arial 10pt entries
//! whose tab and mark runs carry 11 to 12pt sit on a 10pt line).

use docx_edit::{EditCtx, EngineSession, ParaSelector, Position, StyleResolver, seed_from_docx};
use serde_json::{Value, json};
use yrs::Any;

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

/// Liberation Sans single spacing pitch at the given size: the hhea pitch
/// is 2355 of 2048 units, so 10pt measures 15.332px at 96dpi.
fn single(size_pt: f64) -> f64 {
    size_pt * 96.0 / 72.0 * 2355.0 / 2048.0
}

fn document(body: &str, styles: &str) -> Vec<u8> {
    let section = r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#;
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{body}{section}</w:body></w:document>"#)),
        ("word/styles.xml", styles.to_owned()),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn styles(doc_defaults_size: Option<u32>, hyperlink_size: Option<u32>) -> String {
    let doc_size = doc_defaults_size
        .map(|size| format!(r#"<w:sz w:val="{size}"/>"#))
        .unwrap_or_default();
    let link_size = hyperlink_size
        .map(|size| {
            format!(r#"<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="{size}"/></w:rPr>"#)
        })
        .unwrap_or_else(|| {
            r#"<w:rPr><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr>"#.to_owned()
        });
    r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>{doc_size}</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="0"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>{link_size}</w:style></w:styles>"#.to_owned()
}

/// One Word shaped TOC entry: entry text, a leader tab, a page number.
fn entry(text: &str, tab_size: Option<u32>, mark_size: Option<u32>) -> String {
    let tab_rpr = tab_size
        .map(|size| format!(r#"<w:rPr><w:sz w:val="{size}"/></w:rPr>"#))
        .unwrap_or_default();
    let mark = mark_size
        .map(|size| format!(r#"<w:rPr><w:sz w:val="{size}"/></w:rPr>"#))
        .unwrap_or_default();
    r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/>{mark}</w:pPr><w:hyperlink w:anchor="_Toc1"><w:r><w:t>{text}</w:t></w:r><w:r>{tab_rpr}<w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:hyperlink></w:p>"#.to_owned()
}

fn line_heights(bytes: &[u8]) -> Vec<f64> {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(77100);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(
                &json!({
                    "bodyStory": "body", "renderEnv": {},
                    "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
                    "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 10}}
                })
                .to_string(),
            )
            .unwrap(),
    )
    .unwrap();
    output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .filter_map(|block| {
            let has_text = block["block"]["runs"].as_array().is_some_and(|runs| {
                runs.iter().any(|run| {
                    run["text"].as_str().is_some_and(|text| !text.is_empty())
                })
            });
            if !has_text {
                return None;
            }
            let lines = block["measure"]["lines"].as_array()?;
            if lines.is_empty() {
                return None;
            }
            assert_eq!(lines.len(), 1, "entry must stay on one line: {block}");
            lines[0]["lineHeight"].as_f64()
        })
        .collect()
}

fn assert_height(actual: f64, expected_pt: f64) {
    let expected = single(expected_pt);
    assert!(
        (actual - expected).abs() < 0.05,
        "line height {actual:.3}px, Word {expected:.3}px ({expected_pt}pt)"
    );
}

#[test]
fn a_larger_tab_run_does_not_raise_a_toc_entry() {
    // Entry text at the 10pt document default, leader tab run at 14pt.
    let body = entry("Entry one", Some(28), None);
    let heights = line_heights(&document(&body, &styles(Some(20), None)));
    assert_height(heights[0], 10.0);
}

#[test]
fn a_larger_paragraph_mark_does_not_raise_a_toc_entry() {
    // No size in the style chain at all: the mark alone carries 14pt and
    // the entry text falls back to the 10pt default.
    let body = entry("Entry one", None, Some(28));
    let heights = line_heights(&document(&body, &styles(None, None)));
    assert_height(heights[0], 10.0);
}

#[test]
fn a_hyperlink_style_size_does_not_raise_a_toc_entry() {
    // Hyperlink character style at 14pt over a 10pt entry, the shape the
    // Verdana 12pt Hyperlink style takes over 10pt TOC styles in Word.
    let body = r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Entry one</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:hyperlink></w:p>"#.to_owned();
    let heights = line_heights(&document(&body, &styles(Some(20), Some(28))));
    assert_height(heights[0], 10.0);
}

#[test]
fn a_larger_field_result_does_not_raise_a_toc_entry() {
    let body = r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Toc1"><w:r><w:t>Entry one</w:t></w:r><w:r><w:tab/></w:r><w:fldSimple w:instr=" PAGEREF _Toc1 \h "><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>3</w:t></w:r></w:fldSimple></w:hyperlink></w:p>"#.to_owned();
    let heights = line_heights(&document(&body, &styles(Some(20), None)));
    assert_height(heights[0], 10.0);
}

#[test]
fn toc_entries_stay_at_the_entry_size_down_a_page() {
    let body = (0..6)
        .map(|index| entry(&format!("Entry {index}"), Some(28), Some(24)))
        .collect::<String>();
    let heights = line_heights(&document(&body, &styles(Some(20), None)));
    assert_eq!(heights.len(), 6);
    for height in heights {
        assert_height(height, 10.0);
    }
}

/// Control: tab and text at the same size already agree with Word.
#[test]
fn a_tab_at_the_entry_size_keeps_its_line_height() {
    let body = entry("Entry one", Some(24), None);
    let heights = line_heights(&document(&body, &styles(Some(24), None)));
    assert_height(heights[0], 12.0);
}

/// Control: larger entry text still drives the line; only the auxiliary
/// runs are capped, so a 14pt entry keeps its 14pt line.
#[test]
fn larger_entry_text_still_drives_a_toc_line() {
    let body = r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>Entry one</w:t></w:r><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:hyperlink></w:p>"#.to_owned();
    let heights = line_heights(&document(&body, &styles(Some(20), None)));
    assert_height(heights[0], 14.0);
}

/// Control: outside a TOC entry a larger tab run keeps the taller line it
/// has on main; the cap is scoped to TOC entries.
#[test]
fn a_larger_tab_run_still_raises_a_body_line() {
    let body = r#"<w:p><w:r><w:t>Body text</w:t></w:r><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:p>"#.to_owned();
    let heights = line_heights(&document(&body, &styles(Some(20), None)));
    assert_height(heights[0], 14.0);
}

/// Body control: outside a TOC field result Word applies a Hyperlink
/// character style size, so a body hyperlink at 14pt keeps its 14pt line
/// on main and after the fix.
#[test]
fn a_hyperlink_style_size_still_applies_in_body_text() {
    let body = r#"<w:p><w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Body link text</w:t></w:r></w:hyperlink></w:p>"#.to_owned();
    let heights = line_heights(&document(&body, &styles(Some(20), Some(28))));
    assert_height(heights[0], 14.0);
}

/// Body red test: a larger paragraph mark does not size body runs
/// either. Word takes unsized runs at the document default, so this
/// fails on main, where the mark size fills the runs.
#[test]
fn a_larger_paragraph_mark_does_not_raise_a_body_paragraph() {
    let body = r#"<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:r><w:t>Body text</w:t></w:r></w:p>"#.to_owned();
    let heights = line_heights(&document(&body, &styles(None, None)));
    assert_height(heights[0], 10.0);
}

/// Control: a plain paragraph at the default size is unchanged.
#[test]
fn a_plain_paragraph_keeps_its_line_height() {
    let body = r#"<w:p><w:r><w:t>Body text</w:t></w:r></w:p>"#.to_owned();
    let heights = line_heights(&document(&body, &styles(Some(20), None)));
    assert_height(heights[0], 10.0);
}

const CHAIN_STYLES: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="32"/></w:rPr></w:style></w:styles>"#;

/// A Normal paragraph whose mark carries a direct 14pt over the 10pt
/// chain, holding one unsized run.
fn marked_paragraph() -> (Vec<u8>, StyleResolver) {
    let body = r#"<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="Normal"/><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:r><w:t>Switch me</w:t></w:r></w:p>"#;
    let bytes = document(body, CHAIN_STYLES);
    let parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let styles = StyleResolver::from_parts(&parts).unwrap();
    (bytes, styles)
}

fn properties(
    doc: &docx_edit::EditingDoc,
    para_id: &str,
) -> std::collections::BTreeMap<String, Any> {
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .find(|paragraph| paragraph.para_id == para_id)
        .unwrap()
        .properties
}

/// The seed chain defaults must not survive a style switch: after
/// switching to Heading1 the unsized run takes the new 16pt chain, not
/// the stale 10pt one and not the 14pt mark.
#[test]
fn a_style_switch_drops_the_seed_chain_defaults() {
    let (bytes, styles) = marked_paragraph();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(77101);
    let doc = engine.doc();
    seed_from_docx(doc, &bytes).unwrap();
    let before = properties(doc, "00000001");
    assert!(before.contains_key("runDefaultTextFormatting"));

    doc.restyle_paragraphs(
        &EditCtx::local("", ""),
        &ParaSelector::One("00000001".to_owned()),
        "Heading1",
        &styles,
    )
    .unwrap();

    let after = properties(doc, "00000001");
    assert!(!after.contains_key("runDefaultTextFormatting"));
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(
                &json!({
                    "bodyStory": "body", "renderEnv": {},
                    "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
                    "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 10}}
                })
                .to_string(),
            )
            .unwrap(),
    )
    .unwrap();
    let blocks = output["measured"].as_array().unwrap();
    let run = &blocks[0]["block"]["runs"][0];
    assert_eq!(run["fontSize"], 16.0, "run must take the Heading1 chain");
}

/// A split at the end rebuilds the second paragraph from the caret
/// formatting: it carries no seed chain defaults.
#[test]
fn a_split_at_the_end_drops_the_seed_chain_defaults() {
    let (bytes, _styles) = marked_paragraph();
    let doc = docx_edit::EditingDoc::new(77102);
    seed_from_docx(&doc, &bytes).unwrap();
    doc.split_paragraph(&EditCtx::local("", ""), Position::new("body", 9), None)
        .unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs.len(), 2);
    assert!(!paragraphs[1]
        .properties
        .contains_key("runDefaultTextFormatting"));
}

/// A save round trip carries no trace of the seed owned key, and the
/// unchanged document part stays byte identical to the source.
#[test]
fn a_save_round_trip_never_writes_the_seed_chain_defaults() {
    let (bytes, _styles) = marked_paragraph();
    let doc = docx_edit::EditingDoc::new(77103);
    seed_from_docx(&doc, &bytes).unwrap();
    let plan = doc.paragraph_save_plan();
    let package = docx_parse::parse_docx_s9_wire(
        &bytes,
        docx_parse::S9ParseOptions {
            source_ordinals: true,
            ..docx_parse::S9ParseOptions::default()
        },
    )
    .unwrap()
    .document
    .package;
    let request: docx_parse::serializer::S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {
            "content": package.document.content,
            "comments": package.document.comments,
        },
        "headerEntries": package.header_entries.unwrap_or_default(),
        "footnotes": package.footnotes.unwrap_or_default(),
        "footnoteSeparators": package.footnote_separators.unwrap_or_default(),
        "relationshipEntries": package.relationship_entries,
        "options": {"updateModifiedDate": false},
        "paragraphIds": {
            "assignments": plan.assignments.iter().map(|(part, ordinal, para_id)| {
                json!({"part": part, "ordinal": ordinal, "paraId": para_id})
            }).collect::<Vec<_>>(),
            "patchedParts": plan.patched_parts.iter().map(|(part, para_ids)| {
                json!({"part": part, "paraIds": para_ids})
            }).collect::<Vec<_>>(),
        },
    }))
    .unwrap();
    let saved =
        ooxml_opc::unzip_parts(&docx_parse::serializer::write_docx_s13(request, &bytes).unwrap())
            .unwrap();
    let source = ooxml_opc::unzip_parts(&bytes).unwrap();
    let saved_xml = saved
        .iter()
        .find(|(name, _)| name == "word/document.xml")
        .map(|(_, value)| value.clone())
        .unwrap();
    let source_xml = source
        .iter()
        .find(|(name, _)| name == "word/document.xml")
        .map(|(_, value)| value.clone())
        .unwrap();
    let text = String::from_utf8(saved_xml.clone()).unwrap();
    assert!(!text.contains("runDefaultTextFormatting"));
    assert_eq!(saved_xml, source_xml);
}
