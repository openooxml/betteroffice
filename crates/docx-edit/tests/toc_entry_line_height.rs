//! TOC entry line heights and paragraph style font sizes.

use docx_edit::{
    EditCtx, EngineSession, FormatPolicy, ParaSelector, Position, StyleResolver, seed_from_docx,
};
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
    format!(
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>{doc_size}</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="0"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>{link_size}</w:style></w:styles>"#
    )
}

/// One Word shaped TOC entry: entry text, a leader tab, a page number.
fn entry(text: &str, tab_size: Option<u32>, mark_size: Option<u32>) -> String {
    let tab_rpr = tab_size
        .map(|size| format!(r#"<w:rPr><w:sz w:val="{size}"/></w:rPr>"#))
        .unwrap_or_default();
    let mark = mark_size
        .map(|size| format!(r#"<w:rPr><w:sz w:val="{size}"/></w:rPr>"#))
        .unwrap_or_default();
    format!(
        r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/>{mark}</w:pPr><w:hyperlink w:anchor="_Toc1"><w:r><w:t>{text}</w:t></w:r><w:r>{tab_rpr}<w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:hyperlink></w:p>"#
    )
}

fn chain_env(bytes: &[u8]) -> Value {
    let definitions = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package
        .styles
        .unwrap_or_default();
    let doc_default_half = definitions
        .doc_defaults
        .as_ref()
        .and_then(|defaults| defaults.r_pr.as_ref())
        .and_then(|formatting| formatting.font_size)
        .unwrap_or(20.0);
    let styles = &definitions.styles;
    let mut sizes = serde_json::Map::new();
    let mut hyperlink_sizes = serde_json::Map::new();
    for style in styles {
        let hyperlink = style.style_type == "character"
            && [Some(style.style_id.as_str()), style.name.as_deref()]
                .into_iter()
                .flatten()
                .any(|name| {
                    matches!(
                        name.to_ascii_lowercase().as_str(),
                        "hyperlink" | "followedhyperlink"
                    )
                });
        if style.style_type != "paragraph" && !hyperlink {
            continue;
        }
        let mut size = None;
        let mut seen: Vec<&str> = Vec::new();
        let mut current = Some(style);
        while let Some(entry) = current {
            if seen.contains(&entry.style_id.as_str()) {
                break;
            }
            seen.push(&entry.style_id);
            if let Some(found) = entry
                .r_pr
                .as_ref()
                .and_then(|formatting| formatting.font_size)
            {
                size = Some(found);
                break;
            }
            current = entry
                .based_on
                .as_ref()
                .and_then(|id| styles.iter().find(|style| &style.style_id == id));
        }
        if hyperlink {
            if let Some(size) = size {
                hyperlink_sizes.insert(style.style_id.clone(), json!(size / 2.0));
            }
        } else {
            sizes.insert(
                style.style_id.clone(),
                json!(size.unwrap_or(doc_default_half) / 2.0),
            );
        }
    }
    let default_id = styles
        .iter()
        .find(|style| style.style_type == "paragraph" && style.default == Some(true))
        .map(|style| style.style_id.clone())
        .or_else(|| {
            styles
                .iter()
                .any(|style| style.style_type == "paragraph" && style.style_id == "Normal")
                .then(|| "Normal".to_owned())
        });
    json!({
        "paragraphStyleFontSizes": sizes,
        "hyperlinkStyleFontSizes": hyperlink_sizes,
        "docDefaultFontSize": doc_default_half / 2.0,
        "defaultParagraphStyleId": default_id,
    })
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
                    "bodyStory": "body", "renderEnv": chain_env(bytes),
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
                    run["text"]
                        .as_str()
                        .or_else(|| run["fallback"].as_str())
                        .is_some_and(|text| !text.is_empty())
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

fn lowered_first_run(bytes: &[u8], render_env: Value) -> Value {
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
                    "bodyStory": "body", "renderEnv": render_env,
                    "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
                    "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 10}}
                })
                .to_string(),
            )
            .unwrap(),
    )
    .unwrap();
    output["measured"][0]["block"]["runs"][0].clone()
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
fn a_hyperlink_style_size_does_not_drop_a_toc_entry_below_its_chain() {
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="28"/></w:rPr></w:style></w:styles>"#;
    let body = r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Entry one</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:hyperlink></w:p>"#;
    let heights = line_heights(&document(body, styles));
    assert_height(heights[0], 12.0);
}

#[test]
fn a_larger_field_result_does_not_raise_an_unsized_toc_entry() {
    let body = r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:hyperlink w:anchor="_Toc1"><w:r><w:t>Entry one</w:t></w:r><w:r><w:tab/></w:r><w:fldSimple w:instr=" PAGEREF _Toc1 \h "><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>3</w:t></w:r></w:fldSimple></w:hyperlink></w:p>"#;
    let heights = line_heights(&document(body, &styles(None, None)));
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

#[test]
fn a_tab_at_the_entry_size_keeps_its_line_height() {
    let body = entry("Entry one", Some(24), None);
    let heights = line_heights(&document(&body, &styles(Some(24), None)));
    assert_height(heights[0], 12.0);
}

#[test]
fn larger_entry_text_still_drives_a_toc_line() {
    let body = r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>Entry one</w:t></w:r><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:hyperlink></w:p>"#.to_owned();
    let heights = line_heights(&document(&body, &styles(Some(20), None)));
    assert_height(heights[0], 14.0);
}

#[test]
fn a_larger_tab_run_still_raises_a_body_line() {
    let body = r#"<w:p><w:r><w:t>Body text</w:t></w:r><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:p>"#.to_owned();
    let heights = line_heights(&document(&body, &styles(Some(20), None)));
    assert_height(heights[0], 14.0);
}

#[test]
fn a_hyperlink_style_size_still_applies_in_body_text() {
    let body = r#"<w:p><w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Body link text</w:t></w:r></w:hyperlink></w:p>"#.to_owned();
    let heights = line_heights(&document(&body, &styles(Some(20), Some(28))));
    assert_height(heights[0], 14.0);
}

#[test]
fn a_larger_paragraph_mark_does_not_raise_a_body_paragraph() {
    let body = r#"<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:r><w:t>Body text</w:t></w:r></w:p>"#.to_owned();
    let heights = line_heights(&document(&body, &styles(None, None)));
    assert_height(heights[0], 10.0);
}

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

#[test]
fn a_style_switch_measures_the_new_style_size() {
    let (bytes, styles) = marked_paragraph();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(77101);
    let doc = engine.doc();
    seed_from_docx(doc, &bytes).unwrap();
    assert!(!properties(doc, "00000001").contains_key("runDefaultTextFormatting"));

    doc.restyle_paragraphs(
        &EditCtx::local("", ""),
        &ParaSelector::One("00000001".to_owned()),
        "Heading1",
        &styles,
    )
    .unwrap();

    assert!(!properties(doc, "00000001").contains_key("runDefaultTextFormatting"));
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

fn layout_heights(engine: &EngineSession, bytes: &[u8], render_env: Value) -> Vec<f64> {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(
                &json!({
                    "bodyStory": "body", "renderEnv": render_env,
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
                runs.iter()
                    .any(|run| run["text"].as_str().is_some_and(|text| !text.is_empty()))
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

fn heights_with_env(bytes: &[u8], render_env: Value) -> Vec<f64> {
    let engine = EngineSession::new(77100);
    seed_from_docx(engine.doc(), bytes).unwrap();
    layout_heights(&engine, bytes, render_env)
}

const CHAIN12_STYLES: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="28"/></w:rPr></w:style></w:styles>"#;

fn pageref_in_hyperlink_body() -> &'static str {
    r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Entry one</w:t></w:r><w:r><w:tab/></w:r><w:fldSimple w:instr=" PAGEREF _Toc1 \h "><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>3</w:t></w:r></w:fldSimple></w:hyperlink></w:p>"#
}

#[test]
fn a_pageref_field_inside_the_hyperlink_takes_the_10pt_toc_chain_size() {
    let heights = heights_with_env(
        &document(pageref_in_hyperlink_body(), &styles(Some(20), Some(28))),
        json!({"paragraphStyleFontSizes": {"Normal": 10.0, "TOC1": 10.0}, "hyperlinkStyleFontSizes": {"Hyperlink": 14.0}, "docDefaultFontSize": 10.0}),
    );
    assert_height(heights[0], 10.0);
}

#[test]
fn a_pageref_field_inside_the_hyperlink_takes_the_12pt_toc_chain_size() {
    let heights = heights_with_env(
        &document(pageref_in_hyperlink_body(), CHAIN12_STYLES),
        json!({"paragraphStyleFontSizes": {"Normal": 10.0, "TOC1": 12.0}, "hyperlinkStyleFontSizes": {"Hyperlink": 14.0}, "docDefaultFontSize": 10.0}),
    );
    assert_height(heights[0], 12.0);
}

#[test]
fn a_12pt_toc1_entry_with_a_14pt_mark_and_no_field_run_stays_12pt() {
    let body = r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Entry one</w:t></w:r><w:r><w:tab/></w:r><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>3</w:t></w:r></w:hyperlink></w:p>"#;
    let heights = heights_with_env(
        &document(body, CHAIN12_STYLES),
        json!({"paragraphStyleFontSizes": {"Normal": 10.0, "TOC1": 12.0}, "hyperlinkStyleFontSizes": {"Hyperlink": 14.0}, "docDefaultFontSize": 10.0}),
    );
    assert_height(heights[0], 12.0);
}

#[test]
fn a_restyle_to_heading1_measures_the_new_16pt_chain_size() {
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="32"/></w:rPr></w:style></w:styles>"#;
    let body = r#"<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="Normal"/><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:r><w:t>Switch me</w:t></w:r></w:p>"#;
    let bytes = document(body, styles);
    let parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let resolver = StyleResolver::from_parts(&parts).unwrap();
    let engine = EngineSession::new(77104);
    let doc = engine.doc();
    seed_from_docx(doc, &bytes).unwrap();
    doc.restyle_paragraphs(
        &EditCtx::local("", ""),
        &ParaSelector::One("00000001".to_owned()),
        "Heading1",
        &resolver,
    )
    .unwrap();
    let heights = layout_heights(
        &engine,
        &bytes,
        json!({"paragraphStyleFontSizes": {"Normal": 10.0, "Heading1": 16.0}, "docDefaultFontSize": 10.0}),
    );
    assert_height(heights[0], 16.0);
}

#[test]
fn typing_into_an_empty_12pt_toc1_paragraph_takes_the_chain_size() {
    let body = r#"<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="TOC1"/><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr></w:p>"#;
    let bytes = document(body, CHAIN12_STYLES);
    let engine = EngineSession::new(77105);
    let doc = engine.doc();
    seed_from_docx(doc, &bytes).unwrap();
    doc.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", 0),
        "Typed entry",
        FormatPolicy::Plain,
    )
    .unwrap();
    let heights = layout_heights(
        &engine,
        &bytes,
        json!({"paragraphStyleFontSizes": {"Normal": 10.0, "TOC1": 12.0}, "docDefaultFontSize": 10.0}),
    );
    assert_height(heights[0], 12.0);
}

#[test]
fn a_split_keeps_both_halves_at_the_chain_size() {
    let body = r#"<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="Normal"/><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:r><w:t>Body text</w:t></w:r></w:p>"#;
    let bytes = document(body, &styles(None, None));
    let engine = EngineSession::new(77106);
    let doc = engine.doc();
    seed_from_docx(doc, &bytes).unwrap();
    doc.split_paragraph(&EditCtx::local("", ""), Position::new("body", 5), None)
        .unwrap();
    let heights = layout_heights(
        &engine,
        &bytes,
        json!({"paragraphStyleFontSizes": {"Normal": 10.0}, "docDefaultFontSize": 10.0}),
    );
    assert_eq!(heights.len(), 2);
    assert_height(heights[0], 10.0);
    assert_height(heights[1], 10.0);
}

const NORMAL11_STYLES: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style></w:styles>"#;

#[test]
fn a_paragraph_without_a_style_takes_the_default_style_chain_size() {
    let body =
        r#"<w:p w14:paraId="00000001"><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr></w:p>"#;
    let bytes = document(body, NORMAL11_STYLES);
    let engine = EngineSession::new(77107);
    let doc = engine.doc();
    seed_from_docx(doc, &bytes).unwrap();
    doc.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", 0),
        "Body text",
        FormatPolicy::Plain,
    )
    .unwrap();
    let heights = layout_heights(
        &engine,
        &bytes,
        json!({"defaultParagraphStyleId": "Normal", "paragraphStyleFontSizes": {"Normal": 11.0}, "docDefaultFontSize": 10.0}),
    );
    assert_height(heights[0], 11.0);
}

#[test]
fn a_paragraph_with_an_unknown_style_takes_the_default_style_chain_size() {
    let body = r#"<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="UnknownStyle"/><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr></w:p>"#;
    let bytes = document(body, NORMAL11_STYLES);
    let engine = EngineSession::new(77108);
    let doc = engine.doc();
    seed_from_docx(doc, &bytes).unwrap();
    doc.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", 0),
        "Body text",
        FormatPolicy::Plain,
    )
    .unwrap();
    let heights = layout_heights(
        &engine,
        &bytes,
        json!({"defaultParagraphStyleId": "Normal", "paragraphStyleFontSizes": {"Normal": 11.0}, "docDefaultFontSize": 10.0}),
    );
    assert_height(heights[0], 11.0);
}

#[test]
fn a_style_without_basedon_over_a_12pt_normal_measures_10pt() {
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/></w:style></w:styles>"#;
    let body =
        r#"<w:p><w:pPr><w:pStyle w:val="Plain"/></w:pPr><w:r><w:t>Body text</w:t></w:r></w:p>"#;
    let heights = heights_with_env(
        &document(body, styles),
        json!({"defaultParagraphStyleId": "Normal", "paragraphStyleFontSizes": {"Normal": 12.0, "Plain": 10.0}, "docDefaultFontSize": 10.0}),
    );
    assert_height(heights[0], 10.0);
}

#[test]
fn an_unsized_run_beside_a_larger_direct_run_keeps_the_paragraph_default_size() {
    let body = r#"<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Small </w:t></w:r><w:r><w:rPr><w:sz w:val="36"/></w:rPr><w:t>Big</w:t></w:r></w:p>"#;
    let run = lowered_first_run(&document(body, &styles(None, None)), json!({}));
    assert_eq!(run["fontSize"], 10.0);
}

const CS_STYLES: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/><w:szCs w:val="28"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"#;

#[test]
fn an_unsized_complex_script_run_takes_the_cs_chain_size() {
    let body = r#"<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:rPr><w:cs/></w:rPr><w:t>Complex</w:t></w:r></w:p>"#;
    let run = lowered_first_run(
        &document(body, CS_STYLES),
        json!({"paragraphStyleFontSizes": {"Normal": 10.0}, "docDefaultFontSize": 10.0}),
    );
    assert_eq!(run["fontSizeCs"], 14.0);
}

#[test]
fn a_directly_sized_run_without_szcs_keeps_the_main_cs_size() {
    let body = r#"<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:rPr><w:sz w:val="32"/><w:cs/></w:rPr><w:t>Complex</w:t></w:r></w:p>"#;
    let run = lowered_first_run(
        &document(body, CS_STYLES),
        json!({"paragraphStyleFontSizes": {"Normal": 10.0}, "docDefaultFontSize": 10.0}),
    );
    assert_eq!(run["fontSize"], 16.0);
    assert_eq!(run["fontSizeCs"], 14.0);
}

#[test]
fn a_split_at_the_end_has_no_chain_defaults_key() {
    let (bytes, _styles) = marked_paragraph();
    let doc = docx_edit::EditingDoc::new(77102);
    seed_from_docx(&doc, &bytes).unwrap();
    doc.split_paragraph(&EditCtx::local("", ""), Position::new("body", 9), None)
        .unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs.len(), 2);
    assert!(
        !paragraphs[1]
            .properties
            .contains_key("runDefaultTextFormatting")
    );
}

#[test]
fn a_save_round_trip_never_writes_a_chain_defaults_key() {
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

fn lowered_blocks(bytes: &[u8], render_env: Value) -> Value {
    let doc = docx_edit::EditingDoc::new(77109);
    seed_from_docx(&doc, bytes).unwrap();
    let blocks = docx_edit::bridge::yrs_doc_to_layout_blocks(
        &doc,
        "body",
        &serde_json::from_value(render_env).unwrap(),
    )
    .unwrap();
    serde_json::to_value(blocks).unwrap()
}

fn lowered_runs(bytes: &[u8], render_env: Value) -> Vec<Value> {
    lowered_blocks(bytes, render_env)[0]["runs"]
        .as_array()
        .unwrap()
        .clone()
}

#[test]
fn a_toc_field_sizes_all_three_custom_style_entries_and_stops_at_its_end() {
    for with_ids in [true, false] {
        let paragraph = |index: usize, text: &str, begin: bool, end: bool| {
            let id = if with_ids {
                format!(r#" w14:paraId="{index:08X}""#)
            } else {
                String::new()
            };
            let begin = if begin {
                concat!(
                    r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
                    r#"<w:r><w:instrText> TOC \h </w:instrText></w:r>"#,
                    r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                )
            } else {
                ""
            };
            let end = if end {
                r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            } else {
                ""
            };
            format!(
                concat!(
                    r#"<w:p{id}><w:pPr><w:pStyle w:val="ContentsEntry"/></w:pPr>{begin}"#,
                    r#"<w:hyperlink w:anchor="entry"><w:r><w:rPr>"#,
                    r#"<w:rStyle w:val="Hyperlink"/></w:rPr><w:t>{text}</w:t></w:r>"#,
                    r#"<w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:tab/></w:r>"#,
                    r#"<w:fldSimple w:instr=" PAGEREF entry "><w:r><w:rPr>"#,
                    r#"<w:sz w:val="28"/></w:rPr><w:t>3</w:t></w:r></w:fldSimple>"#,
                    r#"</w:hyperlink>{end}</w:p>"#,
                ),
                id = id,
                begin = begin,
                text = text,
                end = end,
            )
        };
        let body = [
            paragraph(1, "Before", false, false),
            paragraph(2, "Entry one", true, false),
            paragraph(3, "Entry two", false, false),
            paragraph(4, "Entry three", false, true),
            paragraph(5, "After", false, false),
        ]
        .concat();
        let styles = styles(Some(20), Some(28))
            .replace("TOC1", "ContentsEntry")
            .replace("toc 1", "Contents Entry");
        let bytes = document(&body, &styles);
        let blocks = lowered_blocks(&bytes, chain_env(&bytes));
        assert_eq!(blocks.as_array().unwrap().len(), 5);
        for index in 1..=3 {
            let runs = blocks[index]["runs"].as_array().unwrap();
            let text = ["Entry one", "Entry two", "Entry three"][index - 1];
            assert!(runs.iter().any(|run| run["text"] == text));
            for run in runs {
                assert_eq!(run["fontSize"], 10.0, "{run}");
            }
        }
        for index in [0, 4] {
            assert_eq!(blocks[index]["runs"][0]["fontSize"], 14.0);
        }
        let heights = line_heights(&bytes);
        assert_eq!(heights.len(), 5);
        for (height, size) in heights.into_iter().zip([14.0, 10.0, 10.0, 10.0, 14.0]) {
            assert_height(height, size);
        }
    }
}

#[test]
fn an_opaque_toc_field_keeps_its_14pt_cached_entry_text() {
    let body = concat!(
        r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>"#,
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
        r#"<w:r><w:instrText> TOC </w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
        r#"<w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>Entry one</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    );
    let bytes = document(body, &styles(Some(20), None));
    let runs = lowered_runs(&bytes, chain_env(&bytes));
    assert_eq!(runs.len(), 1);
    assert_eq!(runs[0]["kind"], "field");
    assert_eq!(runs[0]["fallback"], "Entry one");
    assert_eq!(runs[0]["fontSize"], 14.0);
    assert_height(line_heights(&bytes)[0], 14.0);
}

#[test]
fn opaque_toc_entry_text_sets_the_cap_for_page_fields() {
    for instruction in ["PAGEREF _Toc1", "PAGE", "NUMPAGES", "SECTIONPAGES"] {
        let body = format!(
            concat!(
                r#"<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>"#,
                r#"<w:fldSimple w:instr=" TOC "><w:r><w:rPr><w:sz w:val="28"/>"#,
                r#"</w:rPr><w:t>Entry one</w:t></w:r></w:fldSimple>"#,
                r#"<w:r><w:rPr><w:sz w:val="36"/></w:rPr><w:tab/></w:r>"#,
                r#"<w:fldSimple w:instr=" {instruction} "><w:r><w:rPr>"#,
                r#"<w:sz w:val="36"/></w:rPr><w:t>3</w:t></w:r></w:fldSimple></w:p>"#,
            ),
            instruction = instruction,
        );
        let bytes = document(&body, &styles(Some(20), None));
        let runs = lowered_runs(&bytes, chain_env(&bytes));
        assert_eq!(runs.len(), 3);
        for run in &runs {
            assert_eq!(run["fontSize"], 14.0, "{run}");
        }
        assert_height(line_heights(&bytes)[0], 14.0);
    }
}

#[test]
fn toc_hyperlink_sizes_are_derived_from_the_render_environment() {
    for style_id in ["Hyperlink", "FollowedHyperlink"] {
        for (direct_size, expected) in [("", 10.0), (r#"<w:sz w:val="36"/>"#, 18.0)] {
            let body = format!(
                concat!(
                    r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>"#,
                    r#"<w:r><w:rPr><w:rStyle w:val="{style_id}"/>{direct_size}</w:rPr>"#,
                    r#"<w:t>Entry one</w:t></w:r></w:p>"#,
                ),
                style_id = style_id,
                direct_size = direct_size,
            );
            let styles = styles(Some(20), Some(28)).replace("Hyperlink", style_id);
            let bytes = document(&body, &styles);
            let runs = lowered_runs(&bytes, chain_env(&bytes));
            assert_eq!(runs[0]["fontSize"], expected);
        }
    }
}

#[test]
fn a_direct_size_matching_another_hyperlink_style_stays_direct() {
    let styles = styles(Some(20), Some(28)).replace(
        "</w:styles>",
        concat!(
            r#"<w:style w:type="character" w:styleId="FollowedHyperlink"><w:rPr>"#,
            r#"<w:sz w:val="36"/></w:rPr></w:style></w:styles>"#,
        ),
    );
    let styles = styles.replace(
        r#"<w:sz w:val="28"/>"#,
        r#"<w:sz w:val="28"/><w:color w:val="0000FF"/><w:u w:val="single"/>"#,
    );
    let body = concat!(
        r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>"#,
        r#"<w:hyperlink w:anchor="entry"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/>"#,
        r#"<w:sz w:val="36"/></w:rPr><w:t>Entry one</w:t></w:r></w:hyperlink></w:p>"#,
    );
    let bytes = document(body, &styles);
    let runs = lowered_runs(&bytes, chain_env(&bytes));
    assert_eq!(runs[0]["fontSize"], 18.0);
}

#[test]
fn a_body_pageref_preserves_the_hyperlink_style_size() {
    let body = r#"<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:hyperlink w:anchor="target"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Body link</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> see page </w:t></w:r><w:fldSimple w:instr=" PAGEREF target "><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p>"#;
    let bytes = document(body, &styles(Some(20), Some(28)));
    let runs = lowered_runs(&bytes, chain_env(&bytes));
    assert_eq!(runs[0]["fontSize"], 14.0);
    assert_ne!(runs[0]["hyperlink"]["noDefaultStyle"], true);
    assert_height(line_heights(&bytes)[0], 14.0);
}

#[test]
fn a_toc_field_result_with_a_custom_style_uses_the_entry_size() {
    let body = r#"<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Entry one</w:t></w:r><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:tab/></w:r><w:fldSimple w:instr=" PAGEREF _Toc1 \h "><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>3</w:t></w:r></w:fldSimple></w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#;
    let bytes = document(body, &styles(Some(20), Some(28)));
    let runs = lowered_runs(&bytes, chain_env(&bytes));
    assert!(runs.iter().any(|run| run["text"] == "Entry one"));
    assert!(runs.iter().any(|run| run["kind"] == "tab"));
    for run in &runs {
        assert_eq!(run["fontSize"], 10.0, "{run}");
    }
}

#[test]
fn stripping_a_toc_hyperlink_size_preserves_complex_script_sizes() {
    let styles = styles(Some(20), Some(28)).replace(
        r#"<w:sz w:val="28"/>"#,
        r#"<w:sz w:val="28"/><w:szCs w:val="36"/>"#,
    );
    for (direct_cs, expected) in [("", 18.0), (r#"<w:szCs w:val="38"/>"#, 19.0)] {
        let body = format!(
            r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/>{direct_cs}<w:cs/></w:rPr><w:t>Entry one</w:t></w:r></w:hyperlink></w:p>"#
        );
        let bytes = document(&body, &styles);
        let runs = lowered_runs(&bytes, chain_env(&bytes));
        assert_eq!(runs[0]["fontSize"], 10.0);
        assert_eq!(runs[0]["fontSizeCs"], expected);
    }
}

#[test]
fn capping_toc_auxiliary_sizes_preserves_complex_script_sizes() {
    let body = r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Entry one</w:t></w:r><w:r><w:rPr><w:sz w:val="28"/><w:szCs w:val="32"/><w:cs/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:sz w:val="28"/><w:szCs w:val="34"/><w:cs/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:fldSimple w:instr=" PAGEREF _Toc1 "><w:r><w:rPr><w:sz w:val="28"/><w:szCs w:val="36"/><w:cs/></w:rPr><w:t>3</w:t></w:r></w:fldSimple></w:p>"#;
    let bytes = document(body, &styles(Some(20), None));
    let runs = lowered_runs(&bytes, chain_env(&bytes));
    assert_eq!(runs.len(), 4);
    for (run, size_cs) in runs[1..].iter().zip([16.0, 17.0, 18.0]) {
        assert_eq!(run["fontSize"], 10.0, "{run}");
        assert_eq!(run["fontSizeCs"], size_cs, "{run}");
    }
}
