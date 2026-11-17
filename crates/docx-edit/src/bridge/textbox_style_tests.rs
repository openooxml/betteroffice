use docx_layout::types::{LayoutBlock, ParagraphBlock, Run, RunFormatting};
use serde_json::{Value, json};

use crate::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use crate::seed::fixture;
use crate::{EditCtx, EditingDoc, Position, SegmentContent, seed_from_docx};

const STYLES: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:before="120" w:after="180" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr><w:jc w:val="left"/></w:pPr><w:rPr><w:color w:val="112233"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Base"><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/><w:spacing w:before="240"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="28"/><w:b/><w:lang w:val="fr-FR"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Derived"><w:basedOn w:val="Base"/><w:pPr><w:spacing w:after="360"/></w:pPr><w:rPr><w:i/><w:color w:val="445566"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="EastAsia"><w:basedOn w:val="Derived"/><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr></w:style><w:style w:type="character" w:styleId="AccentBase"><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="32"/><w:color w:val="778899"/></w:rPr></w:style><w:style w:type="character" w:styleId="Accent"><w:basedOn w:val="AccentBase"/><w:rPr><w:b/><w:lang w:val="de-DE"/></w:rPr></w:style><w:style w:type="character" w:styleId="BidiAccent"><w:basedOn w:val="Accent"/><w:rPr><w:lang w:eastAsia="zh-TW" w:bidi="he-IL"/></w:rPr></w:style></w:styles>"#;

#[derive(Clone, Copy)]
enum TextBox {
    Drawing,
    Vml,
    Header,
    Table,
}

fn drawing(content: &str) -> String {
    format!(
        r#"<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2743200" cy="1371600"/><wp:wrapNone/><wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>"#
    )
}

fn document(paragraph: &str, text_box: TextBox) -> Vec<u8> {
    let content = paragraph.replace("Sample", "Boxed");
    let boxed = match text_box {
        TextBox::Vml => format!(
            r#"<w:p><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="Box" style="position:absolute;width:216pt;height:108pt" filled="f" stroked="f"><v:textbox><w:txbxContent>{content}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"#
        ),
        TextBox::Table => drawing(&format!(
            r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{content}</w:tc></w:tr></w:tbl>"#
        )),
        _ => drawing(&content),
    };
    if matches!(text_box, TextBox::Header) {
        fixture::Package::new(&format!(
            r#"{paragraph}<w:sectPr><w:headerReference w:type="default" r:id="rHeader"/></w:sectPr>"#
        ))
        .styles(STYLES)
        .part(
            "header1.xml",
            "rHeader",
            "header",
            "header",
            &format!("<w:hdr {}>{boxed}</w:hdr>", fixture::NS),
        )
        .bytes()
    } else {
        fixture::Package::new(&format!("{paragraph}{boxed}"))
            .styles(STYLES)
            .bytes()
    }
}

fn paragraph(style: &str, run_style: Option<&str>, direct: bool) -> String {
    let run_style = run_style.map_or_else(String::new, |id| format!(r#"<w:rStyle w:val="{id}"/>"#));
    let paragraph_mark = if matches!(style, "Derived" | "EastAsia") {
        r#"<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="36"/></w:rPr>"#
    } else {
        ""
    };
    let (ppr, rpr) = if direct {
        (
            r#"<w:jc w:val="right"/><w:spacing w:before="80" w:after="100" w:line="400" w:lineRule="exact"/>"#,
            r#"<w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/><w:sz w:val="40"/><w:b/><w:i w:val="0"/><w:color w:val="AABBCC"/><w:lang w:val="it-IT" w:eastAsia="ko-KR"/>"#,
        )
    } else {
        ("", "")
    };
    format!(
        r#"<w:p><w:pPr><w:pStyle w:val="{style}"/>{ppr}{paragraph_mark}</w:pPr><w:r><w:rPr>{run_style}{rpr}</w:rPr><w:t>Sample</w:t></w:r></w:p>"#
    )
}

fn text_formatting(paragraph: &ParagraphBlock, text: &str) -> RunFormatting {
    paragraph
        .runs
        .iter()
        .find_map(|run| match run {
            Run::Text(run) if run.text == text => Some(run.fmt.clone()),
            _ => None,
        })
        .unwrap()
}

fn lower_text_box_paragraphs(bytes: &[u8], env: &RenderEnv) -> Vec<ParagraphBlock> {
    let doc = EditingDoc::new(75306);
    seed_from_docx(&doc, bytes).unwrap();
    yrs_doc_to_layout_blocks(&doc, "body", env)
        .unwrap()
        .into_iter()
        .flat_map(|block| match block {
            LayoutBlock::Shape(shape) => shape.inner_text.unwrap_or_default(),
            _ => Vec::new(),
        })
        .collect()
}

fn effective(formatting: &RunFormatting) -> Value {
    json!({
        "font": formatting.font_family,
        "fontSlots": formatting.font_slots,
        "size": formatting.font_size,
        "bold": formatting.bold == Some(true),
        "italic": formatting.italic == Some(true),
        "color": formatting.color,
        "language": formatting.language,
    })
}

/// Matches body formatting except language slots set by the run or its character style.
fn compare(
    style: &str,
    run_style: Option<&str>,
    direct: bool,
    text_box: TextBox,
    own_language_slots: &[&str],
) -> Value {
    let bytes = document(&paragraph(style, run_style, direct), text_box);
    let parsed = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let request =
        crate::layout_request::region_layout_request(&serde_json::to_value(parsed).unwrap(), 24.0);
    let env: RenderEnv = serde_json::from_value(request["renderEnv"].clone()).unwrap();
    let doc = EditingDoc::new(75304);
    seed_from_docx(&doc, &bytes).unwrap();
    let before: Vec<_> = doc
        .list_stories()
        .stories
        .into_iter()
        .map(|story| {
            let segments = doc.story_segments(&story.story).unwrap();
            (story.story, segments)
        })
        .collect();
    let save_plan = doc.paragraph_save_plan();
    let body = yrs_doc_to_layout_blocks(&doc, "body", &env).unwrap();
    let body_paragraph = body
        .iter()
        .find_map(|block| match block {
            LayoutBlock::Paragraph(paragraph)
                if paragraph
                    .runs
                    .iter()
                    .any(|run| matches!(run, Run::Text(run) if run.text == "Sample")) =>
            {
                Some(paragraph)
            }
            _ => None,
        })
        .unwrap();
    let header;
    let blocks = if matches!(text_box, TextBox::Header) {
        let story = doc
            .list_stories()
            .stories
            .into_iter()
            .find(|story| {
                matches!(
                    story.part.as_deref(),
                    Some("word/header1.xml" | "/word/header1.xml")
                )
            })
            .unwrap();
        header = yrs_doc_to_layout_blocks(&doc, &story.story, &env).unwrap();
        &header
    } else {
        &body
    };
    let boxed = blocks
        .iter()
        .find_map(|block| match block {
            LayoutBlock::Shape(shape) => shape.inner_text.as_ref()?.iter().find(|paragraph| {
                paragraph
                    .runs
                    .iter()
                    .any(|run| matches!(run, Run::Text(run) if run.text == "Boxed"))
            }),
            _ => None,
        })
        .unwrap();
    let mut body_formatting = effective(&text_formatting(body_paragraph, "Sample"));
    let box_formatting = effective(&text_formatting(boxed, "Boxed"));
    let mut comparable_box = box_formatting.clone();
    for slot in own_language_slots {
        body_formatting["language"]
            .as_object_mut()
            .unwrap()
            .remove(*slot);
        comparable_box["language"]
            .as_object_mut()
            .unwrap()
            .remove(*slot);
    }
    assert_eq!(comparable_box, body_formatting);
    let body_attrs = body_paragraph.attrs.as_ref().unwrap();
    let box_attrs = boxed.attrs.as_ref().unwrap();
    assert_eq!(box_attrs.alignment, body_attrs.alignment);
    assert_eq!(box_attrs.spacing, body_attrs.spacing);
    assert_eq!(
        box_attrs.default_font_family,
        body_attrs.default_font_family
    );
    assert_eq!(box_attrs.default_font_size, body_attrs.default_font_size);
    if matches!(style, "Derived" | "EastAsia") {
        assert_eq!(box_attrs.default_font_family.as_deref(), Some("Cambria"));
        assert_eq!(box_attrs.default_font_size, Some(18.0));
    }
    for (story, segments) in before {
        assert_eq!(doc.story_segments(&story).unwrap(), segments);
    }
    assert_eq!(doc.paragraph_save_plan(), save_plan);
    box_formatting
}

#[test]
fn text_boxes_resolve_paragraph_styles_through_based_on() {
    for text_box in [
        TextBox::Drawing,
        TextBox::Vml,
        TextBox::Header,
        TextBox::Table,
    ] {
        let formatting = compare("Derived", None, false, text_box, &[]);
        assert_eq!(formatting["font"], "Arial");
        assert_eq!(formatting["size"], 14.0);
        assert_eq!(formatting["bold"], true);
        assert_eq!(formatting["italic"], true);
        assert_eq!(formatting["color"], "#445566");
        assert_eq!(formatting["language"]["latin"], "fr-FR");
    }
}

#[test]
fn text_boxes_resolve_character_styles_and_style_toggles() {
    for text_box in [
        TextBox::Drawing,
        TextBox::Vml,
        TextBox::Header,
        TextBox::Table,
    ] {
        let formatting = compare("Derived", Some("Accent"), false, text_box, &["latin"]);
        assert_eq!(formatting["font"], "Courier New");
        assert_eq!(formatting["size"], 16.0);
        assert_eq!(formatting["bold"], false);
        assert_eq!(formatting["italic"], true);
        assert_eq!(formatting["color"], "#778899");
        assert_eq!(formatting["language"]["latin"], "de-DE");
    }
}

#[test]
fn text_box_direct_formatting_overrides_paragraph_and_character_styles() {
    for text_box in [
        TextBox::Drawing,
        TextBox::Vml,
        TextBox::Header,
        TextBox::Table,
    ] {
        let formatting = compare(
            "Derived",
            Some("Accent"),
            true,
            text_box,
            &["latin", "eastAsia"],
        );
        assert_eq!(formatting["font"], "Verdana");
        assert_eq!(formatting["size"], 20.0);
        assert_eq!(formatting["bold"], true);
        assert_eq!(formatting["italic"], false);
        assert_eq!(formatting["color"], "#AABBCC");
        assert_eq!(formatting["language"]["latin"], "it-IT");
        assert_eq!(formatting["language"]["eastAsia"], "ko-KR");
    }
}

#[test]
fn text_box_styles_merge_individual_language_slots() {
    for text_box in [
        TextBox::Drawing,
        TextBox::Vml,
        TextBox::Header,
        TextBox::Table,
    ] {
        let formatting = compare("EastAsia", None, false, text_box, &[]);
        assert_eq!(
            formatting["language"],
            json!({"latin": "fr-FR", "eastAsia": "ja-JP", "bidi": "ar-SA"})
        );
    }
}

#[test]
fn text_box_character_languages_fill_slots_unset_by_direct_formatting() {
    for text_box in [
        TextBox::Drawing,
        TextBox::Vml,
        TextBox::Header,
        TextBox::Table,
    ] {
        for direct in [false, true] {
            let formatting = compare(
                "EastAsia",
                Some("BidiAccent"),
                direct,
                text_box,
                &["latin", "eastAsia", "bidi"],
            );
            assert_eq!(
                formatting["language"],
                if direct {
                    json!({"latin": "it-IT", "eastAsia": "ko-KR", "bidi": "he-IL"})
                } else {
                    json!({"latin": "de-DE", "eastAsia": "zh-TW", "bidi": "he-IL"})
                }
            );
        }
    }
}

#[test]
fn unknown_text_box_styles_fall_back_like_body_styles() {
    for text_box in [
        TextBox::Drawing,
        TextBox::Vml,
        TextBox::Header,
        TextBox::Table,
    ] {
        let formatting = compare("Missing", Some("MissingRun"), false, text_box, &[]);
        assert_eq!(formatting["font"], "Calibri");
        assert_eq!(formatting["size"], 11.0);
        assert_eq!(formatting["color"], "#112233");
        assert_eq!(formatting["language"]["latin"], "en-US");
    }
}

#[test]
fn edited_text_box_content_keeps_direct_formatting_over_cached_styles() {
    let bytes = document(
        &paragraph("Derived", Some("Accent"), false),
        TextBox::Drawing,
    );
    let doc = EditingDoc::new(75305);
    seed_from_docx(&doc, &bytes).unwrap();
    let env = RenderEnv::default();
    yrs_doc_to_layout_blocks(&doc, "body", &env).unwrap();
    let mut offset = 0;
    let mut target = None;
    for segment in doc.story_segments("body").unwrap() {
        match segment.content {
            SegmentContent::Text(text) => offset += text.encode_utf16().count() as u32,
            SegmentContent::OtherEmbed { kind, payload } if kind == "shape" => {
                let yrs::Any::String(shape) = &payload["shapeJson"] else {
                    panic!("missing shape JSON")
                };
                target = Some((offset, serde_json::from_str::<Value>(shape).unwrap()));
                break;
            }
            _ => offset += 1,
        }
    }
    let (offset, mut shape) = target.unwrap();
    let run = &mut shape["textBody"]["content"][0]["content"][0];
    run["content"][0]["text"] = json!("Edited");
    run["formatting"]["fontFamily"] = json!({"ascii": "Verdana", "hAnsi": "Verdana"});
    run["formatting"]["fontSize"] = json!(36);
    run["formatting"]["bold"] = json!(true);
    run["formatting"]["italic"] = json!(false);
    run["formatting"]["color"] = json!({"rgb": "ABCDEF"});
    run["formatting"]["language"] = json!({"latin": "it-IT", "bidi": "fa-IR"});
    doc.set_embed_attrs(
        &EditCtx::local("", "2030-01-02T03:04:05.006Z"),
        Position::new("body", offset),
        vec![("shapeJson".to_owned(), yrs::Any::from(shape.to_string()))],
    )
    .unwrap();
    let stored = doc.story_segments("body").unwrap();
    for _ in 0..2 {
        let blocks = yrs_doc_to_layout_blocks(&doc, "body", &env).unwrap();
        let boxed = blocks
            .iter()
            .find_map(|block| match block {
                LayoutBlock::Shape(shape) => shape.inner_text.as_ref()?.first(),
                _ => None,
            })
            .unwrap();
        let formatting = effective(&text_formatting(boxed, "Edited"));
        assert_eq!(formatting["font"], "Verdana");
        assert_eq!(formatting["size"], 18.0);
        assert_eq!(formatting["bold"], true);
        assert_eq!(formatting["italic"], false);
        assert_eq!(formatting["color"], "#ABCDEF");
        assert_eq!(
            formatting["language"],
            json!({"latin": "it-IT", "eastAsia": "zh-CN", "bidi": "fa-IR"})
        );
        assert_eq!(doc.story_segments("body").unwrap(), stored);
    }
}

#[test]
fn text_box_paragraphs_leave_borders_out_of_the_paragraph_only_renderer() {
    let content = r#"<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="24" w:space="8" w:color="99CB38"/><w:bottom w:val="single" w:sz="24" w:space="8" w:color="99CB38"/></w:pBdr></w:pPr><w:r><w:t>Boxed</w:t></w:r></w:p>"#;
    let bytes = fixture::Package::new(&drawing(content))
        .styles(STYLES)
        .bytes();
    let doc = EditingDoc::new(75305);
    seed_from_docx(&doc, &bytes).unwrap();
    let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
    let boxed = blocks
        .iter()
        .find_map(|block| match block {
            LayoutBlock::Shape(shape) => shape.inner_text.as_ref()?.first(),
            _ => None,
        })
        .unwrap();
    assert!(boxed.attrs.as_ref().unwrap().borders.is_none());
}

#[test]
fn text_box_direct_run_shading_keeps_theme_fills_tints_and_shades() {
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:shd w:val="clear" w:fill="102030"/></w:rPr></w:style><w:style w:type="character" w:styleId="Shade"><w:rPr><w:shd w:val="clear" w:fill="223344"/></w:rPr></w:style></w:styles>"#;
    let env = RenderEnv {
        theme_colors: [("accent1".to_owned(), "204060".to_owned())].into(),
        ..RenderEnv::default()
    };
    for (properties, expected) in [
        ("", Some("#102030")),
        (r#"<w:rStyle w:val="Shade"/>"#, Some("#223344")),
        (
            r#"<w:shd w:val="clear" w:fill="auto" w:themeFill="accent1"/>"#,
            Some("#204060"),
        ),
        (
            r#"<w:shd w:val="clear" w:fill="AABBCC" w:themeFill="accent1" w:themeFillTint="80"/>"#,
            Some("#8F9FAF"),
        ),
        (
            r#"<w:shd w:val="clear" w:fill="AABBCC" w:themeFill="accent1" w:themeFillShade="80"/>"#,
            Some("#102030"),
        ),
        (
            r#"<w:shd w:val="clear" w:fill="204060" w:themeFillTint="80"/>"#,
            Some("#8F9FAF"),
        ),
        (
            r#"<w:shd w:val="clear" w:fill="204060" w:themeFillShade="80"/>"#,
            Some("#102030"),
        ),
        (r#"<w:shd w:val="clear" w:fill="auto"/>"#, None),
    ] {
        let content =
            format!(r#"<w:p><w:r><w:rPr>{properties}</w:rPr><w:t>Boxed</w:t></w:r></w:p>"#);
        let bytes = fixture::Package::new(&drawing(&content))
            .styles(styles)
            .bytes();
        let paragraphs = lower_text_box_paragraphs(&bytes, &env);
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(
            text_formatting(&paragraphs[0], "Boxed")
                .highlight
                .as_deref(),
            expected,
            "{properties}"
        );
    }
}

#[test]
fn text_box_run_font_hints_select_east_asia_for_non_ascii_latin() {
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:rFonts w:ascii="Latin" w:hAnsi="Latin" w:eastAsia="Hint CJK"/></w:rPr></w:style><w:style w:type="character" w:styleId="Hint"><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:style></w:styles>"#;
    for (properties, hint) in [
        (r#"<w:rFonts w:hint="eastAsia"/>"#, "eastAsia"),
        (r#"<w:rStyle w:val="Hint"/>"#, "eastAsia"),
        (
            r#"<w:rStyle w:val="Hint"/><w:rFonts w:hint="default"/>"#,
            "default",
        ),
    ] {
        let content = format!(r#"<w:p><w:r><w:rPr>{properties}</w:rPr><w:t>é</w:t></w:r></w:p>"#);
        let bytes = fixture::Package::new(&drawing(&content))
            .styles(styles)
            .bytes();
        let paragraphs = lower_text_box_paragraphs(&bytes, &RenderEnv::default());
        assert_eq!(paragraphs.len(), 1);
        let formatting = text_formatting(&paragraphs[0], "é");
        assert_eq!(
            formatting.font_slots.as_ref().unwrap().hint.as_deref(),
            Some(hint)
        );
        let blocks = paragraphs
            .into_iter()
            .map(LayoutBlock::Paragraph)
            .collect::<Vec<_>>();
        let requirements =
            docx_layout::measure_blocks::collect_font_requirements(&blocks, "Calibri");
        assert_eq!(
            requirements.iter().any(|face| {
                face.family == "Hint CJK" && face.characters.iter().any(|ch| ch == 'é')
            }),
            hint == "eastAsia",
            "{properties}"
        );
    }
}

#[test]
fn text_box_style_numbering_does_not_synthesize_indent() {
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="List"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="IndentedList"><w:basedOn w:val="List"/><w:pPr><w:ind w:left="480" w:firstLine="120"/></w:pPr></w:style></w:styles>"#;
    for (style, direct, expected_left, expected_first_line) in [
        ("List", "", None, None),
        (
            "List",
            r#"<w:ind w:left="480" w:firstLine="120"/>"#,
            Some(32.0),
            Some(8.0),
        ),
        ("List", r#"<w:ind w:firstLine="120"/>"#, None, Some(8.0)),
        ("IndentedList", "", Some(32.0), Some(8.0)),
    ] {
        let content = format!(
            r#"<w:p><w:pPr><w:pStyle w:val="{style}"/>{direct}</w:pPr><w:r><w:t>Boxed</w:t></w:r></w:p>"#
        );
        let bytes = fixture::Package::new(&drawing(&content))
            .styles(styles)
            .bytes();
        let paragraphs = lower_text_box_paragraphs(&bytes, &RenderEnv::default());
        assert_eq!(paragraphs.len(), 1);
        let attrs = paragraphs[0].attrs.as_ref().unwrap();
        assert!(attrs.num_pr.is_none());
        assert!(attrs.list_marker.is_none());
        if expected_left.is_some() || expected_first_line.is_some() {
            let indent = attrs.indent.as_ref().unwrap();
            assert_eq!(indent.left, expected_left);
            assert_eq!(indent.first_line, expected_first_line);
            assert!(indent.hanging.is_none());
        } else {
            assert!(attrs.indent.is_none());
        }
    }
}

#[test]
fn shared_styles_survive_payload_edits_and_new_directly_formatted_runs() {
    let content = paragraph("Derived", Some("Accent"), false).replace(
        "</w:p>",
        r#"<w:r><w:rPr><w:rStyle w:val="Accent"/></w:rPr><w:t>Unedited</w:t></w:r></w:p>"#,
    );
    let bytes = document(&content, TextBox::Drawing);
    let source = EditingDoc::new(75360);
    seed_from_docx(&source, &bytes).unwrap();
    let doc = EditingDoc::new(75361);
    doc.apply_update_v1(&source.encode_state_as_update_v1())
        .unwrap();
    let env = RenderEnv::default();
    assert_eq!(
        yrs_doc_to_layout_blocks(&doc, "body", &env).unwrap(),
        yrs_doc_to_layout_blocks(&source, "body", &env).unwrap()
    );
    let original = yrs_doc_to_layout_blocks(&doc, "body", &env).unwrap();
    let original = original
        .iter()
        .find_map(|block| match block {
            LayoutBlock::Shape(shape) => shape.inner_text.as_ref()?.first(),
            _ => None,
        })
        .unwrap();
    let untouched = text_formatting(original, "Unedited");
    assert_eq!(untouched.font_family.as_deref(), Some("Courier New"));
    let mut offset = 0;
    for segment in doc.story_segments("body").unwrap() {
        match segment.content {
            SegmentContent::Text(text) => offset += text.encode_utf16().count() as u32,
            SegmentContent::OtherEmbed { payload, .. } if payload.contains_key("shapeJson") => {
                let yrs::Any::String(json) = &payload["shapeJson"] else {
                    panic!("shape JSON")
                };
                let mut shape: Value = serde_json::from_str(json).unwrap();
                let runs = shape["textBody"]["content"][0]["content"]
                    .as_array_mut()
                    .unwrap();
                runs[0]["content"][0]["text"] = json!("Edited");
                runs[0]["formatting"]["fontSize"] = json!(40);
                runs.push(
                    json!({"type": "run", "formatting": {"styleId": "Accent", "italic": false,
                    "color": {"rgb": "ABCDEF"}}, "content": [{"type": "text", "text": "New"}]}),
                );
                doc.set_embed_attrs(
                    &EditCtx::local("", ""),
                    Position::new("body", offset),
                    vec![("shapeJson".to_owned(), yrs::Any::from(shape.to_string()))],
                )
                .unwrap();
                break;
            }
            _ => offset += 1,
        }
    }
    let blocks = yrs_doc_to_layout_blocks(&doc, "body", &env).unwrap();
    let boxed = blocks
        .iter()
        .find_map(|block| match block {
            LayoutBlock::Shape(shape) => shape.inner_text.as_ref()?.first(),
            _ => None,
        })
        .unwrap();
    assert_eq!(text_formatting(boxed, "Unedited"), untouched);
    let edited = text_formatting(boxed, "Edited");
    assert_eq!(edited.font_family, untouched.font_family);
    assert_eq!(edited.font_size, Some(20.0));
    assert_eq!(edited.italic, Some(true));
    let added = text_formatting(boxed, "New");
    assert_eq!(added.font_family, untouched.font_family);
    assert_eq!(added.font_size, untouched.font_size);
    assert_ne!(added.italic, Some(true));
    assert_eq!(added.color.as_deref(), Some("#ABCDEF"));
}

#[test]
fn state_only_styles_cover_both_formats_stories_tables_and_sdts() {
    use yrs::{Map, ReadTxn, Transact};

    let paragraph = paragraph("Derived", Some("Accent"), false).replace(
        "<w:rPr><w:rFonts w:ascii=\"Cambria\"",
        "<w:rPr><w:rStyle w:val=\"BidiAccent\"/><w:rFonts w:ascii=\"Cambria\"",
    );
    let content = format!(
        r#"{paragraph}<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{paragraph}</w:tc></w:tr></w:tbl><w:sdt><w:sdtPr/><w:sdtContent>{paragraph}</w:sdtContent></w:sdt>"#
    );
    let drawing = drawing(&content);
    let vml = format!(
        r#"<w:p><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="VML" style="position:absolute;width:216pt;height:108pt"><v:textbox><w:txbxContent>{content}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"#
    );
    let bytes = fixture::Package::new(&format!(
        r#"{drawing}{vml}<w:sectPr><w:headerReference w:type="default" r:id="rHeader"/></w:sectPr>"#
    ))
    .styles(STYLES)
    .part(
        "header1.xml",
        "rHeader",
        "header",
        "header",
        &format!("<w:hdr {}>{drawing}{vml}</w:hdr>", fixture::NS),
    )
    .bytes();
    let source = EditingDoc::new(75362);
    seed_from_docx(&source, &bytes).unwrap();
    let peer = EditingDoc::new(75363);
    peer.apply_update_v1(&source.encode_state_as_update_v1())
        .unwrap();
    assert!(peer.source_metadata().is_none());
    let txn = peer.yrs_doc().transact();
    assert_eq!(
        txn.get_map(crate::shape_styles::SHAPE_STYLES)
            .unwrap()
            .len(&txn),
        1
    );
    drop(txn);
    for story in source.list_stories().stories {
        let expected =
            yrs_doc_to_layout_blocks(&source, &story.story, &RenderEnv::default()).unwrap();
        let actual = yrs_doc_to_layout_blocks(&peer, &story.story, &RenderEnv::default()).unwrap();
        assert_eq!(actual, expected);
        let paragraphs: Vec<_> = actual
            .iter()
            .filter_map(|block| match block {
                LayoutBlock::Shape(shape) => shape.inner_text.as_ref(),
                _ => None,
            })
            .flatten()
            .collect();
        assert_eq!(paragraphs.len(), 6);
        for paragraph in paragraphs {
            assert_eq!(
                text_formatting(paragraph, "Sample").font_family.as_deref(),
                Some("Courier New")
            );
        }
    }
}

#[test]
fn styles_arriving_after_first_paint_invalidate_resident_and_preview_lowering() {
    use std::collections::BTreeSet;
    use std::rc::Rc;
    use yrs::{Map, ReadTxn, Transact};

    let bytes = document(
        &paragraph("Derived", Some("Accent"), false),
        TextBox::Drawing,
    );
    for enabled in [false, true] {
        let source = EditingDoc::new(75364);
        seed_from_docx(&source, &bytes).unwrap();
        let json = {
            let mut txn = source.yrs_doc().transact_mut();
            let map = txn.get_map(crate::shape_styles::SHAPE_STYLES).unwrap();
            let json = crate::map_string(&map, &txn, "resolver").unwrap();
            map.remove(&mut txn, "resolver");
            json
        };
        let engine = crate::EngineSession::new(75365);
        engine.set_local_lowering(enabled);
        engine
            .doc()
            .apply_update_v1(&source.encode_state_as_update_v1())
            .unwrap();
        let env = RenderEnv::default();
        let before = engine.lower_story_json("body", &env).unwrap();
        let mut local = crate::bridge::local::LocalLowering::new(enabled);
        let (blocks, map, _, preview) =
            crate::bridge::preview::lower_recorded(engine.doc(), "body", &env, &mut local, true)
                .unwrap();
        assert!(local.matches_source(engine.doc()));
        {
            let mut txn = source.yrs_doc().transact_mut();
            txn.get_map(crate::shape_styles::SHAPE_STYLES)
                .unwrap()
                .insert(&mut txn, "resolver", json);
        }
        engine
            .doc()
            .apply_update_v1(&source.encode_state_as_update_v1())
            .unwrap();
        assert!(!local.matches_source(engine.doc()));
        let current: Vec<_> = blocks.into_iter().map(Rc::new).collect();
        assert!(
            crate::bridge::preview::replay(
                engine.doc(),
                &env,
                &preview.unwrap(),
                &BTreeSet::from(["late-style".to_owned()]),
                &current,
                &map,
                &local
            )
            .is_none()
        );
        let after = engine.lower_story_json("body", &env).unwrap();
        assert_ne!(before, after);
        let cold = crate::EngineSession::new(75366);
        cold.doc()
            .apply_update_v1(&source.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(after, cold.lower_story_json("body", &env).unwrap());
        let mut preview_env = env.clone();
        preview_env.revision_preview.insert(
            "late-style".to_owned(),
            crate::bridge::RevisionPreview::Accepted,
        );
        assert_eq!(
            engine.lower_story_json("body", &preview_env).unwrap(),
            cold.lower_story_json("body", &preview_env).unwrap()
        );
    }
}

#[test]
fn text_box_runs_keep_complex_script_toggles_like_the_direct_path() {
    let content = r#"<w:p><w:pPr><w:pStyle w:val="Derived"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:cs="Arial"/><w:bCs/><w:iCs/><w:cs/></w:rPr><w:t>Complex</w:t></w:r></w:p>"#;
    let bytes = fixture::Package::new(&drawing(content))
        .styles(STYLES)
        .bytes();
    let styled = lower_text_box_paragraphs(&bytes, &RenderEnv::default());
    let formatting = text_formatting(&styled[0], "Complex");
    assert_eq!(formatting.bold_cs, Some(true));
    assert_eq!(formatting.italic_cs, Some(true));
    assert_eq!(formatting.complex_script, Some(true));
}
