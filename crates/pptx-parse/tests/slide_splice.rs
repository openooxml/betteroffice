use std::collections::BTreeMap;

use pptx_parse::{
    BackgroundWrite, DeckWrite, ParagraphWrite, Placeholder, RunProperties, RunWrite, ShapeAdd,
    ShapePatch, ShapeWrite, SlideWrite, TableRowSource, TableStructure, TextTarget, TextWrite,
    parse_pptx, write_pptx_with_edits,
};

const ROOT_START: &str = r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">"#;

const UNTOUCHED: &str = r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Untouched"/><p:cNvSpPr/><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="3000" cy="4000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom><a:extLst><a:ext uri="{11111111-2222-3333-4444-555555555555}"><a14:creationId xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="7"/></a:ext></a:extLst></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Untouched text</a:t></a:r></a:p></p:txBody></p:sp>"#;

const ALTERNATE: &str = r#"<mc:AlternateContent><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:sp><p:nvSpPr><p:cNvPr id="9" name="Choice"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="4" name="Alternate"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Alternate text</a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>"#;

fn shape(id: u32, name: &str, text: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

fn kept_shape(id: u32, name: &str, text: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr name="{name}" id="{id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off y="20" x="10"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

fn deck(slide: &str) -> Vec<u8> {
    let parts: Vec<(String, Vec<u8>)> = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        ),
        (
            "ppt/presentation.xml",
            r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#,
        ),
        (
            "ppt/_rels/presentation.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
        ),
        ("ppt/slides/slide1.xml", slide),
        (
            "ppt/slides/_rels/slide1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
        ),
    ]
    .into_iter()
    .map(|(path, body)| (path.to_owned(), body.as_bytes().to_vec()))
    .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn slide_xml(body: &str) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{ROOT_START}<p:cSld name=\"Custom\">{body}</p:cSld><p:clrMapOvr><a:overrideClrMapping bg1=\"lt1\" tx1=\"dk1\"/></p:clrMapOvr></p:sld>"
    )
}

fn sp_tree(inner: &str) -> String {
    format!(
        "<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{inner}</p:spTree>"
    )
}

fn text_patch(text: &str) -> ShapePatch {
    ShapePatch {
        texts: vec![TextWrite {
            target: TextTarget::Body,
            autofit: None,
            paragraphs: vec![ParagraphWrite {
                source_index: Some(0),
                source_paragraphs: Vec::new(),
                preserved_xml: None,
                rebuild: true,
                properties_changed: false,
                alignment: None,
                level: 0,
                bullet: None,
                runs: vec![RunWrite {
                    text: text.to_owned(),
                    properties: RunProperties::default(),
                }],
            }],
        }],
        ..ShapePatch::default()
    }
}

fn save(bytes: &[u8], shapes: Vec<ShapeWrite>, background: Option<BackgroundWrite>) -> Vec<u8> {
    let package = parse_pptx(bytes).unwrap();
    write_pptx_with_edits(
        &package,
        &DeckWrite {
            slides: vec![SlideWrite::Patch {
                part_path: "ppt/slides/slide1.xml".to_owned(),
                shapes,
                background,
                layout_part_path: None,
            }],
            comments: None,
            notes: None,
        },
    )
    .unwrap()
}

fn saved_slide(saved: &[u8]) -> String {
    let parts: BTreeMap<String, Vec<u8>> =
        ooxml_opc::unzip_parts(saved).unwrap().into_iter().collect();
    String::from_utf8(parts["ppt/slides/slide1.xml"].clone()).unwrap()
}

#[test]
fn a_text_edit_keeps_the_untouched_shape_byte_identical() {
    let edited = shape(2, "Edited", "Before");
    let slide = slide_xml(&sp_tree(&format!(
        "{edited}<!-- keep this comment --><?keep-pi some data?>{UNTOUCHED}{ALTERNATE}"
    )));
    let saved = save(
        &deck(&slide),
        vec![
            ShapeWrite::Patch {
                source_index: 0,
                patch: Box::new(text_patch("Edited text")),
            },
            ShapeWrite::Keep { source_index: 1 },
            ShapeWrite::Keep { source_index: 2 },
        ],
        None,
    );
    let xml = saved_slide(&saved);
    assert!(xml.contains(UNTOUCHED), "{xml}");
    assert!(xml.contains(ALTERNATE), "{xml}");
    assert!(xml.contains("<!-- keep this comment -->"), "{xml}");
    assert!(xml.contains(r#"<?keep-pi some data?>"#), "{xml}");
    assert!(xml.contains(ROOT_START), "{xml}");
    assert!(xml.contains("Edited text"), "{xml}");
    assert!(!xml.contains("<a:t>Before</a:t>"), "{xml}");
}

#[test]
fn a_space_changed_to_more_whitespace_keeps_the_new_text() {
    for (text, expected) in [("  ", "<a:t>  </a:t>"), ("\t", "<a:t>\t</a:t>")] {
        let slide = slide_xml(&sp_tree(&shape(2, "Edited", " ")));
        let saved = save(
            &deck(&slide),
            vec![ShapeWrite::Patch {
                source_index: 0,
                patch: Box::new(text_patch(text)),
            }],
            None,
        );
        let xml = saved_slide(&saved);
        assert!(xml.contains(expected), "{xml}");
        assert!(!xml.contains("<a:t> </a:t>"), "{xml}");
    }
}

#[test]
fn a_space_typed_into_an_empty_text_element_is_written() {
    let source = shape(2, "Edited", "X").replace("<a:t>X</a:t>", "<a:t/>");
    let slide = slide_xml(&sp_tree(&source));
    let saved = save(
        &deck(&slide),
        vec![ShapeWrite::Patch {
            source_index: 0,
            patch: Box::new(text_patch(" ")),
        }],
        None,
    );
    let xml = saved_slide(&saved);
    assert!(xml.contains("<a:t> </a:t>"), "{xml}");
    assert!(!xml.contains("<a:t/>"), "{xml}");
}

#[test]
fn clearing_a_table_cell_space_leaves_an_empty_text_element() {
    let table = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="1000"/></a:tblGrid><a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t> </a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
    let slide = slide_xml(&sp_tree(table));
    let saved = save(
        &deck(&slide),
        vec![ShapeWrite::Patch {
            source_index: 0,
            patch: Box::new(ShapePatch {
                table_structure: Some(TableStructure {
                    generation: "test".to_owned(),
                    story_keys: Vec::new(),
                    grid: vec![1000],
                    columns: vec![0],
                    copied_columns: vec![false],
                    rows: vec![TableRowSource {
                        source_row: 0,
                        copied: false,
                        copied_cells: vec![false],
                        height: 370840,
                        cells: vec![[0, 0]],
                    }],
                }),
                texts: vec![TextWrite {
                    target: TextTarget::TableCell { row: 0, cell: 0 },
                    autofit: None,
                    paragraphs: vec![ParagraphWrite {
                        source_index: Some(0),
                        source_paragraphs: Vec::new(),
                        preserved_xml: None,
                        rebuild: true,
                        properties_changed: false,
                        alignment: None,
                        level: 0,
                        bullet: None,
                        runs: vec![RunWrite {
                            text: String::new(),
                            properties: RunProperties::default(),
                        }],
                    }],
                }],
                ..ShapePatch::default()
            }),
        }],
        None,
    );
    let xml = saved_slide(&saved);
    assert!(xml.contains("<a:t/>"), "{xml}");
    assert!(!xml.contains("<a:t> </a:t>"), "{xml}");
}

#[test]
fn an_added_and_a_removed_shape_keep_the_survivors_byte_identical() {
    let first = kept_shape(2, "First", "First text");
    let removed = shape(3, "Removed", "Removed text");
    let last = kept_shape(4, "Last", "Last text");
    let slide = slide_xml(&sp_tree(&format!(
        "<!-- first stays -->{first}{removed}<!-- last stays -->{last}"
    )));
    let saved = save(
        &deck(&slide),
        vec![
            ShapeWrite::Keep { source_index: 0 },
            ShapeWrite::Add(Box::new(ShapeAdd {
                name: "Added".to_owned(),
                geometry: "rect".to_owned(),
                text_box: false,
                x: 500,
                y: 600,
                width: 700,
                height: 800,
                adjust_values: BTreeMap::new(),
                fill: None,
                outline: None,
                paragraphs: Some(vec![ParagraphWrite {
                    source_index: None,
                    source_paragraphs: Vec::new(),
                    preserved_xml: None,
                    rebuild: true,
                    properties_changed: true,
                    alignment: None,
                    level: 0,
                    bullet: None,
                    runs: vec![RunWrite {
                        text: "Added text".to_owned(),
                        properties: RunProperties::default(),
                    }],
                }]),
                picture: None,
            })),
            ShapeWrite::Keep { source_index: 2 },
        ],
        None,
    );
    let xml = saved_slide(&saved);
    assert!(
        xml.contains(first.as_str()),
        "survivor bytes require the splice: {xml}"
    );
    assert!(
        xml.contains(last.as_str()),
        "survivor bytes require the splice: {xml}"
    );
    assert!(
        xml.contains("<!-- first stays -->"),
        "survivor bytes require the splice: {xml}"
    );
    assert!(
        xml.contains("<!-- last stays -->"),
        "survivor bytes require the splice: {xml}"
    );
    assert!(xml.contains("Added text"), "{xml}");
    assert!(!xml.contains("Removed"), "{xml}");
    assert_eq!(parse_pptx(&saved).unwrap().slides.len(), 1);
}

#[test]
fn a_placeholder_patch_keeps_the_other_shape_byte_identical() {
    let placed = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Placed"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Placed</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let other = kept_shape(3, "Other", "Other text");
    let slide = slide_xml(&sp_tree(&format!("{placed}<!-- other stays -->{other}")));
    let saved = save(
        &deck(&slide),
        vec![
            ShapeWrite::Patch {
                source_index: 0,
                patch: Box::new(ShapePatch {
                    placeholder: Some(Placeholder {
                        placeholder_type: Some("body".to_owned()),
                        index: Some(7),
                        orientation: None,
                        size: None,
                    }),
                    texts: text_patch("Placed text").texts,
                    ..ShapePatch::default()
                }),
            },
            ShapeWrite::Keep { source_index: 1 },
        ],
        None,
    );
    let xml = saved_slide(&saved);
    assert!(
        xml.contains(r#"<p:ph type="body" idx="7"/>"#),
        "placeholder bytes require the splice: {xml}"
    );
    assert!(xml.contains("Placed text"), "{xml}");
    assert!(
        xml.contains(other.as_str()),
        "survivor bytes require the splice: {xml}"
    );
    assert!(
        xml.contains("<!-- other stays -->"),
        "survivor bytes require the splice: {xml}"
    );
}

#[test]
fn a_background_patch_keeps_the_shapes_byte_identical() {
    let first = kept_shape(2, "First", "First text");
    let second = kept_shape(3, "Second", "Second text");
    let body = format!(
        "<p:bg><p:bgPr><a:solidFill><a:srgbClr val=\"112233\"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>{}",
        sp_tree(&format!(
            "<!-- first stays -->{first}<!-- second stays -->{second}"
        ))
    );
    let slide = slide_xml(&body);
    let saved = save(
        &deck(&slide),
        vec![
            ShapeWrite::Keep { source_index: 0 },
            ShapeWrite::Keep { source_index: 1 },
        ],
        Some(BackgroundWrite::Clear),
    );
    let xml = saved_slide(&saved);
    assert!(!xml.contains("<p:bg>"), "{xml}");
    assert!(
        xml.contains(first.as_str()),
        "survivor bytes require the splice: {xml}"
    );
    assert!(
        xml.contains(second.as_str()),
        "survivor bytes require the splice: {xml}"
    );
    assert!(
        xml.contains("<!-- first stays -->"),
        "survivor bytes require the splice: {xml}"
    );
    assert!(
        xml.contains("<!-- second stays -->"),
        "survivor bytes require the splice: {xml}"
    );
    assert_eq!(parse_pptx(&saved).unwrap().slides.len(), 1);
}
