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

const TABLE_ROW: &str = r#"<a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Same</a:t></a:r></a:p></a:txBody><a:tcPr marL="91440" anchor="ctr"></a:tcPr></a:tc></a:tr>"#;

fn table_slide(rows: usize) -> String {
    let table = format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="1000"/></a:tblGrid>{}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
        TABLE_ROW.repeat(rows)
    );
    slide_xml(&sp_tree(&table))
}

fn table_patch(row: usize, text: &str) -> ShapePatch {
    ShapePatch {
        texts: vec![TextWrite {
            target: TextTarget::TableCell { row, cell: 0 },
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

#[test]
fn a_comment_and_pi_before_a_removed_shape_survive() {
    let edited = shape(2, "Edited", "Before");
    let removed = shape(3, "Removed", "Removed text");
    let last = kept_shape(4, "Last", "Last text");
    let slide = slide_xml(&sp_tree(&format!(
        "{edited}<!-- before removed --><?keep-pi some data?>{removed}{last}"
    )));
    let saved = save(
        &deck(&slide),
        vec![
            ShapeWrite::Patch {
                source_index: 0,
                patch: Box::new(text_patch("Edited text")),
            },
            ShapeWrite::Keep { source_index: 2 },
        ],
        None,
    );
    let xml = saved_slide(&saved);
    assert!(xml.contains("<!-- before removed -->"), "{xml}");
    assert!(xml.contains("<?keep-pi some data?>"), "{xml}");
    assert!(
        xml.contains(last.as_str()),
        "survivor bytes require the splice: {xml}"
    );
    assert!(xml.contains("Edited text"), "{xml}");
    assert!(!xml.contains("Removed"), "{xml}");
}

#[test]
fn a_comment_between_two_removed_shapes_survives() {
    let first = kept_shape(2, "First", "First text");
    let removed_one = shape(3, "RemovedOne", "Removed one text");
    let removed_two = shape(4, "RemovedTwo", "Removed two text");
    let last = kept_shape(5, "Last", "Last text");
    let slide = slide_xml(&sp_tree(&format!(
        "{first}{removed_one}<!-- between removed -->{removed_two}{last}"
    )));
    let saved = save(
        &deck(&slide),
        vec![
            ShapeWrite::Patch {
                source_index: 0,
                patch: Box::new(text_patch("First edited")),
            },
            ShapeWrite::Keep { source_index: 3 },
        ],
        None,
    );
    let xml = saved_slide(&saved);
    assert!(xml.contains("<!-- between removed -->"), "{xml}");
    assert!(
        xml.contains(last.as_str()),
        "survivor bytes require the splice: {xml}"
    );
    assert!(xml.contains("First edited"), "{xml}");
    assert!(!xml.contains("RemovedOne"), "{xml}");
    assert!(!xml.contains("RemovedTwo"), "{xml}");
}

#[test]
fn a_removed_last_shape_keeps_the_following_comment() {
    let first = kept_shape(2, "First", "First text");
    let removed = shape(3, "Removed", "Removed text");
    let slide = slide_xml(&sp_tree(&format!("{first}{removed}<!-- tail stays -->")));
    let saved = save(
        &deck(&slide),
        vec![ShapeWrite::Patch {
            source_index: 0,
            patch: Box::new(text_patch("First edited")),
        }],
        None,
    );
    let xml = saved_slide(&saved);
    assert!(xml.contains("<!-- tail stays -->"), "{xml}");
    assert!(xml.contains("First edited"), "{xml}");
    assert!(!xml.contains("Removed"), "{xml}");
}

#[test]
fn a_comment_with_whitespace_before_a_removed_shape_survives() {
    let edited = shape(2, "Edited", "Before");
    let removed = shape(3, "Removed", "Removed text");
    let last = kept_shape(4, "Last", "Last text");
    let slide = slide_xml(&sp_tree(&format!(
        "{edited}\n  <!-- gap stays -->\n  {removed}\n  {last}\n"
    )));
    let saved = save(
        &deck(&slide),
        vec![
            ShapeWrite::Patch {
                source_index: 0,
                patch: Box::new(text_patch("Edited text")),
            },
            ShapeWrite::Keep { source_index: 2 },
        ],
        None,
    );
    let xml = saved_slide(&saved);
    assert!(xml.contains("<!-- gap stays -->"), "{xml}");
    assert!(
        xml.contains(last.as_str()),
        "survivor bytes require the splice: {xml}"
    );
    assert!(!xml.contains("Removed"), "{xml}");
}

#[test]
fn reordered_untouched_shapes_keep_their_source_bytes() {
    let first = kept_shape(2, "Alpha", "Alpha text");
    let second = kept_shape(3, "Bravo", "Bravo text");
    let third = kept_shape(4, "Charlie", "Charlie text");
    let fourth = kept_shape(5, "Delta", "Delta text");
    let edited = shape(6, "Edited", "Before");
    let slide = slide_xml(&sp_tree(&format!(
        "{first}{second}{third}{fourth}{edited}"
    )));
    let saved = save(
        &deck(&slide),
        vec![
            ShapeWrite::Keep { source_index: 3 },
            ShapeWrite::Keep { source_index: 2 },
            ShapeWrite::Keep { source_index: 1 },
            ShapeWrite::Keep { source_index: 0 },
            ShapeWrite::Patch {
                source_index: 4,
                patch: Box::new(text_patch("Edited text")),
            },
        ],
        None,
    );
    let xml = saved_slide(&saved);
    for survivor in [&first, &second, &third, &fourth] {
        assert!(
            xml.contains(survivor.as_str()),
            "reordered survivor bytes require the splice: {xml}"
        );
    }
    assert!(xml.contains("Edited text"), "{xml}");
}

#[test]
fn a_reordered_edited_shape_keeps_the_other_shapes_byte_identical() {
    let first = kept_shape(2, "Alpha", "Alpha text");
    let second = kept_shape(3, "Bravo", "Bravo text");
    let third = kept_shape(4, "Charlie", "Charlie text");
    let fourth = kept_shape(5, "Delta", "Delta text");
    let slide = slide_xml(&sp_tree(&format!(
        "{first}{second}{third}{fourth}"
    )));
    let saved = save(
        &deck(&slide),
        vec![
            ShapeWrite::Keep { source_index: 3 },
            ShapeWrite::Patch {
                source_index: 2,
                patch: Box::new(text_patch("Charlie edited")),
            },
            ShapeWrite::Keep { source_index: 1 },
            ShapeWrite::Keep { source_index: 0 },
        ],
        None,
    );
    let xml = saved_slide(&saved);
    for survivor in [&first, &second, &fourth] {
        assert!(
            xml.contains(survivor.as_str()),
            "reordered survivor bytes require the splice: {xml}"
        );
    }
    assert!(xml.contains("Charlie edited"), "{xml}");
    assert!(!xml.contains("<a:t>Charlie text</a:t>"), "{xml}");
}

#[test]
fn a_large_table_of_identical_rows_splices_and_scales() {
    use std::time::Instant;

    let edited_row = TABLE_ROW.replace("<a:t>Same</a:t>", "<a:t>Edited</a:t>");
    let mut durations = Vec::new();
    for rows in [2_000usize, 4_000] {
        let bytes = deck(&table_slide(rows));
        let mut best = None::<std::time::Duration>;
        let mut saved = Vec::new();
        for _ in 0..3 {
            let started = Instant::now();
            saved = save(
                &bytes,
                vec![ShapeWrite::Patch {
                    source_index: 0,
                    patch: Box::new(table_patch(rows / 2, "Edited")),
                }],
                None,
            );
            let elapsed = started.elapsed();
            best = Some(best.map_or(elapsed, |current| current.min(elapsed)));
        }
        let xml = saved_slide(&saved);
        assert_eq!(xml.matches(TABLE_ROW).count(), rows - 1, "{rows} rows");
        assert_eq!(xml.matches(edited_row.as_str()).count(), 1, "{rows} rows");
        durations.push(best.unwrap());
    }
    let ratio = durations[1].as_secs_f64() / durations[0].as_secs_f64();
    assert!(ratio < 2.8, "scaling ratio {ratio} for {durations:?}");
}

#[test]
fn deleting_one_of_two_shapes_in_an_alternate_content_choice_keeps_the_other() {
    for fallback_id in [4, 3] {
        let deleted = kept_shape(2, "Deleted", "Deleted text");
        let kept = kept_shape(3, "Kept", "Kept text");
        let fallback = shape(fallback_id, "Fallback", "Fallback text");
        let slide = slide_xml(&sp_tree(&format!(
            r#"<mc:AlternateContent><mc:Choice xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" Requires="a">{deleted}<!-- inside wrapper -->{kept}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#
        )));
        let saved = save(
            &deck(&slide),
            vec![ShapeWrite::Keep { source_index: 1 }],
            None,
        );
        let xml = saved_slide(&saved);
        assert!(
            xml.contains(kept.as_str()),
            "kept shape bytes require the splice (fallback id {fallback_id}): {xml}"
        );
        assert!(xml.contains("<!-- inside wrapper -->"), "{xml}");
        assert!(!xml.contains("Deleted"), "{xml}");
        assert_eq!(parse_pptx(&saved).unwrap().slides.len(), 1);
    }
}

#[test]
fn a_moved_one_object_alternate_content_wrapper_keeps_its_bytes() {
    let wrapped = kept_shape(3, "Wrapped", "Wrapped text");
    let fallback = shape(3, "WrappedFallback", "Wrapped fallback text");
    let wrapper = format!(
        r#"<mc:AlternateContent><mc:Choice xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" Requires="a">{wrapped}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#
    );
    let first = kept_shape(4, "First", "First text");
    let edited = shape(5, "Edited", "Before");
    let slide = slide_xml(&sp_tree(&format!("{first}{wrapper}{edited}")));
    let saved = save(
        &deck(&slide),
        vec![
            ShapeWrite::Keep { source_index: 1 },
            ShapeWrite::Keep { source_index: 0 },
            ShapeWrite::Patch {
                source_index: 2,
                patch: Box::new(text_patch("Edited text")),
            },
        ],
        None,
    );
    let xml = saved_slide(&saved);
    assert!(
        xml.contains(wrapper.as_str()),
        "moved wrapper bytes require the splice: {xml}"
    );
    assert!(
        xml.contains(first.as_str()),
        "survivor bytes require the splice: {xml}"
    );
    assert!(xml.contains("Edited text"), "{xml}");
}
