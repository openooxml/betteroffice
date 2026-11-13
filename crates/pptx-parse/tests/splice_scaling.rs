use std::collections::BTreeMap;

use pptx_parse::test_counters::{self, SpliceWork};
use pptx_parse::{
    DeckWrite, ParagraphWrite, RunProperties, RunWrite, ShapePatch, ShapeWrite, SlideWrite,
    TextTarget, TextWrite, parse_pptx, write_pptx_with_edits,
};

const ROWS: [usize; 3] = [2_000, 4_000, 8_000];
const MAX_CONSTANT_WORK: usize = 64;

const TABLE_ROW: &str = r#"<a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Same</a:t></a:r></a:p></a:txBody><a:tcPr marL="91440" anchor="ctr"></a:tcPr></a:tc></a:tr>"#;

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

fn table_slide(rows: usize) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:cSld name="Custom"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="1000"/></a:tblGrid>{}</a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1"/></p:clrMapOvr></p:sld>"#,
        TABLE_ROW.repeat(rows)
    )
}

fn table_edit(row: usize, text: &str) -> DeckWrite {
    let patch = ShapePatch {
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
    };
    DeckWrite {
        slides: vec![SlideWrite::Patch {
            part_path: "ppt/slides/slide1.xml".to_owned(),
            shapes: vec![ShapeWrite::Patch {
                source_index: 0,
                patch: Box::new(patch),
            }],
            background: None,
            layout_part_path: None,
        }],
        comments: None,
        notes: None,
    }
}

fn splice_work(rows: usize) -> SpliceWork {
    let package = parse_pptx(&deck(&table_slide(rows))).unwrap();
    let edit = table_edit(rows / 2, "Edited");
    test_counters::take();
    let saved = write_pptx_with_edits(&package, &edit).unwrap();
    let work = test_counters::take();
    let parts: BTreeMap<String, Vec<u8>> = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .collect();
    let xml = String::from_utf8(parts["ppt/slides/slide1.xml"].clone()).unwrap();
    let edited_row = TABLE_ROW.replace("<a:t>Same</a:t>", "<a:t>Edited</a:t>");
    assert_eq!(xml.matches(TABLE_ROW).count(), rows - 1, "{rows} rows");
    assert_eq!(xml.matches(edited_row.as_str()).count(), 1, "{rows} rows");
    work
}

fn assert_doubles(name: &str, works: &[SpliceWork; 3], count: fn(&SpliceWork) -> usize) -> i64 {
    let [small, middle, large] = works.map(|work| count(&work) as i64);
    let (first, second) = (middle - small, large - middle);
    assert!(first > 0, "{name}: {works:?}");
    assert_eq!(second, 2 * first, "{name}: {works:?}");
    first
}

#[test]
fn editing_one_row_of_a_large_table_splices_in_linear_work() {
    let works = ROWS.map(splice_work);
    for (rows, work) in ROWS.iter().zip(&works) {
        println!("{rows} rows: {work:?}");
    }
    assert_eq!(splice_work(ROWS[0]), works[0], "{works:?}");
    let [small, middle, large] = works;

    assert_eq!(small.dp_cells, middle.dp_cells, "{works:?}");
    assert_eq!(middle.dp_cells, large.dp_cells, "{works:?}");
    assert!(small.dp_cells <= MAX_CONSTANT_WORK, "{works:?}");

    assert_eq!(
        small.fragment_elements, middle.fragment_elements,
        "{works:?}"
    );
    assert_eq!(
        middle.fragment_elements, large.fragment_elements,
        "{works:?}"
    );
    assert!(small.fragment_elements > 0, "{works:?}");
    assert!(small.fragment_elements <= MAX_CONSTANT_WORK, "{works:?}");

    let added_equal_calls = assert_doubles("equal_calls", &works, |work| work.equal_calls);
    assert!(
        added_equal_calls <= 2 * (ROWS[1] - ROWS[0]) as i64,
        "{works:?}"
    );
    assert_doubles("hashed_nodes", &works, |work| work.hashed_nodes);
}
