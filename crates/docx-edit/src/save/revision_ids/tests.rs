use super::*;
use docx_parse::block::BlockSdt;
use docx_parse::write_docx_s13;
use docx_parse::xml::{ParseBudget, ParseLimits, XmlDocument, XmlElement, parse_xml};
use serde_json::json;

const FIRST: u64 = (1 << 53) - 1;
const SECOND: u64 = 1 << 40;
const DOCUMENT: &str = "word/document.xml";

fn request(content: Value) -> S13SaveRequest {
    serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2030-01-02T03:04:05.006Z"},
        "document": {"content": content},
        "options": {"updateModifiedDate": false}
    }))
    .unwrap()
}

fn change(id: u64) -> Value {
    json!({"type": "insertion", "info": {"id": id, "author": "Reviewer"}, "content": []})
}

fn write_document(request: S13SaveRequest) -> XmlDocument {
    let source = ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            DOCUMENT.to_owned(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#.to_vec(),
        ),
    ])
    .unwrap();
    let saved = write_docx_s13(request, &source).unwrap();
    let (_, document) = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(part, _)| part == DOCUMENT)
        .unwrap();
    let limits = ParseLimits::default();
    parse_xml(&document, DOCUMENT, &mut ParseBudget::new(&limits)).unwrap()
}

fn document_ids(document: &XmlDocument) -> BTreeMap<String, Vec<String>> {
    fn visit(element: &XmlElement, path: String, ids: &mut BTreeMap<String, Vec<String>>) {
        for child in element.child_elements() {
            visit(child, format!("{path}/{}", child.local_name()), ids);
        }
        if let Some(id) = element.attributes.get("w:id") {
            ids.entry(path).or_default().push(id.clone());
        }
    }

    let body = document.root().unwrap().child("w", "body").unwrap();
    let mut ids = BTreeMap::new();
    for element in body.child_elements() {
        visit(element, element.local_name().to_owned(), &mut ids);
    }
    ids
}

fn assert_assignment(
    original: S13SaveRequest,
    sites: &[(&str, u64, u32, &str)],
    source_ids: &[(&str, u64)],
) {
    let mut expected = serde_json::to_value(&original).unwrap();
    let mut assignments = BTreeMap::new();
    let mut written: BTreeMap<String, Vec<String>> = BTreeMap::new();
    let mut used = BTreeSet::new();
    for &(xml_path, id) in source_ids {
        used.insert(id);
        written
            .entry(xml_path.to_owned())
            .or_default()
            .push(id.to_string());
    }
    for &(path, placeholder, assigned, xml_path) in sites {
        let value = expected.pointer_mut(path).unwrap();
        assert_eq!(value.as_f64(), Some(placeholder as f64), "{path}");
        *value = json!(f64::from(assigned));
        assert!((1..MAX_ID).contains(&assigned));
        if let Some(prior) = assignments.insert(placeholder, assigned) {
            assert_eq!(prior, assigned, "{path}");
        } else {
            assert!(used.insert(u64::from(assigned)), "reused ID at {path}");
        }
        written
            .entry(xml_path.to_owned())
            .or_default()
            .push(assigned.to_string());
    }
    let expected: S13SaveRequest = serde_json::from_value(expected).unwrap();
    let session = assignments.keys().copied().collect();
    for _ in 0..2 {
        let mut actual = original.clone();
        assert_eq!(
            assign_session_revision_ids(&mut actual, &session).unwrap(),
            assignments
        );
        assert_eq!(actual, expected);
        assert_eq!(document_ids(&write_document(actual)), written);
    }
}

#[test]
fn keeps_out_of_range_source_revision_ids_without_raising_the_base() {
    let limit = u64::from(MAX_ID);
    let large = 4_294_967_296u64;
    let huge = 1_000_000_000_000_000u64;
    let info = |id| json!({"id": id, "author": "Source"});
    let original = request(json!([
        {
            "type": "paragraph",
            "propertyChanges": [{"type": "paragraphPropertyChange", "info": info(13)}],
            "content": [
                {"type": "insertion", "info": info(limit), "content": []},
                {"type": "deletion", "info": info(large), "content": []},
                {"type": "moveFrom", "info": info(huge), "content": []},
                {"type": "moveTo", "info": info(11), "content": []},
                {"type": "run", "propertyChanges": [
                    {"type": "runPropertyChange", "info": info(12)}
                ], "content": [{"type": "text", "text": "Source"}]},
                {"type": "bookmarkStart", "id": 14, "name": "SourceBookmark"},
                change(FIRST),
                {"type": "deletion", "info": {"id": FIRST, "author": "Reviewer"}, "content": []},
                {"type": "bookmarkEnd", "id": 14}
            ]
        },
        {
            "type": "table",
            "columnWidths": [2000],
            "propertyChanges": [{"type": "tablePropertyChange", "info": info(15)}],
            "rows": [{
                "type": "tableRow",
                "propertyChanges": [{"type": "tableRowPropertyChange", "info": info(16)}],
                "structuralChange": {"type": "tableRowInsertion", "info": info(17)},
                "cells": [{
                    "type": "tableCell",
                    "propertyChanges": [{"type": "tableCellPropertyChange", "info": info(18)}],
                    "structuralChange": {"type": "tableCellInsertion", "info": info(19)},
                    "content": [{"type": "paragraph", "content": [change(SECOND)]}]
                }]
            }]
        }
    ]));
    assert_assignment(
        original,
        &[
            ("/document/content/0/content/6/info/id", FIRST, 20, "p/ins"),
            ("/document/content/0/content/7/info/id", FIRST, 20, "p/del"),
            (
                "/document/content/1/rows/0/cells/0/content/0/content/0/info/id",
                SECOND,
                21,
                "tbl/tr/tc/p/ins",
            ),
        ],
        &[
            ("p/ins", limit),
            ("p/del", large),
            ("p/moveFrom", huge),
            ("p/moveTo", 11),
            ("p/r/rPr/rPrChange", 12),
            ("p/pPr/pPrChange", 13),
            ("p/bookmarkStart", 14),
            ("p/bookmarkEnd", 14),
            ("tbl/tblPr/tblPrChange", 15),
            ("tbl/tr/trPr/trPrChange", 16),
            ("tbl/tr/trPr/ins", 17),
            ("tbl/tr/tc/tcPr/tcPrChange", 18),
            ("tbl/tr/tc/tcPr/cellIns", 19),
        ],
    );
}

#[test]
fn bases_assignments_on_kept_raw_xml_and_starts_at_one_without_reservations() {
    for (raw, source_ids, first) in [
        (None, &[][..], 1),
        (
            Some(json!({"type": "rawXml", "xml": "<w:customXml w:id='40'/>"})),
            &[("customXml", 40)][..],
            41,
        ),
    ] {
        let mut content = vec![json!({
            "type": "paragraph",
            "content": [change(FIRST), change(SECOND), change(FIRST)]
        })];
        content.extend(raw);
        assert_assignment(
            request(json!(content)),
            &[
                (
                    "/document/content/0/content/0/info/id",
                    FIRST,
                    first,
                    "p/ins",
                ),
                (
                    "/document/content/0/content/1/info/id",
                    SECOND,
                    first + 1,
                    "p/ins",
                ),
                (
                    "/document/content/0/content/2/info/id",
                    FIRST,
                    first,
                    "p/ins",
                ),
            ],
            source_ids,
        );
    }
}

#[test]
fn ignores_dropped_run_property_attributes_when_assigning_ids() {
    let mut request = request(json!([{
        "type": "paragraph",
        "content": [{
            "type": "insertion",
            "info": {"id": FIRST, "author": "Reviewer"},
            "content": [{
                "type": "run",
                "propertyChanges": [{
                    "type": "runPropertyChange",
                    "info": {"id": FIRST, "author": "Reviewer"},
                    "previousAttributes": {"witness": "w:id='40'"}
                }],
                "content": [{"type": "text", "text": "Revised text"}]
            }]
        }]
    }]));
    let typed = serde_json::to_value(&request).unwrap();
    assert!(
        typed
            .pointer("/document/content/0/content/0/content/0/propertyChanges/0/previousAttributes")
            .is_none()
    );
    let assigned = assign_session_revision_ids(&mut request, &BTreeSet::from([FIRST])).unwrap();
    assert_eq!(assigned, BTreeMap::from([(FIRST, 1)]));

    let document = write_document(request);
    assert_eq!(
        document_ids(&document),
        BTreeMap::from([
            ("p/ins".to_owned(), vec!["1".to_owned()]),
            ("p/ins/r/rPr/rPrChange".to_owned(), vec!["1".to_owned()]),
        ])
    );
    let insertion = document
        .root()
        .unwrap()
        .child("w", "body")
        .unwrap()
        .child("w", "p")
        .unwrap()
        .child("w", "ins")
        .unwrap();
    assert_eq!(insertion.text_content(), "Revised text");
}

#[test]
fn assigns_inline_revisions_before_paragraph_marks_and_properties() {
    let mark = FIRST - 1;
    let deleted_mark = FIRST - 2;
    let properties = FIRST - 3;
    for shared in [false, true] {
        let mut content = vec![change(FIRST), change(SECOND)];
        let mut sites = vec![
            ("/document/content/0/content/0/info/id", FIRST, 8, "p/ins"),
            ("/document/content/0/content/1/info/id", SECOND, 9, "p/ins"),
            ("/document/content/0/pPrIns/id", mark, 10, "p/pPr/rPr/ins"),
            (
                "/document/content/0/pPrDel/id",
                deleted_mark,
                11,
                "p/pPr/rPr/del",
            ),
            (
                "/document/content/0/propertyChanges/0/info/id",
                properties,
                12,
                "p/pPr/pPrChange",
            ),
        ];
        if shared {
            content.push(change(mark));
            sites.push(("/document/content/0/content/2/info/id", mark, 10, "p/ins"));
        }
        content.extend([
            json!({"type": "bookmarkStart", "id": 7, "name": "SourceBookmark"}),
            json!({"type": "bookmarkEnd", "id": 7}),
        ]);
        let original = request(json!([{
            "type": "paragraph",
            "pPrIns": {"id": mark, "author": "Reviewer"},
            "pPrDel": {"id": deleted_mark, "author": "Reviewer"},
            "propertyChanges": [{
                "type": "paragraphPropertyChange", "info": {"id": properties, "author": "Reviewer"}
            }],
            "content": content
        }]));
        assert_assignment(
            original,
            &sites,
            &[("p/bookmarkStart", 7), ("p/bookmarkEnd", 7)],
        );
    }
}

#[test]
fn notes_serialized_annotations_and_comments() {
    let found = collect(
        &json!([
            {"type": "bookmarkStart", "id": 7},
            {"type": "bookmarkEnd", "id": 8},
            {"type": "moveFromRangeStart", "id": 9},
            {"type": "moveFromRangeEnd", "id": 10},
            {"type": "moveToRangeStart", "id": 11},
            {"type": "moveToRangeEnd", "id": 12},
            {"type": "commentRangeStart", "id": 13},
            {"type": "commentRangeEnd", "id": 14},
            {"type": "commentReference", "id": 15},
            {"type": "commentReference", "author": "A", "id": 16, "revisionId": 17},
            {"author": "", "id": 18, "revisionId": 19, "content": []},
            {"author": "A", "id": 20, "revisionId": 0, "content": {}},
            {"author": "A", "id": 21, "revision_id": 99},
            {"type": "bookmarkStart", "author": false, "id": 22},
            {"type": "footnoteRef", "id": 99},
            {"author": "A", "id": 2147483648u32, "revisionId": 2147483649u32},
            {"author": "A", "id": -1, "revisionId": 1.5},
            {"author": "A", "id": "99", "revisionId": null},
            {"author": null, "id": 99},
            {"type": "commentRangeStart", "id": 100},
            {"type": "commentRangeEnd", "id": 101},
            {"type": "commentReference", "id": 102},
            {"author": "A", "id": 103, "revisionId": 104, "content": []},
        ]),
        &BTreeSet::new(),
    )
    .unwrap();
    assert_eq!(found.max, 22);
    assert_eq!(
        found.used,
        [0].into_iter().chain(7..=22).chain(100..=104).collect()
    );
}

#[test]
fn scans_xml_in_values_with_javascript_boundaries() {
    let found = collect(
        &json!({
            "w:id = '99'": [
                "w:id=\u{feff}\"00031' éw:id\u{a0}= '42\"",
                "aw:id='90' _w:id='91' w:id = '-92' w:id\u{85}='93'",
                "w:id='2147483648' w:id='123456789012345678901234567890'",
                "w:id=95 w:id=\"96 w:id='' w:identifier='97'",
            ]
        }),
        &BTreeSet::new(),
    )
    .unwrap();
    assert_eq!(found.max, 42);
    assert_eq!(found.used, BTreeSet::from([31, 42]));
}

#[test]
fn uses_serialized_names_and_does_not_descend_into_session_infos() {
    #[derive(Serialize)]
    struct Record {
        author: String,
        id: u32,
        #[serde(rename = "revisionId")]
        revision_id: u32,
    }

    let found = collect(
        &Record {
            author: "A".into(),
            id: 3,
            revision_id: 5,
        },
        &BTreeSet::from([3]),
    )
    .unwrap();
    assert!(found.sites.is_empty());
    assert_eq!(found.used, BTreeSet::from([3, 5]));

    let infos = [
        docx_parse::TrackedChangeInfo {
            id: 3.0,
            author: "w:id='55'".into(),
            date: None,
        },
        docx_parse::TrackedChangeInfo {
            id: 5.0,
            author: "A".into(),
            date: Some("w:id='12'".into()),
        },
    ];
    let found = collect(&infos, &BTreeSet::from([3])).unwrap();
    assert_eq!(found.sites, BTreeMap::from([(3, 1)]));
    assert_eq!(found.max, 12);
    assert_eq!(found.used, BTreeSet::from([5, 12]));
}

#[test]
fn assigns_in_site_order_and_shares_one_id_across_info_types() {
    let mut request = request(json!([
        {"type": "rawXml", "xml": "<w:customXml w:id='40'/>"},
        {
            "type": "paragraph",
            "content": [
                change(FIRST),
                {"type": "run", "propertyChanges": [
                    {"type": "runPropertyChange", "info": {"id": SECOND, "author": "A"}},
                    {"type": "runPropertyChange", "info": {"id": FIRST, "author": "A"}}
                ], "content": []},
                change(FIRST),
                change(11),
                {"type": "commentRangeStart", "id": 41}
            ]
        }
    ]));
    let ids = assign_session_revision_ids(&mut request, &BTreeSet::from([FIRST, SECOND])).unwrap();
    assert_eq!(ids, BTreeMap::from([(FIRST, 42), (SECOND, 43)]));
    let result = serde_json::to_value(&request).unwrap();
    let content = &result["document"]["content"][1]["content"];
    assert_eq!(content[0]["info"]["id"].as_f64(), Some(42.0));
    assert_eq!(
        content[1]["propertyChanges"][0]["info"]["id"].as_f64(),
        Some(43.0)
    );
    assert_eq!(
        content[1]["propertyChanges"][1]["info"]["id"].as_f64(),
        Some(42.0)
    );
    assert_eq!(content[2]["info"]["id"].as_f64(), Some(42.0));
    assert_eq!(content[3]["info"]["id"].as_f64(), Some(11.0));
}

#[test]
fn assigns_after_content_in_projected_paragraph_and_table_order() {
    let ids: Vec<_> = (FIRST - 10..=FIRST).rev().collect();
    let info = |index: usize| json!({"id": ids[index], "author": "A"});
    let mut request = request(json!([
        {
            "type": "paragraph",
            "propertyChanges": [{"type": "paragraphPropertyChange", "info": info(4)}],
            "pPrIns": info(2),
            "pPrDel": info(3),
            "content": [
                {"type": "run", "propertyChanges": [
                    {"type": "runPropertyChange", "info": info(0)}
                ], "content": []},
                change(ids[1])
            ]
        },
        {
            "type": "table",
            "propertyChanges": [{"type": "tablePropertyChange", "info": info(10)}],
            "rows": [{
                "type": "tableRow",
                "propertyChanges": [{"type": "tableRowPropertyChange", "info": info(9)}],
                "structuralChange": {"type": "tableRowInsertion", "info": info(8)},
                "cells": [{
                    "type": "tableCell",
                    "propertyChanges": [{"type": "tableCellPropertyChange", "info": info(7)}],
                    "structuralChange": {"type": "tableCellInsertion", "info": info(6)},
                    "content": [{"type": "paragraph", "content": [change(ids[5])]}]
                }]
            }]
        }
    ]));
    let assigned =
        assign_session_revision_ids(&mut request, &ids.iter().copied().collect()).unwrap();
    assert_eq!(assigned, ids.into_iter().zip(1..=11).collect());
}

#[test]
fn wraps_and_skips_used_ids() {
    let mut request = request(json!([{
        "type": "paragraph",
        "content": [change(u64::from(MAX_ID - 2)), change(FIRST), change(1), change(SECOND)]
    }]));
    let ids = assign_session_revision_ids(&mut request, &BTreeSet::from([FIRST, SECOND])).unwrap();
    assert_eq!(ids, BTreeMap::from([(FIRST, MAX_ID - 1), (SECOND, 2)]));
}

#[test]
fn errors_on_returning_to_the_base() {
    let mut next = 3;
    assert_eq!(
        next_id(&mut next, 5, &BTreeSet::from([4])),
        Err(RevisionIdsError::Exhausted)
    );
}

#[test]
fn leaves_request_with_no_session_sites_unchanged() {
    let mut request = request(json!([{"type": "paragraph", "content": [change(9)]}]));
    let original = request.clone();
    for session in [BTreeSet::new(), BTreeSet::from([FIRST])] {
        assert!(
            assign_session_revision_ids(&mut request, &session)
                .unwrap()
                .is_empty()
        );
        assert_eq!(request, original);
    }
}

#[test]
fn accepts_assigned_ids_that_equal_placeholders() {
    let mut request = request(json!([{"type": "paragraph", "content": [change(2), change(1)]}]));
    let ids = assign_session_revision_ids(&mut request, &BTreeSet::from([1, 2])).unwrap();
    assert_eq!(ids, BTreeMap::from([(1, 2), (2, 1)]));
    let result = serde_json::to_value(&request).unwrap();
    let content = result["document"]["content"][0]["content"].as_array().unwrap();
    assert_eq!(
        content.iter().map(|site| site["info"]["id"].as_f64()).collect::<Vec<_>>(),
        [Some(1.0), Some(2.0)]
    );
}

#[test]
fn clones_only_shared_blocks_on_a_revision_site_path() {
    let original = request(json!([
        {"type": "paragraph", "content": [{"type": "run", "content": []}]},
        {"type": "table", "rows": [{"type": "tableRow", "cells": [{
            "type": "tableCell", "content": [
                {"type": "paragraph", "content": []},
                {"type": "paragraph", "content": [change(FIRST)]}
            ]
        }]}]}
    ]));
    let session = BTreeSet::from([FIRST]);
    let mut actual = original.clone();
    assert_eq!(
        assign_session_revision_ids(&mut actual, &session).unwrap(),
        BTreeMap::from([(FIRST, 1)])
    );
    let [BlockContent::Paragraph(unchanged), BlockContent::Table(table)] =
        original.document.content.as_slice()
    else {
        panic!("expected a paragraph and a table");
    };
    let [BlockContent::Paragraph(actual_unchanged), BlockContent::Table(actual_table)] =
        actual.document.content.as_slice()
    else {
        panic!("expected a paragraph and a table");
    };
    assert!(Arc::ptr_eq(unchanged, actual_unchanged));
    assert!(!Arc::ptr_eq(table, actual_table));
    let [BlockContent::Paragraph(cell_unchanged), BlockContent::Paragraph(cell_changed)] =
        table.rows[0].cells[0].content.as_slice()
    else {
        panic!("expected two cell paragraphs");
    };
    let [BlockContent::Paragraph(actual_cell_unchanged), BlockContent::Paragraph(actual_cell_changed)] =
        actual_table.rows[0].cells[0].content.as_slice()
    else {
        panic!("expected two cell paragraphs");
    };
    assert!(Arc::ptr_eq(cell_unchanged, actual_cell_unchanged));
    assert!(!Arc::ptr_eq(cell_changed, actual_cell_changed));

    let mut probe = original.clone();
    let found = collect(&probe, &session).unwrap();
    let mut walker = SiteWalker::new(&session, None, &found.blocks_with_sites);
    walker.request(&mut probe);
    assert_eq!(walker.block_make_mut, 2);
}

#[test]
fn bounds_visits_to_nested_shared_blocks() {
    std::thread::Builder::new()
        .stack_size(32 * 1024 * 1024)
        .spawn(|| {
            for depth in [50, 100] {
                let mut original = request(json!([
                    {"type": "paragraph", "content": []},
                    {"type": "paragraph", "content": [change(FIRST)]}
                ]));
                let mut block = original.document.content.pop().unwrap();
                let sibling = original.document.content.pop().unwrap();
                let properties: docx_parse::inline::SdtProperties =
                    serde_json::from_value(json!({"sdtType": "richText"})).unwrap();
                for _ in 0..depth {
                    block = BlockContent::BlockSdt(Arc::new(BlockSdt {
                        node_type: "blockSdt".to_owned(),
                        properties: properties.clone(),
                        content: vec![sibling.clone(), block],
                    }));
                }
                original.document.content.push(block);
                let mut actual = original.clone();

                BLOCK_VISITS.set(0);
                assert_eq!(
                    assign_session_revision_ids(&mut actual, &BTreeSet::from([FIRST])).unwrap(),
                    BTreeMap::from([(FIRST, 1)])
                );
                let visits = BLOCK_VISITS.get();
                assert!(visits <= 3 * (2 * depth + 1), "{visits} visits at depth {depth}");
            }
        })
        .unwrap()
        .join()
        .unwrap();
}

#[test]
fn inserts_each_nested_model_reservation_once() {
    std::thread::Builder::new()
        .stack_size(32 * 1024 * 1024)
        .spawn(|| {
            for count in [1_000u32, 2_000] {
                let mut shape: Shape = serde_json::from_value(model_shape(json!([]))).unwrap();
                let mut model = Value::Null;
                for id in (1..=count).rev() {
                    model = Value::Object(
                        [
                            ("author".to_owned(), json!("Source")),
                            ("id".to_owned(), json!(id)),
                            ("info".to_owned(), model),
                        ]
                        .into_iter()
                        .collect(),
                    );
                }
                shape.drawing_model = Some(Value::Object(
                    [("info".to_owned(), model)].into_iter().collect(),
                ));
                let found = collect(&shape, &BTreeSet::from([FIRST])).unwrap();
                assert_eq!(found.used, (1..=count).collect());
                assert_eq!(found.reservation_inserts, count as usize);
            }
        })
        .unwrap()
        .join()
        .unwrap();
}

fn model_shape(content: Value) -> Value {
    json!({
        "type": "shape", "shapeType": "rect", "size": {"width": 1, "height": 1},
        "textBody": {"content": content}
    })
}

fn assert_text_box_model_order(child: Value, placeholders: &[u64]) {
    let shape = model_shape(json!([{"type": "paragraph", "content": [child]}]));
    let mut request = request(json!([{"type": "paragraph", "content": [{
        "type": "run", "content": [{"type": "shape", "shape": shape}]
    }]}]));
    let mut paths = Vec::new();
    let ids = assign_session_revision_ids_with_paths(
        &mut request,
        &placeholders.iter().copied().collect(),
        &mut paths,
    )
    .unwrap();
    let result = serde_json::to_value(&request).unwrap();
    for (index, ((path, id), &placeholder)) in paths.iter().zip(placeholders).enumerate() {
        let assigned = index as u32 + 1;
        assert_eq!(*id, placeholder, "{path}");
        assert_eq!(ids[&placeholder], assigned, "{path}");
        assert_eq!(result.pointer(path).and_then(Value::as_u64), Some(u64::from(assigned)));
    }
    assert_eq!(paths.len(), placeholders.len());
}

#[test]
fn assigns_text_box_field_inline_sites_before_result_blocks() {
    assert_text_box_model_order(
        json!({
            "type": "complexField", "instruction": "PAGE", "fieldType": "PAGE",
            "fieldCode": [], "fieldResult": [],
            "structuredResult": {
                "inline": [change(FIRST)],
                "blocks": [{"type": "paragraph", "content": [change(SECOND)]}]
            }
        }),
        &[FIRST, SECOND],
    );
}

#[test]
fn assigns_text_box_field_tree_code_and_result_before_child_trees() {
    assert_text_box_model_order(
        json!({
            "type": "complexField", "instruction": "PAGE", "fieldType": "PAGE",
            "fieldCode": [], "fieldResult": [],
            "fieldTree": {
                "code": {"inline": [change(FIRST)]},
                "result": {"inline": [change(SECOND)]},
                "children": [{"result": {"inline": [change(FIRST - 1)]}}]
            }
        }),
        &[FIRST, SECOND, FIRST - 1],
    );
}

#[test]
fn assigns_text_box_hyperlink_children_before_structured_children() {
    assert_text_box_model_order(
        json!({
            "type": "hyperlink", "anchor": "target",
            "children": [change(FIRST)], "structuredChildren": [change(SECOND)]
        }),
        &[FIRST, SECOND],
    );
}

#[test]
fn assigns_text_box_scene_shape_before_child_nodes() {
    let with_revision = |id| model_shape(json!([{"type": "paragraph", "content": [change(id)]}]));
    let mut shape = model_shape(json!([]));
    shape["scene"] = json!({"root": {
        "shape": with_revision(FIRST),
        "children": [{"shape": with_revision(SECOND)}]
    }});
    assert_text_box_model_order(
        json!({"type": "run", "content": [{"type": "shape", "shape": shape}]}),
        &[FIRST, SECOND],
    );
}

#[test]
fn finds_revision_infos_in_serialized_text_boxes() {
    let mut request = request(json!([{"type": "paragraph", "content": [
        {"type": "run", "content": [{"type": "shape", "shape": {
            "type": "shape", "shapeType": "rect", "size": {"width": 1, "height": 1},
            "textBox": true, "textBody": {"content": [
                {"type": "paragraph", "content": [change(FIRST)]}
            ]}
        }}]}
    ]}]));
    let ids = assign_session_revision_ids(&mut request, &BTreeSet::from([FIRST])).unwrap();
    assert_eq!(ids, BTreeMap::from([(FIRST, 1)]));
    let result = serde_json::to_value(&request).unwrap();
    assert_eq!(
        result.pointer(
            "/document/content/0/content/0/content/0/shape/textBody/content/0/content/0/info/id"
        ),
        Some(&json!(1))
    );
}
