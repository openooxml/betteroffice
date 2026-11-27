use super::*;
use docx_parse::block::BlockSdt;
use docx_parse::write_docx_s13;
use docx_parse::xml::{ParseBudget, ParseLimits, XmlDocument, XmlElement, parse_xml};
use serde_json::json;

mod string_paths {
    include!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/support/revision_id_string_paths.rs"
    ));
}

const FIRST: u64 = (1 << 53) - 1;
const SECOND: u64 = 1 << 40;
const DOCUMENT: &str = "word/document.xml";

fn collect_for_test(value: &impl collect::Visit, session: &BTreeSet<u64>) -> Collected {
    collect::<false>(value, &SessionIds::new(session), &NO_VIEWS)
}

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

// The typed run property change has no previous-attributes carrier yet.
#[test]
fn stated_difference_previous_run_attributes_not_reserved() {
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
    let found = collect_for_test(
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
    );
    assert_eq!(found.max, 22);
    assert_eq!(
        found.used,
        [0].into_iter().chain(7..=22).chain(100..=104).collect()
    );
}

#[test]
fn scans_xml_in_values_with_javascript_boundaries() {
    let found = collect_for_test(
        &json!({
            "w:id = '99'": [
                "w:id=\u{feff}\"00031' éw:id\u{a0}= '42\"",
                "aw:id='90' _w:id='91' w:id = '-92' w:id\u{85}='93'",
                "w:id='2147483648' w:id='123456789012345678901234567890'",
                "w:id=95 w:id=\"96 w:id='' w:identifier='97'",
            ]
        }),
        &BTreeSet::new(),
    );
    assert_eq!(found.max, 42);
    assert_eq!(found.used, [31, 42].into_iter().collect());
}

#[test]
fn distinguishes_typed_and_unmarked_json_revision_infos() {
    let found = collect_for_test(
        &json!({"author": "A", "id": 3, "revisionId": 5}),
        &BTreeSet::from([3]),
    );
    assert!(found.sites.is_empty());
    assert_eq!(found.used, [3, 5].into_iter().collect());

    let infos = vec![
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
    let found = collect_for_test(&infos, &BTreeSet::from([3]));
    assert_eq!(found.site_count, 1);
    assert_eq!(found.sites[&3].remaining, 1);
    assert_eq!(found.max, 12);
    assert_eq!(found.used, [5, 12].into_iter().collect());
}

// Patch payloads reach product saves as freshly parsed JSON with no session marks.
#[test]
fn paragraph_patch_payload_infos_are_not_session_sites() {
    let mut wire =
        serde_json::to_value(request(json!([{"type": "paragraph", "content": []}]))).unwrap();
    wire["paragraphIds"] = json!({
        "splicedParts": [{
            "part": DOCUMENT,
            "sha256": "",
            "paragraphPatches": [{
                "ordinal": 0, "sessionKey": "body:0", "resultSha256": "",
                "seed": {
                    "units": [{"kind": "atom", "atomKind": "opaque", "attrs": {},
                        "payload": {"info": {"id": FIRST, "author": "Reviewer"}}}],
                    "properties": {},
                    "unitSha256": ""
                },
                "textEdits": [], "runEdits": [], "paraDeltas": [], "paraInherit": {}
            }]
        }]
    });
    let mut request = serde_json::from_value(wire).unwrap();
    let ids = assign_session_revision_ids(&mut request, &BTreeSet::from([FIRST])).unwrap();
    assert!(ids.is_empty());
    let result = serde_json::to_value(&request).unwrap();
    assert_eq!(
        result.pointer("/paragraphIds/splicedParts/0/paragraphPatches/0/seed/units/0/payload/info/id"),
        Some(&json!(FIRST))
    );
}

#[test]
fn reports_reservations_at_serialized_typed_and_json_paths() {
    let mut wire = serde_json::to_value(request(json!([
        {"type": "rawXml", "xml": "<w:customXml w:id='7'/>"},
        {
            "type": "paragraph",
            "extraAttributes": [{"name": "custom", "value": "w:id='8'"}],
            "pPrIns": {"id": FIRST, "author": "w:id='99'"},
            "pPrDel": {"id": 9, "author": "w:id='10'"},
            "content": [
                {"type": "run", "content": [{"type": "text", "text": "w:id='11'"}]},
                {
                    "type": "complexField",
                    "instruction": "w:id='12'",
                    "fieldType": "PAGE",
                    "fieldCode": [{
                        "type": "run", "content": [{"type": "instrText", "text": "w:id='13'"}]
                    }],
                    "fieldResult": [],
                    "rawFormDataXml": "<w:ffData w:id='14'/>",
                    "sourceRunXml": "<w:r w:id='15'/>"
                },
                {"type": "run", "content": [{
                    "type": "opaqueDrawing",
                    "kind": "drawing",
                    "xml": "<w:drawing w:id='16'/>",
                    "textBox": {
                        "type": "shape", "shapeType": "rect", "size": {"width": 1, "height": 1},
                        "drawingRunXml": "<w:r w:id='17'/>",
                        "drawingModel": {"marker": {"type": "bookmarkEnd", "id": 21}},
                        "textBody": {"content": [{"a~/b": "w:id='18' w:id='18'"}]}
                    }
                }]}
            ]
        },
        {
            "type": "blockSdt",
            "properties": {"sdtType": "richText", "rawEndPropertiesXml": "w:id='19'"},
            "content": []
        }
    ])))
    .unwrap();
    wire["document"]["comments"] = json!([{"id": 80, "author": "Reviewer", "content": []}]);
    wire["headerEntries"] = json!([["rId1", {
        "type": "header", "hdrFtrType": "default", "content": [{"type": "rawXml", "xml": "w:id='20'"}]
    }]]);
    wire["paragraphIds"] = json!({
        "sectionEdits": {"sha256": "", "edits": [{"properties": {"xml": "w:id='21'"}}]},
        "splicedParts": [{
            "part": "word/document.xml",
            "sha256": "",
            "tablePatches": [{"ordinal": 0, "before": null, "after": {"xml": "w:id='21'"}}],
            "paragraphPatches": [{
                "ordinal": 0, "sessionKey": "", "resultSha256": "",
                "seed": {
                    "units": [{"kind": "atom", "atomKind": "opaque", "attrs": {},
                        "payload": {"xml": "w:id='21'"}}],
                    "properties": {},
                    "unitSha256": ""
                },
                "textEdits": [], "runEdits": [], "paraDeltas": [], "paraInherit": {}
            }]
        }]
    });
    let mut request = serde_json::from_value(wire).unwrap();
    let session = BTreeSet::from([FIRST]);
    let found = session_revision_id_reservations(&request, &session).unwrap();
    let expected = [
        ("/document/content/0/xml", 7, 1),
        ("/document/content/1/extraAttributes/0/value", 8, 1),
        ("/document/content/1/pPrDel/id", 9, 1),
        ("/document/content/1/pPrDel/author", 10, 1),
        ("/document/content/1/content/0/content/0/text", 11, 1),
        ("/document/content/1/content/1/instruction", 12, 1),
        (
            "/document/content/1/content/1/fieldCode/0/content/0/text",
            13,
            1,
        ),
        ("/document/content/1/content/1/rawFormDataXml", 14, 1),
        ("/document/content/1/content/1/sourceRunXml", 15, 1),
        ("/document/content/1/content/2/content/0/xml", 16, 1),
        (
            "/document/content/1/content/2/content/0/textBox/drawingRunXml",
            17,
            1,
        ),
        (
            "/document/content/1/content/2/content/0/textBox/drawingModel/marker/id",
            21,
            1,
        ),
        (
            "/document/content/1/content/2/content/0/textBox/textBody/content/0/a~0~1b",
            18,
            2,
        ),
        ("/document/content/2/properties/rawEndPropertiesXml", 19, 1),
        ("/headerEntries/0/1/content/0/xml", 20, 1),
        ("/paragraphIds/sectionEdits/edits/0/properties/xml", 21, 1),
        (
            "/paragraphIds/splicedParts/0/tablePatches/0/after/xml",
            21,
            1,
        ),
        (
            "/paragraphIds/splicedParts/0/paragraphPatches/0/seed/units/0/payload/xml",
            21,
            1,
        ),
        ("/document/comments/0/id", 80, 1),
    ];
    assert_eq!(
        found.sites,
        expected
            .into_iter()
            .map(|(path, id, count)| ((path.to_owned(), id), count))
            .collect()
    );
    assert_eq!(found.used, (7..=21).chain([80]).collect());
    assert_eq!(found.max, 21);

    let mut paths = Vec::new();
    let ids = assign_session_revision_ids_with_paths(&mut request, &session, &mut paths).unwrap();
    assert_eq!(ids, BTreeMap::from([(FIRST, 22)]));
    assert_eq!(paths, [("/document/content/1/pPrIns/id".to_owned(), FIRST)]);
}

fn drawing_string_path_examples() -> Vec<Value> {
    let transforms = [
        "tint", "shade", "hue", "hueMod", "hueOff", "sat", "satMod", "satOff", "lum", "lumMod",
        "lumOff", "red", "redMod", "redOff", "green", "greenMod", "greenOff", "blue", "blueMod",
        "blueOff", "alpha", "alphaMod", "alphaOff",
    ]
    .into_iter()
    .map(|kind| json!({"kind": kind, "value": 0.5}))
    .chain(
        ["comp", "inv", "gray", "gamma", "invGamma"]
            .into_iter()
            .map(|kind| json!({"kind": kind})),
    )
    .collect::<Vec<_>>();
    let image = json!({
        "type": "image", "rId": "rId1", "size": {"width": 1, "height": 1},
        "wrap": {"type": "square", "wrapText": "bothSides"},
        "position": {"horizontal": {"relativeTo": "page"}, "vertical": {"relativeTo": "line"}},
        "hyperlink": {"href": "https://example.com", "doc_location": "target"},
        "effects": {"ordered": [{"kind": "duotone", "colors": ["001122", "334455"]}]}
    });
    let shape = json!({
        "type": "shape", "shapeType": "rect", "size": {"width": 1, "height": 1},
        "fill": {
            "type": "gradient",
            "gradient": {"type": "linear", "stops": [
                {"position": 0, "color": {"rgb": "112233", "transforms": transforms}}
            ]}
        },
        "outline": {"style": "solid", "headEnd": {"type": "triangle"}},
        "fillPaint": {
            "kind": "picture", "picture": image, "tile": {"alignment": "ctr"},
            "stops": [{"color": {"rgb": "334455"}}]
        },
        "strokePaint": {"dash": "dash", "tailEnd": {"type": "arrow"}},
        "effects": [{"kind": "shadow", "color": {"rgb": "556677"}}],
        "textBodyProperties": {"anchor": "ctr", "margins": {}},
        "geometryPath": [
            {"type": "move", "x": 1, "y": 2},
            {"type": "line", "x": 2, "y": 3},
            {"type": "quad", "cpx": 3, "cpy": 4, "x": 5, "y": 6},
            {"type": "cubic", "cp1x": 1, "cp1y": 2, "cp2x": 3, "cp2y": 4, "x": 5, "y": 6},
            {"type": "close"}
        ],
        "scene": {"title": "group", "root": {"kind": "group"}},
        "drawingModel": {
            "type": "insertion",
            "info": {"id": SECOND, "author": "Reviewer", "custom": ["hidden", {"name": "hidden"}]},
            "content": [
                {"type": "run", "content": [{"type": "text", "text": "Drawing"}]},
                {"type": "hyperlink", "children": [
                    {"type": "simpleField", "fieldTree": {"result": {"inline": ["Page"]}}},
                    {"type": "complexField", "structuredCode": {"inline": ["PAGE"]}}
                ]},
                {"type": "shape", "scene": {"root": {"shape": {"name": "Untyped shape"}}}},
                {"type": "table", "rows": [{"type": "tableRow", "cells": [
                    {"type": "tableCell", "content": [{"type": "paragraph", "typeLabel": "Cell"}]}
                ]}]}
            ]
        },
        "textBody": {"anchor": "ctr", "content": [
            {"type": "paragraph", "content": ["", null, false, 1, {"a~/b": "nested text"}]}
        ]}
    });
    let chart = json!({
        "type": "chart", "chartType": "bar",
        "series": [{
            "name": "A", "categories": ["Q1"], "values": [1], "color": "112233",
            "points": [{"color": "223344"}], "marker": {"symbol": "circle"},
            "dataLabels": {"separator": ", ", "points": [{
                "runs": [{"text": "w:id='91'"}, {"field": "VALUE"}], "labels": {}
            }]}
        }],
        "axes": {"category": {
            "axisType": "category", "reversed": false, "hidden": false,
            "line": {"none": false, "color": "334455"}
        }},
        "legend": {"visible": true, "position": "b", "text": {"font": "Arial"}},
        "plotGroups": [{"axisIds": ["1"], "series": [], "varyColors": false, "showDataLabels": true}]
    });
    let mut drawings = vec![
        json!({"type": "drawing", "image": image}),
        json!({"type": "shape", "shape": shape}),
        json!({"type": "chart", "chart": chart}),
        json!({"type": "opaqueDrawing", "kind": "drawing", "xml": "<w:drawing/>", "textBox": shape}),
        json!({"type": "horizontalRule", "rule": {
            "height": 1, "alignment": "left", "noShade": false, "color": "000000", "xml": "<w:pict/>"
        }}),
    ];
    drawings.extend(
        [
            "picture", "shape", "textBox", "group", "chart", "diagram", "other",
        ]
        .into_iter()
        .enumerate()
        .map(|(index, kind)| {
            json!({"type": "drawing", "image": image, "drawing": {
                "type": kind, "placement": if index == 0 { "inline" } else { "anchor" },
                "hidden": false, "decorative": false, "name": "Example"
            }})
        }),
    );
    drawings
}

fn patch_string_path_examples() -> Value {
    let delta = json!({
        "path": ["a~/b"], "before": {"state": "missing"},
        "after": {"state": "present", "value": {"description": "Added"}}
    });
    let mut patch = json!({
        "ordinal": 0, "sessionKey": "body:0", "resultSha256": "result",
        "seed": {
            "units": [
                {"kind": "text", "text": "Before", "attrs": {"color": "red"}},
                {"kind": "atom", "atomKind": "opaque", "payload": {"xml": "<w:r/>"}, "attrs": {}}
            ],
            "properties": {}, "unitSha256": "seed",
            "inherit": {"paragraph": {"style": "Normal"}, "runs": [{}], "paragraphMark": {}}
        },
        "textEdits": [
            {"from": 0, "to": 1, "insert": [{"text": "After", "attrs": {}, "inherit": {}}],
                "donor": {"kind": "at", "at": 0}},
            {"from": 2, "to": 3, "insert": [], "donor": {"kind": "pilcrow"}}
        ],
        "runEdits": [{"from": 0, "to": 1, "deltas": [delta], "target": {}, "inherit": {}}],
        "paraDeltas": [delta], "paraInherit": {}, "markInherit": {"font": "Arial"}
    });
    let current = patch.clone();
    patch["inheritance"] = json!("unknown");
    let legacy = [
        "MISSING_SOURCE",
        "SOURCE_DIGEST_MISMATCH",
        "INVALID_PLAN",
        "SEED_MISMATCH",
        "AMBIGUOUS_ALIGNMENT",
        "AMBIGUOUS_DONOR",
        "INVALID_BOUNDARY",
        "UNSUPPORTED_KEY",
        "UNSUPPORTED_STRUCTURE",
        "PROPERTY_CONFLICT",
        "PATCH_CONFLICT",
        "RESULT_MISMATCH",
    ]
    .into_iter()
    .enumerate()
    .map(|(ordinal, code)| json!({"ordinal": ordinal, "code": code, "detail": "Reason"}))
    .collect::<Vec<_>>();
    json!({
        "removedCommentIds": ["1"],
        "sectionEdits": {"sha256": "sections", "edits": [{"properties": {"type": "nextPage"}}]},
        "assignments": [{"part": DOCUMENT, "ordinal": 0, "paraId": "00000001"}],
        "patchedParts": [{"part": DOCUMENT, "paraIds": [[0, "00000001"]]}],
        "wholeParts": [{"part": "word/header1.xml", "sha256": "header", "changed": [0]}],
        "touchedParts": [{"part": "word/footer1.xml", "touched": [0]}],
        "splicedParts": [{
            "part": DOCUMENT, "sha256": "body", "paragraphPatches": [current, patch],
            "legacyParagraphs": legacy, "tablePatches": [{
                "ordinal": 0, "before": {"type": "table", "rows": []},
                "after": {"type": "table", "rows": ["untyped row"]}
            }]
        }]
    })
}

#[test]
fn scans_all_serialized_string_paths() {
    const TS_SKIPPED_REVISION_INFOS: (&str, [&str; 3]) = (
        "packages/docx/src/docx/sessionRevisionIds.ts:60",
        [
            "/document/content/0/pPrIns",
            "/document/content/0/content/0/content/1/shape/drawingModel/info",
            "/document/content/0/content/0/content/3/textBox/drawingModel/info",
        ],
    );
    let mut run_content = drawing_string_path_examples();
    run_content.extend([
        json!({"type": "text", "text": "Body", "preserveSpace": true}),
        json!({"type": "tab"}),
        json!({"type": "break", "breakType": "textWrapping", "clear": "all"}),
        json!({"type": "symbol", "font": "Symbol", "char": "F0B7"}),
        json!({"type": "footnoteRef", "id": 1}),
        json!({"type": "endnoteRef", "id": 1}),
        json!({"type": "footnoteRefMark"}),
        json!({"type": "endnoteRefMark"}),
        json!({"type": "separator"}),
        json!({"type": "continuationSeparator"}),
        json!({"type": "commentReference", "id": 2}),
        json!({"type": "fieldChar", "charType": "begin", "formData": {
            "name": "Example", "controlType": "dropDown", "listEntries": ["One"]
        }}),
        json!({"type": "instrText", "text": "PAGE"}),
        json!({"type": "softHyphen"}),
        json!({"type": "noBreakHyphen"}),
    ]);
    let run_formatting = json!({
        "underline": {"style": "single"}, "fontFamily": {"ascii": "Arial"},
        "language": {"latin": "en-US"}, "modernEffects": {
            "glow": {"color": "112233"}, "shadow": {"color": "223344"},
            "textFill": {"kind": "gradient", "stops": [{"color": "334455"}]},
            "textOutline": {"color": "445566", "dash": "solid"}
        }
    });
    let sdt = json!({
        "sdtType": "richText", "alias": "Control", "listItems": [{"displayText": "One", "value": "1"}],
        "dateState": {"format": "yyyy-MM-dd"}, "checkedState": {"value": "2612"},
        "gallery": {"gallery": "placeholder"}, "controlState": {"value": "1"},
        "dataBinding": {"xpath": "/root/value", "storeItemID": "item"}
    });
    let raw = json!({"type": "rawXml", "xml": "<w:customXml/>"});
    let empty_paragraph = json!({"type": "paragraph", "content": []});
    let field_result = json!({"inline": [raw], "blocks": [empty_paragraph]});
    let info = json!({"id": 3, "author": "Editor", "date": "2030-01-02"});
    let body = json!({
        "type": "paragraph", "paraId": "00000001",
        "pPrIns": {"id": FIRST, "author": "Skipped", "date": "2030-01-03"},
        "pPrDel": info,
        "propertyChanges": [{"type": "paragraphPropertyChange", "info": info}],
        "formatting": {
            "frame": {"hAnchor": "page"}, "cnfStyle": {"val": "1000"},
            "borders": {"top": {"style": "single"}}, "shading": {"pattern": "clear"},
            "tabs": [{"position": 720, "alignment": "left"}], "runProperties": run_formatting
        },
        "listRendering": {
            "marker": "1.", "level": 0, "numId": 1, "isBullet": false,
            "numFmt": "decimal", "levelNumFmts": ["decimal"]
        },
        "extraAttributes": [{"name": "custom", "value": "body"}],
        "content": [
            {"type": "run", "formatting": run_formatting, "content": run_content,
                "directionScopes": [{"kind": "dir", "id": 0, "val": "ltr"}, {"kind": "bdo", "id": 1}],
                "propertyChanges": [{"type": "runPropertyChange",
                    "info": {"id": 4, "author": "Editor", "rsid": "0123ABCD"}}]},
            change(5),
            {"type": "moveFromRangeStart", "id": 6, "name": "move"},
            {"type": "moveFromRangeEnd", "id": 6},
            {"type": "commentRangeStart", "id": 2},
            {"type": "bookmarkStart", "id": 7, "name": "bookmark"},
            {"type": "bookmarkEnd", "id": 7},
            {"type": "hyperlink", "href": "#bookmark", "children": [change(8)]},
            {"type": "simpleField", "instruction": "PAGE", "fieldType": "PAGE", "content": [],
                "structuredResult": field_result},
            {"type": "complexField", "instruction": "PAGE", "fieldType": "PAGE",
                "fieldCode": [], "fieldResult": [], "fieldTree": {
                    "displayMode": "result", "result": field_result
                }},
            {"type": "inlineSdt", "properties": sdt, "content": []},
            {"type": "mathEquation", "display": "inline", "ommlXml": "<m:oMath/>"},
            raw
        ]
    });
    let table = json!({
        "type": "table", "formatting": {
            "caption": "Table", "floating": {"horzAnchor": "page"},
            "look": {"value": "04A0"}, "cellMargins": {"top": {"type": "dxa", "value": 100}}
        },
        "propertyChanges": [{"type": "tablePropertyChange", "info": info}],
        "rows": [{
            "type": "tableRow", "formatting": {"heightRule": "atLeast"},
            "propertyChanges": [{"type": "tableRowPropertyChange", "info": info}],
            "structuralChange": {"type": "insertion", "info": info},
            "cells": [{
                "type": "tableCell", "formatting": {"verticalAlign": "center"},
                "propertyChanges": [{"type": "tableCellPropertyChange", "info": info}],
                "content": [empty_paragraph]
            }]
        }]
    });
    let sections = json!({
        "orientation": "portrait",
        "headerReferences": [{"type": "default", "rId": "rId1"}],
        "lineNumbers": {"restart": "continuous"}, "pageNumbering": {"format": "decimal"},
        "pageBorders": {"display": "allPages"}, "background": {"themeColor": "accent1"},
        "footnotePr": {"position": "pageBottom", "separators": [{"kind": "separator"}]},
        "docGrid": {"type": "lines"}
    });
    let level = json!({"ilvl": 0, "numFmt": "decimal", "lvlText": "%1."});
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2030-01-02T03:04:05.006Z"},
        "document": {
            "content": [body, table, raw, {"type": "blockSdt", "properties": sdt, "content": [raw]}],
            "sections": [{"id": "main", "properties": sections, "content": [empty_paragraph]}],
            "comments": [{"id": 2, "author": "Reviewer", "content": [empty_paragraph]}]
        },
        "headerEntries": [["rId1", {
            "type": "header", "hdrFtrType": "default", "content": [empty_paragraph],
            "watermark": {"kind": "text", "text": "Draft", "font": "Arial", "color": "AAAAAA",
                "semitransparent": true, "layout": "diagonal"},
            "watermarkBox": {"widthPt": 20, "heightPt": 10, "rotation": 45, "horizontal": "center"}
        }]],
        "footerEntries": [["rId2", {
            "type": "footer", "hdrFtrType": "default", "content": [empty_paragraph],
            "watermark": {"kind": "picture", "relId": "rId3", "scale": 1, "washout": false}
        }]],
        "footnotes": [{"type": "footnote", "id": 1, "noteType": "normal", "content": [empty_paragraph]}],
        "endnotes": [{"type": "endnote", "id": 1, "noteType": "normal", "content": [raw]}],
        "footnoteSeparators": [{"type": "footnote", "id": -1, "noteType": "separator", "content": [],
            "verbatimXml": "<w:footnote w:id='90'/>"}],
        "endnoteSeparators": [{"type": "endnote", "id": -1, "noteType": "separator", "content": [],
            "verbatimXml": "<w:endnote w:id='92'/>"}],
        "relationshipEntries": [
            ["rId1", {"id": "rId1", "type": "header", "target": "header1.xml", "targetMode": "Internal"}],
            ["rId2", {"id": "rId2", "type": "hyperlink", "target": "https://example.com", "targetMode": "External"}]
        ],
        "options": {"modifiedBy": "Editor"},
        "numbering": {
            "abstractNums": [{"abstractNumId": 1, "levels": [level]}],
            "nums": [{"numId": 1, "abstractNumId": 1, "levelOverrides": [{"ilvl": 0, "lvl": level}]}]
        },
        "selective": {"changedParaIds": ["00000001"], "sourceParagraphs": {
            "partSha256": "body", "paragraphs": [{"path": [0], "block": 0}]
        }},
        "paragraphIds": patch_string_path_examples()
    }))
    .unwrap();
    let found =
        session_revision_id_reservations(&request, &BTreeSet::from([FIRST, SECOND])).unwrap();
    let wire = serde_json::to_value(&request).unwrap();
    let (source, skipped) = TS_SKIPPED_REVISION_INFOS;
    for path in skipped {
        assert!(
            wire.pointer(path).is_some(),
            "missing {source} exception at {path}"
        );
    }
    let expected = string_paths::json_string_paths(&wire, &BTreeSet::from(skipped));
    assert_eq!(found.string_paths, expected, "{source}");
    assert!(BTreeSet::from([90, 91, 92]).is_subset(&found.used));
}

#[test]
fn visits_source_paragraphs_exhaustively() {
    let paragraph = docx_parse::serializer::S13SourceParagraph {
        path: vec![0],
        block: 0,
    };
    let found = collect_for_test(&paragraph, &BTreeSet::new());
    assert!(found.used.is_empty());
    assert!(found.sites.is_empty());
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
        next_id(&mut next, 5, &[4].into_iter().collect()),
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
    let content = result["document"]["content"][0]["content"]
        .as_array()
        .unwrap();
    assert_eq!(
        content
            .iter()
            .map(|site| site["info"]["id"].as_f64())
            .collect::<Vec<_>>(),
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
    let [
        BlockContent::Paragraph(unchanged),
        BlockContent::Table(table),
    ] = original.document.content.as_slice()
    else {
        panic!("expected a paragraph and a table");
    };
    let [
        BlockContent::Paragraph(actual_unchanged),
        BlockContent::Table(actual_table),
    ] = actual.document.content.as_slice()
    else {
        panic!("expected a paragraph and a table");
    };
    assert!(Arc::ptr_eq(unchanged, actual_unchanged));
    assert!(!Arc::ptr_eq(table, actual_table));
    let [
        BlockContent::Paragraph(cell_unchanged),
        BlockContent::Paragraph(cell_changed),
    ] = table.rows[0].cells[0].content.as_slice()
    else {
        panic!("expected two cell paragraphs");
    };
    let [
        BlockContent::Paragraph(actual_cell_unchanged),
        BlockContent::Paragraph(actual_cell_changed),
    ] = actual_table.rows[0].cells[0].content.as_slice()
    else {
        panic!("expected two cell paragraphs");
    };
    assert!(Arc::ptr_eq(cell_unchanged, actual_cell_unchanged));
    assert!(!Arc::ptr_eq(cell_changed, actual_cell_changed));

    let mut probe = original.clone();
    let found = collect_for_test(&probe, &session);
    let session = SessionIds::new(&session);
    let mut walker = SiteWalker::new(
        &session,
        None,
        &NO_VIEWS,
        &found.blocks_with_sites,
        &found.body_blocks_with_sites,
        found.site_count,
    );
    WORK.with_borrow_mut(|work| *work = Work::default());
    walker.request(&mut probe);
    assert_eq!(WORK.with_borrow(|work| work.make_mut), 2);
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

                WORK.with_borrow_mut(|work| *work = Work::default());
                assert_eq!(
                    assign_session_revision_ids(&mut actual, &BTreeSet::from([FIRST])).unwrap(),
                    BTreeMap::from([(FIRST, 1)])
                );
                let visits = WORK.with_borrow(|work| {
                    work.collected_blocks + work.body_block_selections + work.block_probes
                });
                assert!(
                    visits <= 3 * (2 * depth + 1),
                    "{visits} visits at depth {depth}"
                );
            }
        })
        .unwrap()
        .join()
        .unwrap();
}

fn assign_flat(paragraphs: usize) -> Work {
    let mut session = BTreeSet::new();
    let mut blocks = Vec::with_capacity(paragraphs + paragraphs / 50);
    for index in 0..paragraphs {
        let mut content = vec![json!({
            "type": "run", "content": [{"type": "text", "text": "Earlier inspection."}]
        })];
        if index % 10 == 0 {
            let id = FIRST - index as u64;
            session.insert(id);
            content.push(change(id));
        }
        blocks.push(json!({"type": "paragraph", "content": content}));
        if index % 50 == 0 {
            blocks.push(json!({"type": "rawXml", "xml": "<w:customXml w:id='42'/>"}));
        }
    }
    let mut request = request(Value::Array(blocks));
    WORK.with_borrow_mut(|work| *work = Work::default());
    let ids = assign_session_revision_ids(&mut request, &session).unwrap();
    assert_eq!(ids.len(), paragraphs / 10);
    assert_eq!(ids.values().copied().min(), Some(43));
    WORK.with_borrow(|work| *work)
}

#[test]
fn bounds_work_for_flat_revision_requests() {
    let paragraphs = 2_000;
    let small = assign_flat(paragraphs);
    let large = assign_flat(10 * paragraphs);
    for (name, small, large, bound) in [
        ("fields", small.fields, large.fields, 16),
        ("strings", small.strings, large.strings, 5),
        ("string bytes", small.string_bytes, large.string_bytes, 128),
        (
            "reservations",
            small.reservation_inserts,
            large.reservation_inserts,
            1,
        ),
        ("site inserts", small.site_inserts, large.site_inserts, 1),
        (
            "collected blocks",
            small.collected_blocks,
            large.collected_blocks,
            1,
        ),
        ("block inserts", small.block_inserts, large.block_inserts, 1),
        (
            "body block selections",
            small.body_block_selections,
            large.body_block_selections,
            1,
        ),
        ("make_mut", small.make_mut, large.make_mut, 1),
        ("id probes", small.id_probes, large.id_probes, 1),
        ("verified sites", small.verified, large.verified, 1),
    ] {
        assert!(
            small > 0 && small <= bound * paragraphs + 128,
            "{name}: {small}"
        );
        assert!(large <= bound * 10 * paragraphs + 128, "{name}: {large}");
        assert!(large <= 10 * small + 32, "{name}: {small} -> {large}");
    }
    for (count, work) in [(paragraphs, small), (10 * paragraphs, large)] {
        assert_eq!(work.reservation_inserts, count / 50);
        assert_eq!(work.site_inserts, count / 10);
        assert_eq!(work.collected_blocks, count);
        assert_eq!(work.block_inserts, count / 10);
        assert_eq!(work.body_block_selections, count / 10);
        assert_eq!(work.block_probes, 0);
        assert_eq!(work.strings, 2 + 2 * count + count / 10 + count / 50);
        assert_eq!(work.make_mut, count / 10);
        assert_eq!(work.id_probes, count / 10);
        assert_eq!(work.verified, count / 10);
    }
}

#[test]
#[ignore]
fn revision_id_work_counts_large() {
    let small = assign_flat(50_000);
    let large = assign_flat(200_000);
    let counters = [
        ("fields", small.fields, large.fields, 14),
        ("strings", small.strings, large.strings, 2),
        ("string_bytes", small.string_bytes, large.string_bytes, 88),
        (
            "reservation_inserts",
            small.reservation_inserts,
            large.reservation_inserts,
            0,
        ),
        ("site_inserts", small.site_inserts, large.site_inserts, 0),
        (
            "collected_blocks",
            small.collected_blocks,
            large.collected_blocks,
            0,
        ),
        (
            "body_block_selections",
            small.body_block_selections,
            large.body_block_selections,
            0,
        ),
        ("block_inserts", small.block_inserts, large.block_inserts, 0),
        ("block_probes", small.block_probes, large.block_probes, 0),
        ("make_mut", small.make_mut, large.make_mut, 0),
        ("id_probes", small.id_probes, large.id_probes, 0),
        ("verified", small.verified, large.verified, 0),
    ];
    for (paragraphs, use_large) in [(50_000, false), (200_000, true)] {
        let mut line = format!("REVISION-ID-WORK paragraphs={paragraphs}");
        for (name, small, large, _) in counters {
            let value = if use_large { large } else { small };
            write!(line, " {name}={value}").unwrap();
        }
        println!("{line}");
    }
    for (name, small, large, fixed) in counters {
        assert_eq!(large + 3 * fixed, 4 * small, "{name} (fixed work: {fixed})");
    }
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
                WORK.with_borrow_mut(|work| *work = Work::default());
                let found = collect_for_test(&shape, &BTreeSet::from([FIRST]));
                assert_eq!(found.used, (1..=count).collect());
                assert_eq!(
                    WORK.with_borrow(|work| work.reservation_inserts),
                    count as usize
                );
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

#[test]
fn probes_model_record_ids_once_and_author_only_when_needed() {
    let mut shape: Shape = serde_json::from_value(model_shape(json!([]))).unwrap();
    shape.drawing_model = Some(json!([
        {"label": "No annotation"},
        {"id": 23, "author": "Source"},
        {"revisionId": 24, "author": "Source"},
        {"id": FIRST, "author": "Session"}
    ]));
    WORK.with_borrow_mut(|work| *work = Work::default());
    let found = collect_for_test(&shape, &BTreeSet::from([FIRST]));
    assert_eq!(found.used, [23, 24].into_iter().collect());
    assert_eq!(found.site_count, 1);
    let work = WORK.with_borrow(|work| *work);
    assert_eq!(work.model_id_lookups, 4);
    assert_eq!(work.model_author_lookups, 3);
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
        assert_eq!(
            result.pointer(path).and_then(Value::as_u64),
            Some(u64::from(assigned))
        );
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
fn assigns_ordered_model_fields_before_extra_fields() {
    assert_text_box_model_order(
        json!({
            "type": "insertion",
            "additional": {"info": {"id": FIRST - 1, "author": "Reviewer"}},
            "content": [change(SECOND)],
            "info": {"id": FIRST, "author": "Reviewer"}
        }),
        &[FIRST, SECOND, FIRST - 1],
    );
}

#[test]
fn assigns_model_revision_infos_under_arbitrary_keys() {
    let mut shape = model_shape(json!([]));
    shape["drawingModel"] = json!({"frame": {"revision": {"id": SECOND, "author": "A"}}});
    let mut request = request(json!([
        {"type": "paragraph", "content": [{
            "type": "run", "content": [{"type": "shape", "shape": shape}]
        }]},
        {"type": "paragraph", "content": [change(FIRST)]}
    ]));
    let ids = assign_session_revision_ids(&mut request, &BTreeSet::from([FIRST, SECOND])).unwrap();
    assert_eq!(ids, BTreeMap::from([(SECOND, 1), (FIRST, 2)]));
    let result = serde_json::to_value(&request).unwrap();
    let body = &result["document"]["content"];
    assert_eq!(
        body.pointer("/0/content/0/content/0/shape/drawingModel/frame/revision/id"),
        Some(&json!(1))
    );
    assert_eq!(body[1]["content"][0]["info"]["id"].as_f64(), Some(2.0));
}

// TypeScript assigns in insertion order, and the typed request keeps keys sorted.
#[test]
fn stated_difference_arbitrary_model_keys_assign_in_sorted_order() {
    let mut request = serde_json::from_str(
        r#"{
            "determinism": {
                "seed": "0000000000000000000000000000000000000000000000000000000000000000",
                "now": "2030-01-02T03:04:05.006Z"
            },
            "document": {"content": [{"type": "paragraph", "content": [{
                "type": "run", "content": [{"type": "shape", "shape": {
                    "type": "shape", "shapeType": "rect", "size": {"width": 1, "height": 1},
                    "drawingModel": {"frame": {
                        "z": {"id": 9007199254740991, "author": "A"},
                        "a": {"id": 1099511627776, "author": "B"}
                    }, "otherSha256": ""}
                }}]
            }]}]}
        }"#,
    )
    .unwrap();
    let ids = assign_session_revision_ids(&mut request, &BTreeSet::from([FIRST, SECOND])).unwrap();
    assert_eq!(ids, BTreeMap::from([(SECOND, 1), (FIRST, 2)]));
    let result = serde_json::to_value(&request).unwrap();
    let frame = result
        .pointer("/document/content/0/content/0/content/0/shape/drawingModel/frame")
        .unwrap();
    assert_eq!(frame["a"]["id"], json!(1));
    assert_eq!(frame["z"]["id"], json!(2));
}

#[test]
fn reserves_parent_id_above_base_when_content_is_a_skipped_view() {
    let mut shape = model_shape(json!([]));
    shape["drawingModel"] = json!({"frame": {"entries": [
        {"id": 90, "author": "Source", "content": []},
        {"id": FIRST, "author": "Reviewer"}
    ]}, "otherSha256": ""});
    let mut request = request(json!([{"type": "paragraph", "content": [{
        "type": "run", "content": [{"type": "shape", "shape": shape}]
    }]}]));
    let skipped_views = BTreeSet::from([
        "/document/content/0/content/0/content/0/shape/drawingModel/frame/entries/0/content"
            .to_owned(),
    ]);
    let session = BTreeSet::from([FIRST]);
    let mut paths = Vec::new();
    let ids = assign_session_revision_ids_with_paths_skipping_views(
        &mut request,
        &session,
        &mut paths,
        &skipped_views,
    )
    .unwrap();
    assert_eq!(ids, BTreeMap::from([(FIRST, 91)]));
}

#[test]
fn leaves_marked_descendants_untouched_when_a_sibling_shares_the_placeholder() {
    let sibling = change(SECOND);
    let mut first = change(FIRST);
    first["info"]["info"] = sibling["info"].clone();
    let shape = model_shape(json!([{
        "type": "paragraph", "content": [first, sibling]
    }]));
    let mut request = request(json!([{"type": "paragraph", "content": [{
        "type": "run", "content": [{"type": "shape", "shape": shape}]
    }]}]));
    let ids = assign_session_revision_ids(&mut request, &BTreeSet::from([FIRST, SECOND])).unwrap();
    assert_eq!(ids, BTreeMap::from([(FIRST, 1), (SECOND, 2)]));
    let result = serde_json::to_value(&request).unwrap();
    let content = result
        .pointer("/document/content/0/content/0/content/0/shape/textBody/content/0/content")
        .unwrap();
    assert_eq!(content[0]["info"]["id"], json!(1));
    assert_eq!(content[0]["info"]["info"]["id"], json!(SECOND));
    assert_eq!(content[1]["info"]["id"], json!(2));
}

#[test]
fn stated_difference_shared_holder_inside_marked_info_is_not_aliased() {
    let content = |frame: Value| {
        let mut shape = model_shape(json!([]));
        shape["drawingModel"] = json!({"frame": frame, "otherSha256": ""});
        json!([{"type": "paragraph", "content": [{
            "type": "run", "content": [{"type": "shape", "shape": shape}]
        }]}])
    };
    let holder = json!({"info": {"id": SECOND, "author": "B"}});
    let body = content(json!({
        "info": {"id": FIRST, "author": "A", "holder": holder.clone()}
    }));
    let mut wire = serde_json::to_value(request(body)).unwrap();
    wire["headerEntries"] = json!([["rId1", {
        "type": "header", "hdrFtrType": "default", "content": content(json!({"holder": holder}))
    }]]);
    let mut request = serde_json::from_value(wire).unwrap();
    let ids = assign_session_revision_ids(&mut request, &BTreeSet::from([FIRST, SECOND])).unwrap();
    assert_eq!(ids, BTreeMap::from([(FIRST, 1), (SECOND, 2)]));
    let result = serde_json::to_value(&request).unwrap();
    let body_shape = result
        .pointer("/document/content/0/content/0/content/0/shape")
        .unwrap();
    let header_shape = result
        .pointer("/headerEntries/0/1/content/0/content/0/content/0/shape")
        .unwrap();
    let body_info = &body_shape["drawingModel"]["frame"]["info"];
    assert_eq!(body_info["id"], json!(1));
    assert_eq!(body_info["holder"]["info"]["id"], json!(SECOND));
    assert_eq!(
        header_shape["drawingModel"]["frame"]["holder"]["info"]["id"],
        json!(2)
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
