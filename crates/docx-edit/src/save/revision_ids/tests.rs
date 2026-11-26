use super::*;
use serde_json::json;

const FIRST: u64 = (1 << 53) - 1;
const SECOND: u64 = 1 << 40;

fn request(content: Value) -> S13SaveRequest {
    serde_json::from_value(json!({
        "determinism": {"seed": "", "now": ""},
        "document": {"content": content}
    }))
    .unwrap()
}

fn change(id: u64) -> Value {
    json!({"type": "insertion", "info": {"id": id, "author": "Reviewer"}, "content": []})
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
    assert_eq!(content[1]["propertyChanges"][0]["info"]["id"].as_f64(), Some(43.0));
    assert_eq!(content[1]["propertyChanges"][1]["info"]["id"].as_f64(), Some(42.0));
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
    let assigned = assign_session_revision_ids(&mut request, &ids.iter().copied().collect()).unwrap();
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
        assert!(assign_session_revision_ids(&mut request, &session).unwrap().is_empty());
        assert_eq!(request, original);
    }
}

#[test]
fn accepts_assigned_ids_that_equal_placeholders() {
    let mut request = request(json!([{"type": "paragraph", "content": [change(2), change(1)]}]));
    let ids = assign_session_revision_ids(&mut request, &BTreeSet::from([1, 2])).unwrap();
    assert_eq!(ids, BTreeMap::from([(1, 2), (2, 1)]));
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
