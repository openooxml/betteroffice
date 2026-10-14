#[path = "support/revision_boundary.rs"]
mod boundary;

use docx_edit::{EditCtx, EditingDoc, GeometryRead, GeometrySentinel};
use serde_json::json;

#[test]
fn outlines_of_existing_block_boundary_fixtures() {
    for (kind, block_size) in [
        ("table", 12),
        ("blockSdt", 10),
        ("pageBreak", 1),
        ("columnBreak", 1),
    ] {
        let doc = EditingDoc::new(84201);
        let split = boundary::seed(&doc, kind, &EditCtx::local("", ""));
        let GeometryRead::Value(outline) = doc.geometry_position_outline("body") else {
            panic!("legacy outline for {kind}");
        };
        assert_eq!(
            serde_json::to_value(&outline["body"]).unwrap(),
            json!({"contentStart": 0, "size": 11 + block_size, "paragraphs": [
                {"paraId": split.first_para_id, "displayStart": 0, "length": 3, "leading": 0},
                {"paraId": split.second_para_id, "displayStart": 5 + block_size, "length": 4, "leading": 1}
            ]}),
            "{kind}"
        );
        if kind == "table" {
            assert_eq!(outline["body:t0:r0c0"].content_start, 8);
            assert_eq!(outline["body:t0:r0c0"].size, 6);
        }
        if kind == "blockSdt" {
            assert_eq!(outline["control"].content_start, 6);
            assert_eq!(outline["control"].size, 8);
        }
    }
}

#[test]
fn suggested_block_boundary_fixtures_keep_canonical_fallback() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        let doc = EditingDoc::new(84202);
        let split = boundary::seed(
            &doc,
            kind,
            &EditCtx::local("Reviewer", "2026-09-29T12:00:00Z").suggesting(),
        );
        assert_eq!(
            doc.owned_revision_ranges(&split.revision_ids),
            GeometryRead::Sentinel(GeometrySentinel::Fallback),
            "{kind}"
        );
        assert!(matches!(
            doc.geometry_position_outline("body"),
            GeometryRead::Value(_)
        ));
    }
}
