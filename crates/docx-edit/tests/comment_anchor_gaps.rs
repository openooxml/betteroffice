//! A comment's anchors are opaque bytes in the comments map, so a replica can integrate a comment
//! before the text its anchors point at. Lowering has to survive that window and pick the comment
//! up once the text arrives.

use std::collections::BTreeMap;

use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Position, StoryRange};
use serde_json::Value;
use yrs::Any;

const DATE: &str = "2026-10-17T00:00:00Z";

fn open(state: &[u8], client: u64) -> EditingDoc {
    let doc = EditingDoc::new(client);
    doc.apply_update_v1(state).unwrap();
    doc
}

fn commented_runs(doc: &EditingDoc, env: &RenderEnv) -> Vec<(String, Value)> {
    let blocks = yrs_doc_to_layout_blocks(doc, "body", env).unwrap();
    serde_json::to_value(blocks)
        .unwrap()
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|block| block["runs"].as_array().cloned().unwrap_or_default())
        .filter(|run| run.get("commentIds").is_some())
        .map(|run| {
            let text = run["text"].as_str().unwrap().to_owned();
            (text, run["commentIds"].clone())
        })
        .collect()
}

#[test]
fn a_comment_that_arrives_before_its_anchored_text_lowers_and_resolves_once_the_text_arrives() {
    let origin = EditingDoc::new(300);
    origin
        .create_story("body", "Hello world", "Normal", "left")
        .unwrap();
    let state = origin.encode_state_as_update_v1();
    let (author, typist, late) = (open(&state, 301), open(&state, 302), open(&state, 303));

    let base = typist.encode_state_vector_v1();
    typist
        .insert_text(
            &EditCtx::local("Ada", DATE),
            Position::new("body", 6),
            "brave ",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let typed = typist.encode_diff_v1(&base).unwrap();
    author.apply_update_v1(&typed).unwrap();

    let before = author.encode_state_vector_v1();
    let comment = author
        .add_comment(
            &[StoryRange::new("body", 6, 11)],
            "Ada",
            DATE,
            Any::from("note"),
        )
        .unwrap();
    let commented = author.encode_diff_v1(&before).unwrap();
    let env = RenderEnv {
        numeric_ids: BTreeMap::from([(comment.clone(), 7.0)]),
        ..RenderEnv::default()
    };
    let expected = vec![("brave".to_owned(), serde_json::json!([7.0]))];
    assert_eq!(commented_runs(&author, &env), expected);

    late.apply_update_v1(&commented).unwrap();
    assert_eq!(commented_runs(&late, &env), Vec::new());
    assert_eq!(late.comment_anchors(&comment).unwrap().len(), 1);

    late.apply_update_v1(&typed).unwrap();
    assert_eq!(commented_runs(&late, &env), expected);
    assert_eq!(
        late.resolve_comment(&comment).unwrap(),
        author.resolve_comment(&comment).unwrap()
    );
}
