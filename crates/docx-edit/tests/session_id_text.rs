use docx_edit::{EditingDoc, StoryRange};
use yrs::Any;

const DATE: &str = "2026-11-25T00:00:00Z";

fn comment(doc: &EditingDoc, body: &str) -> String {
    doc.add_comment(
        &[StoryRange::new("body", 0, 5)],
        "Ada",
        DATE,
        Any::from(body),
    )
    .unwrap()
}

fn restored(body: &str) -> EditingDoc {
    let doc = EditingDoc::new(41);
    doc.create_story("body", "Hello world", "Normal", "left")
        .unwrap();
    comment(&doc, body);
    let restored = EditingDoc::new(41);
    restored
        .apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    restored
}

fn counter(id: &str) -> u64 {
    id.strip_prefix("41:").unwrap().parse().unwrap()
}

#[test]
fn comment_text_at_or_past_the_id_limit_leaves_new_ids_drawable() {
    for value in [
        "4503599627370496",
        "9007199254740991",
        "9007199254740992",
        "18446744073709551615",
    ] {
        let doc = restored(&format!("see 41:{value}"));
        let first = comment(&doc, "next");
        let second = comment(&doc, "again");
        assert!(counter(&first) < 1 << 52, "{value}: {first}");
        assert_ne!(first, second);
    }
}

#[test]
fn comment_text_below_the_limit_still_raises_the_floor() {
    let doc = restored("see 41:4503599627370494");
    assert!(counter(&comment(&doc, "next")) > 4_503_599_627_370_494);
}
