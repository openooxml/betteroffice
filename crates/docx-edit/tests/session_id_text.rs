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

fn restore(doc: &EditingDoc) -> EditingDoc {
    let restored = EditingDoc::new(41);
    restored
        .apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    restored
}

fn commented(body: &str) -> (EditingDoc, String) {
    let doc = EditingDoc::new(41);
    doc.create_story("body", "Hello world", "Normal", "left")
        .unwrap();
    let id = comment(&doc, body);
    (doc, id)
}

#[test]
fn comment_text_at_or_past_the_id_limit_leaves_new_ids_drawable() {
    for value in [
        "4503599627370496",
        "9007199254740991",
        "9007199254740992",
        "18446744073709551615",
    ] {
        let (doc, original) = commented(&format!("see 41:{value}"));
        let doc = restore(&doc);
        let first = comment(&doc, "next");
        let second = comment(&doc, "again");
        assert_ne!(first, second, "{value}");
        assert!(first != original && second != original, "{value}");
        assert!(doc.resolve_comment(&original).is_ok(), "{value}");
    }
}

#[test]
fn ids_drawn_at_the_cap_stay_unique_across_restores() {
    let (doc, original) = commented("see 41:9007199254740990");
    let doc = restore(&doc);
    let last = comment(&doc, "last below the cap");
    let doc = restore(&doc);
    let first = comment(&doc, "next");
    let second = comment(&doc, "again");
    let ids = [&original, &last, &first, &second];
    for (at, id) in ids.iter().enumerate() {
        assert!(!ids[at + 1..].contains(id), "{id} drawn twice");
        assert!(doc.resolve_comment(id).is_ok(), "{id} is live");
    }
}

#[test]
fn many_crafted_ids_are_skipped_at_the_cap() {
    let body: String = (0..1000)
        .map(|n| format!("41:{n} "))
        .chain(["41:9007199254740991".to_owned()])
        .collect();
    let (doc, original) = commented(&body);
    let doc = restore(&doc);
    let first = comment(&doc, "next");
    let second = comment(&doc, "again");
    let counter = |id: &str| id.strip_prefix("41:").unwrap().parse::<u64>().unwrap();
    assert!(counter(&first) >= 1000 && counter(&second) > counter(&first));
    assert!(first != original && second != original);
}
