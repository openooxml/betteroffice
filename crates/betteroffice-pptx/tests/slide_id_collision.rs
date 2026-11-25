use betteroffice_pptx::{EditCtx, Presentation, ReadRequest};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const ADDITIONS: u32 = 300;

fn ctx() -> EditCtx {
    EditCtx::local("save-test")
}

fn slide_text(deck: &Presentation, slide_id: &str) -> Vec<String> {
    deck.read_content(&ReadRequest {
        slide_ids: Some(vec![slide_id.to_owned()]),
    })
    .unwrap()
    .unwrap()
    .stories
    .into_iter()
    .map(|story| story.text)
    .collect()
}

fn all_text(deck: &Presentation) -> Vec<String> {
    deck.read_content(&ReadRequest { slide_ids: None })
        .unwrap()
        .unwrap()
        .stories
        .into_iter()
        .map(|story| story.text)
        .collect()
}

fn assert_save_reopens(deck: &Presentation, calls: u32) {
    let reopened = Presentation::open(&deck.save().unwrap()).unwrap();
    assert_eq!(
        reopened.slides().len(),
        deck.slide_ids().unwrap().len(),
        "slides after {calls} calls"
    );
    assert_eq!(all_text(&reopened), all_text(deck), "texts after {calls} calls");
}

fn add_slides(add: impl Fn(&Presentation, &EditCtx, &[String], u32)) {
    let deck = Presentation::open(FIXTURE).unwrap();
    let ids = deck.slide_ids().unwrap();
    let count = ids.len() as u32;
    let text = slide_text(&deck, &ids[1]);
    assert!(text.iter().any(|story| !story.is_empty()));
    let context = ctx();

    for call in 1..=ADDITIONS {
        add(&deck, &context, &ids, count + call - 1);
        assert_eq!(
            deck.slide_ids().unwrap().len(),
            (count + call) as usize,
            "count after call {call}"
        );
        if call == 258 || call == ADDITIONS {
            assert_eq!(slide_text(&deck, &ids[1]), text, "second slide after call {call}");
            assert_save_reopens(&deck, call);
        }
    }
}

#[test]
fn duplicating_slides_preserves_the_original_slide() {
    add_slides(|deck, context, ids, index| {
        deck.duplicate_slide(context, &ids[0], index).unwrap();
    });
}

#[test]
fn inserting_slides_preserves_the_original_slide() {
    add_slides(|deck, context, _, index| {
        deck.insert_slide(context, index, None).unwrap();
    });
}

#[test]
fn the_first_two_hundred_added_slides_keep_their_ids() {
    let deck = Presentation::open(FIXTURE).unwrap();
    let count = deck.slide_ids().unwrap().len() as u32;
    let context = ctx();
    let added: Vec<_> = (count..count + 200)
        .map(|index| deck.insert_slide(&context, index, None).unwrap().slide_id)
        .collect();
    let expected: Vec<_> = (0..200).map(|counter| format!("slide:1:{counter}")).collect();
    assert_eq!(added, expected);
}

#[test]
fn inserting_slides_skips_only_reserved_ids() {
    let deck = Presentation::open(FIXTURE).unwrap();
    let ids = deck.slide_ids().unwrap();
    assert_eq!(ids[1], "slide:1:257");
    let context = ctx();
    deck.delete_slide(&context, &ids[1]).unwrap();
    let count = ids.len() as u32 - 1;

    let added: Vec<_> = (count..count + ADDITIONS)
        .map(|index| deck.insert_slide(&context, index, None).unwrap().slide_id)
        .collect();
    let expected: Vec<_> = (0..)
        .map(|counter| format!("slide:1:{counter}"))
        .filter(|id| !ids.contains(id))
        .take(ADDITIONS as usize)
        .collect();

    assert_eq!(added, expected);
}
