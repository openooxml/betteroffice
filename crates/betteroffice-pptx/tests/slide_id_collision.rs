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

fn saved_slides(deck: &Presentation) -> usize {
    let bytes = deck.save().unwrap();
    Presentation::open(&bytes).unwrap().slides().len()
}

#[test]
fn duplicating_slides_preserves_the_original_slide() {
    let deck = Presentation::open(FIXTURE).unwrap();
    let ids = deck.slide_ids().unwrap();
    let count = ids.len() as u32;
    let text = slide_text(&deck, &ids[1]);
    assert!(text.iter().any(|story| !story.is_empty()));
    let context = ctx();

    for index in count..count + ADDITIONS {
        deck.duplicate_slide(&context, &ids[0], index).unwrap();
        assert_eq!(deck.slide_ids().unwrap().len(), index as usize + 1);
    }

    assert_eq!(slide_text(&deck, &ids[1]), text);
    assert_eq!(saved_slides(&deck), (count + ADDITIONS) as usize);
}

#[test]
fn inserting_slides_preserves_the_original_slide() {
    let deck = Presentation::open(FIXTURE).unwrap();
    let ids = deck.slide_ids().unwrap();
    let count = ids.len() as u32;
    let text = slide_text(&deck, &ids[1]);
    assert!(text.iter().any(|story| !story.is_empty()));
    let context = ctx();

    for index in count..count + ADDITIONS {
        deck.insert_slide(&context, index, None).unwrap();
        assert_eq!(deck.slide_ids().unwrap().len(), index as usize + 1);
    }

    assert_eq!(slide_text(&deck, &ids[1]), text);
    assert_eq!(saved_slides(&deck), (count + ADDITIONS) as usize);
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
