use betteroffice_pptx::{EditCtx, Presentation};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn ctx() -> EditCtx {
    EditCtx::local("save-test")
}

fn saved_slides(deck: &Presentation) -> usize {
    let bytes = deck.save().unwrap();
    Presentation::open(&bytes).unwrap().slides().len()
}

#[test]
fn save_after_deleting_a_slide() {
    let deck = Presentation::open(FIXTURE).unwrap();
    let ids = deck.slide_ids().unwrap();
    let count = ids.len();
    deck.delete_slide(&ctx(), &ids[count - 1]).unwrap();
    assert_eq!(saved_slides(&deck), count - 1);
    assert_eq!(saved_slides(&deck), count - 1);
}

#[test]
fn save_after_inserting_duplicating_and_moving_slides() {
    let deck = Presentation::open(FIXTURE).unwrap();
    let count = deck.slide_ids().unwrap().len();
    deck.insert_slide(&ctx(), 0, None).unwrap();
    assert_eq!(saved_slides(&deck), count + 1);
    let ids = deck.slide_ids().unwrap();
    deck.duplicate_slide(&ctx(), &ids[1], ids.len() as u32)
        .unwrap();
    assert_eq!(saved_slides(&deck), count + 2);
    let ids = deck.slide_ids().unwrap();
    deck.move_slide(&ctx(), &ids[0], (ids.len() - 1) as u32)
        .unwrap();
    deck.delete_slide(&ctx(), &ids[1]).unwrap();
    assert_eq!(saved_slides(&deck), count + 1);
    assert_eq!(saved_slides(&deck), count + 1);
}
