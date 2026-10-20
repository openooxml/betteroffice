//! A replica that loses a concurrent edit to a peer's must not publish deletions that leave a
//! third replica, which has not seen the winning edit yet, in a state no replica ever held.

use std::cell::RefCell;
use std::rc::Rc;

use docx_edit::*;

const DATE: &str = "2026-10-16T00:00:00Z";
const PEER: u64 = 201;
const LOSER: u64 = 202;
const WINNER: u64 = 203;

struct Race {
    /// The loser's paragraphs right after its own edit.
    seen: Vec<ParagraphSnapshot>,
    /// The peer's paragraphs after everything the loser published, before the winner's edit.
    early: Vec<ParagraphSnapshot>,
}

fn ctx() -> EditCtx {
    EditCtx::local("Ada", DATE)
}

/// Runs `edit(doc, paragraph, is_loser)` on the loser and the winner from the same state. The
/// peer gets everything the loser publishes, in order, before the winner's edit; then all three
/// must converge.
fn race(edit: impl Fn(&EditingDoc, &ParagraphId, bool)) -> Race {
    let origin = EditingDoc::new(100);
    let para = origin
        .create_story("body", "Hello world", "Normal", "left")
        .unwrap();
    let state = origin.encode_state_as_update_v1();
    let open = |client| {
        let doc = EditingDoc::new(client);
        doc.apply_update_v1(&state).unwrap();
        doc
    };
    let (peer, loser, winner) = (open(PEER), open(LOSER), open(WINNER));

    let base = loser.encode_state_vector_v1();
    edit(&loser, &para, true);
    edit(&winner, &para, false);
    let lost = loser.encode_diff_v1(&base).unwrap();
    let won = winner.encode_diff_v1(&base).unwrap();
    let seen = loser.paragraphs("body").unwrap();

    let published = Rc::new(RefCell::new(Vec::new()));
    let sink = Rc::clone(&published);
    let _subscription = loser
        .observe_update_v1(move |update, remote| {
            if !remote {
                sink.borrow_mut().push(update.to_vec());
            }
        })
        .unwrap();
    loser.apply_update_v1(&won).unwrap();

    peer.apply_update_v1(&lost).unwrap();
    for update in published.borrow().iter() {
        peer.apply_update_v1(update).unwrap();
    }
    let early = peer.paragraphs("body").unwrap();

    peer.apply_update_v1(&won).unwrap();
    winner.apply_update_v1(&lost).unwrap();
    for update in published.borrow().iter() {
        winner.apply_update_v1(update).unwrap();
    }
    let converged = loser.story_segments("body").unwrap();
    assert_eq!(winner.story_segments("body").unwrap(), converged);
    assert_eq!(peer.story_segments("body").unwrap(), converged);
    assert_eq!(
        winner.paragraphs("body").unwrap(),
        loser.paragraphs("body").unwrap()
    );
    assert_eq!(
        peer.paragraphs("body").unwrap(),
        loser.paragraphs("body").unwrap()
    );
    assert_eq!(
        peer.encode_state_vector_v1(),
        loser.encode_state_vector_v1()
    );
    Race { seen, early }
}

#[test]
fn a_peer_keeps_the_losers_paragraph_attributes_until_the_winner_arrives() {
    let race = race(|doc, para, loser| {
        let delta = ParaAttrDelta {
            alignment: Patch::Set(if loser { "center" } else { "right" }.into()),
            indent_left: Patch::Set(if loser { 360.0 } else { 720.0 }),
            ..ParaAttrDelta::default()
        };
        doc.set_paragraph_attrs(&ctx(), &ParaSelector::One(para.clone()), &delta)
            .unwrap();
    });
    assert_eq!(race.early, race.seen);
}

#[test]
fn a_peer_keeps_every_paragraph_id_of_a_lost_concurrent_split() {
    let race = race(|doc, _, loser| {
        let at = if loser { 2 } else { 8 };
        doc.split_paragraph(&ctx(), Position::new("body", at), None)
            .unwrap();
    });
    assert!(
        race.early
            .iter()
            .all(|paragraph| !paragraph.para_id.is_empty())
    );
    assert_eq!(race.early, race.seen);
}

#[test]
fn a_concurrent_run_format_converges_and_stays_readable() {
    let race = race(|doc, _, loser| {
        let delta = InlineFormatDelta {
            highlight: Patch::Set(if loser { "yellow" } else { "green" }.into()),
            ..InlineFormatDelta::default()
        };
        doc.format_range(&ctx(), StoryRange::new("body", 0, 5), &delta)
            .unwrap();
    });
    let text = |paragraphs: &[ParagraphSnapshot]| {
        paragraphs
            .iter()
            .map(|paragraph| paragraph.text.clone())
            .collect::<Vec<_>>()
    };
    assert_eq!(text(&race.early), text(&race.seen));
}
