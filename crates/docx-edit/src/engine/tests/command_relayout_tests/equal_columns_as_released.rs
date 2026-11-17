//! Text typed into a one-column section above a continuous section of three
//! equal columns that keeps its side margins, with a page-anchored picture
//! over the second column.

use super::equal_column_pictures::{
    assert_input_matches_fresh, equal_columns, layout_digest, opened, text_end, words,
};
use super::*;

const FIRST: &str = "00002000";
const RELEASED_PAGES_AFTER_TYPING: usize = 4;

fn client_for(lead: usize) -> u64 {
    if lead % 3 == 0 {
        41_000 + 2 * (21 + lead as u64 / 3) - 1
    } else {
        51_000 + lead as u64
    }
}

/// Types `count` words at the end of the one-column section and then removes
/// them, calling `check` while they are there.
fn type_and_remove(
    host: &mut CommandHost<'_>,
    engine: &EngineSession,
    count: usize,
    check: &mut dyn FnMut(&mut CommandHost<'_>),
) {
    let text = format!(" {}", words(count, 700));
    let at = text_end(engine, FIRST);
    host.command("body", at, at, Some(&text));
    check(host);
    host.command("body", at, at + text.chars().count() as u32, None);
}

#[test]
fn equal_columns_after_typing_in_the_section_above_match_a_fresh_open() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    for lead in 12..=18 {
        let client_id = client_for(lead);
        let bytes = equal_columns(3, lead);
        let engine = seeded_engine(client_id, &bytes);
        let mut host = opened(&engine, &bytes).seeded(client_id, bytes.clone());
        for count in [1, 8, 20, 40] {
            type_and_remove(&mut host, &engine, count, &mut |host| {
                assert_input_matches_fresh(
                    host,
                    client_id,
                    &bytes,
                    &format!("lead of {lead} words, {count} words typed above"),
                );
            });
        }
    }
}

#[test]
fn equal_columns_after_typing_in_the_section_above_keep_the_released_page_count() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let lead = 12;
    let client_id = client_for(lead);
    let bytes = equal_columns(3, lead);
    let engine = seeded_engine(client_id, &bytes);
    let mut host = opened(&engine, &bytes).seeded(client_id, bytes.clone());
    let mut pages = Vec::new();
    for count in [1, 8, 20, 40] {
        type_and_remove(&mut host, &engine, count, &mut |_| {
            pages.push(layout_digest(&engine).0);
        });
    }
    assert_eq!(pages.last(), Some(&RELEASED_PAGES_AFTER_TYPING));
}
