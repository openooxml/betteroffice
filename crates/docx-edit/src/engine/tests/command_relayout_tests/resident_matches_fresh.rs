//! Typing through the resident edit bridge, checked after every keystroke
//! against a fresh session built from the same document state.

use super::*;
use serde_json::Value;

const TYPED: [&str; 4] = ["x", "y", " ", "한"];
const SEEDS: u64 = 24;
const GENERATED_KEYS: usize = 40;

/// SplitMix64, so a seed always replays the same document and keystrokes.
struct Rng(u64);

impl Rng {
    fn new(seed: u64) -> Self {
        Self(seed.wrapping_mul(0x1000_0000_01B3))
    }

    fn next(&mut self) -> u64 {
        self.0 = self.0.wrapping_add(0x9E37_79B9_7F4A_7C15);
        let mut z = self.0;
        z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
        z ^ (z >> 31)
    }

    fn below(&mut self, count: usize) -> usize {
        (self.next() % count.max(1) as u64) as usize
    }

    fn chance(&mut self, one_in: usize) -> bool {
        self.below(one_in) == 0
    }
}

/// One paragraph of a story: its id, its pilcrow offset and its text segments
/// with their offsets, in UTF-16 units.
struct Para {
    id: String,
    pilcrow: u32,
    texts: Vec<(u32, String)>,
}

impl Para {
    fn has_text(&self) -> bool {
        self.texts.iter().any(|(_, text)| !text.is_empty())
    }

    fn text_start(&self) -> u32 {
        self.texts
            .iter()
            .find(|(_, text)| !text.is_empty())
            .map_or(self.pilcrow, |(start, _)| *start)
    }

    fn text_end(&self) -> u32 {
        self.texts
            .iter()
            .rev()
            .find(|(_, text)| !text.is_empty())
            .map_or(self.pilcrow, |(start, text)| {
                start + text.encode_utf16().count() as u32
            })
    }

    /// The range of every character, in UTF-16 units.
    fn characters(&self) -> Vec<(u32, u32)> {
        self.texts
            .iter()
            .flat_map(|(start, text)| {
                text.chars().scan(*start, |at, ch| {
                    let from = *at;
                    *at += ch.len_utf16() as u32;
                    Some((from, *at))
                })
            })
            .collect()
    }

    fn boundary(&self, rng: &mut Rng) -> u32 {
        let characters = self.characters();
        if characters.is_empty() || rng.chance(characters.len() + 1) {
            return self.text_end();
        }
        characters[rng.below(characters.len())].0
    }

    fn character(&self, rng: &mut Rng) -> Option<(u32, u32)> {
        let characters = self.characters();
        characters.get(rng.below(characters.len())).copied()
    }
}

fn story_paras(engine: &EngineSession, story: &str) -> Vec<Para> {
    let mut paras = Vec::new();
    let mut texts = Vec::new();
    let mut offset = 0u32;
    for segment in engine.doc().story_segments(story).unwrap() {
        match segment.content {
            crate::SegmentContent::Text(value) => {
                let length = value.encode_utf16().count() as u32;
                texts.push((offset, value));
                offset += length;
            }
            crate::SegmentContent::Pilcrow(properties) => {
                paras.push(Para {
                    id: properties.para_id.to_string(),
                    pilcrow: offset,
                    texts: std::mem::take(&mut texts),
                });
                offset += 1;
            }
            crate::SegmentContent::OtherEmbed { .. } => offset += 1,
        }
    }
    paras
}

fn cell_stories(engine: &EngineSession) -> Vec<String> {
    use yrs::{ReadTxn, Transact};
    let txn = engine.doc().yrs_doc().transact();
    let stories = txn.get_map(crate::STORIES).unwrap();
    let mut cells: Vec<String> = crate::live_stories(&stories, &txn)
        .map(|(story, _)| story.to_string())
        .filter(|story| story.starts_with("body:t"))
        .collect();
    cells.sort();
    cells
}

struct Key {
    story: String,
    start: u32,
    end: u32,
    text: Option<&'static str>,
}

impl Key {
    fn insert(story: &str, at: u32, text: &'static str) -> Self {
        Self {
            story: story.to_owned(),
            start: at,
            end: at,
            text: Some(text),
        }
    }

    fn delete(story: &str, (start, end): (u32, u32)) -> Self {
        Self {
            story: story.to_owned(),
            start,
            end,
            text: None,
        }
    }

    fn describe(&self) -> String {
        match self.text {
            Some(text) => format!("insert {text:?} at {} {}", self.story, self.start),
            None => format!("delete {} {}..{}", self.story, self.start, self.end),
        }
    }
}

fn measured_inputs(session: &EngineSession) -> Vec<Value> {
    let pagination = session.pagination.borrow();
    pagination.input.as_ref().map_or_else(Vec::new, |input| {
        input
            .measured
            .iter()
            .map(|measured| serde_json::to_value(&measured.block).unwrap())
            .collect()
    })
}

fn fresh_inputs(host: &CommandHost) -> Vec<Value> {
    let seed = host
        .seed
        .as_ref()
        .map(|(client_id, bytes)| (*client_id, bytes.as_slice()));
    let cold = cold_session(host.engine, 9636, seed);
    cold.layout_regions_for_trigger(&host.request, None, RelayoutTrigger::Open)
        .unwrap();
    measured_inputs(&cold)
}

fn assert_same_inputs(left: &[Value], right: &[Value], label: &str) {
    assert_eq!(left.len(), right.len(), "{label}: measured block count");
    for (index, (left, right)) in left.iter().zip(right).enumerate() {
        assert_eq!(left, right, "{label}: block {index}");
    }
}

fn resident_matches_fresh(host: &CommandHost, label: &str) {
    assert_same_inputs(&measured_inputs(host.engine), &fresh_inputs(host), label);
}

fn fresh_matches_fresh(host: &CommandHost, label: &str) {
    assert_same_inputs(&fresh_inputs(host), &fresh_inputs(host), label);
}

/// Deferred body patches applied and fallbacks taken on this thread so far.
fn deferred_patches() -> (usize, usize) {
    crate::engine::DEFERRED_PATCHES.with(|counts| counts.get())
}

/// Deferred patches and fallbacks since `before`.
fn deferred_since(before: (usize, usize)) -> (usize, usize) {
    let now = deferred_patches();
    (now.0 - before.0, now.1 - before.1)
}

fn assert_deferred((patched, fallbacks): (usize, usize), expected: usize, label: &str) {
    assert!(patched > 0, "{label}: no deferred patch was applied");
    assert_eq!(fallbacks, expected, "{label}: deferred fallbacks");
}

/// Opens `bytes` as a host does, types up to `keys` keystrokes from
/// `next_key` and runs `check` on the opened state and after every keystroke.
fn type_and_check(
    name: &str,
    client_id: u64,
    bytes: &[u8],
    keys: usize,
    mut next_key: impl FnMut(&EngineSession) -> Option<Key>,
    check: fn(&CommandHost, &str),
) -> (usize, usize) {
    let before = deferred_patches();
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let engine = seeded_engine(client_id, bytes);
    let request = lowering_pages::region_request(&engine, bytes, font).to_string();
    let mut host = CommandHost::open_with(&engine, request).seeded(client_id, bytes.to_vec());
    check(&host, &format!("{name}, opened"));
    let mut typed = Vec::new();
    for step in 1..=keys {
        let Some(key) = next_key(&engine) else {
            continue;
        };
        host.command(&key.story, key.start, key.end, key.text);
        typed.push(key.describe());
        check(
            &host,
            &format!("{name}, keystroke {step} [{}]", typed.join("; ")),
        );
    }
    deferred_since(before)
}

const STYLES: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style></w:styles>"#;
const SECTION: &str = r#"<w:sectPr><w:pgSz w:w="5200" w:h="4200"/><w:pgMar w:top="400" w:right="400" w:bottom="400" w:left="400" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>"#;
const FORMATS: [&str; 7] = [
    "",
    "<w:b/>",
    "<w:i/>",
    r#"<w:color w:val="C00000"/>"#,
    r#"<w:sz w:val="28"/>"#,
    r#"<w:u w:val="single"/>"#,
    "<w:b/><w:i/>",
];
const WORDS: [&str; 12] = [
    "alpha", "beta", "gamma", "delta", "river", "stone", "maple", "copper", "quiet", "north",
    "signal", "harbor",
];

fn formatted(format: &str, text: &str) -> String {
    if format.is_empty() {
        return lowering_fixture::run(text);
    }
    format!(r#"<w:r><w:rPr>{format}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

fn heading(id: &str, bookmark: Option<usize>, content: &str) -> String {
    let content = match bookmark {
        Some(mark) => format!(
            r#"<w:bookmarkStart w:id="{mark}" w:name="_Toc{mark}"/>{content}<w:bookmarkEnd w:id="{mark}"/>"#
        ),
        None => content.to_owned(),
    };
    format!(r#"<w:p w14:paraId="{id}"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>{content}</w:p>"#)
}

fn table(rows: &[[String; 2]]) -> String {
    let rows: String = rows
        .iter()
        .map(|cells| {
            let cells: String = cells
                .iter()
                .map(|content| {
                    format!(
                        r#"<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>{content}</w:tc>"#
                    )
                })
                .collect();
            format!("<w:tr>{cells}</w:tr>")
        })
        .collect();
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>{rows}</w:tbl>"#
    )
}

fn package(body: &str) -> Vec<u8> {
    lowering_fixture::Package::new(&format!("{body}{SECTION}"))
        .styles(STYLES)
        .bytes()
}

/// Two headings, paragraphs of mixed run formatting and a two by
/// two table between them.
fn typing_package() -> Vec<u8> {
    use lowering_fixture::{para, run};
    let cell = |id: &str, format: &str, text: &str| para(id, &formatted(format, text));
    let body = [
        heading("00001001", None, &run("Overview of the plan")),
        para(
            "00001002",
            &[
                run("The first paragraph has "),
                formatted("<w:b/>", "bold words"),
                run(", then "),
                formatted("<w:i/>", "italic ones"),
                run(" and a plain tail."),
            ]
            .concat(),
        ),
        table(&[
            [
                cell("00001101", "", "North"),
                cell("00001102", "<w:b/>", "South"),
            ],
            [
                cell("00001103", "<w:i/>", "East"),
                cell("00001104", "", "West wing"),
            ],
        ]),
        para(
            "00001003",
            &[
                formatted(r#"<w:color w:val="C00000"/>"#, "Red lead words "),
                run("between "),
                formatted(r#"<w:sz w:val="28"/>"#, "larger type"),
                run(" to close."),
            ]
            .concat(),
        ),
        heading("00001004", None, &run("Details")),
        para(
            "00001005",
            &[
                run("Plain start "),
                formatted(r#"<w:u w:val="single"/>"#, "underlined middle"),
                run(" plain end."),
            ]
            .concat(),
        ),
        para(
            "00001006",
            &run("A closing paragraph long enough to wrap over two lines of this narrow page."),
        ),
        para("00001007", ""),
    ]
    .concat();
    package(&body)
}

/// Where a fixed keystroke lands: units after the start of a body paragraph's
/// text, units before its end, or units into a table cell's text.
#[derive(Clone, Copy)]
enum At {
    Body(&'static str, u32),
    BodyEnd(&'static str, u32),
    Cell(usize, u32),
}

const TYPING: [(At, Option<&str>); 14] = [
    (At::BodyEnd("00001003", 0), Some("y")),
    (At::BodyEnd("00001003", 4), None),
    (At::Body("00001002", 5), Some("x")),
    (At::Cell(0, 2), Some("y")),
    (At::Body("00001005", 4), None),
    (At::Body("00001006", 10), Some("x")),
    (At::BodyEnd("00001004", 0), Some("x")),
    (At::Body("00001006", 6), Some("y")),
    (At::Cell(2, 1), None),
    (At::Body("00001002", 12), None),
    (At::Body("00001005", 8), Some("x")),
    (At::Cell(3, 0), Some("y")),
    (At::Body("00001002", 3), Some("😀")),
    (At::BodyEnd("00001007", 0), Some("x")),
];

fn typing_key(engine: &EngineSession, (at, text): (At, Option<&'static str>)) -> Key {
    let (story, id, offset) = match at {
        At::Body(id, offset) => ("body".to_owned(), id, Ok(offset)),
        At::BodyEnd(id, back) => ("body".to_owned(), id, Err(back)),
        At::Cell(index, offset) => (cell_stories(engine)[index].clone(), "", Ok(offset)),
    };
    let paras = story_paras(engine, &story);
    let para = paras
        .iter()
        .find(|para| id.is_empty() || para.id == id)
        .unwrap();
    let at = match offset {
        Ok(offset) => para.text_start() + offset,
        Err(back) => para.text_end() - back,
    };
    match text {
        Some(text) => Key::insert(&story, at, text),
        None => Key::delete(&story, (at, at + 1)),
    }
}

/// Body paragraphs, headings (some bookmarked) and tables of mixed run
/// formatting, at least one table among them.
fn generated_body(rng: &mut Rng) -> String {
    use lowering_fixture::para;
    let mut next_id = 0x2000u32;
    let mut id = || {
        next_id += 1;
        format!("{next_id:08X}")
    };
    let content = |rng: &mut Rng| -> String {
        let runs = 1 + rng.below(4);
        (0..runs)
            .map(|_| {
                let count = 1 + rng.below(4);
                let words: Vec<&str> = (0..count).map(|_| WORDS[rng.below(WORDS.len())]).collect();
                let format = FORMATS[rng.below(FORMATS.len())];
                formatted(format, &format!("{} ", words.join(" ")))
            })
            .collect()
    };
    let blocks = 5 + rng.below(5);
    let table_at = rng.below(blocks);
    let mut body = String::new();
    let mut bookmark = 0;
    for index in 0..blocks {
        if index == table_at || rng.chance(6) {
            let count = 1 + rng.below(2);
            let rows: Vec<[String; 2]> = (0..count)
                .map(|_| [para(&id(), &content(rng)), para(&id(), &content(rng))])
                .collect();
            body += &table(&rows);
        } else if rng.chance(4) {
            bookmark += 1;
            let mark = rng.chance(2).then_some(bookmark);
            body += &heading(&id(), mark, &content(rng));
        } else {
            body += &para(&id(), &content(rng));
        }
    }
    body += &para(&id(), &content(rng));
    body
}

fn generated_key(engine: &EngineSession, rng: &mut Rng) -> Option<Key> {
    let text = TYPED[rng.below(TYPED.len())];
    let kind = rng.below(10);
    if kind >= 6 {
        let cells = cell_stories(engine);
        let story = &cells[rng.below(cells.len())];
        let paras = story_paras(engine, story);
        let para = &paras[rng.below(paras.len())];
        if kind >= 8 && para.has_text() {
            return Some(Key::delete(story, para.character(rng)?));
        }
        return Some(Key::insert(story, para.boundary(rng), text));
    }
    let paras = story_paras(engine, "body");
    let with_text: Vec<&Para> = paras.iter().filter(|para| para.has_text()).collect();
    let para = if with_text.is_empty() {
        &paras[rng.below(paras.len())]
    } else {
        with_text[rng.below(with_text.len())]
    };
    match kind {
        0..=2 => Some(Key::insert("body", para.boundary(rng), text)),
        3 | 4 if para.has_text() => Some(Key::delete("body", para.character(rng)?)),
        _ => Some(Key::insert("body", para.text_end(), text)),
    }
}

/// Paragraphs of short runs, each formatted unlike its neighbours and
/// repeating the same few characters, plus one empty paragraph.
fn boundary_body(rng: &mut Rng) -> String {
    use lowering_fixture::para;
    let mut body: String = (0..5)
        .map(|index| {
            let runs: String = (0..2 + rng.below(5))
                .map(|run| {
                    let text = ["a", "ab", "ba", "b", "a😀"][rng.below(5)];
                    formatted(FORMATS[(index + run) % FORMATS.len()], text)
                })
                .collect();
            para(&format!("{:08X}", 0x3001 + index), &runs)
        })
        .collect();
    body += &para("00003010", "");
    body
}

/// Inserts at the start of a run or at a paragraph's end, or deletes one
/// character, in a body paragraph.
fn boundary_key(engine: &EngineSession, rng: &mut Rng) -> Option<Key> {
    let paras = story_paras(engine, "body");
    let para = &paras[rng.below(paras.len())];
    if para.has_text() && rng.chance(2) {
        return Some(Key::delete("body", para.character(rng)?));
    }
    let at = para
        .texts
        .get(rng.below(para.texts.len() + 1))
        .map_or(para.text_end(), |(start, _)| *start);
    Some(Key::insert("body", at, ["a", "b", "😀"][rng.below(3)]))
}

/// Types the first fixed keys, lets a peer bold a run of a paragraph they
/// typed into, then types into it again; `relayout` lays the peer's update out
/// before that key.
fn peer_formats_between_keys(relayout: bool) -> (usize, usize) {
    let before = deferred_patches();
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let bytes = typing_package();
    let engine = seeded_engine(9812, &bytes);
    let request = lowering_pages::region_request(&engine, &bytes, font).to_string();
    let mut host = CommandHost::open_with(&engine, request).seeded(9812, bytes);
    let label = format!("peer formatting, relayout {relayout}");
    for (step, key) in TYPING[..4].iter().enumerate() {
        let key = typing_key(&engine, *key);
        host.command(&key.story, key.start, key.end, key.text);
        resident_matches_fresh(&host, &format!("{label}, keystroke {}", step + 1));
    }

    let paras = story_paras(&engine, "body");
    let para = paras.iter().find(|para| para.id == "00001003").unwrap();
    let (start, text) = para
        .texts
        .iter()
        .find(|(_, text)| text == "larger type")
        .unwrap();
    let end = start + text.encode_utf16().count() as u32;
    let peer = crate::EditingDoc::new(9813);
    peer.apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    let vector = engine.doc().encode_state_vector_v1();
    peer.format_range(
        &crate::EditCtx::local("", ""),
        crate::StoryRange::new("body", *start, end),
        &crate::InlineFormatDelta {
            bold: crate::Patch::Set(true),
            ..Default::default()
        },
    )
    .unwrap();
    engine
        .doc()
        .apply_update_v1(&peer.encode_diff_v1(&vector).unwrap())
        .unwrap();
    if relayout {
        host.relayout(None);
        resident_matches_fresh(&host, &format!("{label}, peer update"));
    }

    let last = typing_key(&engine, (At::Body("00001003", 2), Some("x")));
    host.command(&last.story, last.start, last.end, last.text);
    resident_matches_fresh(&host, &format!("{label}, key after the peer"));
    deferred_since(before)
}

fn type_fixed(check: fn(&CommandHost, &str)) -> (usize, usize) {
    let bytes = typing_package();
    let mut keys = TYPING.into_iter();
    type_and_check(
        "fixed typing",
        9811,
        &bytes,
        TYPING.len(),
        |engine| keys.next().map(|key| typing_key(engine, key)),
        check,
    )
}

fn type_generated(seed: u64, keys: usize, check: fn(&CommandHost, &str)) -> (usize, usize) {
    let mut rng = Rng::new(seed);
    let bytes = package(&generated_body(&mut rng));
    type_and_check(
        &format!("seed {seed}"),
        9900 + seed,
        &bytes,
        keys,
        |engine| generated_key(engine, &mut rng),
        check,
    )
}

fn type_all_generated(check: fn(&CommandHost, &str)) -> (usize, usize) {
    let before = deferred_patches();
    for seed in 0..SEEDS {
        type_generated(seed, GENERATED_KEYS, check);
    }
    deferred_since(before)
}

#[test]
fn typing_in_formatted_paragraphs_lays_out_as_a_fresh_session_does() {
    assert_deferred(type_fixed(resident_matches_fresh), 1, "fixed typing");
}

/// Fallbacks the generated typing takes. A body keystroke right after a cell
/// keystroke patches in place: the cell keystroke re-lowers only its table,
/// which leaves the body's stored paragraphs current.
const GENERATED_FALLBACKS: usize = 99;

#[test]
fn generated_typing_lays_out_as_a_fresh_session_does() {
    assert_deferred(
        type_all_generated(resident_matches_fresh),
        GENERATED_FALLBACKS,
        "generated typing",
    );
}

#[test]
fn typing_across_run_boundaries_lays_out_as_a_fresh_session_does() {
    let before = deferred_patches();
    for seed in 0..SEEDS {
        let mut rng = Rng::new(seed);
        let bytes = package(&boundary_body(&mut rng));
        type_and_check(
            &format!("boundary seed {seed}"),
            9700 + seed,
            &bytes,
            GENERATED_KEYS,
            |engine| boundary_key(engine, &mut rng),
            resident_matches_fresh,
        );
    }
    assert_deferred(deferred_since(before), 0, "boundary typing");
}

#[test]
fn typing_after_a_peer_formats_the_paragraph_lays_out_as_a_fresh_session_does() {
    assert_deferred(peer_formats_between_keys(false), 1, "peer, no relayout");
    assert_deferred(peer_formats_between_keys(true), 1, "peer, relayout");
}

#[test]
fn fresh_sessions_agree_on_every_typed_state() {
    assert_deferred(type_fixed(fresh_matches_fresh), 1, "fixed typing");
    assert_deferred(
        type_all_generated(fresh_matches_fresh),
        GENERATED_FALLBACKS,
        "generated typing",
    );
}
