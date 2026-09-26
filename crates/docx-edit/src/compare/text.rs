//! Exact text differences of aligned paragraphs in UTF-16 offsets of their projected text.

use std::collections::BTreeMap;
use std::ops::Range;
use std::sync::Arc;

use ooxml_diff::{DiffKind, DiffLimits, LimitFallback};
use unicode_segmentation::UnicodeSegmentation;
use yrs::types::Attrs;
use yrs::{Any, Text, Transact};

use super::align::{Budget, Exhausted};
use super::format::ComplexScript;
use super::options::Granularity;
use crate::ops::ChunkKind;
use crate::ops::text::RichRun;
use crate::target::{AtomKind, ParagraphView, TextAtom};
use crate::{DEL, EditingDoc, INS, story_ref};

/// A token of projected text; atoms compare by kind.
#[derive(Clone, Debug)]
struct Token<'a> {
    text: &'a str,
    atom: Option<AtomKind>,
    units: Range<u32>,
}

impl PartialEq for Token<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.text == other.text && self.atom == other.atom
    }
}

fn tokens<'a>(text: &'a str, atoms: &[TextAtom], granularity: Granularity) -> Vec<Token<'a>> {
    let segments: Vec<&str> = match granularity {
        Granularity::Word => text.split_word_bounds().collect(),
        Granularity::Char => text.graphemes(true).collect(),
    };
    let mut out = Vec::with_capacity(segments.len());
    let mut units = 0u32;
    for segment in segments {
        let mut rest = segment;
        while !rest.is_empty() {
            let (piece, tail) = match rest.find('\u{FFFC}') {
                Some(0) => rest.split_at('\u{FFFC}'.len_utf8()),
                Some(index) => rest.split_at(index),
                None => (rest, ""),
            };
            let len = piece.encode_utf16().count() as u32;
            let atom = (piece == "\u{FFFC}").then(|| {
                atoms
                    .iter()
                    .find(|atom| atom.offset == units)
                    .map_or(AtomKind::Other, |atom| atom.kind)
            });
            out.push(Token {
                text: piece,
                atom,
                units: units..units + len,
            });
            units += len;
            rest = tail;
        }
    }
    out
}

/// One maximal differing stretch: replaced original units and their revised replacement.
#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) struct Hunk {
    pub original: Range<u32>,
    pub revised: Range<u32>,
}

/// The differences between two paragraph texts and the stretches they share.
#[derive(Debug, Default, Eq, PartialEq)]
pub(crate) struct TextDiff {
    pub hunks: Vec<Hunk>,
    pub equal: Vec<Hunk>,
}

fn diff(
    original: &[Token<'_>],
    revised: &[Token<'_>],
    budget: &mut Budget,
) -> Result<TextDiff, Exhausted> {
    let limits = DiffLimits::new(usize::MAX, budget.diff_cells, LimitFallback::Error);
    let result = ooxml_diff::diff_tokens(original, revised, limits).map_err(|_| Exhausted::Diff)?;
    budget.charge_diff(result.cells_used)?;
    let units = |tokens: &[Token<'_>], range: Range<usize>, at: usize| match range.is_empty() {
        true => {
            let offset = tokens.get(at).map_or_else(
                || tokens.last().map_or(0, |token| token.units.end),
                |token| token.units.start,
            );
            offset..offset
        }
        false => tokens[range.start].units.start..tokens[range.end - 1].units.end,
    };
    let mut out = TextDiff::default();
    let mut pending: Option<Hunk> = None;
    for hunk in result.hunks {
        let original_units = units(original, hunk.old.clone(), hunk.old.start);
        let revised_units = units(revised, hunk.new.clone(), hunk.new.start);
        if hunk.kind == DiffKind::Equal {
            out.hunks.extend(pending.take());
            out.equal.push(Hunk {
                original: original_units,
                revised: revised_units,
            });
            continue;
        }
        match pending.as_mut() {
            Some(open) => {
                open.original.end = open.original.end.max(original_units.end);
                open.revised.end = open.revised.end.max(revised_units.end);
            }
            None => {
                pending = Some(Hunk {
                    original: original_units,
                    revised: revised_units,
                })
            }
        }
    }
    out.hunks.extend(pending);
    Ok(out)
}

/// Diffs two paragraph texts at `granularity`.
pub(crate) fn diff_paragraphs(
    (original, original_atoms): (&str, &[TextAtom]),
    (revised, revised_atoms): (&str, &[TextAtom]),
    granularity: Granularity,
    budget: &mut Budget,
) -> Result<TextDiff, Exhausted> {
    let left = tokens(original, original_atoms, granularity);
    let right = tokens(revised, revised_atoms, granularity);
    diff(&left, &right, budget)
}

/// The grapheme stretches a replacement keeps, so formatting on them can be checked.
pub(crate) fn retained_within(
    original: &str,
    revised: &str,
    budget: &mut Budget,
) -> Result<Vec<Hunk>, Exhausted> {
    let left = tokens(original, &[], Granularity::Char);
    let right = tokens(revised, &[], Granularity::Char);
    Ok(diff(&left, &right, budget)?.equal)
}

/// The substring of `text` between two UTF-16 offsets on scalar boundaries.
pub(crate) fn slice(text: &str, range: Range<u32>) -> &str {
    let byte = |target: u32| {
        let mut units = 0u32;
        for (index, ch) in text.char_indices() {
            if units >= target {
                return index;
            }
            units += ch.len_utf16() as u32;
        }
        text.len()
    };
    let start = byte(range.start);
    &text[start..byte(range.end).max(start)]
}

/// Formatting attributes of one unit in the story vocabulary, without revision or link marks.
pub(crate) type UnitAttrs = Arc<BTreeMap<String, Any>>;

/// The formatting of each UTF-16 unit of `paragraph`, read from `doc`'s body story.
/// Session attributes stating complex-script bold and italic, which a save otherwise derives from
/// bold and italic.
pub(crate) const COMPLEX_SCRIPT_BOLD: &str = "boldCs";
pub(crate) const COMPLEX_SCRIPT_ITALIC: &str = "italicCs";

/// States each unit's own complex-script bold and italic on `paragraph` of `story`, outside undo
/// history.
pub(crate) fn state_complex_script(
    doc: &EditingDoc,
    story: &str,
    paragraph: &ParagraphView,
    runs: &ComplexScript,
) {
    let mut txn = doc.yrs_doc().transact_mut_with("system");
    let Ok(text) = story_ref(&txn, story) else {
        return;
    };
    let mut offset = 0;
    for &(units, [bold, italic]) in runs {
        let attrs = Attrs::from([
            (Arc::from(COMPLEX_SCRIPT_BOLD), Any::Bool(bold)),
            (Arc::from(COMPLEX_SCRIPT_ITALIC), Any::Bool(italic)),
        ]);
        for raw in paragraph.raw_ranges(offset..offset + units) {
            text.format(&mut txn, raw.start, raw.end - raw.start, attrs.clone());
        }
        offset += units;
    }
}

pub(crate) fn unit_attrs(
    doc: &EditingDoc,
    story: &str,
    paragraph: &ParagraphView,
) -> Vec<UnitAttrs> {
    let txn = doc.yrs_doc().transact();
    let Ok(text) = story_ref(&txn, story) else {
        return Vec::new();
    };
    let chunks = doc.chunk_snapshot(story, &text, &txn);
    unit_attrs_in(&chunks, paragraph)
}

fn unit_attrs_in(chunks: &[crate::ops::Chunk], paragraph: &ParagraphView) -> Vec<UnitAttrs> {
    let empty: UnitAttrs = Arc::new(BTreeMap::new());
    let mut cache: Vec<Option<UnitAttrs>> = vec![None; chunks.len()];
    (0..paragraph.len())
        .map(|offset| {
            let raw = paragraph.raw_at(offset);
            let index = chunks.partition_point(|chunk| chunk.end() <= raw);
            let Some(chunk) = chunks.get(index).filter(|chunk| chunk.start <= raw) else {
                return Arc::clone(&empty);
            };
            if matches!(chunk.kind, ChunkKind::Pilcrow(_)) {
                return Arc::clone(&empty);
            }
            Arc::clone(cache[index].get_or_insert_with(|| {
                Arc::new(
                    chunk
                        .attrs
                        .iter()
                        .filter(|(key, value)| {
                            !matches!(key.as_str(), INS | DEL | crate::format::HYPERLINK)
                                && **value != Any::Null
                        })
                        .map(|(key, value)| (key.clone(), value.clone()))
                        .collect(),
                )
            }))
        })
        .collect()
}

/// `text`'s units `range` as runs of equal formatting.
pub(crate) fn rich_runs(text: &str, range: Range<u32>, attrs: &[UnitAttrs]) -> Vec<RichRun> {
    let mut runs: Vec<RichRun> = Vec::new();
    let mut run_start = range.start;
    let mut offset = range.start;
    let slice_text = slice(text, range.clone());
    let mut chars = slice_text.chars().peekable();
    let mut current = String::new();
    while let Some(ch) = chars.next() {
        let len = ch.len_utf16() as u32;
        current.push(ch);
        offset += len;
        let boundary =
            chars.peek().is_none() || attrs.get(offset as usize) != attrs.get(run_start as usize);
        if boundary {
            runs.push(RichRun {
                text: std::mem::take(&mut current),
                attrs: attrs
                    .get(run_start as usize)
                    .map(|attrs| (**attrs).clone())
                    .unwrap_or_default(),
            });
            run_start = offset;
        }
    }
    runs
}

/// A canonical string for comparing formatting attributes.
pub(crate) fn attrs_key(attrs: &BTreeMap<String, Any>) -> String {
    let mut out = String::new();
    for (key, value) in attrs {
        out.push_str(key);
        out.push('=');
        any_key(value, &mut out);
        out.push(';');
    }
    out
}

fn any_key(value: &Any, out: &mut String) {
    match value {
        Any::Map(map) => {
            let sorted: BTreeMap<&str, &Any> = map
                .iter()
                .map(|(key, value)| (key.as_str(), value))
                .collect();
            out.push('{');
            for (key, value) in sorted {
                out.push_str(key);
                out.push(':');
                any_key(value, out);
                out.push(',');
            }
            out.push('}');
        }
        Any::Array(values) => {
            out.push('[');
            for value in values.iter() {
                any_key(value, out);
                out.push(',');
            }
            out.push(']');
        }
        other => out.push_str(&format!("{other:?}")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn budget() -> Budget {
        Budget {
            alignment_cells: 250_000,
            diff_cells: 4_000_000,
        }
    }

    fn hunks(original: &str, revised: &str, granularity: Granularity) -> Vec<(String, String)> {
        diff_paragraphs((original, &[]), (revised, &[]), granularity, &mut budget())
            .unwrap()
            .hunks
            .into_iter()
            .map(|hunk| {
                (
                    slice(original, hunk.original).to_owned(),
                    slice(revised, hunk.revised).to_owned(),
                )
            })
            .collect()
    }

    fn pair(left: &str, right: &str) -> (String, String) {
        (left.to_owned(), right.to_owned())
    }

    #[test]
    fn word_diffs_keep_punctuation_and_whitespace_tokens() {
        assert_eq!(
            hunks(
                "The quick brown fox.",
                "The slow brown fox!",
                Granularity::Word
            ),
            vec![pair("quick", "slow"), pair(".", "!")]
        );
        assert_eq!(
            hunks("a  b", "a b", Granularity::Word),
            vec![pair("  ", " ")]
        );
        assert_eq!(
            hunks("keep", "keep more", Granularity::Word),
            vec![pair("", " more")]
        );
        assert_eq!(hunks("same", "same", Granularity::Word), vec![]);
    }

    #[test]
    fn char_diffs_split_on_grapheme_clusters() {
        let family = "\u{1F468}\u{200D}\u{1F469}\u{200D}\u{1F467}";
        assert_eq!(
            hunks(&format!("a{family}b"), "ab", Granularity::Char),
            vec![pair(family, "")]
        );
        assert_eq!(
            hunks("e\u{0301}", "\u{00e9}", Granularity::Char),
            vec![pair("e\u{0301}", "\u{00e9}")]
        );
        let hunk = diff_paragraphs(
            ("x\u{1F600}y", &[]),
            ("x\u{1F601}y", &[]),
            Granularity::Char,
            &mut budget(),
        )
        .unwrap()
        .hunks;
        assert_eq!(
            hunk,
            vec![Hunk {
                original: 1..3,
                revised: 1..3
            }]
        );
    }

    #[test]
    fn atoms_compare_by_kind() {
        let line = [TextAtom {
            offset: 1,
            kind: AtomKind::LineBreak,
        }];
        let image = [TextAtom {
            offset: 1,
            kind: AtomKind::Image,
        }];
        let same = diff_paragraphs(
            ("a\u{FFFC}b", &line),
            ("a\u{FFFC}c", &line),
            Granularity::Word,
            &mut budget(),
        )
        .unwrap();
        assert_eq!(
            same.hunks,
            vec![Hunk {
                original: 2..3,
                revised: 2..3
            }]
        );
        let other = diff_paragraphs(
            ("a\u{FFFC}b", &line),
            ("a\u{FFFC}b", &image),
            Granularity::Word,
            &mut budget(),
        )
        .unwrap();
        assert_eq!(
            other.hunks,
            vec![Hunk {
                original: 1..2,
                revised: 1..2
            }]
        );
    }

    #[test]
    fn diffs_charge_the_budget() {
        let mut tight = Budget {
            alignment_cells: 0,
            diff_cells: 3,
        };
        assert_eq!(
            diff_paragraphs(("a b", &[]), ("c d", &[]), Granularity::Word, &mut tight),
            Err(Exhausted::Diff)
        );
    }

    #[test]
    fn slices_use_utf16_offsets() {
        assert_eq!(slice("a\u{1F600}b", 1..3), "\u{1F600}");
        assert_eq!(slice("abc", 3..3), "");
        assert_eq!(slice("abc", 0..3), "abc");
    }
}
