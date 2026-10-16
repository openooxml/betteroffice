//! Permission range markers, which the model does not keep, carried from a story part's source
//! paragraphs into the paragraphs a save writes for them.

use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::ops::Range;

use crate::paragraph_identity::{attribute, paragraph_occurrences, tags, unescaped};

use super::context::RecordedParagraphs;
use super::paragraph_ids::S13SplicedPart;
use super::splice::{POSITIONED, paragraph_span};

const MARKERS: [&str; 2] = ["w:permStart", "w:permEnd"];

/// The text that stands between two paragraphs read as one.
const SEPARATOR: char = '\u{2029}';

/// The most insertions and deletions [`edits`] looks for before it takes everything between
/// the common prefix and suffix as one edit.
const EDIT_LIMIT: usize = 256;

pub(crate) fn has_markers(xml: &str) -> bool {
    MARKERS.iter().any(|marker| xml.contains(marker))
}

/// A recorded paragraph: written from the source paragraph with this ordinal, or the one at
/// this index of those written without one.
#[derive(Clone, Copy)]
enum Slot {
    Written(u32),
    Inserted(usize),
}

pub(crate) fn token_prefix(source: &str) -> String {
    let mut prefix = "<!--bo-paragraph ".to_owned();
    while source.contains(&prefix) {
        prefix.push('_');
    }
    prefix
}

/// The comment before the paragraph recorded at `position` in writing order.
pub(crate) fn token(prefix: &str, position: usize) -> String {
    format!("{prefix}{position}-->")
}

/// Carries the permission markers of the paragraphs of `source` into those `recorded` holds
/// for `part` (see [`placements`]), and into `serialized`, which holds each right after its
/// [`token`] when `source` holds a marker, and drops the tokens.
pub(crate) fn keep(
    source: &str,
    part: &S13SplicedPart,
    recorded: &mut RecordedParagraphs,
    serialized: &mut String,
) {
    let Some(prefix) = recorded.token_prefix.as_deref() else {
        return;
    };
    let replacements = placements(source, part, recorded).unwrap_or_default();
    let mut output = String::with_capacity(serialized.len() + 256);
    let mut cursor = 0;
    let mut inserted = 0;
    let mut updates = Vec::new();
    for (position, entry) in recorded.order.iter().enumerate() {
        let (slot, old) = match *entry {
            Some(ordinal) => (Slot::Written(ordinal), recorded.written.get(&ordinal)),
            None => {
                inserted += 1;
                (
                    Slot::Inserted(inserted - 1),
                    recorded.inserted.get(inserted - 1),
                )
            }
        };
        let token = token(prefix, position);
        let Some(found) = serialized[cursor..].find(&token) else {
            continue;
        };
        output.push_str(&serialized[cursor..cursor + found]);
        cursor += found + token.len();
        if let Some(new) = replacements.get(&position) {
            if let Some(old) = old.filter(|old| serialized[cursor..].starts_with(old.as_str())) {
                output.push_str(new);
                cursor += old.len();
            }
            updates.push((slot, new.clone()));
        }
    }
    output.push_str(&serialized[cursor..]);
    *serialized = output;
    for (slot, new) in updates {
        match slot {
            Slot::Written(ordinal) => {
                recorded.written.insert(ordinal, new);
            }
            Slot::Inserted(index) => recorded.inserted[index] = new,
        }
    }
}

/// The XML with permission markers of each paragraph `recorded` holds for `part` that takes
/// one, by its position in writing order; `None` when none does. Each paragraph written from a
/// source paragraph takes, as one text, that paragraph and the removed ones whose nearest
/// written paragraph it is (the one before them, else after), and places their markers in
/// itself and the paragraphs written without a source paragraph that follow it, before the
/// next one written from one. Paragraphs of groups without markers are not read.
fn placements(
    source: &str,
    part: &S13SplicedPart,
    recorded: &RecordedParagraphs,
) -> Option<BTreeMap<usize, String>> {
    let occurrences = paragraph_occurrences(source)?;
    let xml = |ordinal: u32| {
        occurrences
            .get(ordinal as usize)
            .filter(|occurrence| occurrence.ordinal == ordinal)
            .and_then(|occurrence| paragraph_span(source, occurrence))
            .map(|span| &source[span])
    };
    let written: BTreeSet<u32> = recorded.written.keys().copied().collect();
    let mut taken: BTreeMap<u32, Vec<u32>> = BTreeMap::new();
    for &ordinal in &part.removed {
        let taker = written
            .range(..ordinal)
            .next_back()
            .or_else(|| written.range(ordinal..).next());
        if let Some(&taker) = taker {
            taken.entry(taker).or_default().push(ordinal);
        }
    }
    let mut sources: BTreeMap<u32, Vec<&str>> = BTreeMap::new();
    for &ordinal in &written {
        let mut ordinals = taken.remove(&ordinal).unwrap_or_default();
        ordinals.push(ordinal);
        ordinals.sort_unstable();
        let paragraphs: Vec<&str> = ordinals.into_iter().filter_map(&xml).collect();
        if paragraphs.iter().any(|paragraph| has_markers(paragraph)) {
            sources.insert(ordinal, paragraphs);
        }
    }
    if sources.is_empty() {
        return None;
    }
    let mut slots = Vec::with_capacity(recorded.order.len());
    let mut groups: BTreeMap<u32, Vec<usize>> = BTreeMap::new();
    let mut leading = Vec::new();
    let mut current = None;
    let mut inserted = 0;
    for entry in &recorded.order {
        let position = slots.len();
        match *entry {
            Some(ordinal) => {
                slots.push(Slot::Written(ordinal));
                current = Some(ordinal);
                if sources.contains_key(&ordinal) {
                    let group = groups.entry(ordinal).or_default();
                    group.append(&mut leading);
                    group.push(position);
                }
                leading.clear();
            }
            None => {
                slots.push(Slot::Inserted(inserted));
                inserted += 1;
                match current {
                    Some(ordinal) => {
                        if let Some(group) = groups.get_mut(&ordinal) {
                            group.push(position);
                        }
                    }
                    None => leading.push(position),
                }
            }
        }
    }
    let text = |slot: Slot| match slot {
        Slot::Written(ordinal) => recorded.written.get(&ordinal).map(String::as_str),
        Slot::Inserted(index) => recorded.inserted.get(index).map(String::as_str),
    };
    let mut replacements: BTreeMap<usize, String> = BTreeMap::new();
    for (ordinal, positions) in &groups {
        let Some(written) = positions
            .iter()
            .map(|&position| text(slots[position]))
            .collect::<Option<Vec<&str>>>()
        else {
            continue;
        };
        for (position, carried) in positions.iter().zip(carry(&sources[ordinal], &written)) {
            if let Some(carried) = carried {
                replacements.insert(*position, carried);
            }
        }
    }
    Some(replacements)
}

/// A permission marker: the offset in its paragraph's text it sits at, how many empty siblings
/// precede it there, its kind and decoded ID, and its markup.
struct Marker<'x> {
    offset: usize,
    rank: usize,
    start: bool,
    id: String,
    xml: &'x str,
}

/// A paragraph read once: its text, positioned run content each as one character; its
/// permission markers at any depth; and the gaps between its children after its properties,
/// each as the byte it starts at and the text offset it holds, in order.
#[derive(Default)]
struct Reading<'x> {
    text: Vec<char>,
    markers: Vec<Marker<'x>>,
    gaps: Vec<(usize, usize)>,
}

fn read(xml: &str) -> Option<Reading<'_>> {
    let mut reading = Reading::default();
    let mut depth = 0usize;
    let mut in_text = 0usize;
    let mut in_properties = 0usize;
    let mut previous = 0;
    let mut child = None;
    let mut open_marker: Option<(usize, usize)> = None;
    let mut empty_run = (0, 0);
    for tag in tags(xml)? {
        if in_text > 0 {
            reading
                .text
                .extend(unescaped(xml, previous..tag.range.start)?.chars());
        }
        previous = tag.range.end;
        let text = matches!(tag.name, "w:t" | "w:delText");
        let properties = tag.name.ends_with("Pr");
        let offset = reading.text.len();
        if tag.end {
            depth = depth.checked_sub(1)?;
            if text {
                in_text = in_text.saturating_sub(1);
            } else if properties {
                in_properties = in_properties.saturating_sub(1);
            }
            if let Some((index, start)) = open_marker.take() {
                reading.markers.get_mut(index)?.xml = &xml[start..tag.range.end];
            }
        } else if depth == 0 {
            if !reading.gaps.is_empty() || tag.name != "w:p" || tag.empty {
                return None;
            }
            reading.gaps.push((tag.range.end, 0));
            depth = 1;
            continue;
        } else {
            if depth == 1 {
                child = Some((tag.name, offset));
            }
            if MARKERS.contains(&tag.name) {
                if !tag.empty {
                    open_marker = Some((reading.markers.len(), tag.range.start));
                }
                reading.markers.push(Marker {
                    offset,
                    rank: if empty_run.0 == offset {
                        empty_run.1
                    } else {
                        0
                    },
                    start: tag.name == "w:permStart",
                    id: attribute(&tag, "w:id")
                        .and_then(|range| unescaped(xml, range))
                        .unwrap_or_default(),
                    xml: &xml[tag.range.clone()],
                });
            }
            if properties {
                in_properties += usize::from(!tag.empty);
            } else if text {
                in_text += usize::from(!tag.empty);
            } else if in_properties == 0
                && let Some(index) = POSITIONED.iter().position(|name| *name == tag.name)
            {
                reading.text.push(char::from_u32(0xE000 + index as u32)?);
            }
            if !tag.empty {
                depth += 1;
                continue;
            }
        }
        if depth != 1 {
            continue;
        }
        let Some((name, start)) = child.take() else {
            continue;
        };
        let end = reading.text.len();
        if name == "w:pPr" && reading.gaps.len() == 1 {
            reading.gaps[0] = (tag.range.end, end);
            continue;
        }
        reading.gaps.push((tag.range.end, end));
        if start != end {
            empty_run = (end, 0);
        } else if !MARKERS.contains(&name) && name != "w:proofErr" {
            empty_run = if empty_run.0 == end {
                (end, empty_run.1 + 1)
            } else {
                (end, 1)
            };
        }
    }
    (depth == 0 && !reading.gaps.is_empty()).then_some(reading)
}

/// The edits that turn `old` into `new`, in order, each as the range of `old` it replaces and
/// the range of `new` it writes: one for each run of changes in a shortest edit script, two
/// joined when the text they leave between them is shorter than each, or one for everything
/// between the common prefix and suffix when that script needs more than [`EDIT_LIMIT`]
/// insertions and deletions.
fn edits(old: &[char], new: &[char]) -> Vec<(Range<usize>, Range<usize>)> {
    let prefix = old.iter().zip(new).take_while(|(a, b)| a == b).count();
    let suffix = old[prefix..]
        .iter()
        .rev()
        .zip(new[prefix..].iter().rev())
        .take_while(|(a, b)| a == b)
        .count();
    let (cut, kept) = (old.len() - suffix, new.len() - suffix);
    if prefix == cut && prefix == kept {
        return Vec::new();
    }
    let whole = vec![(prefix..cut, prefix..kept)];
    if prefix == cut || prefix == kept {
        return whole;
    }
    let (a, b) = (&old[prefix..cut], &new[prefix..kept]);
    let Some(snakes) = snakes(a, b) else {
        return whole;
    };
    let size = |(old, new): &(Range<usize>, Range<usize>)| old.len().max(new.len());
    let mut edits: Vec<(Range<usize>, Range<usize>)> = Vec::new();
    let (mut x, mut y) = (0, 0);
    for (sx, sy, length) in snakes.into_iter().chain([(a.len(), b.len(), 0)]) {
        if (sx, sy) != (x, y) {
            let edit = (prefix + x..prefix + sx, prefix + y..prefix + sy);
            match edits.last_mut() {
                Some(last)
                    if edit.0.start - last.0.end < size(last)
                        && edit.0.start - last.0.end < size(&edit) =>
                {
                    *last = (last.0.start..edit.0.end, last.1.start..edit.1.end);
                }
                _ => edits.push(edit),
            }
        }
        (x, y) = (sx + length, sy + length);
    }
    edits
}

/// The runs `a` and `b` share in a shortest edit script between them, each as where it starts
/// in both and its length, in order; `None` when that script needs more than [`EDIT_LIMIT`]
/// insertions and deletions.
fn snakes(a: &[char], b: &[char]) -> Option<Vec<(usize, usize, usize)>> {
    let (n, m) = (a.len() as isize, b.len() as isize);
    let limit = EDIT_LIMIT.min(a.len() + b.len()) as isize;
    let index = |k: isize| (k + limit + 1) as usize;
    let mut v = vec![0isize; 2 * limit as usize + 3];
    let mut trace = Vec::new();
    for d in 0..=limit {
        trace.push(v.clone());
        for k in (-d..=d).step_by(2) {
            let down = k == -d || (k != d && v[index(k - 1)] < v[index(k + 1)]);
            let mut x = if down {
                v[index(k + 1)]
            } else {
                v[index(k - 1)] + 1
            };
            let mut y = x - k;
            while x < n && y < m && a[x as usize] == b[y as usize] {
                x += 1;
                y += 1;
            }
            v[index(k)] = x;
            if x < n || y < m {
                continue;
            }
            let mut snakes = Vec::new();
            let (mut x, mut y) = (n, m);
            for d in (1..=d).rev() {
                let v = &trace[d as usize];
                let k = x - y;
                let down = k == -d || (k != d && v[index(k - 1)] < v[index(k + 1)]);
                let previous = if down { k + 1 } else { k - 1 };
                let (px, py) = (v[index(previous)], v[index(previous)] - previous);
                let (mx, my) = if down { (px, py + 1) } else { (px + 1, py) };
                if x > mx {
                    snakes.push((mx as usize, my as usize, (x - mx) as usize));
                }
                (x, y) = (px, py);
            }
            if x > 0 {
                snakes.push((0, 0, x as usize));
            }
            snakes.reverse();
            return Some(snakes);
        }
    }
    None
}

/// The written text offsets a marker at `at` in the source text may take, and whether it moves
/// left, toward lower offsets, when no gap lies there. A marker in kept text keeps its place,
/// moving toward the kept text beside it; one at either end of a deletion stays with the kept
/// text on its other side; one in replaced text, or where text was inserted, spans the new
/// text and moves to its start when `start`, else its end.
fn mapped(edits: &[(Range<usize>, Range<usize>)], at: usize, start: bool) -> (usize, usize, bool) {
    let next = edits.partition_point(|(old, _)| old.end < at);
    match edits.get(next) {
        Some((old, new)) if old.start <= at => {
            if old.is_empty() || (old.start < at && at < old.end) {
                (new.start, new.end, start)
            } else if at == old.start {
                (new.start, new.start, true)
            } else {
                (new.end, new.end, false)
            }
        }
        _ => {
            let offset = match next.checked_sub(1).map(|before| &edits[before]) {
                Some((old, new)) => at - old.end + new.end,
                None => at,
            };
            (offset, offset, next == 0 || next < edits.len())
        }
    }
}

/// The XML of each of `written`, the paragraphs a save writes for `sources`, the source
/// paragraphs they take, with the permission markers of `sources` they lack by kind and ID, or
/// `None` for one that takes none. Both are read as one text. Each marker goes between written
/// children where the text it stood at went (see [`mapped`]), after as many of the empty
/// children there as preceded it in the source, and never before a marker that precedes it in
/// the source, copied byte for byte.
fn carry(sources: &[&str], written: &[&str]) -> Vec<Option<String>> {
    let mut carried = vec![None; written.len()];
    let mut text = Vec::new();
    let mut markers = Vec::new();
    for source in sources {
        let Some(reading) = read(source) else {
            continue;
        };
        let shift = text.len();
        markers.extend(reading.markers.into_iter().map(|marker| Marker {
            offset: marker.offset + shift,
            ..marker
        }));
        text.extend(reading.text);
        text.push(SEPARATOR);
    }
    if markers.is_empty() {
        return carried;
    }
    let expanded: Vec<Cow<str>> = written
        .iter()
        .map(|xml| match xml.strip_suffix("/>") {
            Some(open) if xml.matches('<').count() == 1 => Cow::Owned(format!("{open}></w:p>")),
            _ => Cow::Borrowed(*xml),
        })
        .collect();
    let mut new_text = Vec::new();
    let mut gaps = Vec::new();
    let mut present: HashMap<(bool, String), usize> = HashMap::new();
    for (paragraph, xml) in expanded.iter().enumerate() {
        let Some(reading) = read(xml) else {
            continue;
        };
        let shift = new_text.len();
        for marker in reading.markers {
            *present.entry((marker.start, marker.id)).or_default() += 1;
        }
        gaps.extend(
            reading
                .gaps
                .into_iter()
                .map(|(byte, offset)| (paragraph, byte, offset + shift)),
        );
        new_text.extend(reading.text);
        new_text.push(SEPARATOR);
    }
    if gaps.is_empty() {
        return carried;
    }
    let edits = edits(&text, &new_text);
    let mut placed: Vec<Vec<(usize, &str)>> = vec![Vec::new(); written.len()];
    let mut floor = 0;
    for marker in &markers {
        if let Some(count) = present
            .get_mut(&(marker.start, marker.id.clone()))
            .filter(|count| **count > 0)
        {
            *count -= 1;
            continue;
        }
        let (low, high, left) = mapped(&edits, marker.offset, marker.start);
        let after = gaps.partition_point(|(_, _, offset)| *offset < low);
        let until = gaps.partition_point(|(_, _, offset)| *offset <= high);
        let chosen = match (after < until, left) {
            (true, true) => gaps.get(after),
            (true, false) => gaps.get(until - 1),
            (false, true) => after.checked_sub(1).and_then(|index| gaps.get(index)),
            (false, false) => gaps.get(until),
        };
        let offset = chosen.or(gaps.last()).map_or(0, |(_, _, offset)| *offset);
        let first = gaps.partition_point(|(_, _, other)| *other < offset);
        let last = gaps
            .partition_point(|(_, _, other)| *other <= offset)
            .saturating_sub(1);
        floor = (first + marker.rank).min(last).max(floor);
        let (paragraph, byte, _) = gaps[floor];
        placed[paragraph].push((byte, marker.xml));
    }
    for ((slot, xml), insertions) in carried.iter_mut().zip(&expanded).zip(placed) {
        if insertions.is_empty() {
            continue;
        }
        let mut output = String::with_capacity(xml.len() + 64);
        let mut cursor = 0;
        for (byte, marker) in insertions {
            output.push_str(&xml[cursor..byte]);
            output.push_str(marker);
            cursor = byte;
        }
        output.push_str(&xml[cursor..]);
        *slot = Some(output);
    }
    carried
}

#[cfg(test)]
mod tests {
    use super::{carry, edits};

    const START: &str = r#"<w:permStart w:id="1" w:edGrp="everyone"/>"#;
    const END: &str = r#"<w:permEnd w:id="1"/>"#;

    fn run(text: &str) -> String {
        format!("<w:r><w:t>{text}</w:t></w:r>")
    }

    fn paragraph(content: &str) -> String {
        format!("<w:p>{content}</w:p>")
    }

    fn carried(source: &str, written: &str) -> String {
        carry(&[source], &[written])
            .pop()
            .flatten()
            .unwrap_or_else(|| written.to_owned())
    }

    #[test]
    fn text_typed_into_a_run_beside_a_marker_stays_on_that_runs_side() {
        let source = paragraph(&format!("{}{START}{}{END}", run("Hi"), run("X")));
        let joined = format!("{}{}", run("Hi there"), run("X"));
        assert_eq!(
            carried(&source, &paragraph(&joined)),
            paragraph(&format!("{}{START}{}{END}", run("Hi there"), run("X")))
        );
        let extended = format!("{}{}", run("Hi"), run("X and Y"));
        assert_eq!(
            carried(&source, &paragraph(&extended)),
            paragraph(&format!("{}{START}{}{END}", run("Hi"), run("X and Y")))
        );
    }

    #[test]
    fn text_typed_at_a_marker_in_a_run_of_its_own_goes_inside_the_range() {
        let source = paragraph(&format!("{}{START}{}{END}", run("Hi"), run("X")));
        let bold = "<w:r><w:rPr><w:b/></w:rPr><w:t>new</w:t></w:r>";
        assert_eq!(
            carried(
                &source,
                &paragraph(&format!("{}{bold}{}", run("Hi"), run("X")))
            ),
            paragraph(&format!("{}{START}{bold}{}{END}", run("Hi"), run("X")))
        );
    }

    #[test]
    fn a_marker_a_written_run_spans_moves_toward_its_kept_text() {
        let source = paragraph(&format!("{}{START}{}{END}", run("A"), run("B")));
        assert_eq!(
            carried(&source, &paragraph(&run("AB"))),
            paragraph(&format!("{START}{}{END}", run("AB")))
        );
    }

    #[test]
    fn markers_keep_their_place_among_empty_siblings() {
        let source = paragraph(&format!(
            r#"<w:pPr><w:jc w:val="center"/></w:pPr><w:bookmarkStart w:id="0" w:name="a"/>{START}{}{END}<w:bookmarkEnd w:id="0"/>"#,
            run("A")
        ));
        let written = paragraph(&format!(
            r#"<w:pPr><w:jc w:val="left"/></w:pPr><w:bookmarkStart w:id="0" w:name="a"/>{}<w:bookmarkEnd w:id="0"/>"#,
            run("A")
        ));
        assert_eq!(carried(&source, &written), source.replace("center", "left"));
    }

    #[test]
    fn an_empty_written_paragraph_takes_the_markers_of_its_deleted_text() {
        let source = paragraph(&format!("{START}{}{END}", run("A")));
        assert_eq!(
            carried(&source, "<w:p/>"),
            paragraph(&format!("{START}{END}"))
        );
    }

    #[test]
    fn a_marker_nested_in_a_child_the_written_xml_drops_moves_beside_it() {
        let link = |content: &str| format!(r#"<w:hyperlink w:anchor="a">{content}</w:hyperlink>"#);
        let source = paragraph(&link(&format!("{START}{}{END}", run("A"))));
        assert_eq!(
            carried(&source, &paragraph(&link(&run("A")))),
            paragraph(&format!("{START}{}{END}", link(&run("A"))))
        );
        let respelled = source.replace('"', "'");
        assert_eq!(carry(&[source.as_str()], &[respelled.as_str()]), [None]);
    }

    #[test]
    fn markers_of_merged_paragraphs_land_where_their_text_went() {
        let sources = [
            paragraph(&run("One")),
            paragraph(&run("Two")),
            paragraph(&format!("{START}{}{END}", run("Three"))),
        ];
        let sources: Vec<&str> = sources.iter().map(String::as_str).collect();
        let written = paragraph(&format!("{}{}{}", run("One"), run("Two"), run("Three")));
        assert_eq!(
            carry(&sources, &[written.as_str()]),
            [Some(paragraph(&format!(
                "{}{}{START}{}{END}",
                run("One"),
                run("Two"),
                run("Three")
            )))]
        );
    }

    #[test]
    fn a_split_paragraph_gives_each_marker_to_the_paragraph_holding_its_text() {
        let source = paragraph(&format!("{START}{}{END}", run("AB")));
        let (first, second) = (paragraph(&run("A")), paragraph(&run("B")));
        assert_eq!(
            carry(&[source.as_str()], &[first.as_str(), second.as_str()]),
            [
                Some(paragraph(&format!("{START}{}", run("A")))),
                Some(paragraph(&format!("{}{END}", run("B"))))
            ]
        );
    }

    #[test]
    fn edits_keep_unchanged_text_between_separate_changes() {
        let chars = |text: &str| text.chars().collect::<Vec<_>>();
        assert_eq!(
            edits(&chars("AMZ"), &chars("BMY")),
            [(0..1, 0..1), (2..3, 2..3)]
        );
        assert!(edits(&chars("abc"), &chars("abc")).is_empty());
        assert_eq!(edits(&chars("ab"), &chars("axb")), [(1..1, 1..2)]);
    }
}
