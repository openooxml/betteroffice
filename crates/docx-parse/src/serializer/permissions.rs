//! Permission range markers, which the model does not keep, carried from a paragraph's source
//! XML into the XML written for it.

use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashMap};

use crate::paragraph_identity::{paragraph_occurrences, tags, unescaped};

use super::paragraph_ids::S13SplicedPart;
use super::splice::{POSITIONED, paragraph_span};

const MARKERS: [&str; 2] = ["w:permStart", "w:permEnd"];

fn has_markers(xml: &str) -> bool {
    MARKERS.iter().any(|marker| xml.contains(marker))
}

/// For each paragraph `part` writes whose permission markers a save must carry, the source XML
/// of the paragraphs whose text and markers it takes, in order: its own, and those of removed
/// paragraphs that hold a marker, each taken by the nearest written paragraph before it, else
/// the nearest after it.
pub(crate) fn sources(source: &str, part: &S13SplicedPart) -> BTreeMap<u32, Vec<String>> {
    let mut found = BTreeMap::new();
    if !has_markers(source) {
        return found;
    }
    let Some(occurrences) = paragraph_occurrences(source) else {
        return found;
    };
    let xml = |ordinal: u32| {
        occurrences
            .get(ordinal as usize)
            .filter(|occurrence| occurrence.ordinal == ordinal)
            .and_then(|occurrence| paragraph_span(source, occurrence))
            .map(|span| &source[span])
    };
    let written: BTreeSet<u32> = part.paragraphs.iter().copied().collect();
    let mut taken: BTreeMap<u32, BTreeSet<u32>> = BTreeMap::new();
    for &ordinal in &written {
        if xml(ordinal).is_some_and(has_markers) {
            taken.entry(ordinal).or_default().insert(ordinal);
        }
    }
    for &ordinal in &part.removed {
        if !xml(ordinal).is_some_and(has_markers) {
            continue;
        }
        let taker = written
            .range(..ordinal)
            .next_back()
            .or_else(|| written.range(ordinal..).next());
        if let Some(&taker) = taker {
            let paragraphs = taken.entry(taker).or_default();
            paragraphs.insert(taker);
            paragraphs.insert(ordinal);
        }
    }
    for (taker, paragraphs) in taken {
        let paragraphs: Option<Vec<String>> = paragraphs
            .into_iter()
            .map(|ordinal| xml(ordinal).map(str::to_owned))
            .collect();
        if let Some(paragraphs) = paragraphs {
            found.insert(taker, paragraphs);
        }
    }
    found
}

/// A permission marker: the offset in its paragraph's text it sits at, how many empty siblings
/// precede it there, and its markup.
struct Marker<'x> {
    offset: usize,
    rank: usize,
    start: bool,
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

/// `written`, the XML written for a paragraph, with the permission markers of `sources`, the
/// source XML of the paragraphs whose text it holds (see [`sources`]), copied byte for byte
/// unless `written` already holds them. Each goes between written children where the text it
/// stood at went, after as many of the empty children there as preceded it in the source, and
/// none before a marker that precedes it in the source. A marker in text the save changed goes
/// to the edit: a start before the new text, an end after it. A marker no written gap holds
/// moves to the nearest one toward the text it stood beside that the save kept.
pub(crate) fn carry(sources: &[String], written: &str) -> String {
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
    }
    if markers.is_empty() {
        return written.to_owned();
    }
    let written: Cow<str> = match written.strip_suffix("/>") {
        Some(open) if written.matches('<').count() == 1 => Cow::Owned(format!("{open}></w:p>")),
        _ => Cow::Borrowed(written),
    };
    let Some(to) = read(&written) else {
        return written.into_owned();
    };
    let mut present: HashMap<&str, usize> = HashMap::new();
    for marker in &to.markers {
        *present.entry(marker.xml).or_default() += 1;
    }
    let prefix = text
        .iter()
        .zip(&to.text)
        .take_while(|(a, b)| a == b)
        .count();
    let suffix = text[prefix..]
        .iter()
        .rev()
        .zip(to.text[prefix..].iter().rev())
        .take_while(|(a, b)| a == b)
        .count();
    let (cut, kept) = (text.len() - suffix, to.text.len() - suffix);
    let gaps = &to.gaps;
    let mut floor = 0;
    let mut output = String::with_capacity(written.len() + 64);
    let mut cursor = 0;
    for marker in &markers {
        if let Some(count) = present.get_mut(marker.xml).filter(|count| **count > 0) {
            *count -= 1;
            continue;
        }
        let at = marker.offset;
        let (low, high, left) = if at <= prefix && at < cut {
            (at, at, true)
        } else if at >= cut && at > prefix {
            (at - cut + kept, at - cut + kept, false)
        } else {
            (prefix, kept, marker.start)
        };
        let after = gaps.partition_point(|(_, offset)| *offset < low);
        let until = gaps.partition_point(|(_, offset)| *offset <= high);
        let chosen = match (after < until, left) {
            (true, true) => gaps.get(after),
            (true, false) => gaps.get(until - 1),
            (false, true) => after.checked_sub(1).and_then(|index| gaps.get(index)),
            (false, false) => gaps.get(until),
        };
        let offset = chosen.or(gaps.last()).map_or(0, |(_, offset)| *offset);
        let first = gaps.partition_point(|(_, other)| *other < offset);
        let last = gaps
            .partition_point(|(_, other)| *other <= offset)
            .saturating_sub(1);
        floor = (first + marker.rank).min(last).max(floor);
        let position = gaps[floor].0;
        output.push_str(&written[cursor..position]);
        output.push_str(marker.xml);
        cursor = position;
    }
    output.push_str(&written[cursor..]);
    output
}

#[cfg(test)]
mod tests {
    use super::carry;

    const START: &str = r#"<w:permStart w:id="1" w:edGrp="everyone"/>"#;
    const END: &str = r#"<w:permEnd w:id="1"/>"#;

    fn run(text: &str) -> String {
        format!("<w:r><w:t>{text}</w:t></w:r>")
    }

    fn paragraph(content: &str) -> String {
        format!("<w:p>{content}</w:p>")
    }

    fn carried(source: &str, written: &str) -> String {
        carry(&[source.to_owned()], written)
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
        assert_eq!(carried(&source, &source), source);
    }

    #[test]
    fn markers_of_a_merged_paragraph_land_where_its_text_went() {
        let sources = [
            paragraph(&run("One")),
            paragraph(&format!("{}{START}{}", run("Two"), run("Three"))),
        ];
        let written = paragraph(&format!("{}{}{}", run("One"), run("Two"), run("Three")));
        assert_eq!(
            carry(&sources, &written),
            paragraph(&format!(
                "{}{}{START}{}",
                run("One"),
                run("Two"),
                run("Three")
            ))
        );
    }
}
