//! Permission range markers, which the model does not keep, carried from a paragraph's source
//! XML into the XML written for it.

use std::borrow::Cow;
use std::collections::BTreeMap;
use std::ops::Range;

use crate::paragraph_identity::{paragraph_occurrences, tags, unescaped};

use super::paragraph_ids::S13SplicedPart;
use super::splice::{POSITIONED, paragraph_span};

const MARKERS: [&str; 2] = ["w:permStart", "w:permEnd"];

/// The source XML of a paragraph that holds a permission marker.
#[derive(Clone, Debug)]
pub(crate) struct Source {
    pub(crate) xml: String,
    pub(crate) changed: bool,
}

/// The paragraphs of `source` that `part` addresses and that hold a permission marker.
pub(crate) fn sources(source: &str, part: &S13SplicedPart) -> BTreeMap<u32, Source> {
    let mut found = BTreeMap::new();
    if !MARKERS.iter().any(|marker| source.contains(marker)) {
        return found;
    }
    let Some(occurrences) = paragraph_occurrences(source) else {
        return found;
    };
    for &ordinal in &part.paragraphs {
        let Some(span) = occurrences
            .get(ordinal as usize)
            .filter(|occurrence| occurrence.ordinal == ordinal)
            .and_then(|occurrence| paragraph_span(source, occurrence))
        else {
            continue;
        };
        let xml = &source[span];
        if MARKERS.iter().any(|marker| xml.contains(marker)) {
            found.insert(
                ordinal,
                Source {
                    xml: xml.to_owned(),
                    changed: part.changed.contains(&ordinal),
                },
            );
        }
    }
    found
}

/// A child element of a paragraph: its bytes and the text it holds, as offsets into the
/// paragraph's text.
struct Child<'x> {
    name: &'x str,
    span: Range<usize>,
    text: Range<usize>,
}

/// A paragraph's text, positioned run content each as one character, and child elements.
#[derive(Default)]
struct Shape<'x> {
    text: Vec<char>,
    children: Vec<Child<'x>>,
    /// Where the paragraph's content starts: after its start tag and properties.
    content: usize,
    /// Permission markers below the paragraph's children.
    nested: Vec<Range<usize>>,
}

fn shape(xml: &str) -> Option<Shape<'_>> {
    let mut shape = Shape::default();
    let mut depth = 0usize;
    let mut in_text = 0usize;
    let mut in_properties = 0usize;
    let mut previous = 0;
    let mut open = None;
    for tag in tags(xml)? {
        if in_text > 0 {
            shape
                .text
                .extend(unescaped(xml, previous..tag.range.start)?.chars());
        }
        previous = tag.range.end;
        let text = matches!(tag.name, "w:t" | "w:delText");
        let properties = tag.name.ends_with("Pr");
        if tag.end {
            depth = depth.checked_sub(1)?;
            if text {
                in_text = in_text.saturating_sub(1);
            } else if properties {
                in_properties = in_properties.saturating_sub(1);
            }
            if depth == 1 {
                let (name, start, offset) = open.take()?;
                shape.children.push(Child {
                    name,
                    span: start..tag.range.end,
                    text: offset..shape.text.len(),
                });
            }
            continue;
        }
        if depth == 0 {
            if shape.content != 0 || tag.name != "w:p" || tag.empty {
                return None;
            }
            depth = 1;
            shape.content = tag.range.end;
            continue;
        }
        if depth == 1 {
            open = Some((tag.name, tag.range.start, shape.text.len()));
        } else if MARKERS.contains(&tag.name) {
            shape.nested.push(tag.range.clone());
        }
        if properties {
            in_properties += usize::from(!tag.empty);
        } else if text {
            in_text += usize::from(!tag.empty);
        } else if in_properties == 0
            && let Some(index) = POSITIONED.iter().position(|name| *name == tag.name)
        {
            shape.text.push(char::from_u32(0xE000 + index as u32)?);
        }
        if !tag.empty {
            depth += 1;
        } else if depth == 1 {
            let (name, start, offset) = open.take()?;
            shape.children.push(Child {
                name,
                span: start..tag.range.end,
                text: offset..shape.text.len(),
            });
        }
    }
    if depth != 0 {
        return None;
    }
    if let Some(properties) = shape.children.first().filter(|child| child.name == "w:pPr") {
        shape.content = properties.span.end;
    }
    Some(shape)
}

/// `written`, the XML written for a paragraph, with the permission markers of `source`, its
/// source XML, copied byte for byte. Each goes between written children where the text it
/// stood at went, after as many of the empty children there as preceded it in the source;
/// text inserted right at a marker goes to the side of it the written runs keep it on. `None`
/// when a marker sits inside text the save changed, inside a written child, or on either side
/// of text inserted at it, or is nested in a child the written XML drops it from.
pub(crate) fn carry(source: &str, written: &str) -> Option<String> {
    let from = shape(source)?;
    for nested in &from.nested {
        let marker = &source[nested.clone()];
        let wanted = from
            .nested
            .iter()
            .filter(|other| &source[(*other).clone()] == marker)
            .count();
        if written.matches(marker).count() < wanted {
            return None;
        }
    }
    let markers: Vec<usize> = from
        .children
        .iter()
        .enumerate()
        .filter(|(_, child)| MARKERS.contains(&child.name))
        .map(|(index, _)| index)
        .collect();
    if markers.is_empty() {
        return Some(written.to_owned());
    }
    let written: Cow<str> = match written.strip_suffix("/>") {
        Some(open) if written.matches('<').count() == 1 => Cow::Owned(format!("{open}></w:p>")),
        _ => Cow::Borrowed(written),
    };
    let to = shape(&written)?;
    let (old, new) = (&from.text, &to.text);
    let prefix = old.iter().zip(new).take_while(|(a, b)| a == b).count();
    let suffix = old[prefix..]
        .iter()
        .rev()
        .zip(new[prefix..].iter().rev())
        .take_while(|(a, b)| a == b)
        .count();
    let (cut, kept) = (old.len() - suffix, new.len() - suffix);
    let mut gaps = vec![(to.content, 0)];
    gaps.extend(
        to.children
            .iter()
            .filter(|child| child.span.start >= to.content)
            .map(|child| (child.span.end, child.text.end)),
    );
    let mut insertions = Vec::new();
    for &index in &markers {
        let marker = &from.children[index];
        let at = marker.text.start;
        let range = if at <= prefix && at < cut {
            at..=at
        } else if at >= cut && at > prefix {
            at - cut + kept..=at - cut + kept
        } else if at == prefix {
            prefix..=kept
        } else {
            return None;
        };
        let fits: Vec<&(usize, usize)> = gaps
            .iter()
            .filter(|(_, offset)| range.contains(offset))
            .collect();
        let offset = fits.first()?.1;
        if fits.iter().any(|(_, other)| *other != offset) {
            return None;
        }
        let empty = from.children[..index]
            .iter()
            .filter(|child| {
                child.text == (at..at)
                    && child.name != "w:pPr"
                    && child.name != "w:proofErr"
                    && !MARKERS.contains(&child.name)
            })
            .count();
        let (position, _) = fits[empty.min(fits.len() - 1)];
        insertions.push((*position, &source[marker.span.clone()]));
    }
    insertions.sort_by_key(|(position, _)| *position);
    let mut output = String::with_capacity(written.len() + 64);
    let mut cursor = 0;
    for (position, marker) in insertions {
        output.push_str(&written[cursor..position]);
        output.push_str(marker);
        cursor = position;
    }
    output.push_str(&written[cursor..]);
    Some(output)
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

    #[test]
    fn text_typed_into_a_run_beside_a_marker_stays_on_that_runs_side() {
        let source = paragraph(&format!("{}{START}{}{END}", run("Hi"), run("X")));
        assert_eq!(
            carry(
                &source,
                &paragraph(&format!("{}{}", run("Hi there"), run("X")))
            ),
            Some(paragraph(&format!(
                "{}{START}{}{END}",
                run("Hi there"),
                run("X")
            )))
        );
        assert_eq!(
            carry(
                &source,
                &paragraph(&format!("{}{}", run("Hi"), run("X and Y")))
            ),
            Some(paragraph(&format!(
                "{}{START}{}{END}",
                run("Hi"),
                run("X and Y")
            )))
        );
    }

    #[test]
    fn refuses_a_marker_text_was_typed_at_in_a_run_of_its_own() {
        let source = paragraph(&format!("{}{START}{}{END}", run("Hi"), run("X")));
        let bold = "<w:r><w:rPr><w:b/></w:rPr><w:t>new</w:t></w:r>";
        assert_eq!(
            carry(
                &source,
                &paragraph(&format!("{}{bold}{}", run("Hi"), run("X")))
            ),
            None
        );
    }

    #[test]
    fn refuses_a_marker_a_written_run_spans() {
        let source = paragraph(&format!("{}{START}{}{END}", run("A"), run("B")));
        assert_eq!(carry(&source, &paragraph(&run("AB"))), None);
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
        assert_eq!(
            carry(&source, &written),
            Some(source.replace("center", "left"))
        );
    }

    #[test]
    fn an_empty_written_paragraph_takes_the_markers_of_an_empty_range() {
        let source = paragraph(&format!(
            "<w:pPr><w:jc w:val=\"center\"/></w:pPr>{START}{END}"
        ));
        assert_eq!(
            carry(&source, "<w:p/>"),
            Some(paragraph(&format!("{START}{END}")))
        );
    }

    #[test]
    fn refuses_a_marker_nested_in_a_child_the_written_xml_drops_it_from() {
        let source = paragraph(&format!(
            r#"<w:hyperlink w:anchor="a">{START}{}{END}</w:hyperlink>"#,
            run("A")
        ));
        let written = paragraph(&format!(
            r#"<w:hyperlink w:anchor="a">{}</w:hyperlink>"#,
            run("B")
        ));
        assert_eq!(carry(&source, &written), None);
        assert_eq!(carry(&source, &source), Some(source.clone()));
    }
}
