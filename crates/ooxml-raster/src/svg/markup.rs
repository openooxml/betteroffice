//! One forward pass over a document's bytes that bounds what `roxmltree` will
//! build before it builds anything. Reaching nesting through a parsed tree is
//! already too late, and `roxmltree` compares each attribute with every other
//! of its element and looks every name up across the namespaces in scope.

use std::collections::HashSet;

use super::{
    MAX_SVG_ATTRIBUTES, MAX_SVG_DEPTH, MAX_SVG_ELEMENT_ATTRIBUTES, MAX_SVG_NAMESPACES,
    MAX_SVG_NODES, SvgRefusal,
};

/// Refuses markup nested past [`MAX_SVG_DEPTH`], with more than
/// [`MAX_SVG_ELEMENT_ATTRIBUTES`] on one element, or with more than
/// [`MAX_SVG_NAMESPACES`] namespace declarations in scope at one element or
/// distinct across the document. `roxmltree` reserves a node for every `<` and
/// an attribute for every `=` in the text before it parses, so those are held
/// to [`MAX_SVG_NODES`] and [`MAX_SVG_ATTRIBUTES`] wherever they appear.
pub(super) fn scan(bytes: &[u8]) -> Result<(), SvgRefusal> {
    let count = |needle: u8| bytes.iter().filter(|byte| **byte == needle).count();
    if count(b'<') > MAX_SVG_NODES as usize || count(b'=') > MAX_SVG_ATTRIBUTES {
        return Err(SvgRefusal::DocumentTooLarge);
    }
    let mut index = 0;
    let mut scopes: Vec<usize> = Vec::new();
    let mut namespaces: HashSet<(&[u8], &[u8])> = HashSet::new();
    while let Some(open) = bytes[index..].iter().position(|byte| *byte == b'<') {
        index += open + 1;
        let rest = &bytes[index..];
        if rest.starts_with(b"!--") {
            index = after(bytes, index, b"-->")?;
        } else if rest.starts_with(b"![CDATA[") {
            index = after(bytes, index, b"]]>")?;
        } else if rest.starts_with(b"!") {
            return Err(SvgRefusal::DoctypeDeclared);
        } else if rest.starts_with(b"?") {
            index = after(bytes, index, b"?>")?;
        } else if rest.starts_with(b"/") {
            index = after(bytes, index, b">")?;
            scopes.pop();
        } else {
            let tag = Tag::read(bytes, index, &mut namespaces)?;
            index = tag.end;
            if tag.attributes > MAX_SVG_ELEMENT_ATTRIBUTES {
                return Err(SvgRefusal::DocumentTooLarge);
            }
            let scope = scopes.last().copied().unwrap_or(0) + tag.declarations;
            if scope > MAX_SVG_NAMESPACES || namespaces.len() > MAX_SVG_NAMESPACES {
                return Err(SvgRefusal::DocumentTooLarge);
            }
            if !tag.empty {
                scopes.push(scope);
                if scopes.len() > MAX_SVG_DEPTH {
                    return Err(SvgRefusal::TooDeeplyNested);
                }
            }
        }
    }
    Ok(())
}

/// A start tag: where it ends, whether it closed itself, and its attributes.
struct Tag {
    end: usize,
    empty: bool,
    attributes: usize,
    declarations: usize,
}

impl Tag {
    /// Reads the tag whose name starts at `from`, collecting each namespace it
    /// declares. Markup `roxmltree` would reject is refused here.
    fn read<'a>(
        bytes: &'a [u8],
        from: usize,
        namespaces: &mut HashSet<(&'a [u8], &'a [u8])>,
    ) -> Result<Self, SvgRefusal> {
        let byte = |at: usize| bytes.get(at).copied().ok_or(SvgRefusal::Unparsable);
        let mut tag = Tag {
            end: from,
            empty: false,
            attributes: 0,
            declarations: 0,
        };
        let mut at = from;
        while !matches!(byte(at)?, b'/' | b'>') && !is_space(byte(at)?) {
            at += 1;
        }
        loop {
            while is_space(byte(at)?) {
                at += 1;
            }
            match byte(at)? {
                b'>' => {
                    tag.end = at + 1;
                    return Ok(tag);
                }
                b'/' if byte(at + 1)? == b'>' => {
                    tag.end = at + 2;
                    tag.empty = true;
                    return Ok(tag);
                }
                b'/' => return Err(SvgRefusal::Unparsable),
                _ => {}
            }
            let name = at;
            while !matches!(byte(at)?, b'=' | b'/' | b'>') && !is_space(byte(at)?) {
                at += 1;
            }
            let name = &bytes[name..at];
            while is_space(byte(at)?) {
                at += 1;
            }
            if byte(at)? != b'=' {
                return Err(SvgRefusal::Unparsable);
            }
            at += 1;
            while is_space(byte(at)?) {
                at += 1;
            }
            let quote = byte(at)?;
            if !matches!(quote, b'"' | b'\'') {
                return Err(SvgRefusal::Unparsable);
            }
            let value = at + 1;
            at = after(bytes, value, &[quote])?;
            tag.attributes += 1;
            if name == b"xmlns" || name.starts_with(b"xmlns:") {
                tag.declarations += 1;
                namespaces.insert((name, &bytes[value..at - 1]));
            }
        }
    }
}

fn is_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r')
}

/// Index just past the next `needle` at or after `from`.
fn after(bytes: &[u8], from: usize, needle: &[u8]) -> Result<usize, SvgRefusal> {
    bytes[from..]
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|at| from + at + needle.len())
        .ok_or(SvgRefusal::Unparsable)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn well_formed_tags_pass_in_any_spacing() {
        let source = concat!(
            "<?xml version=\"1.0\"?><!-- <g a='1'> --><svg xmlns='http://www.w3.org/2000/svg'\n",
            "  a = \"x > y\"\tb='q\"'/>"
        );
        assert_eq!(scan(source.as_bytes()), Ok(()));
    }

    #[test]
    fn attributes_are_bounded_per_element_and_in_all() {
        let element = |count: usize| {
            let attributes: String = (0..count).map(|index| format!(" a{index}=\"\"")).collect();
            format!("<g{attributes}/>")
        };
        assert_eq!(scan(element(MAX_SVG_ELEMENT_ATTRIBUTES).as_bytes()), Ok(()));
        assert_eq!(
            scan(element(MAX_SVG_ELEMENT_ATTRIBUTES + 1).as_bytes()),
            Err(SvgRefusal::DocumentTooLarge)
        );
        let many = element(MAX_SVG_ELEMENT_ATTRIBUTES)
            .repeat(MAX_SVG_ATTRIBUTES / MAX_SVG_ELEMENT_ATTRIBUTES + 1);
        assert_eq!(scan(many.as_bytes()), Err(SvgRefusal::DocumentTooLarge));
    }

    #[test]
    fn namespaces_are_bounded_in_scope_and_across_the_document() {
        let declare = |count: usize, from: usize| -> String {
            (from..from + count)
                .map(|index| format!(" xmlns:p{index}=\"urn:{index}\""))
                .collect()
        };
        let nested = format!(
            "<g{}><g{}/></g>",
            declare(MAX_SVG_NAMESPACES / 2, 0),
            declare(MAX_SVG_NAMESPACES / 2 + 1, MAX_SVG_NAMESPACES)
        );
        assert_eq!(scan(nested.as_bytes()), Err(SvgRefusal::DocumentTooLarge));
        let spread: String = (0..=MAX_SVG_NAMESPACES)
            .map(|index| format!("<g{}/>", declare(1, index)))
            .collect();
        assert_eq!(scan(spread.as_bytes()), Err(SvgRefusal::DocumentTooLarge));
        let redeclared = "<g xmlns=\"http://www.w3.org/2000/svg\"/>".repeat(10_000);
        assert_eq!(scan(redeclared.as_bytes()), Ok(()));
    }

    #[test]
    fn what_roxmltree_reserves_for_is_bounded_wherever_it_appears() {
        let comment = format!("<svg><!--{}--></svg>", "<".repeat(MAX_SVG_NODES as usize));
        assert_eq!(scan(comment.as_bytes()), Err(SvgRefusal::DocumentTooLarge));
        let text = format!(
            "<svg><desc>{}</desc></svg>",
            "=".repeat(MAX_SVG_ATTRIBUTES + 1)
        );
        assert_eq!(scan(text.as_bytes()), Err(SvgRefusal::DocumentTooLarge));
    }

    #[test]
    fn malformed_tags_are_refused() {
        for source in ["<g a/>", "<g a=b/>", "<g a='b", "<g / >", "<g"] {
            assert_eq!(
                scan(source.as_bytes()),
                Err(SvgRefusal::Unparsable),
                "{source}"
            );
        }
    }
}
