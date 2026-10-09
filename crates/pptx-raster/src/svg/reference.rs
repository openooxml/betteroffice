//! Reference targets read with `svgtypes`, the parser `usvg` itself resolves
//! them with, so a target the audit follows is byte for byte the one `usvg`
//! looks up. Each scan is one forward pass over its text.

use svgtypes::{FuncIRI, IRI, Paint};

use super::SvgRefusal;

/// The target of an `href` that `usvg` follows, as `svgtypes::IRI` reads it.
/// A value that is not a fragment of this document is refused.
pub(super) fn href(value: &str) -> Result<Option<&str>, SvgRefusal> {
    if !value.trim_start_matches(is_space).starts_with('#') {
        return Err(SvgRefusal::ExternalReference);
    }
    Ok(IRI::from_str(value).ok().map(|iri| iri.0))
}

/// The target of a `clip-path`, `mask` or `marker` value, as `usvg` reads it
/// with `svgtypes::FuncIRI` once `none` is set aside.
pub(super) fn func_iri(value: &str) -> Result<Option<&str>, SvgRefusal> {
    let target = match value {
        "none" => None,
        _ => FuncIRI::from_str(value).ok().map(|iri| iri.0),
    };
    local(value, target)
}

/// The paint server a `fill` or `stroke` value names, as `usvg` reads it with
/// `svgtypes::Paint`.
pub(super) fn paint(value: &str) -> Result<Option<&str>, SvgRefusal> {
    let target = match Paint::from_str(value) {
        Ok(Paint::FuncIRI(target, _)) => Some(target),
        _ => None,
    };
    local(value, target)
}

/// Refuses a value that mentions `url(` but does not parse as a same-document
/// reference: `usvg` would follow nothing, and nothing outside is fetched,
/// but such a document is declined rather than drawn half-resolved.
fn local<'a>(value: &str, target: Option<&'a str>) -> Result<Option<&'a str>, SvgRefusal> {
    if target.is_none() && contains_ignore_case(value, b"url(") {
        return Err(SvgRefusal::ExternalReference);
    }
    Ok(target)
}

/// Every target a stylesheet block or `style` attribute can name. Each
/// `url(` must read `url(#target)` with no space, quote or bracket inside, so
/// it cannot open inside another declaration's string or function, and is
/// then parsed by `svgtypes::FuncIRI` as `usvg` parses the declaration.
pub(super) fn css<'a>(text: &'a str, found: &mut Vec<&'a str>) -> Result<(), SvgRefusal> {
    let bytes = text.as_bytes();
    let mut at = 0;
    while let Some(offset) = find_ignore_case(&bytes[at..], b"url(") {
        let start = at + offset;
        let body = start + 4;
        if bytes.get(body) != Some(&b'#') {
            return Err(SvgRefusal::ExternalReference);
        }
        let end = bytes[body..]
            .iter()
            .position(|byte| matches!(byte, b')' | b'(' | b'\'' | b'"' | b' '))
            .map(|length| body + length)
            .filter(|end| bytes[*end] == b')')
            .ok_or(SvgRefusal::UnsupportedStyle)?;
        let target = FuncIRI::from_str(&text[start..=end])
            .map_err(|_| SvgRefusal::UnsupportedStyle)?
            .0;
        found.push(target);
        at = end + 1;
    }
    Ok(())
}

/// `svgtypes`' whitespace.
fn is_space(c: char) -> bool {
    matches!(c, ' ' | '\t' | '\n' | '\r')
}

/// Whether `haystack` contains `needle`, ASCII case folded.
pub(super) fn contains_ignore_case(haystack: &str, needle: &[u8]) -> bool {
    find_ignore_case(haystack.as_bytes(), needle).is_some()
}

fn find_ignore_case(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .position(|window| window.eq_ignore_ascii_case(needle))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn targets_are_read_exactly_as_svgtypes_reads_them() {
        assert_eq!(func_iri("url(#a\t)"), Ok(Some("a\t")));
        assert_eq!(func_iri(" url( '#a\t ' ) "), Ok(Some("a")));
        assert_eq!(func_iri("url(#a b)"), Err(SvgRefusal::ExternalReference));
        assert_eq!(func_iri("none"), Ok(None));
        assert_eq!(paint(" url(#g\t) red "), Ok(Some("g\t")));
        assert_eq!(
            paint("url(#g) nonsense"),
            Err(SvgRefusal::ExternalReference)
        );
        assert_eq!(paint("#f00"), Ok(None));
        assert_eq!(href("#a\t"), Ok(Some("a\t")));
        assert_eq!(href(" #a"), Ok(Some("a")));
        assert_eq!(href("#a b"), Ok(None));
    }

    #[test]
    fn a_reference_outside_the_document_is_refused() {
        for value in ["url(https://x/p.svg#g)", "URL(#g)", "Url( data:x)"] {
            assert_eq!(paint(value), Err(SvgRefusal::ExternalReference), "{value}");
            assert_eq!(
                func_iri(value),
                Err(SvgRefusal::ExternalReference),
                "{value}"
            );
        }
        assert_eq!(href("sprite.svg#icon"), Err(SvgRefusal::ExternalReference));
    }

    #[test]
    fn stylesheet_targets_are_read_one_url_at_a_time() {
        let mut found = Vec::new();
        css(
            "fill:url(#a\t);stroke:url(#b) red;clip-path:url(#c)",
            &mut found,
        )
        .unwrap();
        assert_eq!(found, ["a\t", "b", "c"]);
        for (text, refusal) in [
            ("fill:url('#a')", SvgRefusal::ExternalReference),
            ("fill:uRl(https://x)", SvgRefusal::ExternalReference),
            ("fill:url(#a", SvgRefusal::UnsupportedStyle),
            ("font:'url(#a';fill:url(#b)", SvgRefusal::UnsupportedStyle),
            ("fill:url(#a b)", SvgRefusal::UnsupportedStyle),
            ("fill:URL(#a)", SvgRefusal::UnsupportedStyle),
        ] {
            assert_eq!(css(text, &mut Vec::new()), Err(refusal), "{text}");
        }
    }
}
