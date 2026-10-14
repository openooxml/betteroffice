//! Reads and patches captured content-control property XML (`w:sdtPr`) as the session keeps it.

use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::inline::{SdtProperties, parse_sdt_properties};
use crate::serializer::raw::CONTENT_FRAGMENT_PREFIX;
use crate::xml::{ParseBudget, ParseError, ParseLimits, parse_xml_strict};

const W_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const FRAGMENT_SUFFIX: &str = "</s11:root>";

fn fragment(xml: &str) -> String {
    format!("{CONTENT_FRAGMENT_PREFIX}{xml}{FRAGMENT_SUFFIX}")
}

/// Parses one captured `w:sdtPr` element the way the parser reads it in place.
pub fn parse_sdt_properties_xml(xml: &str) -> Result<SdtProperties, ParseError> {
    let source = fragment(xml);
    let limits = ParseLimits::default();
    let document = parse_xml_strict(
        source.as_bytes(),
        "sdt-properties.xml",
        &mut ParseBudget::new(&limits),
    )?;
    let element = document
        .root()
        .and_then(|root| root.child_elements().next())
        .filter(|element| element.local_name() == "sdtPr")
        .ok_or_else(|| {
            ParseError::Canonical("captured content-control properties are not w:sdtPr".to_owned())
        })?;
    Ok(parse_sdt_properties(Some(element), None, None))
}

/// Removes every WordprocessingML `showingPlcHdr` child of a captured `w:sdtPr`, leaving all
/// other bytes as they were. A child whose prefix the fragment leaves unbound counts as
/// WordprocessingML when it shares the prefix of the `sdtPr` it was captured under. `None` when
/// the properties show no placeholder.
pub fn clear_showing_placeholder_xml(xml: &str) -> Result<Option<String>, ParseError> {
    let source = fragment(xml);
    let offset = CONTENT_FRAGMENT_PREFIX.len();
    let mut reader = NsReader::from_str(&source);
    let malformed = |error: quick_xml::Error| ParseError::Canonical(error.to_string());
    let mut depth = 0usize;
    let mut open: Option<usize> = None;
    let mut root: Option<Vec<u8>> = None;
    let mut cuts: Vec<(usize, usize)> = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event().map_err(malformed)?;
        let in_w = match &namespace {
            ResolveResult::Bound(bound) => bound.as_ref() == W_NAMESPACE,
            ResolveResult::Unknown(prefix) => root.as_deref() == Some(prefix.as_slice()),
            ResolveResult::Unbound => false,
        };
        let unbound = match namespace {
            ResolveResult::Unknown(prefix) => Some(prefix),
            _ => None,
        };
        let placeholder =
            |depth: usize, local: &[u8]| in_w && depth == 3 && local == b"showingPlcHdr";
        match event {
            Event::Start(start) => {
                depth += 1;
                if depth == 2 {
                    root = unbound;
                } else if open.is_none() && placeholder(depth, start.local_name().as_ref()) {
                    open = Some(before);
                }
            }
            Event::Empty(empty) => {
                if open.is_none() && placeholder(depth + 1, empty.local_name().as_ref()) {
                    cuts.push((before, reader.buffer_position() as usize));
                }
            }
            Event::End(_) => {
                if depth == 3
                    && let Some(start) = open.take()
                {
                    cuts.push((start, reader.buffer_position() as usize));
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if cuts.is_empty() {
        return Ok(None);
    }
    let mut patched = String::with_capacity(xml.len());
    let mut cursor = offset;
    for (start, end) in cuts {
        patched.push_str(&source[cursor..start]);
        cursor = end;
    }
    patched.push_str(&source[cursor..source.len() - FRAGMENT_SUFFIX.len()]);
    Ok(Some(patched))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn multi_line_reads_the_attribute_and_defaults_to_false() {
        let single =
            parse_sdt_properties_xml("<w:sdtPr><w:tag w:val=\"a\"/><w:text/></w:sdtPr>").unwrap();
        assert_eq!(single.sdt_type, "plainText");
        assert_eq!(single.multi_line, Some(false));
        let multi =
            parse_sdt_properties_xml("<w:sdtPr><w:text w:multiLine=\"true\"/></w:sdtPr>").unwrap();
        assert_eq!(multi.multi_line, Some(true));
        let rich = parse_sdt_properties_xml("<w:sdtPr><w:richText/></w:sdtPr>").unwrap();
        assert_eq!(rich.multi_line, None);
    }

    #[test]
    fn placeholder_flags_are_cut_out_of_untouched_bytes() {
        let raw = "<w:sdtPr><w:alias w:val=\"A &amp; B\"/><w:showingPlcHdr/><w:tag  w:val=\"t\"/><w:showingPlcHdr w:val=\"1\"></w:showingPlcHdr><w15:appearance w15:val=\"hidden\"/></w:sdtPr>";
        assert_eq!(
            clear_showing_placeholder_xml(raw).unwrap().as_deref(),
            Some(
                "<w:sdtPr><w:alias w:val=\"A &amp; B\"/><w:tag  w:val=\"t\"/><w15:appearance w15:val=\"hidden\"/></w:sdtPr>"
            )
        );
        assert_eq!(
            clear_showing_placeholder_xml("<w:sdtPr><w:tag w:val=\"t\"/></w:sdtPr>").unwrap(),
            None
        );
    }

    #[test]
    fn placeholder_matching_is_namespace_aware() {
        let foreign = "<w:sdtPr><x:showingPlcHdr xmlns:x=\"urn:other\"/></w:sdtPr>";
        assert_eq!(clear_showing_placeholder_xml(foreign).unwrap(), None);
        let aliased = "<q:sdtPr xmlns:q=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><q:showingPlcHdr/><q:tag q:val=\"t\"/></q:sdtPr>";
        assert_eq!(
            clear_showing_placeholder_xml(aliased).unwrap().as_deref(),
            Some(
                "<q:sdtPr xmlns:q=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><q:tag q:val=\"t\"/></q:sdtPr>"
            )
        );
        let nested = "<w:sdtPr><w:placeholder><w:showingPlcHdr/></w:placeholder></w:sdtPr>";
        assert_eq!(clear_showing_placeholder_xml(nested).unwrap(), None);
        let captured =
            "<x:sdtPr><x:tag x:val=\"t\"/><x:showingPlcHdr/><y:showingPlcHdr/></x:sdtPr>";
        assert_eq!(
            clear_showing_placeholder_xml(captured).unwrap().as_deref(),
            Some("<x:sdtPr><x:tag x:val=\"t\"/><y:showingPlcHdr/></x:sdtPr>")
        );
    }

    #[test]
    fn synthesized_plain_text_keeps_multi_line() {
        let mut properties = parse_sdt_properties_xml(
            "<w:sdtPr><w:tag w:val=\"a\"/><w:text w:multiLine=\"1\"/></w:sdtPr>",
        )
        .unwrap();
        properties.raw_properties_xml = None;
        let synthesized = crate::serializer::paragraph::synthesize_sdt_properties(&properties);
        assert!(
            synthesized.contains("<w:text w:multiLine=\"1\"/>"),
            "{synthesized}"
        );
        assert_eq!(
            parse_sdt_properties_xml(&synthesized).unwrap().multi_line,
            Some(true)
        );
    }
}
