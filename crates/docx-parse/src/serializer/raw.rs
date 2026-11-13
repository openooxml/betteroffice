//! Validation for the narrowly allowed lossless XML replay fields.

use crate::xml::{ParseBudget, ParseError, ParseLimits, XmlElement, parse_xml_strict};

/// Complete namespace context used by WordprocessingML content fragments.
/// The wrapper is never emitted; it only makes captured prefixed subtrees
/// independently parseable under the safe XML budget.
pub(crate) const CONTENT_FRAGMENT_PREFIX: &str = concat!(
    "<s11:root xmlns:s11=\"urn:openooxml:serializer-fragment\" ",
    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
    "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ",
    "xmlns:v=\"urn:schemas-microsoft-com:vml\" ",
    "xmlns:o=\"urn:schemas-microsoft-com:office:office\" ",
    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" ",
    "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" ",
    "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ",
    "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" ",
    "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" ",
    "xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" ",
    "xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" ",
    "xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" ",
    "xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" ",
    "xmlns:w16du=\"http://schemas.microsoft.com/office/word/2023/wordml/word16du\" ",
    "xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" ",
    "xmlns:w16sdtfl=\"http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock\" ",
    "xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" ",
    "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" ",
    "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" ",
    "xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" ",
    "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
);

/// A fragment wrapper binding exactly the story-root prefixes.
const STORY_ROOT_FRAGMENT_PREFIX: &str = concat!(
    "<s11:root xmlns:s11=\"urn:openooxml:serializer-fragment\" ",
    "xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" ",
    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ",
    "xmlns:o=\"urn:schemas-microsoft-com:office:office\" ",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
    "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ",
    "xmlns:v=\"urn:schemas-microsoft-com:vml\" ",
    "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" ",
    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" ",
    "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" ",
    "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ",
    "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" ",
    "xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" ",
    "xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" ",
    "xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" ",
    "xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" ",
    "xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" ",
    "xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" ",
    "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" ",
    "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" ",
    "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
);

/// A replayed foreign fragment must still be exactly one well-formed
/// element; the resolver-free reader accepts its authored prefixes as-is.
pub(crate) fn validate_replayed_fragment(xml: &str) -> Result<(), ParseError> {
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let document = parse_xml_strict(xml.as_bytes(), "replayed-fragment", &mut budget)?;
    let root = document.root().ok_or_else(|| {
        ParseError::Canonical("replayed fragment must contain exactly one element".to_owned())
    })?;
    if matches!(root.name.as_str(), "w:bookmarkStart" | "w:bookmarkEnd") {
        for name in root.attributes.keys() {
            if let Some((prefix, local)) = name.split_once(':')
                && !matches!(prefix, "w" | "xml" | "xmlns")
                && matches!(local, "id" | "name")
                && !root.attributes.contains_key(&format!("xmlns:{prefix}"))
            {
                return Err(ParseError::Canonical(format!(
                    "unbound block bookmark attribute {name}"
                )));
            }
        }
    }
    Ok(())
}

pub(crate) fn validate_raw_subtree(
    xml: &str,
    expected_prefix: &str,
    expected_local_name: &str,
) -> Result<(), ParseError> {
    with_raw_subtree(xml, CONTENT_FRAGMENT_PREFIX, |_, child| {
        require_name(child, expected_prefix, expected_local_name)
    })
}

/// Validates a kept `w:ffData` whose prefixes are each declared in it or by every story root.
pub(crate) fn validate_form_field_data(xml: &str) -> Result<(), ParseError> {
    with_raw_subtree(xml, STORY_ROOT_FRAGMENT_PREFIX, |root, child| {
        require_name(child, "w", "ffData")?;
        let mut scope: Vec<&str> = root
            .attributes
            .keys()
            .filter_map(|name| name.strip_prefix("xmlns:"))
            .filter(|prefix| super::parts::is_story_root_prefix(prefix))
            .collect();
        require_bound_prefixes(child, &mut scope)
    })
}

fn require_bound_prefixes<'a>(
    element: &'a XmlElement,
    scope: &mut Vec<&'a str>,
) -> Result<(), ParseError> {
    let outer = scope.len();
    scope.extend(
        element
            .attributes
            .keys()
            .filter_map(|name| name.strip_prefix("xmlns:")),
    );
    let names = std::iter::once(&element.name).chain(element.attributes.keys());
    for (prefix, _) in names.filter_map(|name| name.split_once(':')) {
        if !matches!(prefix, "xml" | "xmlns") && !scope.contains(&prefix) {
            return Err(ParseError::Canonical(format!("unbound prefix {prefix}")));
        }
    }
    for child in element.child_elements() {
        require_bound_prefixes(child, scope)?;
    }
    scope.truncate(outer);
    Ok(())
}

fn with_raw_subtree(
    xml: &str,
    wrapper: &str,
    check: impl for<'e> FnOnce(&'e XmlElement, &'e XmlElement) -> Result<(), ParseError>,
) -> Result<(), ParseError> {
    let source = format!("{wrapper}{xml}</s11:root>");
    let limits = ParseLimits::default();
    let document = parse_xml_strict(
        source.as_bytes(),
        "serializer-validated-raw.xml",
        &mut ParseBudget::new(&limits),
    )?;
    let root = document.root().ok_or_else(|| {
        ParseError::Canonical("validated XML fragment has no wrapper root".to_owned())
    })?;
    let mut children = root.child_elements();
    let child = children
        .next()
        .ok_or_else(|| ParseError::Canonical("validated XML fragment has no subtree".to_owned()))?;
    if children.next().is_some() || root.children.len() != 1 {
        return Err(ParseError::Canonical(
            "validated XML fragment must contain exactly one subtree".to_owned(),
        ));
    }
    check(root, child)
}

pub(crate) fn validate_math_subtree(xml: &str) -> Result<(), ParseError> {
    let source = format!("{CONTENT_FRAGMENT_PREFIX}{xml}</s11:root>");
    let limits = ParseLimits::default();
    let document = parse_xml_strict(
        source.as_bytes(),
        "serializer-validated-math.xml",
        &mut ParseBudget::new(&limits),
    )?;
    let root = document.root().ok_or_else(|| {
        ParseError::Canonical("validated math fragment has no wrapper root".to_owned())
    })?;
    let mut children = root.child_elements();
    let child = children.next().ok_or_else(|| {
        ParseError::Canonical("validated math fragment has no subtree".to_owned())
    })?;
    if children.next().is_some() || root.children.len() != 1 {
        return Err(ParseError::Canonical(
            "validated math fragment must contain exactly one subtree".to_owned(),
        ));
    }
    if !child.name.starts_with("m:") || !matches!(child.local_name(), "oMath" | "oMathPara") {
        return Err(ParseError::Canonical(format!(
            "validated math fragment expected m:oMath or m:oMathPara, found {}",
            child.name
        )));
    }
    Ok(())
}

fn require_name(
    element: &XmlElement,
    expected_prefix: &str,
    expected_local_name: &str,
) -> Result<(), ParseError> {
    let expected = format!("{expected_prefix}:{expected_local_name}");
    if element.name != expected {
        return Err(ParseError::Canonical(format!(
            "validated XML fragment expected {expected}, found {}",
            element.name
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_one_expected_subtree_and_rejects_siblings_or_doctype() {
        validate_raw_subtree("<w:sdtPr><w:tag w:val=\"safe\"/></w:sdtPr>", "w", "sdtPr").unwrap();
        assert!(validate_raw_subtree("<w:sdtPr/><w:sdtPr/>", "w", "sdtPr").is_err());
        assert!(
            validate_raw_subtree(
                "<!DOCTYPE x [<!ENTITY p SYSTEM 'file:///etc/passwd'>]><w:sdtPr/>",
                "w",
                "sdtPr"
            )
            .is_err()
        );
    }

    #[test]
    fn form_field_data_requires_each_prefix_bound_in_the_fragment_or_by_every_story_root() {
        validate_form_field_data(
            r#"<w:ffData xmlns:ext="urn:x"><w:name w:val="T"/><ext:foo w14:x="1" xml:space="preserve"/></w:ffData>"#,
        )
        .unwrap();
        for (xml, prefix) in [
            (r#"<w:ffData><ext:foo/></w:ffData>"#, "ext"),
            (r#"<w:ffData><w:name ext:val="T"/></w:ffData>"#, "ext"),
            (
                r#"<w:ffData><w:name><ext:foo xmlns:ext="urn:x"/></w:name><ext:bar/></w:ffData>"#,
                "ext",
            ),
            (r#"<w:ffData><w16du:foo/></w:ffData>"#, "w16du"),
        ] {
            let error = validate_form_field_data(xml).unwrap_err().to_string();
            assert!(
                error.contains(&format!("unbound prefix {prefix}")),
                "{xml}: {error}"
            );
        }
    }
}

#[cfg(test)]
mod replay_tests {
    use super::*;

    #[test]
    fn replayed_bookmark_identities_require_bound_aliases() {
        for xml in [
            r#"<w:bookmarkStart b:id="7" w:name="Range"/>"#,
            r#"<w:bookmarkEnd b:id="7"/>"#,
            r#"<w:bookmarkStart w:id="7" b:name="Range"/>"#,
        ] {
            assert!(validate_replayed_fragment(xml).is_err(), "{xml}");
        }
        validate_replayed_fragment(
            r#"<w:bookmarkStart xmlns:b="http://schemas.openxmlformats.org/wordprocessingml/2006/main" b:id="7" b:name="Range"/>"#,
        ).unwrap();
    }

    #[test]
    fn replay_validation_rejects_unescaped_emitted_values() {
        for xml in [
            r#"<x:block x:q="a&b"/>"#,
            r#"<x:block x:q="a<b"/>"#,
            "<x:block>a&b</x:block>",
            "<x:block/><x:block/>",
        ] {
            assert!(validate_replayed_fragment(xml).is_err(), "{xml}");
        }
        validate_replayed_fragment(
            r#"<x:block x:q="a&amp;&lt;&gt;&quot;b">a&amp;&lt;&gt;b</x:block>"#,
        )
        .unwrap();
    }
}
