//! Validation for the narrowly allowed lossless XML replay fields.

use crate::xml::{ParseBudget, ParseError, ParseLimits, XmlElement, parse_xml};

/// Complete namespace context used by WordprocessingML content fragments.
/// The wrapper is never emitted; it only makes captured prefixed subtrees
/// independently parseable under the safe XML budget.
pub(super) const CONTENT_FRAGMENT_PREFIX: &str = concat!(
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

pub(crate) fn validate_raw_subtree(
    xml: &str,
    expected_prefix: &str,
    expected_local_name: &str,
) -> Result<(), ParseError> {
    let source = format!("{CONTENT_FRAGMENT_PREFIX}{xml}</s11:root>");
    let limits = ParseLimits::default();
    let document = parse_xml(
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
    require_name(child, expected_prefix, expected_local_name)
}

pub(crate) fn validate_math_subtree(xml: &str) -> Result<(), ParseError> {
    let source = format!("{CONTENT_FRAGMENT_PREFIX}{xml}</s11:root>");
    let limits = ParseLimits::default();
    let document = parse_xml(
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
}
