//! Structural fingerprint: "is this the same tree?"
//!
//! Insignificant: prefix choice, attribute order, quote style, empty-element
//! spelling, inter-element whitespace outside `xml:space="preserve"`, and
//! exactly the deviations `DECLARED_NORMALIZATIONS` names. Significant:
//! element order, text, attribute values, and the non-standard namespace
//! URIs in scope. Parse-time node identity never enters the projection.

use serde_json::{Value, json};

use crate::registry::STANDARD_WML_NAMESPACE_URIS;
use crate::xml::{XML_NAMESPACE, XmlElement, XmlNode, is_xml_whitespace};

/// Canonical projection of one part's tree, serialized as JSON.
pub fn structural_fingerprint(root: &XmlElement) -> String {
    project_element(root, false, &[]).to_string()
}

/// The projection with named attributes excluded everywhere. For the
/// asymmetric normalizations the report layer owns: it re-compares with an
/// exclusion to classify a difference, never to hide a loss.
pub fn structural_fingerprint_excluding(
    root: &XmlElement,
    excluded_attributes: &[(&str, &str)],
) -> String {
    project_element(root, false, excluded_attributes).to_string()
}

/// FNV-1a 64 of the fingerprint, for compact labels in digests and diffs.
pub fn short_fingerprint(root: &XmlElement) -> String {
    let mut hash: u64 = 0xcbf2_9ce4_8422_2325;
    for byte in structural_fingerprint(root).bytes() {
        hash ^= u64::from(byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    format!("{hash:016x}")
}

fn project_element(
    element: &XmlElement,
    inherited_preserve: bool,
    excluded_attributes: &[(&str, &str)],
) -> Value {
    let preserve = match element.attribute(XML_NAMESPACE, "space") {
        Some("preserve") => true,
        Some("default") => false,
        _ => inherited_preserve,
    };
    let mut attributes: Vec<(&str, &str, &str)> = element
        .attributes
        .iter()
        .filter(|attribute| {
            !excluded_attributes.iter().any(|(namespace, local)| {
                attribute.namespace == *namespace && attribute.local == *local
            })
        })
        .map(|attribute| {
            (
                attribute.namespace.as_str(),
                attribute.local.as_str(),
                attribute.value.as_str(),
            )
        })
        .collect();
    attributes.sort_unstable();
    // Declarations compare by URI, never by prefix spelling.
    let mut bindings: Vec<&str> = element
        .bindings
        .iter()
        .map(|(_, uri)| uri.as_str())
        .filter(|uri| !STANDARD_WML_NAMESPACE_URIS.contains(uri))
        .collect();
    bindings.sort_unstable();
    bindings.dedup();
    let children: Vec<Value> = significant_children(element, preserve)
        .map(|child| match child {
            XmlNode::Element(child) => project_element(child, preserve, excluded_attributes),
            XmlNode::Text(text) => json!(["t", text]),
        })
        .collect();
    json!([
        "e",
        element.namespace,
        element.local,
        attributes
            .iter()
            .map(|(namespace, local, value)| json!([namespace, local, value]))
            .collect::<Vec<_>>(),
        bindings,
        children,
    ])
}

/// Excludes insignificant XML whitespace, including in empty WML containers.
pub(crate) fn significant_children(
    element: &XmlElement,
    preserve: bool,
) -> impl Iterator<Item = &XmlNode> {
    let preserve = match element.attribute(XML_NAMESPACE, "space") {
        Some("preserve") => true,
        Some("default") => false,
        _ => preserve,
    };
    let has_element_children = element.element_children().next().is_some()
        || (element.namespace == crate::wml::W
            && matches!(
                element.local.as_str(),
                "pPr"
                    | "rPr"
                    | "tblPr"
                    | "trPr"
                    | "tcPr"
                    | "sectPr"
                    | "tblGrid"
                    | "p"
                    | "r"
                    | "body"
                    | "document"
                    | "hdr"
                    | "ftr"
            ));
    element.children.iter().filter(move |child| match child {
        XmlNode::Element(_) => true,
        XmlNode::Text(text) => {
            preserve || !has_element_children || !text.chars().all(is_xml_whitespace)
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{XmlLimits, parse_part};

    fn fingerprint(xml: &str) -> String {
        let root = parse_part(xml.as_bytes(), "test.xml", &XmlLimits::default()).unwrap();
        structural_fingerprint(&root)
    }

    #[test]
    fn prefix_choice_is_insignificant() {
        assert_eq!(
            fingerprint(r#"<w:p xmlns:w="ns"><w:r/></w:p>"#),
            fingerprint(r#"<x:p xmlns:x="ns"><x:r/></x:p>"#)
        );
    }

    #[test]
    fn attribute_order_is_insignificant() {
        assert_eq!(
            fingerprint(r#"<p a="1" b="2"/>"#),
            fingerprint(r#"<p b="2" a="1"/>"#)
        );
    }

    #[test]
    fn element_order_is_significant() {
        assert_ne!(
            fingerprint("<p><a/><b/></p>"),
            fingerprint("<p><b/><a/></p>")
        );
    }

    #[test]
    fn attribute_value_is_significant() {
        assert_ne!(fingerprint(r#"<p a="1"/>"#), fingerprint(r#"<p a="2"/>"#));
    }

    #[test]
    fn interelement_whitespace_is_insignificant() {
        assert_eq!(
            fingerprint("<p>\n  <a/>\n  <b/>\n</p>"),
            fingerprint("<p><a/><b/></p>")
        );
    }

    #[test]
    fn preserved_whitespace_is_significant() {
        assert_ne!(
            fingerprint(r#"<t xml:space="preserve"><a/> </t>"#),
            fingerprint(r#"<t xml:space="preserve"><a/></t>"#)
        );
    }

    #[test]
    fn text_only_content_keeps_whitespace() {
        assert_ne!(fingerprint("<t> </t>"), fingerprint("<t></t>"));
    }

    #[test]
    fn binding_set_is_significant() {
        assert_ne!(
            fingerprint(r#"<p xmlns:a="ns-a"/>"#),
            fingerprint(r#"<p xmlns:a="ns-a" xmlns:b="ns-b"/>"#)
        );
    }

    #[test]
    fn nested_default_reverts_preserve() {
        assert_eq!(
            fingerprint(r#"<o xml:space="preserve"><i xml:space="default"><a/> </i></o>"#),
            fingerprint(r#"<o xml:space="preserve"><i xml:space="default"><a/></i></o>"#)
        );
    }

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    #[test]
    fn standard_declaration_boilerplate_is_forgiven() {
        assert_eq!(
            fingerprint(&format!(
                r#"<w:document xmlns:w="{W}"><w:body/></w:document>"#
            )),
            fingerprint(&format!(
                r#"<w:document xmlns:w="{W}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body/></w:document>"#
            ))
        );
    }

    #[test]
    fn a_custom_declaration_is_not_forgiven() {
        assert_ne!(
            fingerprint(&format!(r#"<w:document xmlns:w="{W}"/>"#)),
            fingerprint(&format!(
                r#"<w:document xmlns:w="{W}" xmlns:c="urn:custom"/>"#
            ))
        );
    }

    #[test]
    fn root_and_nested_mc_ignorable_are_recorded() {
        let mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        assert_ne!(
            fingerprint(&format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="{mc}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"/>"#
            )),
            fingerprint(&format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="{mc}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/>"#
            ))
        );
        assert_ne!(
            fingerprint(&format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="{mc}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:p mc:Ignorable="w14"/></w:document>"#
            )),
            fingerprint(&format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="{mc}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:p/></w:document>"#
            ))
        );
    }
}
