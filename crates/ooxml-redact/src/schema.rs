use quick_xml::name::ResolveResult;

const XSD_NAMESPACE: &[u8] = b"http://www.w3.org/2001/XMLSchema";

pub(crate) fn is_schema_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(uri) if uri.as_ref() == XSD_NAMESPACE)
}

pub(crate) fn preserve_attribute(element: &str, attribute: &str) -> bool {
    matches!(
        (element, attribute),
        ("schema", "targetNamespace")
            | ("schema", "elementFormDefault")
            | ("schema", "attributeFormDefault")
            | ("schema", "blockDefault")
            | ("schema", "finalDefault")
            | ("import", "namespace")
            | ("group", "name")
            | ("group", "ref")
            | ("group", "minOccurs")
            | ("group", "maxOccurs")
            | ("attributeGroup", "name")
            | ("attributeGroup", "ref")
            | ("element", "name")
            | ("element", "ref")
            | ("element", "type")
            | ("element", "nillable")
            | ("element", "minOccurs")
            | ("element", "maxOccurs")
            | ("element", "abstract")
            | ("element", "block")
            | ("element", "final")
            | ("element", "form")
            | ("element", "substitutionGroup")
            | ("attribute", "name")
            | ("attribute", "ref")
            | ("attribute", "type")
            | ("attribute", "form")
            | ("attribute", "use")
            | ("complexType", "name")
            | ("complexType", "mixed")
            | ("complexType", "abstract")
            | ("complexType", "block")
            | ("complexType", "final")
            | ("simpleType", "name")
            | ("simpleType", "final")
            | ("restriction", "base")
            | ("extension", "base")
            | ("list", "itemType")
            | ("union", "memberTypes")
            | ("any", "namespace")
            | ("any", "processContents")
            | ("any", "minOccurs")
            | ("any", "maxOccurs")
            | ("anyAttribute", "namespace")
            | ("anyAttribute", "processContents")
            | ("sequence", "minOccurs")
            | ("sequence", "maxOccurs")
            | ("choice", "minOccurs")
            | ("choice", "maxOccurs")
            | ("all", "minOccurs")
            | ("all", "maxOccurs")
            | ("key", "name")
            | ("keyref", "name")
            | ("keyref", "refer")
            | ("unique", "name")
            | ("selector", "xpath")
            | ("field", "xpath")
            | ("length", "value")
            | ("minLength", "value")
            | ("maxLength", "value")
            | ("totalDigits", "value")
            | ("fractionDigits", "value")
            | ("whiteSpace", "value")
            | ("length", "fixed")
            | ("minLength", "fixed")
            | ("maxLength", "fixed")
            | ("totalDigits", "fixed")
            | ("fractionDigits", "fixed")
            | ("whiteSpace", "fixed")
    )
}

pub(crate) fn drop_element(element: &str) -> bool {
    matches!(
        element,
        "annotation"
            | "documentation"
            | "appinfo"
            | "enumeration"
            | "pattern"
            | "minInclusive"
            | "maxInclusive"
            | "minExclusive"
            | "maxExclusive"
    )
}

pub(crate) fn drop_attribute(element: &str, attribute: &str) -> bool {
    matches!(attribute, "schemaLocation" | "source")
        || matches!(element, "element" | "attribute") && matches!(attribute, "default" | "fixed")
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::name::{Namespace, ResolveResult};

    #[test]
    fn recognizes_only_xsd_namespace() {
        assert!(is_schema_namespace(&ResolveResult::Bound(Namespace(
            XSD_NAMESPACE
        ))));
        assert!(!is_schema_namespace(&ResolveResult::Unbound));
        assert!(!is_schema_namespace(&ResolveResult::Bound(Namespace(
            b"http://schemas.microsoft.com/office/2006/metadata/properties",
        ))));
    }

    #[test]
    fn preserves_schema_structure_and_typed_facets() {
        assert!(preserve_attribute("element", "name"));
        assert!(preserve_attribute("element", "minOccurs"));
        assert!(preserve_attribute("sequence", "maxOccurs"));
        assert!(preserve_attribute("keyref", "refer"));
        assert!(preserve_attribute("field", "xpath"));
        assert!(preserve_attribute("maxLength", "value"));
        assert!(preserve_attribute("maxLength", "fixed"));
        assert!(preserve_attribute("schema", "targetNamespace"));
        assert!(!preserve_attribute("element", "default"));
        assert!(!preserve_attribute("import", "schemaLocation"));
    }

    #[test]
    fn drops_optional_annotation_and_literal_content() {
        assert!(drop_element("annotation"));
        assert!(drop_element("enumeration"));
        assert!(drop_element("pattern"));
        assert!(drop_element("minInclusive"));
        assert!(drop_attribute("element", "default"));
        assert!(drop_attribute("attribute", "fixed"));
        assert!(drop_attribute("import", "schemaLocation"));
        assert!(!drop_attribute("maxLength", "fixed"));
        assert!(!drop_element("element"));
        assert!(!drop_attribute("element", "name"));
    }
}
