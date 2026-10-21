#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn document(paragraph: &str) -> String {
        format!(r#"<x:document xmlns:x="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:z="urn:opaque"><x:body>{paragraph}<x:sectPr z:flag='kept'><x:headerReference r:id='rId9' x:type="default"/><z:unknown a='b'/><x:pgSz x:h='15840' x:w="12240" z:keep='yes'/></x:sectPr></x:body></x:document>"#)
    }

    #[test]
    fn section_patch_keeps_exact_spans_and_clones_reference_rows() {
        for properties in ["", "<x:pPr/>", "<x:pPr z:a='b'><x:keepNext/><x:pPrChange/></x:pPr>"] {
            let paragraph = format!("<x:p z:tag='same'>{properties}<x:r z:r='same'><x:t>Text</x:t></x:r></x:p>");
            let source = document(&paragraph);
            let edit = SectionEdit { paragraph: Some(0), donor: None, properties: json!({}) };
            let saved = patch_sections(&source, &source, &[edit]).unwrap();
            let raw = "<x:sectPr z:flag='kept'><x:headerReference r:id='rId9' x:type=\"default\"/><z:unknown a='b'/><x:pgSz x:h='15840' x:w=\"12240\" z:keep='yes'/></x:sectPr>";
            assert_eq!(saved.matches(raw).count(), 2);
            assert!(saved.contains("<x:r z:r='same'><x:t>Text</x:t></x:r>"));
            assert!(saved.ends_with(&source[source.find("</x:p>").unwrap()..]));
            if !properties.is_empty() && properties != "<x:pPr/>" {
                assert!(saved.contains(&format!("<x:keepNext/>{raw}<x:pPrChange/>")));
            }
        }
    }

    #[test]
    fn final_section_patch_changes_only_requested_attribute_values() {
        let source = document("<x:p><x:r><x:t>Kept</x:t></x:r></x:p>");
        let saved = patch_sections(&source, &source, &[SectionEdit {
            paragraph: None, donor: None,
            properties: json!({"pageWidth":15840,"pageHeight":12240,"orientation":"landscape"}),
        }]).unwrap();
        assert_eq!(saved, source.replace("x:h='15840' x:w=\"12240\" z:keep='yes'", "x:h='12240' x:w=\"15840\" z:keep='yes' x:orient=\"landscape\""));
    }

    #[test]
    fn section_patch_refuses_table_paragraphs_and_ambiguous_properties() {
        let source = document("<x:tbl><x:tr><x:tc><x:p/></x:tc></x:tr></x:tbl>");
        assert!(patch_sections(&source, &source, &[SectionEdit { paragraph: Some(0), donor: None, properties: json!({}) }]).is_none());
        let source = document("<x:p><x:pPr/><x:pPr/></x:p>");
        assert!(patch_sections(&source, &source, &[SectionEdit { paragraph: Some(0), donor: None, properties: json!({}) }]).is_none());
    }
}
