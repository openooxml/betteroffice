//! Small XML writer with TypeScript-compatible escaping and ordering.
//!
//! Element and attribute names are `&'static str` by construction: model data
//! can only enter through [`XmlWriter::attribute`] and [`XmlWriter::text`],
//! both of which escape it. This keeps attacker-derived strings out of markup
//! positions while allowing serializer call sites to control byte ordering.

use std::fmt::Write as _;

/// Match the incumbent TypeScript `escapeXml` replacement semantics exactly.
pub fn escape_xml(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

/// Match `String(Math.round(value))`, including negative ties and `-0`.
///
/// The caller supplies an `Option` because TypeScript's `intAttr` also maps
/// `null`/`undefined` to zero. JSON cannot represent non-finite numbers; the
/// explicit finite check keeps native callers compatible as well.
pub fn int_attr(value: Option<f64>) -> String {
    let Some(value) = value.filter(|value| value.is_finite()) else {
        return "0".to_owned();
    };
    let rounded = (value + 0.5).floor();
    if rounded == 0.0 {
        return "0".to_owned();
    }
    js_number(rounded)
}

/// ECMAScript-compatible shortest-round-trip number formatting.
pub fn js_number(value: f64) -> String {
    if value == 0.0 {
        return "0".to_owned();
    }
    let mut buffer = ryu_js::Buffer::new();
    buffer.format(value).to_owned()
}

/// An ordering-preserving XML writer.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct XmlWriter {
    output: String,
    open_start_tag: bool,
    elements: Vec<&'static str>,
}

impl XmlWriter {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn with_capacity(capacity: usize) -> Self {
        Self {
            output: String::with_capacity(capacity),
            ..Self::default()
        }
    }

    /// Append an XML declaration with the incumbent serializer's exact bytes.
    pub fn declaration(&mut self) -> &mut Self {
        self.close_start_tag();
        self.output
            .push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        self
    }

    /// Begin an element. Attributes may be appended until content starts.
    pub fn start_element(&mut self, name: &'static str) -> &mut Self {
        self.close_start_tag();
        self.output.push('<');
        self.output.push_str(name);
        self.elements.push(name);
        self.open_start_tag = true;
        self
    }

    /// Append an escaped attribute to the current start tag.
    pub fn attribute(&mut self, name: &'static str, value: &str) -> &mut Self {
        assert!(
            self.open_start_tag,
            "XML attributes must immediately follow start_element"
        );
        self.output.push(' ');
        self.output.push_str(name);
        self.output.push_str("=\"");
        write_escaped(value, &mut self.output);
        self.output.push('"');
        self
    }

    /// Append escaped character data.
    pub fn text(&mut self, value: &str) -> &mut Self {
        self.close_start_tag();
        write_escaped(value, &mut self.output);
        self
    }

    /// Append XML produced by another `XmlWriter` in this serializer crate.
    /// Attacker-derived raw XML must go through `serializer::raw` instead.
    pub(crate) fn append_serialized(&mut self, xml: &str) -> &mut Self {
        self.close_start_tag();
        self.output.push_str(xml);
        self
    }

    /// Close the current element, using `/>` when it has no content.
    pub fn end_element(&mut self) -> &mut Self {
        let name = self
            .elements
            .pop()
            .expect("end_element requires a matching start_element");
        if self.open_start_tag {
            self.output.push_str("/>");
            self.open_start_tag = false;
        } else {
            self.output.push_str("</");
            self.output.push_str(name);
            self.output.push('>');
        }
        self
    }

    /// Close an attribute-less empty element as `<name />`.
    ///
    /// This odd spelling is retained only for the incumbent `w:col` case,
    /// whose template includes a literal space even when it has no attributes.
    pub(crate) fn end_empty_element_with_space(&mut self) -> &mut Self {
        let _name = self
            .elements
            .pop()
            .expect("end_empty_element_with_space requires a matching start_element");
        assert!(
            self.open_start_tag,
            "end_empty_element_with_space requires an empty element"
        );
        self.output.push_str(" />");
        self.open_start_tag = false;
        self
    }

    /// Complete the document or fragment and return its exact bytes.
    pub fn finish(mut self) -> String {
        self.close_start_tag();
        assert!(
            self.elements.is_empty(),
            "finish requires every XML element to be closed"
        );
        self.output
    }

    fn close_start_tag(&mut self) {
        if self.open_start_tag {
            self.output.push('>');
            self.open_start_tag = false;
        }
    }
}

fn write_escaped(value: &str, output: &mut String) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => {
                // Writing a char to String is infallible.
                let _ = output.write_char(character);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn escape_xml_matches_typescript_replacement_order() {
        assert_eq!(
            escape_xml("<&>\"' &amp;"),
            "&lt;&amp;&gt;&quot;&apos; &amp;amp;"
        );
    }

    #[test]
    fn writer_escapes_every_model_data_position_and_preserves_order() {
        let mut writer = XmlWriter::new();
        writer
            .start_element("w:root")
            .attribute("second", "<&\"'")
            .attribute("first", "safe")
            .start_element("w:text")
            .text("<&>\"'")
            .end_element()
            .start_element("w:empty")
            .end_element()
            .end_element();
        assert_eq!(
            writer.finish(),
            "<w:root second=\"&lt;&amp;&quot;&apos;\" first=\"safe\"><w:text>&lt;&amp;&gt;&quot;&apos;</w:text><w:empty/></w:root>"
        );
    }

    #[test]
    fn declaration_matches_incumbent_bytes() {
        let mut writer = XmlWriter::new();
        writer.declaration().start_element("w:root").end_element();
        assert_eq!(
            writer.finish(),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:root/>"
        );
    }

    #[test]
    fn integer_attributes_match_javascript_math_round() {
        assert_eq!(int_attr(None), "0");
        assert_eq!(int_attr(Some(f64::NAN)), "0");
        assert_eq!(int_attr(Some(f64::INFINITY)), "0");
        assert_eq!(int_attr(Some(1_008.000_000_000_000_1)), "1008");
        assert_eq!(int_attr(Some(1.5)), "2");
        assert_eq!(int_attr(Some(-1.5)), "-1");
        assert_eq!(int_attr(Some(-0.5)), "0");
    }

    #[test]
    #[should_panic(expected = "XML attributes must immediately follow start_element")]
    fn attributes_cannot_be_injected_after_content() {
        let mut writer = XmlWriter::new();
        writer.start_element("w:p").text("content");
        writer.attribute("unsafe", "value");
    }
}
