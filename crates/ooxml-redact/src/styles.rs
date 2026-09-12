use std::collections::{HashMap, HashSet};

use quick_xml::NsReader;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;

use crate::RedactError;

const WORD_TRANSITIONAL: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORD_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";

#[derive(Default)]
pub(crate) struct StyleMap {
    replacements: HashMap<String, String>,
    names: HashMap<String, String>,
}

impl StyleMap {
    pub(crate) fn collect(path: &str, bytes: &[u8]) -> Result<Self, RedactError> {
        let mut reader = NsReader::from_reader(bytes);
        let mut stack = Vec::new();
        let mut styles = Vec::new();
        let mut current = None;

        loop {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| xml_error(path, error))?;
            match event {
                Event::Start(start) => {
                    let local =
                        String::from_utf8_lossy(start.name().local_name().as_ref()).into_owned();
                    let word = is_word_namespace(&namespace);
                    stack.push((word, local.clone()));
                    if word && local.eq_ignore_ascii_case("style") {
                        current = Some(parse_style_attributes(&reader, &start, path)?);
                    } else if word
                        && local.eq_ignore_ascii_case("name")
                        && let Some(style) = current.as_mut()
                    {
                        style.name = word_attribute(&reader, &start, "val", path)?;
                    }
                }
                Event::Empty(start) => {
                    let local_name = start.name().local_name();
                    let local = String::from_utf8_lossy(local_name.as_ref());
                    if is_word_namespace(&namespace) && local.eq_ignore_ascii_case("style") {
                        let style = parse_style_attributes(&reader, &start, path)?;
                        styles.push(style);
                    } else if is_word_namespace(&namespace)
                        && local.eq_ignore_ascii_case("name")
                        && let Some(style) = current.as_mut()
                    {
                        style.name = word_attribute(&reader, &start, "val", path)?;
                    }
                }
                Event::End(_) => {
                    let ended = stack.pop();
                    if ended
                        .as_ref()
                        .is_some_and(|(word, local)| *word && local.eq_ignore_ascii_case("style"))
                        && let Some(style) = current.take()
                    {
                        styles.push(style);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        let mut reserved = HashSet::new();
        for style in &styles {
            if let Some(id) = &style.id {
                reserved.insert(key(id));
            }
            if let Some(name) = &style.name {
                reserved.insert(key(name));
            }
        }

        let mut replacements = HashMap::new();
        let mut names = HashMap::new();
        let mut index = 1;
        for style in styles {
            let Some(id) = style.id else { continue };
            let builtin_name = style.name.as_deref().unwrap_or(&id);
            let builtin = (!style.custom || style.default) && is_builtin_name(builtin_name);
            if builtin && is_builtin_name(&id) {
                let name = builtin_name.to_owned();
                names.insert(id, name);
                continue;
            }
            let mut replacement = format!("RedactedStyle{index}");
            while reserved.contains(&key(&replacement)) {
                index += 1;
                replacement = format!("RedactedStyle{index}");
            }
            reserved.insert(key(&replacement));
            names.insert(
                id.clone(),
                if builtin {
                    builtin_name.to_owned()
                } else {
                    replacement.clone()
                },
            );
            replacements.insert(id, replacement);
            index += 1;
        }
        Ok(Self {
            replacements,
            names,
        })
    }

    pub(crate) fn replacement(
        &self,
        element: &str,
        attribute: &str,
        value: &str,
        current_style_id: Option<&str>,
    ) -> Option<String> {
        let element = element.rsplit(':').next().unwrap_or(element);
        let attribute = attribute.rsplit(':').next().unwrap_or(attribute);
        let style_element = element.eq_ignore_ascii_case("style");
        let name_element = element.eq_ignore_ascii_case("name");
        let alias_element = element.eq_ignore_ascii_case("aliases");
        if style_element && attribute.eq_ignore_ascii_case("styleId") {
            return self.replacements.get(value).cloned();
        }
        if (name_element || alias_element) && attribute.eq_ignore_ascii_case("val") {
            return current_style_id.and_then(|id| self.names.get(id).cloned());
        }
        let reference_element = matches_ignore_case(
            element,
            &[
                "basedOn",
                "next",
                "link",
                "pStyle",
                "rStyle",
                "tblStyle",
                "numStyleLink",
                "styleLink",
            ],
        );
        if reference_element && attribute.eq_ignore_ascii_case("val") {
            return self.replacements.get(value).cloned();
        }
        None
    }
}

struct StyleDefinition {
    id: Option<String>,
    custom: bool,
    default: bool,
    name: Option<String>,
}

fn parse_style_attributes<R: std::io::BufRead>(
    reader: &NsReader<R>,
    start: &quick_xml::events::BytesStart<'_>,
    path: &str,
) -> Result<StyleDefinition, RedactError> {
    let mut style = StyleDefinition {
        id: None,
        custom: false,
        default: false,
        name: None,
    };
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| xml_error(path, error))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !is_word_namespace(&namespace) {
            continue;
        }
        let name = String::from_utf8_lossy(local.as_ref());
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| xml_error(path, error))?;
        if name.eq_ignore_ascii_case("styleId") {
            style.id = Some(value.into_owned());
        } else if name.eq_ignore_ascii_case("customStyle") {
            style.custom = is_true(&value);
        } else if name.eq_ignore_ascii_case("default") {
            style.default = is_true(&value);
        }
    }
    Ok(style)
}

fn word_attribute<R: std::io::BufRead>(
    reader: &NsReader<R>,
    start: &quick_xml::events::BytesStart<'_>,
    expected: &str,
    path: &str,
) -> Result<Option<String>, RedactError> {
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| xml_error(path, error))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if is_word_namespace(&namespace)
            && String::from_utf8_lossy(local.as_ref()).eq_ignore_ascii_case(expected)
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| xml_error(path, error));
        }
    }
    Ok(None)
}

fn is_word_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(namespace) if namespace.as_ref() == WORD_TRANSITIONAL || namespace.as_ref() == WORD_STRICT)
}

fn is_true(value: &str) -> bool {
    matches!(
        value.trim().to_ascii_lowercase().as_str(),
        "1" | "true" | "on"
    )
}

fn is_builtin_name(value: &str) -> bool {
    let name: String = value
        .chars()
        .filter(|character| *character == ' ' || character.is_ascii_alphanumeric())
        .collect();
    if name != value {
        return false;
    }
    let name = name.replace(' ', "").to_ascii_lowercase();
    if ["heading", "toc", "index"].iter().any(|prefix| {
        name.strip_prefix(prefix).is_some_and(|suffix| {
            matches!(suffix, "1" | "2" | "3" | "4" | "5" | "6" | "7" | "8" | "9")
        })
    }) {
        return true;
    }
    matches!(
        name.as_str(),
        "normal"
            | "defaultparagraphfont"
            | "tablenormal"
            | "normaltable"
            | "nolist"
            | "title"
            | "subtitle"
            | "caption"
            | "header"
            | "footer"
            | "footnotetext"
            | "footnotereference"
            | "endnotetext"
            | "endnotereference"
            | "commenttext"
            | "commentreference"
            | "commentsubject"
            | "annotationtext"
            | "annotationreference"
            | "annotationsubject"
            | "balloontext"
            | "bodytext"
            | "bodytext2"
            | "bodytext3"
            | "bodytextindent"
            | "bodytextindent2"
            | "bodytextindent3"
            | "bodytextfirstindent"
            | "bodytextfirstindent2"
            | "list"
            | "list2"
            | "list3"
            | "list4"
            | "list5"
            | "listparagraph"
            | "listbullet"
            | "listbullet2"
            | "listbullet3"
            | "listbullet4"
            | "listbullet5"
            | "listnumber"
            | "listnumber2"
            | "listnumber3"
            | "listnumber4"
            | "listnumber5"
            | "listcontinue"
            | "listcontinue2"
            | "listcontinue3"
            | "listcontinue4"
            | "listcontinue5"
            | "nospacing"
            | "quote"
            | "intensequote"
            | "emphasis"
            | "intenseemphasis"
            | "strong"
            | "subtleemphasis"
            | "subtlereference"
            | "intensereference"
            | "booktitle"
            | "bibliography"
            | "hyperlink"
            | "followedhyperlink"
            | "tocheading"
            | "tableoffigures"
            | "tableofauthorities"
            | "toaheading"
            | "indexheading"
            | "linenumber"
            | "pagenumber"
            | "envelopeaddress"
            | "envelopereturn"
            | "macrotext"
            | "htmlnormal"
            | "htmlcode"
            | "htmlpreformatted"
            | "htmltypewriter"
            | "htmlkeyboard"
            | "htmlsample"
            | "htmlvariable"
            | "htmlacronym"
            | "htmladdress"
            | "htmldefinition"
            | "tablegrid"
    )
}

fn key(value: &str) -> String {
    value.to_ascii_lowercase()
}

fn matches_ignore_case(value: &str, candidates: &[&str]) -> bool {
    candidates
        .iter()
        .any(|candidate| value.eq_ignore_ascii_case(candidate))
}

fn xml_error(path: &str, error: impl std::fmt::Display) -> RedactError {
    RedactError::Xml {
        part: path.to_owned(),
        message: error.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::StyleMap;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    #[test]
    fn preserves_builtin_and_default_styles() {
        let xml = format!(
            r#"<w:styles xmlns:w="{W}"><w:style w:styleId="Normal" w:default="1"/><w:style w:styleId="Custom" w:customStyle="1"/></w:styles>"#
        );
        let map = StyleMap::collect("styles.xml", xml.as_bytes()).unwrap();
        assert_eq!(map.replacement("style", "styleId", "Normal", None), None);
        assert_eq!(
            map.replacement("style", "styleId", "Custom", None),
            Some("RedactedStyle1".into())
        );
    }

    #[test]
    fn avoids_case_insensitive_collisions_and_supports_forward_references() {
        let xml = format!(
            r#"<w:styles xmlns:w="{W}"><w:style w:styleId="redactedstyle1" w:name="taken"/><w:style w:styleId="Custom" w:customStyle="on"><w:name w:val="Project"/></w:style></w:styles>"#
        );
        let map = StyleMap::collect("styles.xml", xml.as_bytes()).unwrap();
        assert_eq!(map.replacement("pStyle", "val", "custom", None), None);
        assert_eq!(
            map.replacement("pStyle", "val", "Custom", None),
            Some("RedactedStyle3".into())
        );
        assert_eq!(
            map.replacement("name", "val", "ignored", Some("Custom")),
            Some("RedactedStyle3".into())
        );
    }

    #[test]
    fn aliases_use_the_canonical_anonymous_name() {
        let xml = format!(
            r#"<w:styles xmlns:w="{W}"><w:style w:styleId="Custom" w:customStyle="true"><w:name w:val="Project"/><w:aliases w:val="Sensitive,Other"/></w:style></w:styles>"#
        );
        let map = StyleMap::collect("styles.xml", xml.as_bytes()).unwrap();
        assert_eq!(
            map.replacement("aliases", "val", "Sensitive,Other", Some("Custom")),
            Some("RedactedStyle1".into())
        );
    }

    #[test]
    fn resolves_strict_namespaces_after_rebinding() {
        let strict = "http://purl.oclc.org/ooxml/wordprocessingml/main";
        let xml = format!(
            r#"<w:styles xmlns:w="{strict}"><w:style w:styleId="Custom" w:customStyle="1"><w:name w:val="Project"/><w:basedOn w:val="Custom"/></w:style><foreign xmlns:w="urn:foreign"><w:style w:styleId="Ignored" w:customStyle="1"/></foreign></w:styles>"#
        );
        let map = StyleMap::collect("styles.xml", xml.as_bytes()).unwrap();
        assert_eq!(
            map.replacement("basedOn", "val", "Custom", None),
            Some("RedactedStyle1".into())
        );
        assert_eq!(map.replacement("style", "styleId", "Ignored", None), None);
    }

    #[test]
    fn anonymizes_unrecognized_default_style_names() {
        let xml = format!(
            r#"<w:styles xmlns:w="{W}"><w:style w:styleId="Default" w:customStyle="1" w:default="on"><w:name w:val="Default"/></w:style></w:styles>"#
        );
        let map = StyleMap::collect("styles.xml", xml.as_bytes()).unwrap();
        assert_eq!(
            map.replacement("style", "styleId", "Default", None),
            Some("RedactedStyle1".into())
        );
    }
}
