//! WML semantic digest: "did a save→reopen keep the meaning?"
//!
//! The fingerprint cannot answer that — a serializer that drops content still
//! fingerprints consistently against itself. So the digest records meaning
//! (text, containment, nested property values, unknown subtrees) across a real
//! reopen, and diffs it as paths: a bare `false` reproduces the silent drop it
//! exists to catch.

use std::collections::BTreeMap;

use crate::error::FidelityError;
use crate::fingerprint::short_fingerprint;
use crate::xml::{XmlAttribute, XmlElement, XmlLimits, XmlNode, parse_part};
use crate::{Part, is_xml_part};

pub const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PACKAGE_RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const MAX_DEPTH: usize = 256;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SemanticDigest {
    /// One story per XML part, ordered by part name.
    pub stories: Vec<StoryDigest>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StoryDigest {
    pub part: String,
    pub root_attributes: Vec<XmlAttribute>,
    pub blocks: Vec<BlockRecord>,
    pub structure: Vec<String>,
    /// True for parts whose entries carry no authored order (relationships,
    /// content types): structure compares as a set, not pairwise.
    pub unordered: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BlockRecord {
    pub path: String,
    /// The paragraph element's own attributes, sorted: managed identity like
    /// `w14:paraId` and any foreign attribute are content, not plumbing.
    pub attributes: Vec<String>,
    /// Paragraph text with tabs and hard breaks as characters.
    pub text: String,
    pub paragraph_properties: Vec<String>,
    /// Per run, in run order.
    pub run_properties: Vec<Vec<String>>,
    /// Non-text meaning in document order: markers, links, drawings, unknowns.
    pub generic_structure: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Difference {
    pub path: String,
    pub before: String,
    pub after: String,
}

pub fn semantic_digest(parts: &[Part]) -> Result<SemanticDigest, FidelityError> {
    let limits = XmlLimits::default();
    let mut named: Vec<(&String, &Vec<u8>)> = parts
        .iter()
        .filter(|(name, _)| is_xml_part(name))
        .map(|(name, bytes)| (name, bytes))
        .collect();
    named.sort_by_key(|(name, _)| name.as_str());
    let mut stories = Vec::with_capacity(named.len());
    for (name, bytes) in named {
        let root = parse_part(bytes, name, &limits)?;
        stories.push(digest_part(name, &root));
    }
    Ok(SemanticDigest { stories })
}

pub fn diff_digests(before: &SemanticDigest, after: &SemanticDigest) -> Vec<Difference> {
    let mut differences = Vec::new();
    let before_parts: BTreeMap<&str, &StoryDigest> = before
        .stories
        .iter()
        .map(|story| (story.part.as_str(), story))
        .collect();
    let after_parts: BTreeMap<&str, &StoryDigest> = after
        .stories
        .iter()
        .map(|story| (story.part.as_str(), story))
        .collect();
    let parts: std::collections::BTreeSet<&str> = before_parts
        .keys()
        .chain(after_parts.keys())
        .copied()
        .collect();
    for part in parts {
        match (before_parts.get(part), after_parts.get(part)) {
            (Some(before), Some(after)) => diff_story(part, before, after, &mut differences),
            (Some(_), None) => differences.push(difference(part, "present", "absent")),
            (None, Some(_)) => differences.push(difference(part, "absent", "present")),
            (None, None) => unreachable!(),
        }
    }
    differences
}

fn diff_story(part: &str, before: &StoryDigest, after: &StoryDigest, out: &mut Vec<Difference>) {
    let mut after_attributes = after.root_attributes.clone();
    crate::registry::normalize_root_ignorable(&before.root_attributes, &mut after_attributes);
    let before_attributes = sorted_attribute_tokens(&before.root_attributes).join(",");
    let after_attributes = sorted_attribute_tokens(&after_attributes).join(",");
    if before_attributes != after_attributes {
        out.push(difference(
            &format!("{part} root.attributes"),
            &before_attributes,
            &after_attributes,
        ));
    }
    if before.blocks.len() != after.blocks.len() {
        out.push(difference(
            &format!("{part} blocks"),
            &before.blocks.len().to_string(),
            &after.blocks.len().to_string(),
        ));
    }
    for (index, pair) in before.blocks.iter().zip(after.blocks.iter()).enumerate() {
        let (before, after) = pair;
        diff_field(part, index, "path", &before.path, &after.path, out);
        diff_field(
            part,
            index,
            "attributes",
            &before.attributes.join(","),
            &after.attributes.join(","),
            out,
        );
        diff_field(part, index, "text", &before.text, &after.text, out);
        diff_field(
            part,
            index,
            "paragraphProperties",
            &before.paragraph_properties.join(","),
            &after.paragraph_properties.join(","),
            out,
        );
        diff_field(
            part,
            index,
            "runProperties",
            &join_nested(&before.run_properties),
            &join_nested(&after.run_properties),
            out,
        );
        diff_field(
            part,
            index,
            "genericStructure",
            &before.generic_structure.join(","),
            &after.generic_structure.join(","),
            out,
        );
    }
    if before.unordered || after.unordered {
        let before_set: std::collections::BTreeSet<&str> =
            before.structure.iter().map(String::as_str).collect();
        let after_set: std::collections::BTreeSet<&str> =
            after.structure.iter().map(String::as_str).collect();
        for line in before_set.difference(&after_set) {
            out.push(difference(&format!("{part} entry"), line, "absent"));
        }
        for line in after_set.difference(&before_set) {
            out.push(difference(&format!("{part} entry"), "absent", line));
        }
        return;
    }
    let lines = before.structure.len().max(after.structure.len());
    for index in 0..lines {
        let before_line = before.structure.get(index).map(String::as_str);
        let after_line = after.structure.get(index).map(String::as_str);
        if before_line != after_line {
            out.push(difference(
                &format!("{part} structure[{index}]"),
                before_line.unwrap_or("absent"),
                after_line.unwrap_or("absent"),
            ));
        }
    }
}

fn diff_field(
    part: &str,
    index: usize,
    field: &str,
    before: &str,
    after: &str,
    out: &mut Vec<Difference>,
) {
    if before != after {
        out.push(difference(
            &format!("{part} block[{index}].{field}"),
            before,
            after,
        ));
    }
}

fn difference(path: &str, before: &str, after: &str) -> Difference {
    Difference {
        path: path.to_owned(),
        before: before.to_owned(),
        after: after.to_owned(),
    }
}

fn join_nested(nested: &[Vec<String>]) -> String {
    nested
        .iter()
        .map(|tokens| tokens.join(","))
        .collect::<Vec<_>>()
        .join(";")
}

struct StoryState {
    blocks: Vec<BlockRecord>,
    structure: Vec<String>,
}

fn digest_part(part: &str, root: &XmlElement) -> StoryDigest {
    let mut state = StoryState {
        blocks: Vec::new(),
        structure: Vec::new(),
    };
    let container = if root.is(W, "document") {
        root.child(W, "body").unwrap_or(root)
    } else {
        root
    };
    // Relationship and content-type entries carry no authored order, so they
    // digest as an identity set: entry tokens without positions, sorted (the
    // declared "relationship-and-content-type-order" normalization).
    let unordered =
        root.is(PACKAGE_RELATIONSHIPS, "Relationships") || root.is(CONTENT_TYPES, "Types");
    if unordered {
        state.structure = container
            .element_children()
            .map(element_token)
            .collect::<Vec<_>>();
        state.structure.sort_unstable();
    } else {
        collect_blocks(container, "", &mut state, 0);
    }
    StoryDigest {
        part: part.to_owned(),
        root_attributes: root.attributes.clone(),
        blocks: state.blocks,
        structure: state.structure,
        unordered,
    }
}

fn collect_blocks(container: &XmlElement, path: &str, state: &mut StoryState, depth: usize) {
    if depth > MAX_DEPTH {
        return;
    }
    let mut index = 0usize;
    for child in &container.children {
        match child {
            XmlNode::Text(text) => {
                if !text.chars().all(char::is_whitespace) {
                    state.structure.push(format!("{path}/#text {text}"));
                }
            }
            XmlNode::Element(element) => {
                let child_path = if path.is_empty() {
                    index.to_string()
                } else {
                    format!("{path}/{index}")
                };
                index += 1;
                if element.is(W, "p") {
                    let record = block_record(element, &child_path, state, depth + 1);
                    state.blocks.push(record);
                    state.structure.push(format!("{child_path} p"));
                } else {
                    state
                        .structure
                        .push(format!("{child_path} {}", element_token(element)));
                    collect_blocks(element, &child_path, state, depth + 1);
                }
            }
        }
    }
}

fn block_record(
    paragraph: &XmlElement,
    path: &str,
    state: &mut StoryState,
    depth: usize,
) -> BlockRecord {
    let paragraph_properties = paragraph
        .child(W, "pPr")
        .map(|properties| property_tokens(properties, depth))
        .unwrap_or_default();
    let mut record = BlockRecord {
        path: path.to_owned(),
        attributes: attribute_tokens(paragraph),
        text: String::new(),
        paragraph_properties,
        run_properties: Vec::new(),
        generic_structure: Vec::new(),
    };
    inline_walk(paragraph, &mut record, path, state, depth);
    record
}

fn inline_walk(
    container: &XmlElement,
    record: &mut BlockRecord,
    path: &str,
    state: &mut StoryState,
    depth: usize,
) {
    if depth > MAX_DEPTH {
        return;
    }
    for element in container.element_children() {
        if element.is(W, "pPr") {
            continue;
        }
        if element.is(W, "r") {
            digest_run(element, record, path, state, depth);
        } else if is_transparent_container(element) {
            record.generic_structure.push(element_token(element));
            if element.is(W, "sdt") {
                if let Some(properties) = element.child(W, "sdtPr") {
                    record.generic_structure.push(subtree_token(properties));
                }
                if let Some(content) = element.child(W, "sdtContent") {
                    inline_walk(content, record, path, state, depth + 1);
                }
            } else {
                inline_walk(element, record, path, state, depth + 1);
            }
        } else if element.children.is_empty() {
            record.generic_structure.push(element_token(element));
        } else {
            record.generic_structure.push(subtree_token(element));
            scan_nested_stories(element, path, state, depth + 1);
        }
    }
}

fn digest_run(
    run: &XmlElement,
    record: &mut BlockRecord,
    path: &str,
    state: &mut StoryState,
    depth: usize,
) {
    record.run_properties.push(
        run.child(W, "rPr")
            .map(|properties| property_tokens(properties, depth))
            .unwrap_or_default(),
    );
    for element in run.element_children() {
        if element.is(W, "rPr") {
            continue;
        }
        if element.namespace == W {
            match element.local.as_str() {
                "t" | "delText" | "instrText" | "delInstrText" => {
                    push_text(element, &mut record.text);
                    continue;
                }
                "tab" | "ptab" => {
                    record.text.push('\t');
                    continue;
                }
                "br" | "cr" => {
                    record.text.push('\n');
                    if !element.attributes.is_empty() {
                        record.generic_structure.push(element_token(element));
                    }
                    continue;
                }
                "noBreakHyphen" => {
                    record.text.push('\u{2011}');
                    continue;
                }
                "softHyphen" => {
                    record.text.push('\u{00AD}');
                    continue;
                }
                _ => {}
            }
        }
        if element.children.is_empty() {
            record.generic_structure.push(element_token(element));
        } else {
            record.generic_structure.push(subtree_token(element));
            scan_nested_stories(element, path, state, depth + 1);
        }
    }
}

/// Containers whose runs join the paragraph's text stream.
fn is_transparent_container(element: &XmlElement) -> bool {
    element.namespace == W
        && matches!(
            element.local.as_str(),
            "hyperlink"
                | "ins"
                | "del"
                | "moveFrom"
                | "moveTo"
                | "smartTag"
                | "sdt"
                | "dir"
                | "bdo"
                | "fldSimple"
        )
}

/// Text boxes carry paragraphs inside otherwise-opaque subtrees; digest them
/// as nested stories so reach stays total.
fn scan_nested_stories(element: &XmlElement, path: &str, state: &mut StoryState, depth: usize) {
    if depth > MAX_DEPTH {
        return;
    }
    let mut nested = 0usize;
    for child in element.element_children() {
        if child.is(W, "txbxContent") {
            let nested_path = format!("{path}/txbx{nested}");
            nested += 1;
            collect_blocks(child, &nested_path, state, depth + 1);
        } else {
            scan_nested_stories(child, path, state, depth + 1);
        }
    }
}

/// Property tokens read children, not just the first level: `w:numPr` with
/// numId 3 and `w:numPr` with numId 99 must not digest as the same bare name.
fn property_tokens(properties: &XmlElement, depth: usize) -> Vec<String> {
    let mut tokens: Vec<String> = properties
        .children
        .iter()
        .filter_map(|child| match child {
            XmlNode::Element(element) => Some(property_token(element, depth)),
            XmlNode::Text(text) => {
                (!text.chars().all(char::is_whitespace)).then(|| format!("#{text}"))
            }
        })
        .collect();
    tokens.sort_unstable();
    tokens
}

fn property_token(element: &XmlElement, depth: usize) -> String {
    let head = element_token(element);
    if depth >= MAX_DEPTH || element.children.is_empty() {
        return head;
    }
    let nested = property_tokens(element, depth + 1);
    if nested.is_empty() {
        head
    } else {
        format!("{head}[{}]", nested.join(","))
    }
}

/// Name plus sorted attributes, no children.
pub(crate) fn element_token(element: &XmlElement) -> String {
    let name = qualified_name(element);
    let attributes = attribute_tokens(element);
    if attributes.is_empty() {
        return name;
    }
    format!("{name}({})", attributes.join(","))
}

fn attribute_tokens(element: &XmlElement) -> Vec<String> {
    sorted_attribute_tokens(&element.attributes)
}

fn sorted_attribute_tokens(attributes: &[XmlAttribute]) -> Vec<String> {
    let mut attributes: Vec<String> = attributes
        .iter()
        .map(|attribute| {
            if attribute.namespace.is_empty() {
                format!("{}={}", attribute.local, attribute.value)
            } else {
                format!(
                    "{{{}}}{}={}",
                    attribute.namespace, attribute.local, attribute.value
                )
            }
        })
        .collect();
    attributes.sort_unstable();
    attributes
}

/// Name plus a whole-subtree fingerprint; descend no further.
fn subtree_token(element: &XmlElement) -> String {
    format!("{}#{}", qualified_name(element), short_fingerprint(element))
}

fn qualified_name(element: &XmlElement) -> String {
    if element.namespace == W {
        format!("w:{}", element.local)
    } else if element.namespace.is_empty() {
        element.local.clone()
    } else {
        format!("{{{}}}{}", element.namespace, element.local)
    }
}

fn push_text(element: &XmlElement, text: &mut String) {
    for child in &element.children {
        match child {
            XmlNode::Text(value) => text.push_str(value),
            XmlNode::Element(nested) => push_text(nested, text),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn document(body: &str) -> Vec<Part> {
        let xml = format!(r#"<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#);
        vec![("word/document.xml".to_owned(), xml.into_bytes())]
    }

    fn digest(body: &str) -> SemanticDigest {
        semantic_digest(&document(body)).unwrap()
    }

    fn assert_differs(before: &str, after: &str) {
        assert_ne!(
            diff_digests(&digest(before), &digest(after)),
            vec![],
            "documents that mean different things digested equal"
        );
    }

    #[test]
    fn text_tabs_and_breaks_are_characters() {
        let d =
            digest(r#"<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r></w:p>"#);
        assert_eq!(d.stories[0].blocks[0].text, "a\tb\nc");
    }

    #[test]
    fn lexical_noise_digests_equal() {
        let with_prefix = digest(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#);
        let xml = format!(
            r#"<x:document xmlns:x="{W}"><x:body>
              <x:p><x:r><x:t>x</x:t></x:r></x:p>
            </x:body></x:document>"#
        );
        let renamed =
            semantic_digest(&[("word/document.xml".to_owned(), xml.into_bytes())]).unwrap();
        assert_eq!(diff_digests(&with_prefix, &renamed), vec![]);
    }

    #[test]
    fn paragraph_attributes_are_meaning() {
        assert_differs(
            r#"<w:p w14:paraId="11111111" xmlns:w14="ns-w14"><w:r><w:t>x</w:t></w:r></w:p>"#,
            r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#,
        );
    }

    #[test]
    fn numbering_identity_is_meaning() {
        assert_differs(
            r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:p>"#,
            r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="5"/><w:numId w:val="99"/></w:numPr></w:pPr></w:p>"#,
        );
    }

    #[test]
    fn an_emptied_property_container_is_meaning() {
        assert_differs(
            r#"<w:p><w:pPr><w:pBdr><w:top w:val="single"/></w:pBdr></w:pPr></w:p>"#,
            r#"<w:p><w:pPr><w:pBdr/></w:pPr></w:p>"#,
        );
    }

    #[test]
    fn lost_tab_stops_are_meaning() {
        assert_differs(
            r#"<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="9360" w:leader="dot"/></w:tabs></w:pPr></w:p>"#,
            r#"<w:p><w:pPr><w:tabs/></w:pPr></w:p>"#,
        );
    }

    #[test]
    fn changed_section_setup_is_meaning() {
        assert_differs(
            r#"<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:cols w:num="2"/></w:sectPr>"#,
            r#"<w:p/><w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:cols w:num="1"/></w:sectPr>"#,
        );
    }

    #[test]
    fn an_emptied_content_control_is_meaning() {
        assert_differs(
            r#"<w:sdt><w:sdtPr><w:alias w:val="Field"/><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>value</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
            r#"<w:sdt><w:sdtPr><w:alias w:val="Field"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#,
        );
    }

    #[test]
    fn drawing_geometry_is_meaning() {
        assert_differs(
            r#"<w:p><w:r><w:drawing><wp:inline xmlns:wp="ns-wp"><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Picture 1"/></wp:inline></w:drawing></w:r></w:p>"#,
            r#"<w:p><w:r><w:drawing><wp:inline xmlns:wp="ns-wp"><wp:extent cx="457200" cy="457200"/><wp:docPr id="1" name="Picture 1"/></wp:inline></w:drawing></w:r></w:p>"#,
        );
    }

    /// An unknown child of a known property container is meaning the model
    /// cannot type; the digest must still see it go.
    #[test]
    fn foreign_property_children_are_meaning() {
        assert_differs(
            r#"<w:p><w:pPr><w:jc w:val="center"/><cx:mark xmlns:cx="urn:custom-x" cx:v="1"/></w:pPr></w:p>"#,
            r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>"#,
        );
    }

    #[test]
    fn a_gutted_style_definition_is_meaning() {
        let styles = |body: &str| {
            semantic_digest(&[(
                "word/styles.xml".to_owned(),
                format!(r#"<w:styles xmlns:w="{W}">{body}</w:styles>"#).into_bytes(),
            )])
            .unwrap()
        };
        let full = styles(
            r#"<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>"#,
        );
        let gutted = styles(
            r#"<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>"#,
        );
        assert_ne!(diff_digests(&full, &gutted), vec![]);
    }

    #[test]
    fn a_flattened_table_is_meaning() {
        assert_differs(
            r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
            r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#,
        );
    }

    #[test]
    fn a_dropped_bookmark_is_meaning() {
        assert_differs(
            r#"<w:p><w:bookmarkStart w:id="1" w:name="a"/><w:r><w:t>x</w:t></w:r></w:p>"#,
            r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#,
        );
    }

    #[test]
    fn paragraph_mark_properties_are_meaning() {
        assert_differs(
            r#"<w:p><w:pPr><w:rPr><w:b/><w:sz w:val="48"/></w:rPr></w:pPr></w:p>"#,
            r#"<w:p><w:pPr><w:rPr/></w:pPr></w:p>"#,
        );
    }

    #[test]
    fn hyperlink_identity_and_content_are_meaning() {
        assert_differs(
            r##"<w:p><w:hyperlink w:anchor="a"><w:r><w:t>go</w:t></w:r></w:hyperlink></w:p>"##,
            r##"<w:p><w:hyperlink w:anchor="b"><w:r><w:t>go</w:t></w:r></w:hyperlink></w:p>"##,
        );
        let linked = digest(
            r##"<w:p><w:hyperlink w:anchor="a"><w:r><w:t>go</w:t></w:r></w:hyperlink></w:p>"##,
        );
        assert_eq!(linked.stories[0].blocks[0].text, "go");
    }

    #[test]
    fn a_gutted_definition_part_is_meaning() {
        let full = format!(
            r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum></w:numbering>"#
        );
        let gutted = format!(
            r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum></w:numbering>"#
        );
        let before =
            semantic_digest(&[("word/numbering.xml".to_owned(), full.into_bytes())]).unwrap();
        let after =
            semantic_digest(&[("word/numbering.xml".to_owned(), gutted.into_bytes())]).unwrap();
        assert_ne!(diff_digests(&before, &after), vec![]);
    }

    #[test]
    fn text_box_paragraphs_are_reached() {
        let d = digest(
            r#"<w:p><w:r><w:pict><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p>"#,
        );
        let texts: Vec<&str> = d.stories[0]
            .blocks
            .iter()
            .map(|block| block.text.as_str())
            .collect();
        assert!(texts.contains(&"boxed"));
    }

    #[test]
    fn a_dropped_story_part_names_itself() {
        let two = vec![
            document(r#"<w:p/>"#).remove(0),
            (
                "word/header1.xml".to_owned(),
                format!(r#"<w:hdr xmlns:w="{W}"><w:p/></w:hdr>"#).into_bytes(),
            ),
        ];
        let before = semantic_digest(&two).unwrap();
        let after = semantic_digest(&two[..1]).unwrap();
        assert_eq!(
            diff_digests(&before, &after),
            vec![Difference {
                path: "word/header1.xml".to_owned(),
                before: "present".to_owned(),
                after: "absent".to_owned(),
            }]
        );
    }

    #[test]
    fn diff_names_the_lost_field() {
        let before = digest(r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r></w:p>"#);
        let after = digest(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#);
        let differences = diff_digests(&before, &after);
        assert_eq!(differences.len(), 1);
        assert_eq!(
            differences[0].path,
            "word/document.xml block[0].runProperties"
        );
        assert_eq!(differences[0].before, "w:b");
    }
}
