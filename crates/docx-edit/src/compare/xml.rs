//! Namespace-aware source XML: canonical forms without bookkeeping, and body paragraph models.

use std::collections::HashSet;
use std::rc::Rc;

use docx_parse::{XmlElement, XmlNode};

use super::CompareDiagnosticCode;
use crate::target::AtomKind;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W_STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const XML: &str = "http://www.w3.org/XML/1998/namespace";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Elements that record tracked changes, wherever they appear.
const REVISION_ELEMENTS: [&str; 24] = [
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "rPrChange",
    "pPrChange",
    "sectPrChange",
    "tblPrChange",
    "tblPrExChange",
    "trPrChange",
    "tcPrChange",
    "tblGridChange",
    "numberingChange",
    "cellIns",
    "cellDel",
    "cellMerge",
    "customXmlInsRangeStart",
    "customXmlDelRangeStart",
    "customXmlMoveFromRangeStart",
    "customXmlMoveToRangeStart",
];

/// The in-scope namespace bindings of a walk.
#[derive(Default)]
pub(crate) struct Namespaces {
    bindings: Vec<(String, String)>,
}

impl Namespaces {
    pub fn enter(&mut self, element: &XmlElement) -> usize {
        let mark = self.bindings.len();
        for (name, value) in &element.attributes {
            if name == "xmlns" {
                self.bindings.push((String::new(), value.clone()));
            } else if let Some(prefix) = name.strip_prefix("xmlns:") {
                self.bindings.push((prefix.to_owned(), value.clone()));
            }
        }
        mark
    }

    pub fn leave(&mut self, mark: usize) {
        self.bindings.truncate(mark);
    }

    fn uri(&self, prefix: &str) -> Option<&str> {
        if prefix == "xml" {
            return Some(XML);
        }
        self.bindings
            .iter()
            .rev()
            .find(|(bound, _)| bound == prefix)
            .map(|(_, uri)| uri.as_str())
    }

    /// The namespace and local name of an element or attribute name.
    pub fn resolve<'a>(&'a self, name: &'a str, attribute: bool) -> (Option<&'a str>, &'a str) {
        match name.split_once(':') {
            Some((prefix, local)) => (self.uri(prefix), local),
            None if attribute => (None, name),
            None => (self.uri(""), name),
        }
    }

    /// The WordprocessingML local name of `element`, if it is in that namespace.
    pub fn w_local<'e>(&self, element: &'e XmlElement) -> Option<&'e str> {
        let (prefix, local) = element
            .name
            .split_once(':')
            .unwrap_or(("", element.name.as_str()));
        matches!(self.uri(prefix), Some(W | W_STRICT)).then_some(local)
    }
}

fn is_bookkeeping_attribute(uri: Option<&str>, local: &str, element_w: Option<&str>) -> bool {
    match uri {
        Some(W | W_STRICT) => local.starts_with("rsid"),
        Some(W14) => matches!(local, "paraId" | "textId"),
        Some(MC) => local == "Ignorable",
        Some(XML) => local == "space" && matches!(element_w, Some("t" | "delText" | "instrText")),
        _ => false,
    }
}

fn is_bookkeeping_element(local: Option<&str>) -> bool {
    matches!(local, Some("proofErr" | "lastRenderedPageBreak"))
}

fn is_text_element(local: Option<&str>) -> bool {
    matches!(local, Some("t" | "delText" | "instrText" | "delInstrText"))
}

fn escape(text: &str, out: &mut String) {
    for ch in text.chars() {
        match ch {
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '&' => out.push_str("&amp;"),
            '"' => out.push_str("&quot;"),
            _ => out.push(ch),
        }
    }
}

/// Descendants a skeleton writes as placeholders, by element-child path.
struct Holes<'a> {
    paths: &'a HashSet<Vec<u32>>,
    path: Vec<u32>,
}

/// Appends a canonical form of `element`: names by namespace, attributes sorted, and namespace
/// declarations, revision-session and paragraph ids, ignorable-namespace lists, proofing marks,
/// layout caches and whitespace between elements left out.
pub(crate) fn canonical(element: &XmlElement, ns: &mut Namespaces, out: &mut String) {
    write_canonical(element, ns, out, &mut None);
}

pub(crate) fn canonical_string(element: &XmlElement, ns: &mut Namespaces) -> String {
    let mut out = String::new();
    canonical(element, ns, &mut out);
    out
}

/// [`canonical`] with each descendant at an element-child path in `holes` written as
/// `<#block/>`.
pub(crate) fn canonical_skeleton(
    element: &XmlElement,
    ns: &mut Namespaces,
    holes: &HashSet<Vec<u32>>,
) -> String {
    let mut out = String::new();
    write_canonical(
        element,
        ns,
        &mut out,
        &mut Some(Holes {
            paths: holes,
            path: Vec::new(),
        }),
    );
    out
}

fn write_canonical(
    element: &XmlElement,
    ns: &mut Namespaces,
    out: &mut String,
    holes: &mut Option<Holes<'_>>,
) {
    let mark = ns.enter(element);
    let local = ns.w_local(element).map(str::to_owned);
    if is_bookkeeping_element(local.as_deref()) {
        ns.leave(mark);
        return;
    }
    let (uri, name) = ns.resolve(&element.name, false);
    out.push('<');
    if let Some(uri) = uri {
        out.push('{');
        out.push_str(uri);
        out.push('}');
    }
    out.push_str(name);
    let mut attributes: Vec<(String, &str)> = element
        .attributes
        .iter()
        .filter(|(key, _)| key.as_str() != "xmlns" && !key.starts_with("xmlns:"))
        .filter_map(|(key, value)| {
            let (uri, attribute) = ns.resolve(key, true);
            (!is_bookkeeping_attribute(uri, attribute, local.as_deref())).then(|| {
                (
                    uri.map_or_else(
                        || attribute.to_owned(),
                        |uri| format!("{{{uri}}}{attribute}"),
                    ),
                    value.as_str(),
                )
            })
        })
        .collect();
    attributes.sort();
    for (key, value) in attributes {
        out.push(' ');
        out.push_str(&key);
        out.push_str("=\"");
        escape(value, out);
        out.push('"');
    }
    out.push('>');
    let significant = is_text_element(local.as_deref());
    let mut index = 0u32;
    for child in &element.children {
        match child {
            XmlNode::Element(child) => {
                let hole = holes.as_mut().is_some_and(|holes| {
                    holes.path.push(index);
                    holes.paths.contains(&holes.path)
                });
                if hole {
                    out.push_str("<#block/>");
                } else {
                    write_canonical(child, ns, out, holes);
                }
                if let Some(holes) = holes.as_mut() {
                    holes.path.pop();
                }
                index += 1;
            }
            XmlNode::Text(text) | XmlNode::CData(text) => {
                if significant || !text.trim().is_empty() {
                    escape(text, out);
                }
            }
        }
    }
    out.push_str("</>");
    ns.leave(mark);
}

/// Whether a WordprocessingML element records a tracked change.
pub(crate) fn is_revision(local: &str) -> bool {
    REVISION_ELEMENTS.contains(&local)
}

/// The UTF-16 units of a WordprocessingML text element's content.
pub(crate) fn w_text_units(local: &str, element: &XmlElement) -> Option<usize> {
    is_text_element(Some(local)).then(|| {
        element
            .children
            .iter()
            .map(|child| match child {
                XmlNode::Text(text) | XmlNode::CData(text) => text.encode_utf16().count(),
                XmlNode::Element(_) => 0,
            })
            .sum()
    })
}

/// The largest numeric WordprocessingML `id` attribute below and on `element`.
pub(crate) fn highest_id(element: &XmlElement, ns: &mut Namespaces) -> u32 {
    let mark = ns.enter(element);
    let own = element
        .attributes
        .iter()
        .filter(|(key, _)| matches!(ns.resolve(key, true), (Some(W | W_STRICT), "id")))
        .filter_map(|(_, value)| value.trim().parse::<u32>().ok())
        .max()
        .unwrap_or(0);
    let highest = element
        .child_elements()
        .map(|child| highest_id(child, ns))
        .fold(own, u32::max);
    ns.leave(mark);
    highest
}

/// One UTF-16 unit of a paragraph's text with the canonical run properties it carries.
#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) struct Unit {
    pub rpr: Rc<str>,
    pub atom: Option<AtomKind>,
}

/// Content of a paragraph in source order: text units, or another child in canonical form.
#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) enum Item {
    Text { text: String, rpr: Rc<str> },
    Atom { kind: String, rpr: Rc<str> },
    Other { name: String, canonical: String },
}

/// Why a paragraph's content cannot be edited through the session and its save projection.
#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) struct Blocker {
    pub code: CompareDiagnosticCode,
    pub message: String,
}

/// A body paragraph as its source XML defines it.
#[derive(Clone, Debug)]
pub(crate) struct SourceParagraph {
    pub ppr: String,
    pub items: Vec<Item>,
    /// The projected text units, as the batch projection counts them.
    pub text: String,
    pub units: Vec<Unit>,
    pub para_id: Option<String>,
    /// The first content the comparison cannot edit, if any.
    pub blocker: Option<Blocker>,
}

impl SourceParagraph {
    /// Whether the two paragraphs have the same properties and content, run boundaries aside.
    pub fn same_content(&self, other: &Self) -> bool {
        self.ppr == other.ppr && merged(&self.items) == merged(&other.items)
    }

    /// Classifies how two paragraphs with equal projected text differ.
    pub fn difference(&self, other: &Self) -> (CompareDiagnosticCode, String) {
        if self.ppr != other.ppr {
            return (
                CompareDiagnosticCode::FormattingChange,
                "paragraph properties differ".to_owned(),
            );
        }
        let (left, right) = (merged(&self.items), merged(&other.items));
        let first = left
            .iter()
            .map(Some)
            .chain(std::iter::repeat(None))
            .zip(right.iter().map(Some).chain(std::iter::repeat(None)))
            .take(left.len().max(right.len()))
            .find(|(a, b)| a != b);
        match first {
            Some((Some(MergedItem::Unit((a, _))), Some(MergedItem::Unit((b, _))))) if a == b => (
                CompareDiagnosticCode::FormattingChange,
                "run formatting differs on unchanged text".to_owned(),
            ),
            Some((Some(MergedItem::Other(name, _)), _) | (_, Some(MergedItem::Other(name, _)))) => {
                (
                    content_code(name),
                    format!("{} content differs", element_label(name)),
                )
            }
            _ => (
                CompareDiagnosticCode::UnsupportedContent,
                "inline content differs".to_owned(),
            ),
        }
    }
}

#[derive(Debug, Eq, PartialEq)]
enum MergedItem<'a> {
    Unit((&'a str, &'a str)),
    Other(&'a str, &'a str),
}

fn merged(items: &[Item]) -> Vec<MergedItem<'_>> {
    let mut out = Vec::new();
    for item in items {
        match item {
            Item::Text { text, rpr } => {
                let mut start = 0;
                for (index, ch) in text.char_indices() {
                    let end = index + ch.len_utf8();
                    out.push(MergedItem::Unit((&text[start..end], rpr)));
                    start = end;
                }
            }
            Item::Atom { kind, rpr } => out.push(MergedItem::Unit((kind, rpr))),
            Item::Other { name, canonical } => out.push(MergedItem::Other(name, canonical)),
        }
    }
    out
}

/// The diagnostic an unequal non-text child of a paragraph reports.
pub(crate) fn content_code(name: &str) -> CompareDiagnosticCode {
    match name {
        "fldSimple" | "fldChar" | "instrText" | "fldData" | "hyperlink" => {
            CompareDiagnosticCode::FieldChange
        }
        "drawing" | "pict" | "object" | "AlternateContent" | "oMath" | "oMathPara" | "sym"
        | "ruby" | "ptab" => CompareDiagnosticCode::ObjectChange,
        "sdt" => CompareDiagnosticCode::ContentControlChange,
        _ => CompareDiagnosticCode::UnsupportedContent,
    }
}

fn element_label(name: &str) -> String {
    match name {
        "fldSimple" | "fldChar" | "instrText" | "fldData" => "field".to_owned(),
        "hyperlink" => "hyperlink".to_owned(),
        "sdt" => "content control".to_owned(),
        "drawing" | "pict" | "object" | "AlternateContent" => "drawing or object".to_owned(),
        _ => format!("w:{name}"),
    }
}

/// Run properties the session keeps and its save projection writes back, with the attributes
/// each may carry.
const RUN_PROPERTIES: [(&str, &[&str]); 29] = [
    ("rStyle", &["val"]),
    (
        "rFonts",
        &[
            "ascii",
            "hAnsi",
            "eastAsia",
            "cs",
            "asciiTheme",
            "hAnsiTheme",
            "eastAsiaTheme",
            "cstheme",
        ],
    ),
    ("b", &["val"]),
    ("bCs", &["val"]),
    ("i", &["val"]),
    ("iCs", &["val"]),
    ("caps", &["val"]),
    ("smallCaps", &["val"]),
    ("strike", &["val"]),
    ("dstrike", &["val"]),
    ("color", &["val", "themeColor", "themeTint", "themeShade"]),
    ("sz", &["val"]),
    ("szCs", &["val"]),
    ("highlight", &["val"]),
    ("u", &["val", "color"]),
    ("vertAlign", &["val"]),
    ("vanish", &["val"]),
    ("emboss", &["val"]),
    ("imprint", &["val"]),
    ("shadow", &["val"]),
    ("outline", &["val"]),
    ("rtl", &["val"]),
    ("spacing", &["val"]),
    ("position", &["val"]),
    ("w", &["val"]),
    ("kern", &["val"]),
    ("em", &["val"]),
    ("shd", &["val", "color", "fill"]),
    ("effect", &["val"]),
];

/// Whether a `w:shd` is the plain fill the session keeps: a clear pattern, an automatic colour
/// and an explicit RGB fill.
fn plain_fill(element: &XmlElement, ns: &Namespaces) -> bool {
    let value = |name: &str| {
        element
            .attributes
            .iter()
            .find(|(key, _)| ns.resolve(key, true) == (Some(W), name))
            .map(|(_, value)| value.as_str())
    };
    value("val") == Some("clear")
        && value("color") == Some("auto")
        && value("fill").is_some_and(|fill| {
            fill.len() == 6 && fill.bytes().all(|byte| byte.is_ascii_hexdigit())
        })
}

/// On/off run properties, which the projection writes only when on.
const TOGGLES: [&str; 14] = [
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "vanish",
    "emboss",
    "imprint",
    "shadow",
    "outline",
    "rtl",
];

fn toggle_on(element: &XmlElement, ns: &Namespaces) -> bool {
    !element.attributes.iter().any(|(key, value)| {
        matches!(ns.resolve(key, true), (Some(W | W_STRICT), "val"))
            && matches!(value.trim(), "0" | "false" | "off")
    })
}

/// Why the children of a `w:rPr` cannot pass through the save projection intact.
fn properties_blocker(element: &XmlElement, ns: &mut Namespaces) -> Option<String> {
    let mark = ns.enter(element);
    let mut found = None;
    for child in element.child_elements() {
        let child_mark = ns.enter(child);
        let local = ns.w_local(child);
        let entry = local.and_then(|local| RUN_PROPERTIES.iter().find(|(name, _)| *name == local));
        let problem = match (local, entry) {
            (Some(local), Some((_, attributes))) => {
                let bad = child.attributes.iter().find(|(key, _)| {
                    if key.as_str() == "xmlns" || key.starts_with("xmlns:") {
                        return false;
                    }
                    match ns.resolve(key, true) {
                        (Some(W | W_STRICT), name) => !attributes.contains(&name),
                        _ => true,
                    }
                });
                if let Some((key, _)) = bad {
                    Some(format!("w:{local} {key}"))
                } else if local == "shd" && !plain_fill(child, ns) {
                    Some("w:shd other than a clear fill".to_owned())
                } else if TOGGLES.contains(&local) && !toggle_on(child, ns) {
                    Some(format!("w:{local} turned off"))
                } else {
                    None
                }
            }
            (Some(local), None) => Some(format!("w:{local}")),
            (None, _) => Some(child.name.clone()),
        };
        ns.leave(child_mark);
        if problem.is_some() {
            found = problem;
            break;
        }
    }
    ns.leave(mark);
    found
}

/// Reads a body `w:p` element.
pub(crate) fn read_paragraph(element: &XmlElement, ns: &mut Namespaces) -> SourceParagraph {
    let mark = ns.enter(element);
    let mut paragraph = SourceParagraph {
        ppr: String::new(),
        items: Vec::new(),
        text: String::new(),
        units: Vec::new(),
        para_id: None,
        blocker: None,
    };
    let block = |paragraph: &mut SourceParagraph, code, message: String| {
        if paragraph.blocker.is_none() {
            paragraph.blocker = Some(Blocker { code, message });
        }
    };
    paragraph.para_id = element
        .attributes
        .iter()
        .find(|(key, _)| ns.resolve(key, true) == (Some(W14), "paraId"))
        .map(|(_, value)| value.clone());
    let empty: Rc<str> = Rc::from("");
    for child in element.child_elements() {
        let child_mark = ns.enter(child);
        let local = ns.w_local(child).map(str::to_owned);
        match local.as_deref() {
            Some("pPr") => paragraph.ppr = canonical_string(child, ns),
            Some("r") => read_run(child, ns, &mut paragraph, &empty),
            name if is_bookkeeping_element(name) => {}
            _ => {
                let name = local.unwrap_or_else(|| child.local_name().to_owned());
                block(
                    &mut paragraph,
                    CompareDiagnosticCode::UnsupportedContent,
                    format!("the paragraph contains {}", element_label(&name)),
                );
                let canonical = canonical_string(child, ns);
                paragraph.items.push(Item::Other { name, canonical });
            }
        }
        ns.leave(child_mark);
    }
    ns.leave(mark);
    paragraph
}

fn read_run(
    run: &XmlElement,
    ns: &mut Namespaces,
    paragraph: &mut SourceParagraph,
    empty: &Rc<str>,
) {
    let mut rpr = Rc::clone(empty);
    for (key, _) in &run.attributes {
        if key == "xmlns" || key.starts_with("xmlns:") {
            continue;
        }
        let (uri, local) = ns.resolve(key, true);
        if !is_bookkeeping_attribute(uri, local, Some("r")) && paragraph.blocker.is_none() {
            paragraph.blocker = Some(Blocker {
                code: CompareDiagnosticCode::UnsupportedContent,
                message: format!("a run carries attribute {key}"),
            });
        }
    }
    for child in run.child_elements() {
        let mark = ns.enter(child);
        if ns.w_local(child) == Some("rPr") {
            rpr = Rc::from(canonical_string(child, ns));
            if let Some(problem) = properties_blocker(child, ns)
                && paragraph.blocker.is_none()
            {
                paragraph.blocker = Some(Blocker {
                    code: CompareDiagnosticCode::UnsupportedFormatting,
                    message: format!("run property {problem} is not kept by edits"),
                });
            }
        }
        ns.leave(mark);
    }
    for child in run.child_elements() {
        let mark = ns.enter(child);
        let local = ns.w_local(child).map(str::to_owned);
        match local.as_deref() {
            Some("rPr") => {}
            Some("t") => {
                let text = run_text(child);
                for ch in text.chars() {
                    for _ in 0..ch.len_utf16() {
                        paragraph.units.push(Unit {
                            rpr: Rc::clone(&rpr),
                            atom: None,
                        });
                    }
                }
                paragraph.text.push_str(&text);
                paragraph.items.push(Item::Text {
                    text,
                    rpr: Rc::clone(&rpr),
                });
            }
            Some(name @ ("tab" | "softHyphen" | "noBreakHyphen")) => {
                let ch = match name {
                    "tab" => '\t',
                    "softHyphen" => '\u{00ad}',
                    _ => '\u{2011}',
                };
                paragraph.units.push(Unit {
                    rpr: Rc::clone(&rpr),
                    atom: None,
                });
                paragraph.text.push(ch);
                paragraph.items.push(Item::Text {
                    text: ch.to_string(),
                    rpr: Rc::clone(&rpr),
                });
            }
            Some("br") if line_break(child, ns) => {
                paragraph.units.push(Unit {
                    rpr: Rc::clone(&rpr),
                    atom: Some(AtomKind::LineBreak),
                });
                paragraph.text.push('\u{FFFC}');
                paragraph.items.push(Item::Atom {
                    kind: "br".to_owned(),
                    rpr: Rc::clone(&rpr),
                });
            }
            name if is_bookkeeping_element(name) => {}
            _ => {
                let name = local.unwrap_or_else(|| child.local_name().to_owned());
                if paragraph.blocker.is_none() {
                    paragraph.blocker = Some(Blocker {
                        code: CompareDiagnosticCode::UnsupportedContent,
                        message: format!("the paragraph contains {}", element_label(&name)),
                    });
                }
                let mut canonical = String::from(&*rpr);
                canonical.push_str(&canonical_string(child, ns));
                paragraph.items.push(Item::Other { name, canonical });
            }
        }
        ns.leave(mark);
    }
}

/// A plain line break: no type other than text wrapping and no clear.
fn line_break(element: &XmlElement, ns: &Namespaces) -> bool {
    element.attributes.iter().all(|(key, value)| {
        key == "xmlns"
            || key.starts_with("xmlns:")
            || matches!(
                ns.resolve(key, true),
                (Some(W | W_STRICT), "type") if value == "textWrapping"
            )
    })
}

/// A `w:t`'s text as the parser reads it: line feeds and carriage returns become spaces.
pub(crate) fn run_text(element: &XmlElement) -> String {
    let mut raw = String::new();
    for child in &element.children {
        if let XmlNode::Text(text) = child {
            raw.push_str(text);
        }
    }
    let mut output = String::with_capacity(raw.len());
    let mut characters = raw.chars().peekable();
    while let Some(ch) = characters.next() {
        if ch == '\r' && characters.peek() == Some(&'\n') {
            characters.next();
        }
        output.push(if matches!(ch, '\r' | '\n') { ' ' } else { ch });
    }
    output
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> docx_parse::XmlDocument {
        let limits = docx_parse::ParseLimits::default();
        docx_parse::parse_xml(
            xml.as_bytes(),
            "test.xml",
            &mut docx_parse::ParseBudget::new(&limits),
        )
        .unwrap()
    }

    const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;

    fn paragraph(body: &str) -> SourceParagraph {
        let document = parse(&format!("<w:p {NS}>{body}</w:p>"));
        read_paragraph(document.root().unwrap(), &mut Namespaces::default())
    }

    #[test]
    fn canonical_forms_ignore_prefixes_bookkeeping_and_run_boundaries() {
        let left = paragraph(
            r#"<w:r w:rsidR="00AB"><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Hel</w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:b/></w:rPr><w:t>lo</w:t></w:r>"#,
        );
        let document = parse(
            r#"<x:p xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="0000AAAA"><x:r><x:rPr><x:b/></x:rPr><x:t>Hello</x:t></x:r></x:p>"#,
        );
        let right = read_paragraph(document.root().unwrap(), &mut Namespaces::default());
        assert!(left.same_content(&right));
        assert_eq!(right.para_id.as_deref(), Some("0000AAAA"));
        assert_eq!(left.text, "Hello");
        let italic = paragraph(r#"<w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r>"#);
        assert!(!left.same_content(&italic));
        assert_eq!(
            left.difference(&italic).0,
            CompareDiagnosticCode::FormattingChange
        );
    }

    #[test]
    fn paragraphs_project_text_like_the_session() {
        let read = paragraph(
            "<w:r><w:t>a\u{1F600}</w:t><w:tab/><w:softHyphen/><w:noBreakHyphen/><w:br/><w:t>b\r\nc</w:t></w:r>",
        );
        assert_eq!(read.text, "a\u{1F600}\t\u{00ad}\u{2011}\u{FFFC}b c");
        assert_eq!(read.units.len(), read.text.encode_utf16().count());
        assert!(read.blocker.is_none());
        assert_eq!(read.units[6].atom, Some(AtomKind::LineBreak));
    }

    #[test]
    fn unsupported_content_and_formatting_block_edits() {
        let blocked = |body: &str| paragraph(body).blocker.map(|blocker| blocker.code);
        assert_eq!(
            blocked(r#"<w:hyperlink w:anchor="x"><w:r><w:t>a</w:t></w:r></w:hyperlink>"#),
            Some(CompareDiagnosticCode::UnsupportedContent)
        );
        assert_eq!(
            blocked(r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#),
            Some(CompareDiagnosticCode::UnsupportedContent)
        );
        assert_eq!(
            blocked(r#"<w:r><w:br w:type="page"/></w:r>"#),
            Some(CompareDiagnosticCode::UnsupportedContent)
        );
        assert_eq!(
            blocked(r#"<w:bookmarkStart w:id="1" w:name="x"/><w:r><w:t>a</w:t></w:r>"#),
            Some(CompareDiagnosticCode::UnsupportedContent)
        );
        assert_eq!(
            blocked(r#"<w:r><w:rPr><w:bdr w:val="single"/></w:rPr><w:t>a</w:t></w:r>"#),
            Some(CompareDiagnosticCode::UnsupportedFormatting)
        );
        assert_eq!(
            blocked(r#"<w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>a</w:t></w:r>"#),
            Some(CompareDiagnosticCode::UnsupportedFormatting)
        );
        assert_eq!(
            blocked(r#"<w:r><w:rPr><w:bCs/><w:iCs/></w:rPr><w:t>a</w:t></w:r>"#),
            None
        );
        assert_eq!(
            blocked(r#"<w:pPr><w:textDirection w:val="tbRl"/></w:pPr><w:r><w:t>a</w:t></w:r>"#),
            None
        );
        assert_eq!(
            blocked(
                r#"<w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:after="120"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:color w:val="FF0000"/></w:rPr><w:t>a</w:t></w:r>"#
            ),
            None
        );
    }

    #[test]
    fn skeletons_keep_wrappers_and_mark_blocks() {
        let wrapped = |element: &str| {
            parse(&format!(
                r#"<w:body {NS}><w:customXml w:element="{element}"><w:p/></w:customXml><w:p/></w:body>"#
            ))
        };
        let holes = HashSet::from([vec![0, 0], vec![1]]);
        let skeleton = |document: &docx_parse::XmlDocument| {
            canonical_skeleton(document.root().unwrap(), &mut Namespaces::default(), &holes)
        };
        assert_eq!(skeleton(&wrapped("a")), skeleton(&wrapped("a")));
        assert_ne!(skeleton(&wrapped("a")), skeleton(&wrapped("b")));
        assert_eq!(skeleton(&wrapped("a")).matches("<#block/>").count(), 2);
        assert!(is_revision("cellIns") && !is_revision("p"));
    }
}
