//! What the source package says about content controls before parsing and seeding lose any of
//! it: whether a control's content is only text a fill may replace, and the controls the session
//! does not hold at all (inside raw XML blocks and comment bodies).

use docx_parse::XmlElement;

use crate::read_types::Anchor;

/// Element names one safety record keeps.
const MAX_UNSUPPORTED: usize = 8;

/// What a control's source content holds beyond text a fill may replace.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub(crate) struct ControlSafety {
    /// Qualified names of source elements a text fill would destroy.
    pub unsupported: Vec<String>,
    /// The content carries tracked revisions.
    pub revisions: bool,
    /// The content holds something without a faithful text projection.
    pub non_text: bool,
    /// The lock or data binding the parser reads differs from what the namespace-resolved XML
    /// says, so the control's policy cannot be classified reliably.
    pub uncertain: bool,
}

impl ControlSafety {
    /// Whether replacing the content destroys nothing but text.
    pub fn safe(&self) -> bool {
        self.unsupported.is_empty() && !self.revisions
    }

    pub(crate) fn merge(&mut self, other: ControlSafety) {
        for name in other.unsupported {
            self.note(name);
        }
        self.revisions |= other.revisions;
        self.non_text |= other.non_text;
        self.uncertain |= other.uncertain;
    }

    fn note(&mut self, name: impl Into<String>) {
        let name = name.into();
        if self.unsupported.len() < MAX_UNSUPPORTED && !self.unsupported.contains(&name) {
            self.unsupported.push(name);
        }
    }

    fn unsupported(&mut self, element: &XmlElement) {
        self.note(element.name.clone());
    }
}

/// The key a control's safety is recorded under: its captured `w:sdtPr` with any
/// `w:showingPlcHdr` removed, so filling the control keeps finding it.
pub(crate) fn safety_key(raw_properties: Option<&str>) -> String {
    let Some(raw) = raw_properties else {
        return String::new();
    };
    docx_parse::clear_showing_placeholder_xml(raw)
        .ok()
        .flatten()
        .unwrap_or_else(|| raw.to_owned())
}

const W_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// In-scope namespace declarations, innermost last.
type Scope<'a> = Vec<(&'a str, &'a str)>;

fn declare<'a>(element: &'a XmlElement, scope: &mut Scope<'a>) -> usize {
    let before = scope.len();
    for (name, value) in &element.attributes {
        if name == "xmlns" {
            scope.push(("", value));
        } else if let Some(prefix) = name.strip_prefix("xmlns:") {
            scope.push((prefix, value));
        }
    }
    before
}

fn resolve<'a>(scope: &Scope<'a>, prefix: &str) -> Option<&'a str> {
    scope
        .iter()
        .rev()
        .find(|(bound, _)| *bound == prefix)
        .map(|(_, uri)| *uri)
}

/// Whether `element` is the WordprocessingML element `local` in `scope`.
fn is_w(element: &XmlElement, scope: &Scope<'_>, local: &str) -> bool {
    element.local_name() == local
        && resolve(scope, element.namespace_prefix().unwrap_or("")) == Some(W_NAMESPACE)
}

/// The WordprocessingML attribute `local` of `element`, resolved through `scope`.
fn w_attribute<'a>(element: &'a XmlElement, scope: &Scope<'_>, local: &str) -> Option<&'a str> {
    element.attributes.iter().find_map(|(name, value)| {
        let (prefix, name) = name.split_once(':')?;
        (name == local && prefix != "xmlns" && resolve(scope, prefix) == Some(W_NAMESPACE))
            .then_some(value.as_str())
    })
}

/// Whether the lock and data binding of `sdt`, resolved by namespace in `scope`, match what the
/// parser reads from it.
fn policy_agrees<'a>(sdt: &'a XmlElement, scope: &mut Scope<'a>) -> bool {
    let parsed = docx_parse::parse_sdt_properties(sdt.child_by_local_name("sdtPr"), None, None);
    let mut lock = None;
    let mut bound = false;
    for properties in sdt.child_elements() {
        let mark = declare(properties, scope);
        if is_w(properties, scope, "sdtPr") {
            for child in properties.child_elements() {
                let inner = declare(child, scope);
                if is_w(child, scope, "lock") {
                    lock = Some(
                        w_attribute(child, scope, "val")
                            .unwrap_or("unlocked")
                            .to_owned(),
                    );
                } else if is_w(child, scope, "dataBinding") {
                    bound = true;
                }
                scope.truncate(inner);
            }
        }
        scope.truncate(mark);
    }
    parsed.lock == lock && parsed.data_binding.is_some() == bound
}

/// A WordprocessingML `w:sdt` of a part.
pub(crate) struct Classified<'a> {
    pub element: &'a XmlElement,
    /// Element-child ordinals from the part's root.
    pub path: Vec<u32>,
    /// The `w14:paraId` of the paragraph holding it, or of a block control's first paragraph.
    pub paragraph: Option<&'a str>,
    pub safety: ControlSafety,
}

/// Classifies every WordprocessingML `w:sdt` at or below `element`, which sits at `path` inside
/// `paragraph`, resolving namespaces the way a consumer of the part would.
pub(crate) fn classify_controls<'a>(
    element: &'a XmlElement,
    scope: &mut Scope<'a>,
    path: &mut Vec<u32>,
    paragraph: Option<&'a str>,
    output: &mut Vec<Classified<'a>>,
) {
    let mark = declare(element, scope);
    let paragraph = if is_w(element, scope, "p") {
        element.attribute(Some("w14"), "paraId")
    } else {
        paragraph
    };
    if is_w(element, scope, "sdt") {
        let mut safety = classify(element);
        safety.uncertain = !policy_agrees(element, &mut scope.clone());
        output.push(Classified {
            element,
            path: path.clone(),
            paragraph: paragraph.or_else(|| {
                element
                    .child_by_local_name("sdtContent")?
                    .child_elements()
                    .find(|child| child.local_name() == "p")?
                    .attribute(Some("w14"), "paraId")
            }),
            safety,
        });
    }
    for (index, child) in element.child_elements().enumerate() {
        path.push(index as u32);
        classify_controls(child, scope, path, paragraph, output);
        path.pop();
    }
    scope.truncate(mark);
}

/// Classifies the `w:sdtContent` of a `w:sdt` element.
pub(crate) fn classify(sdt: &XmlElement) -> ControlSafety {
    let mut safety = ControlSafety::default();
    if let Some(content) = sdt.child_by_local_name("sdtContent") {
        for child in content.child_elements() {
            match child.local_name() {
                "r" => classify_run(child, &mut safety),
                "p" => classify_paragraph(child, &mut safety),
                "proofErr" => {}
                "ins" | "del" | "moveFrom" | "moveTo" => {
                    safety.revisions = true;
                    safety.unsupported(child);
                }
                "fldSimple" | "oMath" | "oMathPara" | "tbl" | "sdt" => {
                    safety.non_text = true;
                    safety.unsupported(child);
                }
                _ => safety.unsupported(child),
            }
        }
    }
    safety
}

fn classify_paragraph(paragraph: &XmlElement, safety: &mut ControlSafety) {
    for child in paragraph.child_elements() {
        match child.local_name() {
            "pPr" => {
                for property in child.child_elements() {
                    match property.local_name() {
                        "sectPr" => safety.unsupported(property),
                        "pPrChange" => {
                            safety.revisions = true;
                            safety.unsupported(property);
                        }
                        "rPr" => {
                            if property.child_elements().any(|mark| {
                                matches!(
                                    mark.local_name(),
                                    "ins" | "del" | "moveFrom" | "moveTo" | "rPrChange"
                                )
                            }) {
                                safety.revisions = true;
                                safety.note("w:pPr/w:rPr");
                            }
                        }
                        _ => {}
                    }
                }
            }
            "r" => classify_run(child, safety),
            "proofErr" => {}
            "ins" | "del" | "moveFrom" | "moveTo" => {
                safety.revisions = true;
                safety.unsupported(child);
            }
            "fldSimple" | "oMath" | "oMathPara" | "sdt" => {
                safety.non_text = true;
                safety.unsupported(child);
            }
            _ => safety.unsupported(child),
        }
    }
}

fn classify_run(run: &XmlElement, safety: &mut ControlSafety) {
    for child in run.child_elements() {
        match child.local_name() {
            "t"
            | "tab"
            | "cr"
            | "softHyphen"
            | "noBreakHyphen"
            | "sym"
            | "lastRenderedPageBreak" => {}
            "rPr" => {
                if child
                    .child_elements()
                    .any(|mark| mark.local_name() == "rPrChange")
                {
                    safety.revisions = true;
                    safety.note("w:rPrChange");
                }
            }
            "br" => {
                if child
                    .attribute(Some("w"), "type")
                    .is_some_and(|kind| kind != "textWrapping")
                {
                    safety.non_text = true;
                    safety.unsupported(child);
                }
            }
            "delText" | "delInstrText" => {
                safety.revisions = true;
                safety.unsupported(child);
            }
            _ => {
                safety.non_text = true;
                safety.unsupported(child);
            }
        }
    }
}

/// A content control only the source package holds.
pub(crate) struct SourceControl {
    /// The story it belongs to: the story of its raw block, or `comment:{id}`.
    pub story: String,
    /// The raw block of `story` holding it, if any.
    pub raw_block: Option<usize>,
    pub anchor: Anchor,
    /// The nearest source control containing it, as an index into the same list.
    pub parent: Option<usize>,
    pub block: bool,
    pub properties: docx_parse::SdtProperties,
}

/// Where scanning a source element for controls records them.
pub(crate) struct ControlScan<'a> {
    pub story: &'a str,
    pub raw_block: Option<usize>,
    /// The part and its digest the element was located in; `None` anchors controls to the raw
    /// block of `story` instead.
    pub part: Option<(&'a str, &'a str)>,
}

const BLOCK_PARENTS: [&str; 9] = [
    "body",
    "comment",
    "tc",
    "txbxContent",
    "hdr",
    "ftr",
    "footnote",
    "endnote",
    "docPartBody",
];

/// Records every `w:sdt` at or below `element` (reached through `path`), outermost first.
pub(crate) fn scan_controls(
    element: &XmlElement,
    parent_name: &str,
    path: &mut Vec<u32>,
    scan: &ControlScan<'_>,
    parent: Option<usize>,
    output: &mut Vec<SourceControl>,
) {
    let mut inner = parent;
    if element.local_name() == "sdt" {
        let content = element.child_by_local_name("sdtContent");
        let block = content.is_some_and(|content| {
            content
                .child_elements()
                .any(|child| matches!(child.local_name(), "p" | "tbl" | "tr" | "tc"))
        }) || BLOCK_PARENTS.contains(&parent_name);
        let properties = element.child_by_local_name("sdtPr");
        output.push(SourceControl {
            story: scan.story.to_owned(),
            raw_block: scan.raw_block,
            anchor: match scan.part {
                Some((part, part_sha256)) => Anchor::SourcePart {
                    part: part.to_owned(),
                    part_sha256: part_sha256.to_owned(),
                    path: path.clone(),
                },
                None => Anchor::Control {
                    story: scan.story.to_owned(),
                    control_id: format!(
                        "{}|raw{}#{}",
                        scan.story,
                        scan.raw_block.unwrap_or_default(),
                        path.iter()
                            .map(u32::to_string)
                            .collect::<Vec<_>>()
                            .join("/")
                    ),
                },
            },
            parent,
            block,
            properties: docx_parse::parse_sdt_properties(properties, None, None),
        });
        inner = Some(output.len() - 1);
    }
    for (index, child) in element.child_elements().enumerate() {
        path.push(index as u32);
        scan_controls(child, element.local_name(), path, scan, inner, output);
        path.pop();
    }
}

/// Whether raw XML text may hold a content control.
pub(crate) fn may_hold_controls(xml: &str) -> bool {
    xml.contains("sdtContent")
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sdt(content: &str) -> XmlElement {
        let xml = format!(
            "<w:sdt xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:sdtPr><w:tag w:val=\"t\"/></w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>"
        );
        let limits = docx_parse::ParseLimits::default();
        docx_parse::parse_xml(
            xml.as_bytes(),
            "test.xml",
            &mut docx_parse::ParseBudget::new(&limits),
        )
        .unwrap()
        .root()
        .unwrap()
        .clone()
    }

    #[test]
    fn text_runs_are_safe_and_markup_is_not() {
        assert!(
            classify(&sdt(
                "<w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t><w:tab/><w:br/></w:r><w:proofErr/>"
            ))
            .safe()
        );
        let bookmark = classify(&sdt(
            "<w:bookmarkStart w:id=\"1\" w:name=\"b\"/><w:r><w:t>a</w:t></w:r>",
        ));
        assert_eq!(bookmark.unsupported, vec!["w:bookmarkStart"]);
        assert!(!bookmark.non_text && !bookmark.revisions);
        let tracked = classify(&sdt("<w:ins w:id=\"1\"><w:r><w:t>a</w:t></w:r></w:ins>"));
        assert!(tracked.revisions && !tracked.safe());
        let page = classify(&sdt("<w:r><w:br w:type=\"page\"/></w:r>"));
        assert!(page.non_text && !page.safe());
        let block = classify(&sdt(
            "<w:p><w:pPr><w:pStyle w:val=\"Normal\"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p><w:p/>",
        ));
        assert!(block.safe());
        let table = classify(&sdt("<w:tbl/>"));
        assert!(table.non_text && !table.safe());
    }

    #[test]
    fn policy_is_resolved_by_namespace() {
        let parse = |xml: &str| {
            let limits = docx_parse::ParseLimits::default();
            docx_parse::parse_xml(
                xml.as_bytes(),
                "test.xml",
                &mut docx_parse::ParseBudget::new(&limits),
            )
            .unwrap()
            .root()
            .unwrap()
            .clone()
        };
        let classified = |xml: &str| {
            let root = parse(xml);
            let mut out = Vec::new();
            classify_controls(&root, &mut Vec::new(), &mut Vec::new(), None, &mut out);
            out.into_iter()
                .map(|classified| classified.safety.uncertain)
                .collect::<Vec<_>>()
        };
        let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        assert_eq!(
            classified(&format!(
                "<q:body xmlns:q=\"{w}\"><q:sdt><q:sdtPr><q:lock q:val=\"contentLocked\"/></q:sdtPr><q:sdtContent/></q:sdt></q:body>"
            )),
            [false]
        );
        assert_eq!(
            classified(&format!(
                "<w:body xmlns:w=\"{w}\" xmlns:q=\"{w}\"><w:sdt><w:sdtPr><w:lock q:val=\"contentLocked\"/></w:sdtPr><w:sdtContent/></w:sdt></w:body>"
            )),
            [true]
        );
        assert_eq!(
            classified("<w:body xmlns:w=\"urn:other\"><w:sdt/></w:body>"),
            Vec::<bool>::new()
        );
    }

    #[test]
    fn safety_keys_ignore_the_placeholder_flag() {
        assert_eq!(
            safety_key(Some(
                "<w:sdtPr><w:showingPlcHdr/><w:tag w:val=\"t\"/></w:sdtPr>"
            )),
            safety_key(Some("<w:sdtPr><w:tag w:val=\"t\"/></w:sdtPr>"))
        );
        assert_eq!(safety_key(None), "");
    }
}
