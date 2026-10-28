use std::collections::BTreeMap;

use docx_parse::paragraph::{HexIdAllocator, parse_paragraph};
use docx_parse::xml::{ParseBudget, ParseLimits, XmlElement, XmlNode, parse_xml};
use yrs::{Text, Transact};

use crate::{EditingDoc, story_ref};

#[derive(Default)]
struct Field {
    instruction: String,
    result: bool,
}

pub(crate) fn end_suffixes(bytes: &[u8], path: &str) -> BTreeMap<u32, u32> {
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    budget.record_source_ordinals();
    let Ok(document) = parse_xml(bytes, path, &mut budget) else {
        return BTreeMap::new();
    };
    let Some(root) = document.root() else {
        return BTreeMap::new();
    };
    let mut ends = BTreeMap::new();
    collect(root, None, &mut Vec::new(), &mut ends);
    ends.into_iter()
        .filter_map(|(ordinal, (paragraph, marker))| {
            let mut found = false;
            let suffix = suffix_after(paragraph, marker, &mut found)?;
            let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64)).ok()?;
            let paragraph = parse_paragraph(
                &suffix,
                None,
                None,
                None,
                None,
                None,
                path,
                &mut budget,
                &mut ids,
                None,
                false,
                0,
            )
            .ok()?;
            let blocks = [serde_json::to_value(paragraph).ok()?];
            let scratch = EditingDoc::new(0);
            crate::seed::seed_blocks(&scratch, None, &[("body".to_owned(), &blocks)]).ok()?;
            let txn = scratch.yrs_doc().transact();
            let width = story_ref(&txn, "body").ok()?.len(&txn).checked_sub(1)?;
            Some((ordinal, width))
        })
        .collect()
}

fn collect<'a>(
    node: &'a XmlElement,
    paragraph: Option<&'a XmlElement>,
    fields: &mut Vec<Field>,
    ends: &mut BTreeMap<u32, (&'a XmlElement, &'a XmlElement)>,
) {
    let paragraph = if node.local_name() == "p" {
        Some(node)
    } else {
        paragraph
    };
    match node.local_name() {
        "fldChar" => match node.attribute(Some("w"), "fldCharType") {
            Some("begin") => fields.push(Field::default()),
            Some("separate") => {
                if let Some(field) = fields.last_mut() {
                    field.result = true;
                }
            }
            Some("end") => {
                if let Some(field) = fields.pop()
                    && field.result
                    && field
                        .instruction
                        .split_whitespace()
                        .next()
                        .is_some_and(|kind| kind.eq_ignore_ascii_case("TOC"))
                    && let Some(paragraph) = paragraph
                    && let Some(ordinal) = paragraph.paragraph_ordinal
                {
                    ends.insert(ordinal, (paragraph, node));
                }
            }
            _ => {}
        },
        "instrText" => {
            if let Some(field) = fields.last_mut()
                && !field.result
            {
                field.instruction.push_str(&node.text_content());
            }
        }
        _ => {}
    }
    for child in node.child_elements() {
        collect(child, paragraph, fields, ends);
    }
}

fn suffix_after(node: &XmlElement, marker: &XmlElement, found: &mut bool) -> Option<XmlElement> {
    if std::ptr::eq(node, marker) {
        *found = true;
        return None;
    }
    if *found {
        return Some(node.clone());
    }
    let mut suffix = node.clone();
    suffix.children = node
        .children
        .iter()
        .filter_map(|child| match child {
            XmlNode::Element(element) => suffix_after(element, marker, found).map(XmlNode::Element),
            _ => (*found).then(|| child.clone()),
        })
        .collect();
    (!suffix.children.is_empty() || node.local_name() == "p").then_some(suffix)
}
