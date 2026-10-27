use std::collections::HashSet;

use crate::comment_patch::{Span, spans};
use crate::xml::{
    ParseBudget, ParseLimits, XmlElement, XmlNode, parse_xml, serialize_xml_fragment,
};

pub(crate) fn splice_slide(
    source: &[u8],
    before: &XmlElement,
    after: &XmlElement,
) -> Option<Vec<u8>> {
    if before == after {
        return Some(source.to_vec());
    }
    if before.name != after.name {
        return None;
    }
    let span = spans(source, "slide").ok()?;
    if span.children.len() != before.child_elements().count() {
        return None;
    }
    let mut root_prefixes = HashSet::new();
    for key in before.attributes.keys() {
        if let Some(prefix) = key.strip_prefix("xmlns:") {
            root_prefixes.insert(prefix.to_owned());
        }
    }
    root_prefixes.insert("xml".to_owned());
    let splicer = Splicer {
        source,
        root_prefixes,
    };
    let range_start = span.range.start;
    let range_end = span.range.end;
    let body = splicer.splice_strict(before, after, &span)?;
    let mut output = Vec::with_capacity(source.len());
    output.extend_from_slice(&source[..range_start]);
    output.extend_from_slice(&body);
    output.extend_from_slice(&source[range_end..]);
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let reparsed = parse_xml(&output, "slide", &mut budget).ok()?;
    if !equivalent(&reparsed, after) {
        return None;
    }
    Some(output)
}

struct Splicer<'a> {
    source: &'a [u8],
    root_prefixes: HashSet<String>,
}

enum Op {
    Match(usize, usize),
    Delete(usize),
    Insert(usize),
}

impl Splicer<'_> {
    fn splice_or_fragment(
        &self,
        before: &XmlElement,
        after: &XmlElement,
        span: &Span,
    ) -> Option<Vec<u8>> {
        match self.splice_strict(before, after, span) {
            Some(bytes) => Some(bytes),
            None => self.checked_fragment(after),
        }
    }

    fn splice_strict(
        &self,
        before: &XmlElement,
        after: &XmlElement,
        span: &Span,
    ) -> Option<Vec<u8>> {
        if before == after {
            return Some(self.source[span.range.clone()].to_vec());
        }
        if before.name != after.name {
            return None;
        }
        if span.children.len() != before.child_elements().count() {
            return None;
        }
        let inner = self.splice_children(before, after, span)?;
        let opening = &self.source[span.range.start..span.opening_end];
        let mut start =
            crate::attribute_patch::patch_start(opening, &before.attributes, &after.attributes)?;
        let self_closing = opening.ends_with(b"/>");
        if after.children.is_empty() {
            if self_closing {
                return Some(start);
            }
            let mut output = start;
            output.extend_from_slice(&inner);
            output.extend_from_slice(&self.source[span.closing_start..span.range.end]);
            return Some(output);
        }
        if self_closing {
            if !start.ends_with(b"/>") {
                return None;
            }
            start.truncate(start.len() - 2);
            start.push(b'>');
            let mut output = start;
            output.extend_from_slice(&inner);
            output.extend_from_slice(format!("</{}>", after.name).as_bytes());
            Some(output)
        } else {
            let mut output = start;
            output.extend_from_slice(&inner);
            output.extend_from_slice(&self.source[span.closing_start..span.range.end]);
            Some(output)
        }
    }

    fn splice_children(
        &self,
        before: &XmlElement,
        after: &XmlElement,
        span: &Span,
    ) -> Option<Vec<u8>> {
        if before.children == after.children {
            return Some(self.source[span.opening_end..span.closing_start].to_vec());
        }
        let before_texts: Vec<&str> = before
            .children
            .iter()
            .filter_map(|node| match node {
                XmlNode::Text(text) => Some(text.as_str()),
                XmlNode::Element(_) => None,
            })
            .collect();
        let after_texts: Vec<&str> = after
            .children
            .iter()
            .filter_map(|node| match node {
                XmlNode::Text(text) => Some(text.as_str()),
                XmlNode::Element(_) => None,
            })
            .collect();
        if before_texts != after_texts {
            return None;
        }
        let meaningful = before
            .children
            .iter()
            .chain(after.children.iter())
            .any(|node| matches!(node, XmlNode::Text(text) if !text.trim().is_empty()));
        if meaningful {
            return None;
        }
        let before_elems: Vec<&XmlElement> = before.child_elements().collect();
        let after_elems: Vec<&XmlElement> = after.child_elements().collect();
        if span.children.len() != before_elems.len() {
            return None;
        }
        if before_elems
            .len()
            .checked_mul(after_elems.len())
            .is_none_or(|size| size > 1_000_000)
        {
            return None;
        }
        let ops = align(&before_elems, &after_elems);
        let deleted: Vec<usize> = ops
            .iter()
            .filter_map(|op| match op {
                Op::Delete(index) => Some(*index),
                _ => None,
            })
            .collect();
        let mut reused: HashSet<usize> = HashSet::new();
        let mut output = Vec::new();
        let mut cursor = span.opening_end;
        for op in &ops {
            match *op {
                Op::Match(before_index, after_index) => {
                    let child_span = &span.children[before_index];
                    output.extend_from_slice(&self.source[cursor..child_span.range.start]);
                    output.extend_from_slice(&self.splice_or_fragment(
                        before_elems[before_index],
                        after_elems[after_index],
                        child_span,
                    )?);
                    cursor = child_span.range.end;
                }
                Op::Delete(before_index) => {
                    cursor = span.children[before_index].range.end;
                }
                Op::Insert(after_index) => {
                    let patched = after_elems[after_index];
                    let reused_index = deleted
                        .iter()
                        .copied()
                        .find(|index| !reused.contains(index) && before_elems[*index] == patched);
                    if let Some(index) = reused_index {
                        reused.insert(index);
                        let child_span = &span.children[index];
                        output.extend_from_slice(&self.source[child_span.range.clone()]);
                    } else {
                        output.extend_from_slice(&self.checked_fragment(patched)?);
                    }
                }
            }
        }
        output.extend_from_slice(&self.source[cursor..span.closing_start]);
        Some(output)
    }

    fn checked_fragment(&self, element: &XmlElement) -> Option<Vec<u8>> {
        let mut declared = self.root_prefixes.clone();
        collect_declared(element, &mut declared);
        let mut used = HashSet::new();
        collect_used(element, &mut used);
        if used.iter().all(|prefix| declared.contains(prefix)) {
            Some(serialize_xml_fragment(element))
        } else {
            None
        }
    }
}

fn align(before: &[&XmlElement], after: &[&XmlElement]) -> Vec<Op> {
    let n = before.len();
    let m = after.len();
    let mut dp = vec![vec![0i32; m + 1]; n + 1];
    for (i, row) in dp.iter_mut().enumerate().take(n) {
        row[m] = -((n - i) as i32);
    }
    for (j, cell) in dp[n].iter_mut().enumerate().take(m) {
        *cell = -((m - j) as i32);
    }
    for i in (0..n).rev() {
        for j in (0..m).rev() {
            let mut best = i32::MIN / 2;
            if before[i].name == after[j].name {
                let score = if before[i] == after[j] { 100 } else { 10 };
                best = best.max(score + dp[i + 1][j + 1]);
            }
            best = best.max(-1 + dp[i + 1][j]);
            best = best.max(-1 + dp[i][j + 1]);
            dp[i][j] = best;
        }
    }
    let mut ops = Vec::new();
    let mut i = 0;
    let mut j = 0;
    while i < n || j < m {
        if i < n && j < m && before[i].name == after[j].name {
            let score = if before[i] == after[j] { 100 } else { 10 };
            if dp[i][j] == score + dp[i + 1][j + 1] {
                ops.push(Op::Match(i, j));
                i += 1;
                j += 1;
                continue;
            }
        }
        if i < n && dp[i][j] == -1 + dp[i + 1][j] {
            ops.push(Op::Delete(i));
            i += 1;
            continue;
        }
        ops.push(Op::Insert(j));
        j += 1;
    }
    ops
}

fn collect_declared(element: &XmlElement, declared: &mut HashSet<String>) {
    for key in element.attributes.keys() {
        if let Some(prefix) = key.strip_prefix("xmlns:") {
            declared.insert(prefix.to_owned());
        }
    }
    for child in element.child_elements() {
        collect_declared(child, declared);
    }
}

fn collect_used(element: &XmlElement, used: &mut HashSet<String>) {
    if let Some((prefix, _)) = element.name.split_once(':') {
        used.insert(prefix.to_owned());
    }
    for key in element.attributes.keys() {
        if key == "xmlns" || key.starts_with("xmlns:") {
            continue;
        }
        if let Some((prefix, _)) = key.split_once(':') {
            used.insert(prefix.to_owned());
        }
    }
    for child in element.child_elements() {
        collect_used(child, used);
    }
}

fn equivalent(before: &XmlElement, after: &XmlElement) -> bool {
    if before.name != after.name || before.attributes != after.attributes {
        return false;
    }
    if before.children.len() != after.children.len() {
        return false;
    }
    before
        .children
        .iter()
        .zip(after.children.iter())
        .all(|(left, right)| match (left, right) {
            (XmlNode::Element(left), XmlNode::Element(right)) => equivalent(left, right),
            (XmlNode::Text(left), XmlNode::Text(right)) => left == right,
            _ => false,
        })
}
