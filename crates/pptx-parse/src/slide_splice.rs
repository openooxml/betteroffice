use std::cell::Cell;
use std::collections::{HashMap, HashSet, VecDeque};

use crate::comment_patch::{Span, spans};
use crate::xml::{
    ParseBudget, ParseLimits, XmlElement, XmlNode, parse_xml, serialize_xml_fragment,
};

pub(crate) fn splice_slide(
    source: &[u8],
    before: &XmlElement,
    after: &XmlElement,
) -> Option<Vec<u8>> {
    splice_slide_with_limit(source, before, after, None)
}

fn splice_slide_with_limit(
    source: &[u8],
    before: &XmlElement,
    after: &XmlElement,
    limit: Option<usize>,
) -> Option<Vec<u8>> {
    let mut root_prefixes = HashSet::new();
    for key in before.attributes.keys() {
        if let Some(prefix) = key.strip_prefix("xmlns:") {
            root_prefixes.insert(prefix.to_owned());
        }
    }
    root_prefixes.insert("xml".to_owned());
    let before_weight = weight(before);
    let total_weight = before_weight.saturating_add(weight(after));
    let limit = limit.unwrap_or_else(|| {
        total_weight
            .saturating_mul(SPLICE_WORK_PER_NODE)
            .saturating_add(SPLICE_WORK_SLACK)
    });
    let splicer = Splicer {
        source,
        root_prefixes,
        budget: Cell::new(limit),
        exhausted: Cell::new(false),
    };
    let output = splicer.splice_strict(before, after, before_weight);
    #[cfg(feature = "test-counters")]
    crate::test_counters::bump(|work| {
        work.record_budget_peak(limit - splicer.budget.get(), total_weight);
    });
    if splicer.exhausted.get() {
        return None;
    }
    output
}

const SPLICE_WORK_PER_NODE: usize = 64;
const SPLICE_WORK_SLACK: usize = 4_096;
const BYTES_PER_WEIGHT_UNIT: usize = 64;

struct Splicer<'a> {
    source: &'a [u8],
    root_prefixes: HashSet<String>,
    budget: Cell<usize>,
    exhausted: Cell<bool>,
}

impl Splicer<'_> {
    fn charge(&self, work: usize) -> Option<()> {
        #[cfg(feature = "test-counters")]
        crate::test_counters::bump(|counted| {
            counted.budget_work = counted.budget_work.saturating_add(work);
        });
        let Some(left) = self.budget.get().checked_sub(work) else {
            self.exhausted.set(true);
            #[cfg(feature = "test-counters")]
            crate::test_counters::bump(|counted| {
                counted.budget_exhausted = counted.budget_exhausted.saturating_add(1);
            });
            return None;
        };
        self.budget.set(left);
        Some(())
    }

    fn classes<'e>(
        &self,
        elements: [&[&'e XmlElement]; 2],
        hashes: [&[(u64, usize)]; 2],
    ) -> Option<[Vec<usize>; 2]> {
        let mut representatives: HashMap<u64, Vec<(usize, &'e XmlElement)>> = HashMap::new();
        let mut next = 0;
        let mut assign = |element: &'e XmlElement, (hash, weight): (u64, usize)| -> Option<usize> {
            let bucket = representatives.entry(hash).or_default();
            for &(class, representative) in bucket.iter() {
                self.charge(weight)?;
                if representative == element {
                    return Some(class);
                }
            }
            bucket.push((next, element));
            next += 1;
            Some(next - 1)
        };
        let mut classes = [Vec::new(), Vec::new()];
        for ((side, elements), hashes) in classes.iter_mut().zip(elements).zip(hashes) {
            for (element, hash) in elements.iter().copied().zip(hashes.iter().copied()) {
                side.push(assign(element, hash)?);
            }
        }
        Some(classes)
    }
}

enum Op {
    Match(usize, usize),
    Delete(usize),
    Insert(usize),
}

impl Splicer<'_> {
    fn splice_strict(
        &self,
        before: &XmlElement,
        after: &XmlElement,
        before_weight: usize,
    ) -> Option<Vec<u8>> {
        self.charge(before_weight)?;
        if before == after {
            return Some(self.source.to_vec());
        }
        let span = spans(self.source, "slide").ok()?;
        let body = self.splice_unequal(before, after, &span)?;
        let mut output = Vec::with_capacity(self.source.len());
        output.extend_from_slice(&self.source[..span.range.start]);
        output.extend_from_slice(&body);
        output.extend_from_slice(&self.source[span.range.end..]);
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let reparsed = parse_xml(&output, "slide", &mut budget).ok()?;
        if !equivalent(&reparsed, after) {
            return None;
        }
        Some(output)
    }

    fn splice_or_fragment(
        &self,
        before: &XmlElement,
        after: &XmlElement,
        span: &Span,
    ) -> Option<Vec<u8>> {
        match self.splice_unequal(before, after, span) {
            Some(bytes) => Some(bytes),
            None if self.exhausted.get() => None,
            None => self.checked_fragment(after),
        }
    }

    fn splice_unequal(
        &self,
        before: &XmlElement,
        after: &XmlElement,
        span: &Span,
    ) -> Option<Vec<u8>> {
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
        if before.children.len() == after.children.len() {
            self.charge(children_weight(before, 0))?;
            if before.children == after.children {
                return Some(self.source[span.opening_end..span.closing_start].to_vec());
            }
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
        let before_hashes: Vec<(u64, usize)> = before_elems
            .iter()
            .map(|element| structural_hash(element))
            .collect();
        let after_hashes: Vec<(u64, usize)> = after_elems
            .iter()
            .map(|element| structural_hash(element))
            .collect();
        self.charge(
            before_hashes
                .iter()
                .chain(&after_hashes)
                .map(|(_, weight)| weight)
                .sum(),
        )?;
        let [before_classes, after_classes] = self.classes(
            [&before_elems, &after_elems],
            [&before_hashes, &after_hashes],
        )?;
        let ops = align(&before_elems, &after_elems, &before_classes, &after_classes)?;
        let after_texts: HashSet<&str> = after_texts.into_iter().collect();
        let mut deleted: HashMap<usize, VecDeque<usize>> = HashMap::new();
        for op in &ops {
            if let Op::Delete(index) = *op {
                deleted
                    .entry(before_classes[index])
                    .or_default()
                    .push_back(index);
            }
        }
        let mut output = Vec::new();
        let mut cursor = span.opening_end;
        for op in &ops {
            match *op {
                Op::Match(before_index, after_index) => {
                    let child_span = &span.children[before_index];
                    output.extend_from_slice(&self.source[cursor..child_span.range.start]);
                    if before_classes[before_index] == after_classes[after_index] {
                        output.extend_from_slice(&self.source[child_span.range.clone()]);
                    } else {
                        output.extend_from_slice(&self.splice_or_fragment(
                            before_elems[before_index],
                            after_elems[after_index],
                            child_span,
                        )?);
                    }
                    cursor = child_span.range.end;
                }
                Op::Delete(before_index) => {
                    let child_span = &span.children[before_index];
                    let gap = &self.source[cursor..child_span.range.start];
                    let gap_text = std::str::from_utf8(gap).unwrap_or("");
                    let whitespace_only = !gap.is_empty() && gap_text.trim().is_empty();
                    if !whitespace_only || after_texts.contains(gap_text) {
                        output.extend_from_slice(gap);
                    }
                    cursor = child_span.range.end;
                }
                Op::Insert(after_index) => {
                    let patched = after_elems[after_index];
                    let reused_index = deleted
                        .get_mut(&after_classes[after_index])
                        .and_then(VecDeque::pop_front);
                    if let Some(index) = reused_index {
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
            #[cfg(feature = "test-counters")]
            crate::test_counters::bump(|work| {
                work.fragment_elements = work
                    .fragment_elements
                    .saturating_add(element_count(element));
            });
            Some(serialize_xml_fragment(element))
        } else {
            None
        }
    }
}

fn align<'e>(
    before: &[&'e XmlElement],
    after: &[&'e XmlElement],
    before_classes: &[usize],
    after_classes: &[usize],
) -> Option<Vec<Op>> {
    let n = before.len();
    let m = after.len();
    let mut names: HashMap<&'e str, usize> = HashMap::new();
    let mut ids: HashMap<&'e str, usize> = HashMap::new();
    let mut intern = |element: &'e XmlElement| {
        let next = names.len();
        let name = *names.entry(element.name.as_str()).or_insert(next);
        let id = drawing_object_id(element).map(|id| {
            let next = ids.len();
            *ids.entry(id).or_insert(next)
        });
        (name, id)
    };
    let before_keys: Vec<(usize, Option<usize>)> =
        before.iter().copied().map(&mut intern).collect();
    let after_keys: Vec<(usize, Option<usize>)> = after.iter().copied().map(&mut intern).collect();
    let equal = |i: usize, j: usize| {
        #[cfg(feature = "test-counters")]
        crate::test_counters::bump(|work| work.equal_calls = work.equal_calls.saturating_add(1));
        before_classes[i] == after_classes[j]
    };
    let can_match = |i: usize, j: usize| {
        let (before_name, before_id) = before_keys[i];
        let (after_name, after_id) = after_keys[j];
        before_name == after_name
            && match (before_id, after_id) {
                (Some(left), Some(right)) => left == right,
                _ => true,
            }
    };
    let mut ops = Vec::new();
    let mut prefix = 0;
    while prefix < n && prefix < m && equal(prefix, prefix) {
        ops.push(Op::Match(prefix, prefix));
        prefix += 1;
    }
    let mut suffix = 0;
    while suffix < n - prefix && suffix < m - prefix && equal(n - 1 - suffix, m - 1 - suffix) {
        suffix += 1;
    }
    let middle_n = n - prefix - suffix;
    let middle_m = m - prefix - suffix;
    #[cfg(feature = "test-counters")]
    crate::test_counters::bump(|work| {
        work.dp_cells = work
            .dp_cells
            .saturating_add(middle_n.saturating_mul(middle_m));
    });
    if middle_n
        .checked_mul(middle_m)
        .is_none_or(|size| size > 1_000_000)
    {
        return None;
    }
    let mut dp = vec![vec![0i32; middle_m + 1]; middle_n + 1];
    for (i, row) in dp.iter_mut().enumerate().take(middle_n) {
        row[middle_m] = -((middle_n - i) as i32);
    }
    for (j, cell) in dp[middle_n].iter_mut().enumerate().take(middle_m) {
        *cell = -((middle_m - j) as i32);
    }
    for i in (0..middle_n).rev() {
        for j in (0..middle_m).rev() {
            let mut best = i32::MIN / 2;
            if can_match(prefix + i, prefix + j) {
                let score = if equal(prefix + i, prefix + j) {
                    100
                } else {
                    10
                };
                best = best.max(score + dp[i + 1][j + 1]);
            }
            best = best.max(-1 + dp[i + 1][j]);
            best = best.max(-1 + dp[i][j + 1]);
            dp[i][j] = best;
        }
    }
    let mut i = 0;
    let mut j = 0;
    while i < middle_n || j < middle_m {
        if i < middle_n && j < middle_m && can_match(prefix + i, prefix + j) {
            let score = if equal(prefix + i, prefix + j) {
                100
            } else {
                10
            };
            if dp[i][j] == score + dp[i + 1][j + 1] {
                ops.push(Op::Match(prefix + i, prefix + j));
                i += 1;
                j += 1;
                continue;
            }
        }
        if i < middle_n && dp[i][j] == -1 + dp[i + 1][j] {
            ops.push(Op::Delete(prefix + i));
            i += 1;
            continue;
        }
        ops.push(Op::Insert(prefix + j));
        j += 1;
    }
    for offset in 0..suffix {
        ops.push(Op::Match(n - suffix + offset, m - suffix + offset));
    }
    Some(ops)
}

fn drawing_object_id(element: &XmlElement) -> Option<&str> {
    if element.local_name() == "AlternateContent" {
        let mut id = None;
        for branch in element.child_elements() {
            if !matches!(branch.local_name(), "Choice" | "Fallback") {
                return None;
            }
            let mut children = branch.child_elements();
            let child = children.next()?;
            if children.next().is_some() {
                return None;
            }
            let branch_id = drawing_object_id(child)?;
            if id.is_some_and(|current| current != branch_id) {
                return None;
            }
            id = Some(branch_id);
        }
        return id;
    }
    let non_visual = if matches!(
        element.local_name(),
        "nvSpPr" | "nvPicPr" | "nvGrpSpPr" | "nvGraphicFramePr" | "nvCxnSpPr"
    ) {
        Some(element)
    } else {
        element.child_elements().find(|child| {
            matches!(
                child.local_name(),
                "nvSpPr" | "nvPicPr" | "nvGrpSpPr" | "nvGraphicFramePr" | "nvCxnSpPr"
            )
        })
    }?;
    non_visual
        .child_elements()
        .find(|child| child.local_name() == "cNvPr")?
        .attribute("id")
}

fn structural_hash(element: &XmlElement) -> (u64, usize) {
    fn mix(hash: &mut u64, bytes: &[u8]) {
        for byte in (bytes.len() as u64).to_le_bytes() {
            *hash ^= u64::from(byte);
            *hash = hash.wrapping_mul(0x100_0000_01b3);
        }
        for byte in bytes {
            *hash ^= u64::from(*byte);
            *hash = hash.wrapping_mul(0x100_0000_01b3);
        }
    }
    fn hash_element(hash: &mut u64, element: &XmlElement) -> usize {
        #[cfg(feature = "test-counters")]
        crate::test_counters::bump(|work| work.hashed_nodes = work.hashed_nodes.saturating_add(1));
        let mut bytes = element.name.len();
        let mut weight = 1;
        mix(hash, element.name.as_bytes());
        mix(hash, &(element.attributes.len() as u64).to_le_bytes());
        for (key, value) in &element.attributes {
            bytes += key.len() + value.len();
            mix(hash, key.as_bytes());
            mix(hash, value.as_bytes());
        }
        mix(hash, &(element.children.len() as u64).to_le_bytes());
        for child in &element.children {
            match child {
                XmlNode::Text(text) => {
                    bytes += text.len();
                    weight += 1;
                    mix(hash, b"text");
                    mix(hash, text.as_bytes());
                }
                XmlNode::Element(child) => {
                    mix(hash, b"element");
                    weight += hash_element(hash, child);
                }
            }
        }
        weight + bytes / BYTES_PER_WEIGHT_UNIT
    }
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    let weight = hash_element(&mut hash, element);
    (hash, weight)
}

fn weight(element: &XmlElement) -> usize {
    let mut bytes = element.name.len();
    for (key, value) in &element.attributes {
        bytes += key.len() + value.len();
    }
    1 + children_weight(element, bytes)
}

fn children_weight(element: &XmlElement, mut bytes: usize) -> usize {
    let mut descendants = 0;
    for child in &element.children {
        match child {
            XmlNode::Text(text) => {
                bytes += text.len();
                descendants += 1;
            }
            XmlNode::Element(child) => descendants += weight(child),
        }
    }
    descendants + bytes / BYTES_PER_WEIGHT_UNIT
}

#[cfg(feature = "test-counters")]
fn element_count(element: &XmlElement) -> usize {
    1 + element.child_elements().map(element_count).sum::<usize>()
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
    enum Child<'a> {
        Element(&'a XmlElement),
        Text(String),
    }
    fn merged(element: &XmlElement) -> Vec<Child<'_>> {
        let mut children: Vec<Child<'_>> = Vec::new();
        for child in &element.children {
            match (children.last_mut(), child) {
                (Some(Child::Text(previous)), XmlNode::Text(text)) => previous.push_str(text),
                (_, XmlNode::Text(text)) => children.push(Child::Text(text.clone())),
                (_, XmlNode::Element(child)) => children.push(Child::Element(child)),
            }
        }
        children
    }
    let left = merged(before);
    let right = merged(after);
    if left.len() != right.len() {
        return false;
    }
    left.iter()
        .zip(right.iter())
        .all(|(left, right)| match (left, right) {
            (Child::Element(left), Child::Element(right)) => equivalent(left, right),
            (Child::Text(left), Child::Text(right)) => left == right,
            _ => false,
        })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn budget_exhaustion_regenerates_the_whole_slide() {
        let source = concat!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
            "<p:cSld><p:spTree>",
            r#"<p:sp><p:nvSpPr><p:cNvPr id="1" name="First"/></p:nvSpPr></p:sp>"#,
            r#"<p:sp><p:nvSpPr><!--kept--><p:cNvPr id="2" name="Middle"/></p:nvSpPr></p:sp>"#,
            r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Last"/></p:nvSpPr></p:sp>"#,
            "</p:spTree></p:cSld></p:sld>",
        );
        let limits = ParseLimits::default();
        let parse =
            |bytes: &[u8]| parse_xml(bytes, "slide", &mut ParseBudget::new(&limits)).unwrap();
        let before = parse(source.as_bytes());
        let after = parse(source.replace("Middle", "Edited").as_bytes());
        assert!(splice_slide_with_limit(source.as_bytes(), &before, &after, Some(1)).is_none());
        let bytes = splice_slide_with_limit(source.as_bytes(), &before, &after, None).unwrap();
        assert!(equivalent(&parse(&bytes), &after));
        assert!(String::from_utf8_lossy(&bytes).contains("<!--kept-->"));
        let mut spliced = 0;
        let full = (weight(&before) + weight(&after)) * SPLICE_WORK_PER_NODE + SPLICE_WORK_SLACK;
        for limit in 0..=full {
            if let Some(partial) =
                splice_slide_with_limit(source.as_bytes(), &before, &after, Some(limit))
            {
                assert_eq!(partial, bytes, "limit {limit}");
                spliced += 1;
            }
        }
        assert!(spliced > 0);
        let root = weight(&before);
        assert!((root..full).any(|limit| {
            splice_slide_with_limit(source.as_bytes(), &before, &after, Some(limit)).is_none()
        }));
    }
}
