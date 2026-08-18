//! Element census: counts by qualified name across every XML part.
//!
//! Catches the drops nobody predicted. Blind to attribute values by
//! construction; identifier multisets get their own guards elsewhere.

use std::collections::BTreeMap;

use crate::error::FidelityError;
use crate::is_xml_part;
use crate::xml::{XmlElement, XmlLimits, XmlNode, parse_part};

/// Counts keyed by `(namespace, local)` over every XML part of a package.
pub type Census = BTreeMap<(String, String), usize>;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Loss {
    pub namespace: String,
    pub local: String,
    pub before: usize,
    pub after: usize,
}

pub fn element_census(parts: &[(String, Vec<u8>)]) -> Result<Census, FidelityError> {
    let limits = XmlLimits::default();
    let mut census = Census::new();
    for (name, bytes) in parts {
        if !is_xml_part(name) {
            continue;
        }
        let root = parse_part(bytes, name, &limits)?;
        count(&root, &mut census);
    }
    Ok(census)
}

/// Names whose count shrank; growth is an edit's business, loss never is.
pub fn losses(before: &Census, after: &Census) -> Vec<Loss> {
    before
        .iter()
        .filter_map(|((namespace, local), &count_before)| {
            let count_after = after.get(&(namespace.clone(), local.clone())).copied();
            let count_after = count_after.unwrap_or(0);
            (count_after < count_before).then(|| Loss {
                namespace: namespace.clone(),
                local: local.clone(),
                before: count_before,
                after: count_after,
            })
        })
        .collect()
}

fn count(element: &XmlElement, census: &mut Census) {
    *census
        .entry((element.namespace.clone(), element.local.clone()))
        .or_insert(0) += 1;
    for child in &element.children {
        if let XmlNode::Element(child) = child {
            count(child, census);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn package(document: &str) -> Vec<(String, Vec<u8>)> {
        vec![
            ("word/document.xml".to_owned(), document.as_bytes().to_vec()),
            ("word/media/image1.png".to_owned(), vec![0x89, 0x50]),
        ]
    }

    #[test]
    fn counts_elements_across_xml_parts_only() {
        let census = element_census(&package(r#"<w:d xmlns:w="ns"><w:p/><w:p/></w:d>"#)).unwrap();
        assert_eq!(census.get(&("ns".to_owned(), "p".to_owned())), Some(&2));
        assert_eq!(census.len(), 2);
    }

    #[test]
    fn a_shrunken_count_is_a_loss() {
        let before = element_census(&package(r#"<w:d xmlns:w="ns"><w:p/><w:p/></w:d>"#)).unwrap();
        let after = element_census(&package(r#"<w:d xmlns:w="ns"><w:p/></w:d>"#)).unwrap();
        assert_eq!(
            losses(&before, &after),
            vec![Loss {
                namespace: "ns".to_owned(),
                local: "p".to_owned(),
                before: 2,
                after: 1,
            }]
        );
    }

    #[test]
    fn growth_is_not_a_loss() {
        let before = element_census(&package(r#"<w:d xmlns:w="ns"><w:p/></w:d>"#)).unwrap();
        let after = element_census(&package(r#"<w:d xmlns:w="ns"><w:p/><w:p/></w:d>"#)).unwrap();
        assert_eq!(losses(&before, &after), vec![]);
    }

    #[test]
    fn a_vanished_name_reports_zero() {
        let before =
            element_census(&package(r#"<w:d xmlns:w="ns"><w:bookmarkStart/></w:d>"#)).unwrap();
        let after = element_census(&package(r#"<w:d xmlns:w="ns"/>"#)).unwrap();
        assert_eq!(losses(&before, &after).len(), 1);
        assert_eq!(losses(&before, &after)[0].after, 0);
    }
}
