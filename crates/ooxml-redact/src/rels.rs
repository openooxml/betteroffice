//! One reading of OPC relationship markup, shared by the XML rewriter and the
//! scrubber so the two cannot disagree about which targets leave the package.

/// Whether the attribute name carries no namespace prefix. Only unqualified
/// OPC attributes take part in relationship decisions.
pub(crate) fn is_unqualified(key: &str) -> bool {
    !key.contains(':')
}

pub(crate) fn attribute_local(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

/// The value of an unqualified attribute with this name. OPC names are
/// case-sensitive, so the exact spelling a consumer reads wins over a tolerated
/// variant appearing earlier.
pub(crate) fn unqualified_value<'a>(
    attributes: &'a [(String, String)],
    expected: &str,
) -> Option<&'a str> {
    let mut variant = None;
    for (name, value) in attributes {
        if !is_unqualified(name) {
            continue;
        }
        if name == expected {
            return Some(value.as_str());
        }
        if variant.is_none() && name.eq_ignore_ascii_case(expected) {
            variant = Some(value.as_str());
        }
    }
    variant
}

/// Whether a relationship's attributes mark it as pointing outside the
/// package. `package_part` enables reading the target's shape, which only a
/// `.rels` part's consumers resolve; a shape is read from the exact-case
/// `Target` those consumers use.
pub(crate) fn external_relationship(attributes: &[(String, String)], package_part: bool) -> bool {
    attributes.iter().any(|(key, value)| {
        if !is_unqualified(key) {
            return false;
        }
        let local = attribute_local(key);
        local.eq_ignore_ascii_case("TargetMode") && value.trim().eq_ignore_ascii_case("External")
            || package_part && local == "Target" && external_target(value)
    })
}

/// Whether a relationship target points outside the package. Query and
/// fragment are dropped first: neither names a part, and either may carry a
/// URI of its own.
pub(crate) fn external_target(target: &str) -> bool {
    let lower = target
        .trim()
        .split(['?', '#'])
        .next()
        .unwrap_or_default()
        .to_ascii_lowercase();
    lower.starts_with("//")
        || lower.starts_with(r"\\")
        || lower
            .split_once(':')
            .is_some_and(|(scheme, _)| is_uri_scheme(scheme))
}

fn is_uri_scheme(scheme: &str) -> bool {
    let mut chars = scheme.chars();
    if !matches!(chars.next(), Some(first) if first.is_ascii_alphabetic()) {
        return false;
    }
    chars.all(|character| character.is_ascii_alphanumeric() || matches!(character, '+' | '-' | '.'))
}
