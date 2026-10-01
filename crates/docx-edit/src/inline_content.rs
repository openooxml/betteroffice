use std::collections::HashMap;

use yrs::Any;

pub(crate) fn visit(content: &Any, visitor: &mut impl FnMut(&HashMap<String, Any>)) {
    let Any::Array(children) = content else {
        return;
    };
    for child in children.iter() {
        let Any::Map(child) = child else { continue };
        visitor(child);
        if let Some(Any::Map(payload)) = child.get("payload")
            && let Some(content) = payload.get("content")
        {
            visit(content, visitor);
        }
    }
}

fn append_text(child: &HashMap<String, Any>, text: &mut String) {
    match child.get("kind").and_then(|value| match value {
        Any::String(value) => Some(value.as_ref()),
        _ => None,
    }) {
        Some("text") => {
            if let Some(Any::String(value)) = child.get("text") {
                text.push_str(value);
            }
        }
        Some("tab") => text.push('\t'),
        Some("break") => text.push('\n'),
        Some("sdt") => {}
        _ => text.push('\u{FFFC}'),
    }
}

pub(crate) fn revision_text(content: &Any, id: &str, key: &str, inherited: bool) -> String {
    let Any::Array(children) = content else {
        return String::new();
    };
    let mut text = String::new();
    for child in children.iter() {
        let Any::Map(child) = child else { continue };
        let matched = inherited
            || matches!(child.get("attrs"), Some(Any::Map(attrs))
            if attrs.get(key).and_then(crate::queries::revision_parts).is_some_and(|(revision, ..)| revision == id));
        if let Some(Any::Map(payload)) = child.get("payload")
            && let Some(content) = payload.get("content")
        {
            text.push_str(&revision_text(content, id, key, matched));
        } else if matched {
            append_text(child, &mut text);
        }
    }
    text
}
