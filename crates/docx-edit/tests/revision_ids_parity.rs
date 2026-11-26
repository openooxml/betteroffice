use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::Path;

use docx_edit::save::revision_ids::assign_session_revision_ids;
use docx_parse::S13SaveRequest;
use serde::Deserialize;
use serde_json::Value;

#[derive(Deserialize)]
struct Case {
    request: String,
    session: Vec<u64>,
    wire: String,
}

struct Site {
    pointer: String,
    placeholder: u64,
    id: u32,
}

fn word_number(value: &Value) -> Option<u64> {
    let id = value.as_f64()?;
    ((0.0..((1u64 << 53) as f64)).contains(&id) && id.fract() == 0.0).then_some(id as u64)
}

fn sites(
    before: &Value,
    wire: &Value,
    session: &BTreeSet<u64>,
    pointer: &mut String,
    candidates: &mut Vec<Site>,
) {
    match before {
        Value::Object(record) => {
            if record.get("author").is_some_and(Value::is_string)
                && let Some(placeholder) = record.get("id").and_then(word_number)
                && session.contains(&placeholder)
                && let Some(id) = wire
                    .pointer(pointer)
                    .and_then(|info| info.get("id"))
                    .and_then(word_number)
                    .and_then(|id| u32::try_from(id).ok())
                && (1..(1 << 31)).contains(&id)
            {
                candidates.push(Site {
                    pointer: format!("{pointer}/id"),
                    placeholder,
                    id,
                });
            }
            for (key, child) in record {
                let len = pointer.len();
                pointer.push('/');
                pointer.push_str(&key.replace('~', "~0").replace('/', "~1"));
                sites(child, wire, session, pointer, candidates);
                pointer.truncate(len);
            }
        }
        Value::Array(items) => {
            for (index, child) in items.iter().enumerate() {
                let len = pointer.len();
                pointer.push('/');
                pointer.push_str(&index.to_string());
                sites(child, wire, session, pointer, candidates);
                pointer.truncate(len);
            }
        }
        _ => {}
    }
}

fn check_case(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let case: Case = serde_json::from_slice(&fs::read(path)?)?;
    let before: Value = serde_json::from_str(&case.request)?;
    let wire: Value = serde_json::from_str(&case.wire)?;
    let session = case.session.iter().copied().collect();
    let mut candidates = Vec::new();
    sites(
        &before,
        &wire,
        &session,
        &mut String::new(),
        &mut candidates,
    );
    let order: BTreeMap<_, _> = case
        .session
        .iter()
        .enumerate()
        .map(|(position, &placeholder)| (placeholder, position))
        .collect();
    candidates.sort_by_key(|site| order[&site.placeholder]);

    let mut ts_ids = BTreeMap::new();
    for site in candidates
        .iter()
        .filter(|site| site.placeholder != u64::from(site.id))
    {
        if let Some(previous) = ts_ids.insert(site.placeholder, site.id)
            && previous != site.id
        {
            return Err(format!(
                "first differing site {}: TypeScript assigned {previous} and {} to {}",
                site.pointer, site.id, site.placeholder
            )
            .into());
        }
    }
    for site in &candidates {
        ts_ids.entry(site.placeholder).or_insert(site.id);
    }
    if let Some(id) = case.session.iter().find(|id| !ts_ids.contains_key(id)) {
        return Err(format!("TypeScript wire has no recorded site for {id}").into());
    }

    let mut request: S13SaveRequest = serde_json::from_str(&case.request)?;
    let actual_ids = assign_session_revision_ids(&mut request, &session).map_err(|error| {
        if let docx_edit::save::revision_ids::RevisionIdsError::Unassigned(id) = error
            && let Some(site) = candidates.iter().find(|site| site.placeholder == id)
        {
            return format!("first differing site {}: {error}", site.pointer);
        }
        error.to_string()
    })?;
    let actual = serde_json::to_string(&request)?;
    let expected = serde_json::to_string(&serde_json::from_str::<S13SaveRequest>(&case.wire)?)?;
    if actual == expected && actual_ids == ts_ids {
        return Ok(());
    }

    let actual_tree: Value = serde_json::from_str(&actual)?;
    let expected_tree: Value = serde_json::from_str(&expected)?;
    for site in &candidates {
        let actual_id = actual_tree.pointer(&site.pointer);
        let expected_id = expected_tree.pointer(&site.pointer);
        if actual_id != expected_id {
            return Err(format!(
                "first differing site {} ({}): Rust {actual_id:?}, TypeScript {expected_id:?}",
                site.pointer, site.placeholder
            )
            .into());
        }
    }
    if let Some((&id, expected)) = ts_ids
        .iter()
        .find(|&(id, expected)| actual_ids.get(id) != Some(expected))
    {
        let site = candidates.iter().find(|site| site.placeholder == id).unwrap();
        return Err(format!(
            "first differing site {} ({id}): Rust map {:?}, TypeScript map {expected}",
            site.pointer,
            actual_ids.get(&id)
        )
        .into());
    }
    let byte = actual
        .bytes()
        .zip(expected.bytes())
        .take_while(|(a, b)| a == b)
        .count();
    Err(format!(
        "request or map mismatch at byte {byte}: Rust {actual_ids:?}, TypeScript {ts_ids:?}"
    )
    .into())
}

#[test]
#[ignore = "requires product save cases from DOCX_REVISION_IDS_DIR"]
fn revision_ids_parity() {
    let root = std::env::var_os("DOCX_REVISION_IDS_DIR").expect("DOCX_REVISION_IDS_DIR is not set");
    let mut paths = fs::read_dir(root)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .filter(|path| path.extension().is_some_and(|extension| extension == "json"))
        .collect::<Vec<_>>();
    paths.sort();
    let count = paths.len();
    eprintln!("checking {count} DOCX revision ID cases");
    assert!(count > 0, "no DOCX revision ID cases were recorded");
    let mut failed = 0;
    for path in paths {
        if let Err(error) = check_case(&path) {
            eprintln!("{}: {error}", path.display());
            failed += 1;
        }
    }
    assert_eq!(failed, 0, "{failed} of {count} DOCX revision ID cases failed");
}
