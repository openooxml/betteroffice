use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::io::{self, Write};
use std::path::Path;

use docx_edit::save::revision_ids::{
    RevisionIdsError, assign_session_revision_ids, session_revision_id_reservations,
};
use docx_parse::S13SaveRequest;
use serde::Deserialize;
use serde_json::Value;

#[derive(Deserialize)]
struct Case {
    document: String,
    request: String,
    session: BTreeSet<u64>,
    wire: String,
    sites: Vec<Site>,
    used: BTreeSet<u32>,
    max: u32,
    #[serde(rename = "outOfRange")]
    out_of_range: BTreeSet<String>,
}

#[derive(Deserialize)]
struct Site {
    path: String,
    placeholder: u64,
    id: u32,
}

#[derive(Default)]
struct DocumentCensus {
    out_of_range: BTreeSet<String>,
    reservation_mismatch: bool,
    dropped_site: bool,
}

fn word_number(value: &Value) -> Option<u64> {
    let id = value.as_f64()?;
    ((0.0..((1u64 << 53) as f64)).contains(&id) && id.fract() == 0.0).then_some(id as u64)
}

fn read_case(path: &Path) -> Result<Case, Box<dyn std::error::Error>> {
    Ok(serde_json::from_slice(&fs::read(path)?)?)
}

fn check_case(case: &Case, census: &mut DocumentCensus) -> Result<(), Box<dyn std::error::Error>> {
    let mut request: S13SaveRequest = serde_json::from_str(&case.request)?;
    let typed_before = serde_json::to_value(&request)?;
    let dropped = case
        .sites
        .iter()
        .find(|site| typed_before.pointer(&site.path).is_none());
    census.dropped_site |= dropped.is_some();

    let (used, max) = session_revision_id_reservations(&request, &case.session)?;
    let reservation_mismatch = used != case.used || max != case.max;
    census.reservation_mismatch |= reservation_mismatch;
    if let Some(site) = dropped {
        return Err(format!("site under a dropped field: {}", site.path).into());
    }

    let mut ts_ids = BTreeMap::new();
    for site in &case.sites {
        let id = typed_before.pointer(&site.path).unwrap();
        if word_number(id) != Some(site.placeholder) {
            return Err(format!(
                "first differing site {}: typed request has {id}, recorded placeholder is {}",
                site.path, site.placeholder
            )
            .into());
        }
        if let Some(previous) = ts_ids.insert(site.placeholder, site.id)
            && previous != site.id
        {
            return Err(format!(
                "first differing site {}: TypeScript assigned {previous} and {} to {}",
                site.path, site.id, site.placeholder
            )
            .into());
        }
    }
    drop(typed_before);
    if !ts_ids.keys().eq(case.session.iter()) {
        return Err("recorded sites and session placeholders differ".into());
    }

    if reservation_mismatch {
        let only_ts = case.used.difference(&used).collect::<Vec<_>>();
        let only_rust = used.difference(&case.used).collect::<Vec<_>>();
        return Err(format!(
            "reserved IDs differ: TypeScript-only {only_ts:?}, Rust-only {only_rust:?}, \
             TypeScript max {}, Rust max {max}",
            case.max
        )
        .into());
    }

    let actual_ids = assign_session_revision_ids(&mut request, &case.session).map_err(|error| {
        if let RevisionIdsError::Unassigned(id) = error
            && let Some(site) = case.sites.iter().find(|site| site.placeholder == id)
        {
            return format!("first differing site {}: {error}", site.path);
        }
        error.to_string()
    })?;
    let actual = serde_json::to_string(&request)?;
    let actual_tree: Value = serde_json::from_str(&actual)?;
    for site in &case.sites {
        let actual_id = actual_tree.pointer(&site.path).and_then(word_number);
        if actual_id != Some(u64::from(site.id)) {
            return Err(format!(
                "first differing site {} ({}): Rust {actual_id:?}, TypeScript {}",
                site.path, site.placeholder, site.id
            )
            .into());
        }
    }
    if let Some(site) = case
        .sites
        .iter()
        .find(|site| actual_ids.get(&site.placeholder) != Some(&site.id))
    {
        return Err(format!(
            "first differing site {} ({}): Rust map {:?}, TypeScript map {}",
            site.path,
            site.placeholder,
            actual_ids.get(&site.placeholder),
            site.id
        )
        .into());
    }
    let expected = serde_json::to_string(&serde_json::from_str::<S13SaveRequest>(&case.wire)?)?;
    if actual == expected && actual_ids == ts_ids {
        return Ok(());
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
    let census_only = std::env::var("DOCX_REVISION_IDS_CENSUS").is_ok_and(|mode| mode == "1");
    let mut paths = fs::read_dir(root)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .filter(|path| path.extension().is_some_and(|extension| extension == "json"))
        .collect::<Vec<_>>();
    paths.sort();
    let count = paths.len();
    let mut docs: BTreeMap<String, DocumentCensus> = BTreeMap::new();
    let mut failed = 0;
    for path in paths {
        let result = read_case(&path).and_then(|case| {
            let census = docs.entry(case.document.clone()).or_default();
            census.out_of_range.extend(case.out_of_range.iter().cloned());
            check_case(&case, census)
        });
        if let Err(error) = result {
            if !census_only {
                eprintln!("{}: {error}", path.display());
            }
            failed += 1;
        }
    }
    let case_a_docs = docs
        .values()
        .filter(|doc| !doc.out_of_range.is_empty())
        .count();
    let case_a_ids: usize = docs.values().map(|doc| doc.out_of_range.len()).sum();
    let case_b_docs = docs.values().filter(|doc| doc.reservation_mismatch).count();
    let dropped_site_docs = docs.values().filter(|doc| doc.dropped_site).count();
    writeln!(
        io::stderr().lock(),
        "REVISION-ID-CENSUS docs={} cases={count} case_a_docs={case_a_docs} \
         case_a_ids={case_a_ids} case_b_docs={case_b_docs} dropped_site_docs={dropped_site_docs}",
        docs.len()
    )
    .unwrap();
    if !census_only {
        assert!(count > 0, "no DOCX revision ID cases were recorded");
        assert_eq!(failed, 0, "{failed} of {count} DOCX revision ID cases failed");
    }
}
