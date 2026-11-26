use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::io::{self, Write};
use std::path::Path;

use docx_edit::save::revision_ids::{
    RevisionIdsError, assign_session_revision_ids_with_paths, session_revision_id_reservations,
};
use docx_parse::S13SaveRequest;
use serde::Deserialize;
use serde_json::{Value, json};

#[derive(Deserialize)]
struct Case {
    document: String,
    request: String,
    session: BTreeSet<u64>,
    wire: String,
    sites: Vec<Site>,
    reservations: Vec<Reservation>,
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

#[derive(Deserialize)]
struct Reservation {
    path: String,
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

    let reservations = session_revision_id_reservations(&request, &case.session)?;
    let mut ts_reservations = BTreeMap::new();
    for reservation in &case.reservations {
        let site = (reservation.path.clone(), reservation.id);
        *ts_reservations.entry(site).or_insert(0) += 1;
    }
    let differing_reservation = ts_reservations
        .keys()
        .find(|&site| ts_reservations.get(site) != reservations.sites.get(site))
        .or_else(|| {
            reservations
                .sites
                .keys()
                .find(|&site| !ts_reservations.contains_key(site))
        });
    let reservation_mismatch = reservations.used != case.used || reservations.max != case.max;
    census.reservation_mismatch |= reservation_mismatch || differing_reservation.is_some();
    if let Some(site) = dropped {
        return Err(format!("site under a dropped field: {}", site.path).into());
    }
    if let Some(site) = differing_reservation {
        let (path, id) = site;
        let ts_count = ts_reservations.get(site).copied().unwrap_or_default();
        let rust_count = reservations.sites.get(site).copied().unwrap_or_default();
        return Err(format!(
            "reservation at {path:?} for ID {id} differs: TypeScript {ts_count}, Rust {rust_count}"
        )
        .into());
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
        let only_ts = case.used.difference(&reservations.used).collect::<Vec<_>>();
        let only_rust = reservations.used.difference(&case.used).collect::<Vec<_>>();
        return Err(format!(
            "reserved IDs differ: TypeScript-only {only_ts:?}, Rust-only {only_rust:?}, \
             TypeScript max {}, Rust max {}",
            case.max, reservations.max
        )
        .into());
    }

    let mut paths = Vec::new();
    let actual_ids =
        assign_session_revision_ids_with_paths(&mut request, &case.session, &mut paths);
    for index in 0..paths.len().max(case.sites.len()) {
        let native = paths.get(index).map(|(path, id)| (path.as_str(), *id));
        let ts = case
            .sites
            .get(index)
            .map(|site| (site.path.as_str(), site.placeholder));
        if native != ts {
            return Err(format!(
                "first differing site at index {index}: TypeScript {ts:?}, Rust {native:?}"
            )
            .into());
        }
    }
    let actual_ids = actual_ids.map_err(|error| {
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
fn detects_dropped_reservation_with_matching_used_ids_and_max() {
    let placeholder = (1u64 << 53) - 1;
    let path = "/document/content/1/content/0/propertyChanges/0";
    let witness = format!("{path}/previousAttributes/witness");
    let request = json!({
        "determinism": {"seed": "0".repeat(64), "now": "2030-01-02T03:04:05.006Z"},
        "document": {"content": [
            {"type": "rawXml", "xml": "<w:customXml w:id='40'/>"},
            {"type": "paragraph", "content": [{
                "type": "run",
                "propertyChanges": [{
                    "type": "runPropertyChange",
                    "info": {"id": placeholder, "author": "Reviewer"},
                    "previousAttributes": {"witness": "w:id='40'"}
                }],
                "content": [{"type": "text", "text": "Revised text"}]
            }]}
        ]}
    });
    let path = format!("{path}/info/id");
    let mut wire = request.clone();
    *wire.pointer_mut(&path).unwrap() = json!(41);
    let case = Case {
        document: "dropped-reservation".to_owned(),
        request: request.to_string(),
        session: BTreeSet::from([placeholder]),
        wire: wire.to_string(),
        sites: vec![Site {
            path,
            placeholder,
            id: 41,
        }],
        reservations: vec![
            Reservation {
                path: "/document/content/0/xml".to_owned(),
                id: 40,
            },
            Reservation {
                path: witness.clone(),
                id: 40,
            },
        ],
        used: BTreeSet::from([40]),
        max: 40,
        out_of_range: BTreeSet::new(),
    };
    let mut census = DocumentCensus::default();
    let error = check_case(&case, &mut census).unwrap_err();
    assert_eq!(
        error.to_string(),
        format!("reservation at {witness:?} for ID 40 differs: TypeScript 1, Rust 0")
    );
    assert!(census.reservation_mismatch);
}

#[test]
fn detects_reordered_and_missing_sites_with_matching_wire_and_ids() {
    let paragraph = json!({"type": "paragraph", "content": [{
        "type": "insertion", "info": {"id": 1, "author": "Reviewer"}, "content": []
    }]});
    let request = json!({
        "determinism": {"seed": "0".repeat(64), "now": "2030-01-02T03:04:05.006Z"},
        "document": {"content": [paragraph]},
        "headerEntries": [["rId1", {"type": "header", "hdrFtrType": "default", "content": [paragraph]}]]
    });
    let body = "/document/content/0/content/0/info/id";
    let header = "/headerEntries/0/1/content/0/content/0/info/id";
    let mut case = Case {
        document: "identity-assignment".to_owned(),
        request: request.to_string(),
        wire: request.to_string(),
        session: BTreeSet::from([1]),
        sites: [body, header]
            .into_iter()
            .map(|path| Site {
                path: path.to_owned(),
                placeholder: 1,
                id: 1,
            })
            .collect(),
        reservations: Vec::new(),
        used: BTreeSet::new(),
        max: 0,
        out_of_range: BTreeSet::new(),
    };
    let mut census = DocumentCensus::default();
    check_case(&case, &mut census).unwrap();

    case.sites.swap(0, 1);
    let error = check_case(&case, &mut census).unwrap_err();
    assert_eq!(
        error.to_string(),
        format!(
            "first differing site at index 0: TypeScript Some(({header:?}, 1)), Rust Some(({body:?}, 1))"
        )
    );

    case.sites.remove(0);
    let error = check_case(&case, &mut census).unwrap_err();
    assert_eq!(
        error.to_string(),
        format!("first differing site at index 1: TypeScript None, Rust Some(({header:?}, 1))")
    );
}

#[test]
#[ignore = "requires product save cases from DOCX_REVISION_IDS_DIR"]
fn revision_ids_parity() {
    let census_only = std::env::var("DOCX_REVISION_IDS_CENSUS").is_ok_and(|mode| mode == "1");
    let mut paths = Vec::new();
    if let Some(root) = std::env::var_os("DOCX_REVISION_IDS_DIR").filter(|root| !root.is_empty()) {
        match fs::read_dir(root) {
            Ok(entries) => {
                paths.extend(entries.map(|entry| entry.unwrap().path()).filter(|path| {
                    path.extension()
                        .is_some_and(|extension| extension == "json")
                }))
            }
            Err(error) if error.kind() == io::ErrorKind::NotFound => {}
            Err(error) => panic!("reading DOCX_REVISION_IDS_DIR: {error}"),
        }
    }
    paths.sort();
    let count = paths.len();
    let mut docs: BTreeMap<String, DocumentCensus> = BTreeMap::new();
    let mut failed = 0;
    for path in paths {
        let result = read_case(&path).and_then(|case| {
            let census = docs.entry(case.document.clone()).or_default();
            census
                .out_of_range
                .extend(case.out_of_range.iter().cloned());
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
    assert!(count > 0, "no DOCX revision ID cases were recorded");
    if !census_only {
        assert_eq!(
            failed, 0,
            "{failed} of {count} DOCX revision ID cases failed"
        );
    }
}
