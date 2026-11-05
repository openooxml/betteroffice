use serde::Deserialize;

use super::*;

const EDIT_PREFIX: &str = "set:";
const TOMBSTONE_PREFIX: &str = "del:";
const WORKBOOK_SCOPE: &str = "*";
/// A filter and its name stay out of the shared document together.
pub(crate) const FILTER_DATABASE_NAME: &str = "_xlnm._FilterDatabase";

/// One name's value in the shared document, for a source name edited or a name
/// no source holds. `seq` orders new names and means nothing else.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct NameEdit {
    name: String,
    formula: String,
    hidden: bool,
    seq: u64,
}

/// The edits and tombstones the document holds over the source table, by
/// identity, `{scope}/{folded}`. A live edit outranks a live tombstone, so a
/// name set while a collaborator deleted it stays.
#[derive(Default)]
pub(super) struct NameOverlay {
    edits: BTreeMap<String, NameEdit>,
    tombstones: BTreeSet<String>,
}

fn identity(scope: &str, name: &str) -> String {
    format!("{scope}/{}", name.to_ascii_lowercase())
}

fn scope_key(defined: &DefinedName, keys: &[String]) -> Option<String> {
    match defined.local_sheet {
        None => Some(WORKBOOK_SCOPE.to_string()),
        Some(sheet) => keys.get(sheet.0 as usize).cloned(),
    }
}

fn name_identity(defined: &DefinedName, keys: &[String]) -> Option<String> {
    scope_key(defined, keys).map(|scope| identity(&scope, &defined.name))
}

fn valid_scope(scope: &str) -> bool {
    scope == WORKBOOK_SCOPE
        || base_sheet_index(scope).is_some()
        || scope.starts_with(APPENDED_SHEET_PREFIX)
}

fn parse_identity(key: &str, prefix: &str) -> Result<(String, String), String> {
    let entry = key
        .strip_prefix(prefix)
        .ok_or_else(|| format!("invalid defined name key {key}"))?;
    let (_, folded) = entry
        .split_once('/')
        .filter(|(scope, folded)| valid_scope(scope) && !folded.is_empty())
        .ok_or_else(|| format!("invalid defined name key {key}"))?;
    if folded.eq_ignore_ascii_case(FILTER_DATABASE_NAME) {
        return Err(format!("defined name key {key} is not replicated"));
    }
    Ok((entry.to_string(), folded.to_string()))
}

impl NameOverlay {
    pub(super) fn read<T: ReadTxn>(txn: &T) -> Result<Self, String> {
        let mut overlay = Self::default();
        let Some(names) = txn.get_map(DEFINED_NAMES) else {
            return Ok(overlay);
        };
        for (key, value) in names.iter(txn) {
            if key.starts_with(TOMBSTONE_PREFIX) {
                let (entry, _) = parse_identity(key, TOMBSTONE_PREFIX)?;
                if !matches!(value, Out::Any(Any::Bool(true))) {
                    return Err(format!("defined name tombstone {key} is not true"));
                }
                overlay.tombstones.insert(entry);
                continue;
            }
            let (entry, folded) = parse_identity(key, EDIT_PREFIX)?;
            let Out::Any(Any::String(json)) = value else {
                return Err(format!("defined name {key} is not a string"));
            };
            let edit: NameEdit = serde_json::from_str(&json)
                .map_err(|error| format!("invalid defined name {key}: {error}"))?;
            if edit.name.to_ascii_lowercase() != folded {
                return Err(format!("defined name {key} holds {}", edit.name));
            }
            overlay.edits.insert(entry, edit);
        }
        Ok(overlay)
    }

    fn is_empty(&self) -> bool {
        self.edits.is_empty() && self.tombstones.is_empty()
    }

    /// Turns `names`, the source table, into what it reads as under this
    /// overlay. Source names keep their place and new names follow in `seq`
    /// order, ties broken by identity. A name scoped to a sheet `keys` does not
    /// hold is left out, not removed: it returns with the sheet.
    pub(super) fn apply(
        mut self,
        names: &mut Vec<DefinedName>,
        keys: &[String],
    ) -> Result<(), String> {
        if self.is_empty() {
            return Ok(());
        }
        let baseline = std::mem::take(names);
        names.reserve(baseline.len() + self.edits.len());
        let mut overridden = HashSet::new();
        for source in baseline {
            let Some(identity) = name_identity(&source, keys) else {
                names.push(source);
                continue;
            };
            if overridden.contains(&identity) {
                return Err(format!(
                    "defined name {} is ambiguous in its source",
                    source.name
                ));
            }
            let tombstoned = self.tombstones.remove(&identity);
            match self.edits.remove(&identity) {
                Some(edit) => {
                    names.push(DefinedName {
                        name: edit.name,
                        formula: edit.formula,
                        local_sheet: source.local_sheet,
                        hidden: edit.hidden,
                    });
                    overridden.insert(identity);
                }
                None if tombstoned => {
                    overridden.insert(identity);
                }
                None => names.push(source),
            }
        }
        if let Some(identity) = self.tombstones.first() {
            return Err(format!("defined name tombstone {identity} has no source"));
        }
        let mut added = self.edits.into_iter().collect::<Vec<_>>();
        added.sort_by(|(left, a), (right, b)| (a.seq, left).cmp(&(b.seq, right)));
        for (identity, edit) in added {
            let (scope, _) = identity
                .split_once('/')
                .ok_or_else(|| format!("invalid defined name identity {identity}"))?;
            let local_sheet = if scope == WORKBOOK_SCOPE {
                None
            } else {
                let Some(index) = keys.iter().position(|key| key == scope) else {
                    continue;
                };
                Some(SheetId(u32::try_from(index).map_err(|_| {
                    "defined name scope is out of range".to_string()
                })?))
            };
            names.push(DefinedName {
                name: edit.name,
                formula: edit.formula,
                local_sheet,
                hidden: edit.hidden,
            });
        }
        Ok(())
    }

    fn next_seq(&self) -> u64 {
        self.edits
            .values()
            .map(|edit| edit.seq.saturating_add(1))
            .max()
            .unwrap_or(0)
    }
}

fn by_identity<'a>(
    names: &'a [DefinedName],
    keys: &[String],
) -> BTreeMap<String, (usize, &'a DefinedName)> {
    let mut table = BTreeMap::new();
    for (index, defined) in names.iter().enumerate() {
        if let Some(identity) = name_identity(defined, keys) {
            table.entry(identity).or_insert((index, defined));
        }
    }
    table
}

/// Writes what moved between two projections, each read against its own sheet
/// keys. A name the projection leaves out is in neither, so it is never touched,
/// and nothing stored is derived from a position in the table.
pub(super) fn sync_defined_names(
    txn: &mut TransactionMut<'_>,
    baseline: &[DefinedName],
    (before, before_keys): (&[DefinedName], &[String]),
    (after, after_keys): (&[DefinedName], &[String]),
) -> Result<(), String> {
    let stored = NameOverlay::read(txn)?;
    let names = txn.get_or_insert_map(DEFINED_NAMES);
    let sources = by_identity(baseline, before_keys);
    let before = by_identity(before, before_keys);
    let after = by_identity(after, after_keys);
    for identity in before.keys().filter(|key| !after.contains_key(*key)) {
        names.remove(txn, &format!("{EDIT_PREFIX}{identity}"));
        if sources.contains_key(identity) {
            names.try_update(txn, format!("{TOMBSTONE_PREFIX}{identity}"), true);
        }
    }
    let mut changed = after
        .iter()
        .filter(|(identity, (_, defined))| {
            before
                .get(*identity)
                .is_none_or(|(_, previous)| previous != defined)
        })
        .collect::<Vec<_>>();
    changed.sort_by_key(|(_, (index, _))| *index);
    let mut next = stored.next_seq();
    for (identity, (_, defined)) in changed {
        let seq = match stored.edits.get(identity) {
            Some(edit) => edit.seq,
            None if sources.contains_key(identity) => 0,
            None => {
                let seq = next;
                next = next.saturating_add(1);
                seq
            }
        };
        let edit = NameEdit {
            name: defined.name.clone(),
            formula: defined.formula.clone(),
            hidden: defined.hidden,
            seq,
        };
        let json = serde_json::to_string(&edit)
            .map_err(|error| format!("cannot encode defined name: {error}"))?;
        names.insert(txn, format!("{EDIT_PREFIX}{identity}"), json);
        names.remove(txn, &format!("{TOMBSTONE_PREFIX}{identity}"));
    }
    Ok(())
}

/// See [`WorkbookAuthority::dormant_defined_names`].
pub(super) fn dormant_names<T: ReadTxn>(
    txn: &T,
    keys: &[String],
) -> Result<Vec<DefinedName>, String> {
    let shown = keys.iter().map(String::as_str).collect::<HashSet<_>>();
    NameOverlay::read(txn)?
        .edits
        .into_iter()
        .filter(|(identity, _)| {
            identity
                .split_once('/')
                .is_some_and(|(scope, _)| scope != WORKBOOK_SCOPE && !shown.contains(scope))
        })
        .enumerate()
        .map(|(index, (_, edit))| -> Result<DefinedName, String> {
            let sheet = u32::try_from(keys.len() + index)
                .map_err(|_| "defined name scope is out of range".to_string())?;
            Ok(DefinedName {
                name: edit.name,
                formula: edit.formula,
                local_sheet: Some(SheetId(sheet)),
                hidden: edit.hidden,
            })
        })
        .collect()
}

/// The stored keys of every name scoped to one of `sheets`.
pub(super) fn scoped_name_keys<T: ReadTxn>(
    names: &MapRef,
    txn: &T,
    sheets: &[&str],
) -> Vec<String> {
    names
        .keys(txn)
        .filter(|key| {
            key.strip_prefix(EDIT_PREFIX)
                .and_then(|entry| entry.split_once('/'))
                .is_some_and(|(scope, _)| sheets.contains(&scope))
        })
        .map(str::to_string)
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn name(name: &str, formula: &str, sheet: Option<u32>) -> DefinedName {
        DefinedName {
            name: name.into(),
            formula: formula.into(),
            local_sheet: sheet.map(SheetId),
            hidden: false,
        }
    }

    fn edit(name: &str, formula: &str, seq: u64) -> NameEdit {
        NameEdit {
            name: name.into(),
            formula: formula.into(),
            hidden: false,
            seq,
        }
    }

    fn keys() -> Vec<String> {
        vec!["sheet:0".to_string(), "replica:7:4:0".to_string()]
    }

    fn overlay(writes: &[(&str, Any)]) -> Result<NameOverlay, String> {
        let doc = Doc::new();
        let mut txn = doc.transact_mut();
        let names = txn.get_or_insert_map(DEFINED_NAMES);
        for (key, value) in writes {
            names.insert(&mut txn, *key, value.clone());
        }
        NameOverlay::read(&txn)
    }

    fn project(
        overlay: NameOverlay,
        baseline: &[DefinedName],
        keys: &[String],
    ) -> Result<Vec<DefinedName>, String> {
        let mut names = baseline.to_vec();
        overlay.apply(&mut names, keys)?;
        Ok(names)
    }

    fn json(edit: &NameEdit) -> Any {
        Any::String(serde_json::to_string(edit).unwrap().into())
    }

    #[test]
    fn an_empty_overlay_is_the_baseline() {
        let baseline = vec![name("Rate", "1", None), name("RATE", "2", None)];
        let projected = project(overlay(&[]).unwrap(), &baseline, &keys()).unwrap();
        assert_eq!(projected, baseline);
    }

    #[test]
    fn source_names_keep_their_place_and_new_names_follow_in_seq_order() {
        let baseline = vec![name("B", "1", None), name("A", "2", None)];
        let written = overlay(&[
            ("set:*/z", json(&edit("Z", "9", 1))),
            ("set:*/y", json(&edit("Y", "8", 2))),
            ("set:*/x", json(&edit("X", "7", 1))),
            ("set:*/a", json(&edit("A", "5", 0))),
        ])
        .unwrap();
        let projected = project(written, &baseline, &keys()).unwrap();
        assert_eq!(
            projected,
            vec![
                name("B", "1", None),
                name("A", "5", None),
                name("X", "7", None),
                name("Z", "9", None),
                name("Y", "8", None),
            ]
        );
    }

    #[test]
    fn an_edit_outranks_a_tombstone_and_a_tombstone_outranks_the_source() {
        let baseline = vec![name("Kept", "1", None), name("Gone", "2", None)];
        let written = overlay(&[
            ("set:*/kept", json(&edit("Kept", "3", 0))),
            ("del:*/kept", Any::Bool(true)),
            ("del:*/gone", Any::Bool(true)),
        ])
        .unwrap();
        let projected = project(written, &baseline, &keys()).unwrap();
        assert_eq!(projected, vec![name("Kept", "3", None)]);
    }

    #[test]
    fn a_name_on_an_inactive_sheet_is_left_out_and_returns_with_it() {
        let writes = [("set:replica:9:1:0/n", json(&edit("N", "1", 0)))];
        let projected = project(overlay(&writes).unwrap(), &[], &keys()).unwrap();
        assert!(projected.is_empty());
        let mut active = keys();
        active.push("replica:9:1:0".to_string());
        let projected = project(overlay(&writes).unwrap(), &[], &active).unwrap();
        assert_eq!(projected, vec![name("N", "1", Some(2))]);
    }

    #[test]
    fn only_names_no_shown_sheet_holds_are_dormant() {
        let writes = [
            ("set:*/shown", json(&edit("Shown", "1", 0))),
            ("set:sheet:0/local", json(&edit("Local", "2", 0))),
            ("set:replica:9:1:0/hidden", json(&edit("Hidden", "3", 0))),
            ("set:sheet:9/missing", json(&edit("Missing", "4", 0))),
        ];
        let doc = Doc::new();
        let mut txn = doc.transact_mut();
        let names = txn.get_or_insert_map(DEFINED_NAMES);
        for (key, value) in &writes {
            names.insert(&mut txn, *key, value.clone());
        }
        assert_eq!(
            dormant_names(&txn, &keys()).unwrap(),
            vec![name("Hidden", "3", Some(2)), name("Missing", "4", Some(3))]
        );
    }

    #[test]
    fn a_slash_in_a_name_stays_in_the_name() {
        let written = overlay(&[("set:*/a/b", json(&edit("A/B", "1", 0)))]).unwrap();
        let projected = project(written, &[], &keys()).unwrap();
        assert_eq!(projected, vec![name("A/B", "1", None)]);
    }

    #[test]
    fn entries_no_honest_peer_writes_are_refused() {
        for (key, value) in [
            ("rate", json(&edit("Rate", "1", 0))),
            ("set:rate", json(&edit("Rate", "1", 0))),
            ("set:nowhere/rate", json(&edit("Rate", "1", 0))),
            ("set:*/", json(&edit("", "1", 0))),
            ("set:*/rate", Any::Bool(true)),
            ("set:*/rate", Any::String("{".into())),
            ("set:*/rate", json(&edit("Other", "1", 0))),
            ("del:*/rate", Any::Bool(false)),
            ("del:rate", Any::Bool(true)),
            (
                "set:sheet:0/_xlnm._filterdatabase",
                json(&edit("_xlnm._FilterDatabase", "1", 0)),
            ),
            ("del:sheet:0/_xlnm._filterdatabase", Any::Bool(true)),
        ] {
            assert!(overlay(&[(key, value.clone())]).is_err(), "{key} {value:?}");
        }
    }

    #[test]
    fn a_tombstone_needs_a_source_name() {
        let written = overlay(&[("del:*/rate", Any::Bool(true))]).unwrap();
        let failure = project(written, &[], &keys()).unwrap_err();
        assert!(failure.contains("has no source"), "{failure}");
    }

    #[test]
    fn an_entry_on_a_repeated_source_identity_is_refused() {
        let baseline = vec![name("Rate", "1", None), name("RATE", "2", None)];
        for write in [
            ("set:*/rate", json(&edit("Rate", "3", 0))),
            ("del:*/rate", Any::Bool(true)),
        ] {
            let failure = project(overlay(&[write]).unwrap(), &baseline, &keys()).unwrap_err();
            assert!(failure.contains("ambiguous"), "{failure}");
        }
    }

    #[test]
    fn diffs_touch_only_what_moved_and_never_a_hidden_name() {
        let doc = Doc::new();
        let baseline = vec![name("Rate", "1", None), name("Tax", "2", None)];
        let keys = keys();
        let mut txn = doc.transact_mut();
        let names = txn.get_or_insert_map(DEFINED_NAMES);
        names.insert(
            &mut txn,
            "set:replica:9:1:0/hidden",
            serde_json::to_string(&edit("Hidden", "1", 4)).unwrap(),
        );
        sync_defined_names(
            &mut txn,
            &baseline,
            (&baseline, &keys),
            (
                &[
                    name("Rate", "5", None),
                    name("Fresh", "6", None),
                    name("Local", "7", Some(1)),
                ],
                &keys,
            ),
        )
        .unwrap();
        let stored = names
            .keys(&txn)
            .map(str::to_string)
            .collect::<BTreeSet<_>>();
        assert_eq!(
            stored,
            [
                "del:*/tax",
                "set:*/fresh",
                "set:*/rate",
                "set:replica:7:4:0/local",
                "set:replica:9:1:0/hidden",
            ]
            .map(str::to_string)
            .into_iter()
            .collect()
        );
        let overlay = NameOverlay::read(&txn).unwrap();
        assert_eq!(overlay.edits["*/rate"].seq, 0);
        assert_eq!(overlay.edits["*/fresh"].seq, 5);
        assert_eq!(overlay.edits["replica:7:4:0/local"].seq, 6);
        assert_eq!(overlay.edits["replica:9:1:0/hidden"].seq, 4);
    }

    #[test]
    fn setting_a_tombstoned_name_clears_the_tombstone() {
        let doc = Doc::new();
        let baseline = vec![name("Rate", "1", None)];
        let keys = keys();
        let mut txn = doc.transact_mut();
        sync_defined_names(&mut txn, &baseline, (&baseline, &keys), (&[], &keys)).unwrap();
        sync_defined_names(
            &mut txn,
            &baseline,
            (&[], &keys),
            (&[name("Rate", "2", None)], &keys),
        )
        .unwrap();
        let names = txn.get_map(DEFINED_NAMES).unwrap();
        assert_eq!(
            names.keys(&txn).map(str::to_string).collect::<Vec<_>>(),
            ["set:*/rate"]
        );
    }

    #[test]
    fn deleting_a_new_name_leaves_no_tombstone() {
        let doc = Doc::new();
        let keys = keys();
        let added = [name("Fresh", "1", None)];
        let mut txn = doc.transact_mut();
        sync_defined_names(&mut txn, &[], (&[], &keys), (&added, &keys)).unwrap();
        sync_defined_names(&mut txn, &[], (&added, &keys), (&[], &keys)).unwrap();
        assert_eq!(txn.get_map(DEFINED_NAMES).unwrap().len(&txn), 0);
    }

    #[test]
    fn scoped_keys_match_whole_sheet_keys_only() {
        let doc = Doc::new();
        let mut txn = doc.transact_mut();
        let names = txn.get_or_insert_map(DEFINED_NAMES);
        for key in [
            "set:replica:1:2:0/a",
            "set:replica:1:2:01/b",
            "set:*/c",
            "del:replica:1:2:0/d",
        ] {
            names.insert(&mut txn, key, "x");
        }
        assert_eq!(
            scoped_name_keys(&names, &txn, &["replica:1:2:0"]),
            ["set:replica:1:2:0/a"]
        );
    }
}
