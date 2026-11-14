//! Conditional formats in a collaborative document. The blob each sheet was
//! seeded with stays the baseline every replica opened, and an edit is one key
//! on the sheet's map: `cf:{client}:{clock}:{n}` holds a rule a replica added,
//! and is that rule's id, and `cf:{id}` set to `false` removes a baseline rule.
//! Group, order and priority are worked out from those on every read and never
//! stored. No operation edits a rule in place, so a key has one writer's value.

use serde::Deserialize;
use xlsx_model::{ConditionalFormatRule, DifferentialFormat};

use super::*;

pub(super) const RULE_PREFIX: &str = "cf:";
const MAX_RULE_ENTRY_BYTES: usize = 64 * 1024;

/// `(client, clock, n)`: the clock this client had before the batch wrote, and
/// the rule's place in that batch.
type RuleId = (u64, u64, u64);

/// One added rule in the shared document. `seq` orders rules and means nothing
/// else, so two replicas may pick the same one.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct AddedRule {
    sqref: String,
    #[serde(rename = "type")]
    rule_type: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    operator: Option<String>,
    formulas: Vec<String>,
    format: DifferentialFormat,
    seq: u64,
}

impl AddedRule {
    /// What a peer has to accept, so the author is held to it before writing.
    fn validate(&self) -> Result<(), String> {
        let range = CellRange::parse_a1(&self.sqref)
            .map_err(|_| format!("invalid conditional format range {}", self.sqref))?;
        if range.to_a1() != self.sqref
            || self.sqref.contains('$')
            || range.end.row >= MAX_ROWS
            || range.end.col >= MAX_COLS
        {
            return Err(format!(
                "noncanonical or out of bounds conditional format range {}",
                self.sqref
            ));
        }
        let cleaned = xlsx_ops::validate_conditional_format(
            &self.rule_type,
            self.operator.as_deref(),
            &self.formulas,
            &self.format,
        )
        .map_err(|error| error.to_string())?;
        if cleaned != self.formulas {
            return Err("conditional format formulas are not in stored form".to_string());
        }
        Ok(())
    }

    fn encode(&self) -> Result<String, String> {
        let json = serde_json::to_string(self)
            .map_err(|error| format!("cannot encode conditional format: {error}"))?;
        if json.len() > MAX_RULE_ENTRY_BYTES {
            return Err("conditional format exceeds its size limit".to_string());
        }
        Ok(json)
    }

    fn decode(key: &str, json: &str) -> Result<Self, String> {
        if json.len() > MAX_RULE_ENTRY_BYTES {
            return Err(format!("conditional format {key} exceeds its size limit"));
        }
        let rule: Self = serde_json::from_str(json)
            .map_err(|error| format!("invalid conditional format {key}: {error}"))?;
        rule.validate()?;
        if rule.encode()? != json {
            return Err(format!("conditional format {key} is not in stored form"));
        }
        Ok(rule)
    }

    fn into_rule(self, id: RuleId, priority: u32) -> (String, ConditionalFormatRule) {
        let rule = ConditionalFormatRule {
            id: rule_key(id),
            rule_type: self.rule_type,
            operator: self.operator,
            formulas: self.formulas,
            priority,
            dxf_id: None,
            format: self.format,
            raw_xml: None,
        };
        (self.sqref, rule)
    }
}

fn rule_key((client, clock, index): RuleId) -> String {
    format!("{RULE_PREFIX}{client}:{clock}:{index}")
}

fn parse_rule_key(key: &str) -> Option<RuleId> {
    let mut parts = key.strip_prefix(RULE_PREFIX)?.split(':');
    let mut number = || parts.next()?.parse::<u64>().ok();
    let id = (number()?, number()?, number()?);
    (rule_key(id) == key).then_some(id)
}

/// The rules a sheet's map holds over its baseline.
#[derive(Default)]
pub(super) struct RuleOverlay {
    added: BTreeMap<RuleId, AddedRule>,
    removed: BTreeSet<String>,
}

impl RuleOverlay {
    pub(super) fn read<T: ReadTxn>(sheet: &MapRef, txn: &T) -> Result<Self, String> {
        let mut overlay = Self::default();
        for (key, value) in sheet.iter(txn) {
            let Some(entry) = key.strip_prefix(RULE_PREFIX) else {
                continue;
            };
            if entry.contains(':') {
                let id = parse_rule_key(key)
                    .ok_or_else(|| format!("invalid conditional format key {key}"))?;
                let Out::Any(Any::String(json)) = value else {
                    return Err(format!("conditional format {key} is not a string"));
                };
                overlay.added.insert(id, AddedRule::decode(key, &json)?);
            } else {
                if entry.is_empty() || !matches!(value, Out::Any(Any::Bool(false))) {
                    return Err(format!("invalid conditional format removal {key}"));
                }
                overlay.removed.insert(entry.to_string());
            }
        }
        Ok(overlay)
    }

    /// Turns `groups`, a sheet's baseline, into what it reads as under this
    /// overlay. Baseline rules keep their priority, and added rules follow the
    /// highest one of the whole baseline in `seq` order, so removing a rule
    /// renumbers nothing. The sum saturates: two adds each legal on their own
    /// must not fail their merge.
    pub(super) fn apply(
        self,
        mut groups: Vec<ConditionalFormatting>,
    ) -> Result<Vec<ConditionalFormatting>, String> {
        if self.added.is_empty() && self.removed.is_empty() {
            return Ok(groups);
        }
        let baseline = groups
            .iter()
            .flat_map(|group| &group.rules)
            .map(|rule| (rule.id.as_str(), rule.editable()))
            .collect::<HashMap<_, _>>();
        for id in &self.removed {
            match baseline.get(id.as_str()) {
                Some(true) => {}
                Some(false) => return Err(format!("conditional format {id} is not editable")),
                None => return Err(format!("conditional format {id} has no source rule")),
            }
        }
        let first = groups
            .iter()
            .flat_map(|group| &group.rules)
            .map(|rule| rule.priority)
            .max()
            .unwrap_or(0)
            .saturating_add(1);
        for group in &mut groups {
            group.rules.retain(|rule| !self.removed.contains(&rule.id));
        }
        let mut places = HashMap::new();
        for (index, group) in groups.iter().enumerate() {
            places.entry(group.sqref.clone()).or_insert(index);
        }
        let mut added = self.added.into_iter().collect::<Vec<_>>();
        added.sort_by_key(|(id, rule)| (rule.seq, *id));
        for (index, (id, entry)) in added.into_iter().enumerate() {
            let offset = u32::try_from(index).unwrap_or(u32::MAX);
            let (sqref, rule) = entry.into_rule(id, first.saturating_add(offset));
            let place = *places.entry(sqref.clone()).or_insert_with(|| {
                groups.push(ConditionalFormatting {
                    sqref,
                    rules: Vec::new(),
                });
                groups.len() - 1
            });
            groups[place].rules.push(rule);
        }
        groups.retain(|group| !group.rules.is_empty());
        Ok(groups)
    }

    fn next_seq(&self) -> u64 {
        self.added
            .values()
            .map(|rule| rule.seq.saturating_add(1))
            .max()
            .unwrap_or(0)
    }
}

/// What a sheet's rules read as with no overlay.
pub(super) fn baseline_conditional_formats<T: ReadTxn>(
    sheet: &MapRef,
    txn: &T,
    fallback: &[ConditionalFormatting],
) -> Result<Vec<ConditionalFormatting>, String> {
    match sheet.get(txn, CONDITIONAL_FORMATS) {
        Some(Out::Any(Any::String(json))) => serde_json::from_str(&json)
            .map_err(|error| format!("sheet conditional formats are invalid: {error}")),
        Some(_) => Err("sheet conditional formats are not a string".to_string()),
        None => Ok(fallback.to_vec()),
    }
}

/// Hands out the ids of one batch. An operation previews a rule under an id of
/// its own, which can equal a removed baseline rule's, so an id is never taken
/// from there. The clock moves on with every batch that writes a rule, also
/// across an undo and a reopen under the same client id.
pub(super) struct RuleIds {
    client: u64,
    clock: u64,
    next: u64,
}

impl RuleIds {
    pub(super) fn new(client: u64, clock: u32) -> Self {
        Self {
            client,
            clock: u64::from(clock),
            next: 0,
        }
    }

    fn mint(&mut self) -> String {
        let key = rule_key((self.client, self.clock, self.next));
        self.next += 1;
        key
    }
}

fn plain_sqref(sqref: &str) -> Result<String, String> {
    let mut range = CellRange::parse_a1(sqref)
        .map_err(|_| format!("invalid conditional format range {sqref}"))?;
    for at in [&mut range.start, &mut range.end] {
        at.abs_row = false;
        at.abs_col = false;
    }
    Ok(range.to_a1())
}

fn rules_by_id(groups: &[ConditionalFormatting]) -> BTreeMap<&str, (&str, &ConditionalFormatRule)> {
    groups
        .iter()
        .flat_map(|group| {
            group
                .rules
                .iter()
                .map(move |rule| (rule.id.as_str(), (group.sqref.as_str(), rule)))
        })
        .collect()
}

/// Writes what moved between the stored projection and `sheet`. Rules are
/// compared whole, because a batch that removes a baseline rule and then adds
/// one previews the new rule under the removed rule's id.
pub(super) fn sync_conditional_formats(
    txn: &mut TransactionMut<'_>,
    sheet_map: &MapRef,
    sheet: &Sheet,
    fallback: &[ConditionalFormatting],
    ids: &mut RuleIds,
) -> Result<(), String> {
    let stored = RuleOverlay::read(sheet_map, txn)?;
    let mut seq = stored.next_seq();
    let projected = stored.apply(baseline_conditional_formats(sheet_map, txn, fallback)?)?;
    let before = rules_by_id(&projected);
    let after = rules_by_id(&sheet.conditional_formats);
    for (id, rule) in &before {
        if after.get(id) == Some(rule) {
            continue;
        }
        if parse_rule_key(id).is_none() {
            sheet_map.try_update(txn, format!("{RULE_PREFIX}{id}"), false);
        } else if after.contains_key(id) {
            return Err(format!("conditional format {id} cannot change"));
        } else {
            sheet_map.remove(txn, id);
        }
    }
    let mut added = after
        .iter()
        .filter(|(id, rule)| before.get(*id) != Some(*rule))
        .map(|(_, rule)| *rule)
        .collect::<Vec<_>>();
    added.sort_by_key(|(_, rule)| rule.priority);
    for (sqref, rule) in added {
        let entry = AddedRule {
            sqref: plain_sqref(sqref)?,
            rule_type: rule.rule_type.clone(),
            operator: rule.operator.clone(),
            formulas: rule.formulas.clone(),
            format: rule.format.clone(),
            seq,
        };
        entry.validate()?;
        sheet_map.insert(txn, ids.mint(), entry.encode()?);
        seq = seq.saturating_add(1);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn group(sqref: &str, rules: &[(&str, u32, &str)]) -> ConditionalFormatting {
        ConditionalFormatting {
            sqref: sqref.to_string(),
            rules: rules
                .iter()
                .map(|(id, priority, kind)| ConditionalFormatRule {
                    id: (*id).to_string(),
                    rule_type: (*kind).to_string(),
                    operator: None,
                    formulas: vec!["A1>0".to_string()],
                    priority: *priority,
                    dxf_id: Some(0),
                    format: DifferentialFormat::default(),
                    raw_xml: None,
                })
                .collect(),
        }
    }

    fn added(sqref: &str, seq: u64) -> String {
        serde_json::to_string(&AddedRule {
            sqref: sqref.to_string(),
            rule_type: "expression".to_string(),
            operator: None,
            formulas: vec!["A1>0".to_string()],
            format: DifferentialFormat::default(),
            seq,
        })
        .unwrap()
    }

    fn overlay(writes: &[(&str, Any)]) -> Result<RuleOverlay, String> {
        let doc = Doc::new();
        let mut txn = doc.transact_mut();
        let sheets = txn.get_or_insert_map(SHEETS);
        let sheet: MapRef = sheets.get_or_init(&mut txn, "sheet:0");
        for (key, value) in writes {
            sheet.insert(&mut txn, *key, value.clone());
        }
        RuleOverlay::read(&sheet, &txn)
    }

    fn rules(groups: &[ConditionalFormatting]) -> Vec<(String, String, u32)> {
        groups
            .iter()
            .flat_map(|group| {
                group
                    .rules
                    .iter()
                    .map(move |rule| (group.sqref.clone(), rule.id.clone(), rule.priority))
            })
            .collect()
    }

    fn tuple(sqref: &str, id: &str, priority: u32) -> (String, String, u32) {
        (sqref.to_string(), id.to_string(), priority)
    }

    #[test]
    fn an_empty_overlay_returns_the_baseline_untouched() {
        let baseline = vec![group("A1:A5", &[("cf1", 9, "cellIs")])];
        assert_eq!(
            RuleOverlay::default().apply(baseline.clone()).unwrap(),
            baseline
        );
    }

    #[test]
    fn priorities_follow_the_whole_baseline_and_ties_go_by_client_then_clock() {
        let baseline = vec![
            group("A1:A5", &[("cf1", 2, "expression"), ("cf2", 7, "cellIs")]),
            group("B1:B5", &[("cf3", 4, "expression")]),
        ];
        let rule = |sqref, seq| Any::from(added(sqref, seq));
        let groups = overlay(&[
            ("cf:cf2", Any::Bool(false)),
            ("cf:cf3", Any::Bool(false)),
            ("cf:9:1:0", rule("B1:B5", 0)),
            ("cf:3:8:1", rule("C1", 1)),
            ("cf:3:8:0", rule("A1:A5", 0)),
        ])
        .unwrap()
        .apply(baseline)
        .unwrap();
        assert_eq!(
            rules(&groups),
            [
                tuple("A1:A5", "cf1", 2),
                tuple("A1:A5", "cf:3:8:0", 8),
                tuple("B1:B5", "cf:9:1:0", 9),
                tuple("C1", "cf:3:8:1", 10),
            ]
        );
    }

    #[test]
    fn priorities_saturate_rather_than_fail_a_merge() {
        let baseline = vec![group("A1", &[("cf1", u32::MAX - 1, "expression")])];
        let groups = overlay(&[
            ("cf:1:0:0", Any::from(added("B1", 0))),
            ("cf:2:0:0", Any::from(added("C1", 0))),
        ])
        .unwrap()
        .apply(baseline)
        .unwrap();
        assert_eq!(
            rules(&groups),
            [
                tuple("A1", "cf1", u32::MAX - 1),
                tuple("B1", "cf:1:0:0", u32::MAX),
                tuple("C1", "cf:2:0:0", u32::MAX),
            ]
        );
    }

    #[test]
    fn a_rule_key_has_one_spelling() {
        assert_eq!(parse_rule_key("cf:1:2:3"), Some((1, 2, 3)));
        for key in [
            "cf:01:2:3",
            "cf:1:2",
            "cf:1:2:3:4",
            "cf:+1:2:3",
            "cf:a:2:3",
            "cf:1:2:",
            "cf:",
        ] {
            assert_eq!(parse_rule_key(key), None, "{key}");
        }
    }

    #[test]
    fn stored_shapes_are_checked_on_read() {
        let good = added("A1:B2", 0);
        let refused = [
            ("cf:01:2:3", Any::from(good.clone())),
            ("cf:1:2", Any::from(good.clone())),
            ("cf:1:2:3", Any::Bool(false)),
            ("cf:cf7", Any::from(good.clone())),
            ("cf:cf7", Any::Bool(true)),
            ("cf:", Any::Bool(false)),
            ("cf:1:2:3", Any::from(good.replace("A1:B2", "$A$1:B2"))),
            ("cf:1:2:3", Any::from(good.replace("A1:B2", "B2:A1"))),
            ("cf:1:2:3", Any::from(good.replace("A1:B2", "A1:B2 C3"))),
            ("cf:1:2:3", Any::from(good.replace("A1:B2", "A1:XFE1"))),
            ("cf:1:2:3", Any::from(good.replace("A1>0", "=A1>0"))),
            ("cf:1:2:3", Any::from(good.replace("expression", "dataBar"))),
            ("cf:1:2:3", Any::from(format!("{good} "))),
            (
                "cf:1:2:3",
                Any::from(good.replace(r#""seq":0"#, r#""seq":0,"priority":1"#)),
            ),
            (
                "cf:1:2:3",
                Any::from(good.replace(r#""format":{}"#, r#""format":{"fill":"red"}"#)),
            ),
            (
                "cf:1:2:3",
                Any::from(good.replace(r#""format":{}"#, r#""format":{"x":1}"#)),
            ),
            (
                "cf:1:2:3",
                Any::from(good.replace("A1>0", &"x".repeat(MAX_RULE_ENTRY_BYTES))),
            ),
        ];
        for (key, value) in refused {
            assert!(overlay(&[(key, value.clone())]).is_err(), "{key} {value:?}");
        }
        assert!(overlay(&[("cf:1:2:3", Any::from(good))]).is_ok());
    }

    #[test]
    fn a_rule_a_peer_would_refuse_is_never_written() {
        let doc = Doc::new();
        let mut txn = doc.transact_mut();
        let sheets = txn.get_or_insert_map(SHEETS);
        let map: MapRef = sheets.get_or_init(&mut txn, "sheet:0");
        let mut sheet = Sheet::new("Data");
        sheet.conditional_formats = vec![group("A1:A5", &[("cf1", 1, "expression")])];
        sheet.conditional_formats[0].rules[0].formulas = vec!["=A1>0".to_string()];
        let mut ids = RuleIds::new(1, 0);
        assert!(sync_conditional_formats(&mut txn, &map, &sheet, &[], &mut ids).is_err());
        assert_eq!(map.len(&txn), 0);
        sheet.conditional_formats[0].rules[0].formulas = vec!["A1>0".to_string()];
        sync_conditional_formats(&mut txn, &map, &sheet, &[], &mut ids).unwrap();
        assert_eq!(map.len(&txn), 1);
    }

    #[test]
    fn a_removal_needs_an_editable_baseline_rule() {
        let baseline = vec![group(
            "A1:A5",
            &[("cf1", 1, "expression"), ("cf2", 2, "dataBar")],
        )];
        let removal = |id: &str| {
            let key = format!("{RULE_PREFIX}{id}");
            overlay(&[(key.as_str(), Any::Bool(false))])
                .unwrap()
                .apply(baseline.clone())
        };
        assert_eq!(rules(&removal("cf1").unwrap()), [tuple("A1:A5", "cf2", 2)]);
        assert!(removal("cf2").is_err());
        assert!(removal("cf3").is_err());
    }
}
