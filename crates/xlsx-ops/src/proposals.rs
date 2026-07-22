//! HITL agent-proposal store: pending agent edits held out of the workbook
//! until a human accepts or rejects. display texts are captured upstream.

use serde::{Deserialize, Serialize};
use xlsx_model::CellValue;

use crate::{CellState, NumberFormatMutation};

#[derive(Debug, Clone, PartialEq)]
pub struct ProposalGhost {
    pub sheet: u32,
    pub row: u32,
    pub col: u32,
    pub old_text: String,
    pub new_text: String,
    pub alignment_value: CellValue,
}

/// one cell edit inside a proposal. `input` is the raw editor string re-applied
/// on accept and stays off the wire; `old_text`/`new_text` are display texts.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProposedEdit {
    pub sheet: u32,
    pub row: u32,
    pub col: u32,
    #[serde(skip)]
    pub input: String,
    #[serde(skip)]
    pub old_state: CellState,
    #[serde(skip)]
    pub number_format: Option<NumberFormatMutation>,
    pub a1: String,
    pub old_text: String,
    pub new_text: String,
}

/// a pending agent proposal; serializes with `edits` presented as `cells`.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Proposal {
    pub id: String,
    pub agent_id: String,
    pub note: Option<String>,
    #[serde(rename = "cells")]
    pub edits: Vec<ProposedEdit>,
    #[serde(skip)]
    pub ghosts: Vec<ProposalGhost>,
}

/// the session's live set of pending proposals, with a monotonic id counter.
#[derive(Debug, Clone, Default)]
pub struct ProposalSet {
    next_id: u64,
    proposals: Vec<Proposal>,
}

impl ProposalSet {
    pub fn new() -> Self {
        Self::default()
    }

    /// allocate the next proposal id (`p1`, `p2`, ...).
    pub fn next_id(&mut self) -> String {
        self.next_id += 1;
        format!("p{}", self.next_id)
    }

    pub fn add(&mut self, proposal: Proposal) {
        self.proposals.push(proposal);
    }

    pub fn list(&self) -> &[Proposal] {
        &self.proposals
    }

    /// remove and return the proposal with `id`, if present.
    pub fn take(&mut self, id: &str) -> Option<Proposal> {
        let pos = self.proposals.iter().position(|p| p.id == id)?;
        Some(self.proposals.remove(pos))
    }

    /// drop the proposal with `id`; reports whether one was removed.
    pub fn remove(&mut self, id: &str) -> bool {
        let before = self.proposals.len();
        self.proposals.retain(|p| p.id != id);
        before != self.proposals.len()
    }

    pub fn clear(&mut self) {
        self.proposals.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn edit(a1: &str, old: &str, new: &str) -> ProposedEdit {
        ProposedEdit {
            sheet: 0,
            row: 0,
            col: 0,
            input: new.to_string(),
            old_state: CellState::default(),
            number_format: None,
            a1: a1.to_string(),
            old_text: old.to_string(),
            new_text: new.to_string(),
        }
    }

    fn proposal(id: &str) -> Proposal {
        Proposal {
            id: id.to_string(),
            agent_id: "agent-1".to_string(),
            note: None,
            edits: vec![edit("A1", "", "hi")],
            ghosts: Vec::new(),
        }
    }

    #[test]
    fn ids_increment() {
        let mut set = ProposalSet::new();
        assert_eq!(set.next_id(), "p1");
        assert_eq!(set.next_id(), "p2");
        assert_eq!(set.next_id(), "p3");
    }

    #[test]
    fn add_list_take_remove() {
        let mut set = ProposalSet::new();
        set.add(proposal("p1"));
        set.add(proposal("p2"));
        assert_eq!(set.list().len(), 2);

        let taken = set.take("p1").unwrap();
        assert_eq!(taken.id, "p1");
        assert_eq!(set.list().len(), 1);
        assert!(set.take("p1").is_none());

        assert!(set.remove("p2"));
        assert!(!set.remove("p2"));
        assert!(set.list().is_empty());
    }

    #[test]
    fn wire_shape_uses_cells_and_hides_input() {
        let p = proposal("p1");
        let json = serde_json::to_string(&p).unwrap();
        let v: serde_json::Value = serde_json::from_str(&json).unwrap();
        assert_eq!(v["id"], "p1");
        assert_eq!(v["agentId"], "agent-1");
        assert!(v["note"].is_null());
        let cell = &v["cells"][0];
        assert_eq!(cell["a1"], "A1");
        assert_eq!(cell["oldText"], "");
        assert_eq!(cell["newText"], "hi");
        assert!(cell.get("input").is_none());
    }
}
