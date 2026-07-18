//! The per-op edit context (op-contract R3).
//!
//! Every mutating op takes an [`EditCtx`] instead of per-op author/suggesting parameters. The
//! context carries durable authorship (`author` + `now_iso`, host-supplied so native, WASM, and
//! agent peers share one clock policy) and the yrs transaction origin used for undo tracking and
//! authority policy. `Origin::System` edits (for example paraId re-uniquing) are excluded from
//! undo because the undo manager tracks only the local origin.

use crate::Author;

/// Who is making this edit, for yrs transaction-origin purposes.
///
/// Only [`EditOrigin::Local`] transactions are tracked by the undo manager; agent, remote, and
/// system edits never enter the local undo stack (op-contract "Undo" section).
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum EditOrigin {
    /// The local human user. Tracked by undo.
    Local,
    /// An agent peer writing through the op surface. Untracked by undo.
    Agent,
    /// A remote collaborator's replayed edit. Untracked by undo.
    Remote,
    /// Schema maintenance (paraId re-uniquing and similar). Untracked by undo.
    System,
}

/// Marker for suggesting (tracked-changes) mode.
///
/// S1 stamps `ins`/`del`/`pPrIns`/`pPrDel` revisions directly; `rPrChange` payloads arrive in S4,
/// which is why this struct is currently empty but kept as a struct (not a bool) so S4 can add
/// fields without changing every op signature.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct SuggestCtx {}

/// Context shared by every mutating op (op-contract R3).
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct EditCtx {
    /// Durable display author stamped into revision values.
    pub author: String,
    /// yrs transaction origin class; see [`EditOrigin`].
    pub origin: EditOrigin,
    /// `Some` puts the op into suggesting (tracked-changes) mode.
    pub suggesting: Option<SuggestCtx>,
    /// ISO-8601 timestamp supplied by the host clock.
    pub now_iso: String,
}

impl EditCtx {
    /// A direct (non-suggesting) local edit.
    pub fn local(author: impl Into<String>, now_iso: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            origin: EditOrigin::Local,
            suggesting: None,
            now_iso: now_iso.into(),
        }
    }

    /// A schema-maintenance edit, excluded from undo.
    pub fn system(now_iso: impl Into<String>) -> Self {
        Self {
            author: String::new(),
            origin: EditOrigin::System,
            suggesting: None,
            now_iso: now_iso.into(),
        }
    }

    /// Switches this context into suggesting (tracked-changes) mode.
    pub fn suggesting(mut self) -> Self {
        self.suggesting = Some(SuggestCtx::default());
        self
    }

    pub(crate) fn is_suggesting(&self) -> bool {
        self.suggesting.is_some()
    }

    pub(crate) fn revision_author(&self) -> Author {
        Author::new(self.author.clone(), self.now_iso.clone())
    }
}

impl crate::EditingDoc {
    /// Opens a mutable transaction whose yrs origin encodes the context's [`EditOrigin`].
    ///
    /// Local edits use the replica's client ID as origin — the one origin the undo manager
    /// tracks. Agent/remote/system edits use string origins and therefore never enter the local
    /// undo stack.
    pub(crate) fn transact_for(&self, ctx: &EditCtx) -> yrs::TransactionMut<'_> {
        use yrs::Transact;
        match ctx.origin {
            EditOrigin::Local => self.yrs_doc().transact_mut_with(self.client_id()),
            EditOrigin::Agent => self.yrs_doc().transact_mut_with("agent"),
            EditOrigin::Remote => self.yrs_doc().transact_mut_with("remote"),
            EditOrigin::System => self.yrs_doc().transact_mut_with("system"),
        }
    }
}
