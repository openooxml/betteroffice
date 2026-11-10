//! A table cell whose story a table of a live story names, while that story is missing (an undo
//! deleted its entry) or removed, reads as one empty paragraph. Reads never write: they answer
//! from a scratch document holding exactly that paragraph. The first local edit into the cell
//! creates the story with the same paragraph, outside the undo history.

use yrs::types::text::YChange;
use yrs::{Any, Map, MapPrelim, MapRef, Out, ReadTxn, Text, TextRef, Transact, TransactionMut};

use crate::identity::IdAllocator;
use crate::{
    EditingDoc, KIND_KEY, PARA_ID, PILCROW_KIND, ParagraphIdOrigin, STORIES, map_string, out_len,
    story_ref,
};

/// The parent story of a cell story id shaped `<parent>:t<n>:r<n>c<n>`, optionally followed by
/// `~<client>`.
pub(crate) fn cell_parent(story_id: &str) -> Option<&str> {
    let bare = match story_id.rsplit_once('~') {
        Some((bare, client)) if digits(client) => bare,
        _ => story_id,
    };
    let (rest, cell) = bare.rsplit_once(':')?;
    let (row, column) = cell.strip_prefix('r')?.split_once('c')?;
    if !digits(row) || !digits(column) {
        return None;
    }
    let (parent, table) = rest.rsplit_once(':')?;
    table
        .strip_prefix('t')
        .is_some_and(digits)
        .then_some(parent)
}

fn digits(value: &str) -> bool {
    !value.is_empty() && value.bytes().all(|byte| byte.is_ascii_digit())
}

fn names_cell(payload: &std::collections::BTreeMap<String, Any>, story_id: &str) -> bool {
    let Some(Any::Array(rows)) = payload.get("rows") else {
        return false;
    };
    rows.iter().any(|row| {
        let Any::Map(row) = row else {
            return false;
        };
        let Some(Any::Array(cells)) = row.get("cells") else {
            return false;
        };
        cells.iter().any(|cell| {
            matches!(cell, Any::Map(cell)
                if matches!(cell.get("story"), Some(Any::String(named)) if named.as_ref() == story_id))
        })
    })
}

/// The live parent story whose table names `story_id` as a cell, and that table embed's index.
pub(crate) fn naming_table<T: ReadTxn>(txn: &T, story_id: &str) -> Option<(TextRef, u32)> {
    let parent = story_ref(txn, cell_parent(story_id)?).ok()?;
    let mut offset = 0u32;
    for diff in parent.diff(txn, YChange::identity) {
        let len = out_len(&diff.insert);
        if let Out::YMap(map) = &diff.insert
            && map_string(map, txn, KIND_KEY).as_deref() == Some("table")
            && names_cell(&crate::embed_payload(map, txn), story_id)
        {
            return Some((parent, offset));
        }
        offset += len;
    }
    None
}

/// Whether `story_id` is a cell that a table of a live story names while its own story is
/// missing or removed.
pub(crate) fn is_empty_cell<T: ReadTxn>(txn: &T, story_id: &str) -> bool {
    story_ref(txn, story_id).is_err() && naming_table(txn, story_id).is_some()
}

/// The key of the one paragraph an empty cell reads as and is created with: the same on every
/// replica, so an edit addressed to it lands whichever replica's creation the map keeps.
pub(crate) fn empty_cell_para_key(story_id: &str) -> String {
    format!("{story_id}#p0")
}

fn write_empty_cell(txn: &mut TransactionMut<'_>, story_id: &str) -> MapRef {
    let stories = txn
        .get_map(STORIES)
        .expect("stories root is declared by EditingDoc::new");
    let story = crate::insert_story(txn, &stories, story_id);
    let pilcrow = story.insert_embed_with_attributes(
        txn,
        0,
        MapPrelim::default(),
        crate::insertion_attrs(None, None),
    );
    pilcrow.insert(txn, KIND_KEY, PILCROW_KIND);
    pilcrow.insert(txn, PARA_ID, empty_cell_para_key(story_id));
    pilcrow
}

/// A scratch document whose only story is `story_id` holding one empty paragraph: what an empty
/// cell reads as.
pub(crate) struct EmptyCell {
    doc: EditingDoc,
}

impl EmptyCell {
    pub(crate) fn new(story_id: &str) -> Self {
        let doc = EditingDoc::new(1);
        {
            let mut txn = doc.yrs_doc().transact_mut();
            write_empty_cell(&mut txn, story_id);
        }
        Self { doc }
    }

    pub(crate) fn doc(&self) -> &EditingDoc {
        &self.doc
    }
}

impl EditingDoc {
    /// The scratch document an empty cell reads from, or `None` when `story_id` is not one.
    pub(crate) fn empty_cell(&self, story_id: &str) -> Option<EmptyCell> {
        let txn = self.yrs_doc().transact();
        is_empty_cell(&txn, story_id).then(|| EmptyCell::new(story_id))
    }

    /// Creates the story of an empty cell with its one empty paragraph, under the `system`
    /// origin so it is never an undo step of its own. Returns whether it created one; a story
    /// that exists, or a cell no live table names, is left alone.
    ///
    /// A cell that a table of a live story names while its own story is missing or removed reads
    /// as one empty paragraph, but the write methods of this type take existing stories only and
    /// answer such a cell with an unknown-story error. Call this first to write into it; the
    /// session write methods do so themselves.
    pub fn ensure_cell_story(&self, story_id: &str) -> bool {
        if self.empty_cell(story_id).is_none() {
            return false;
        }
        let mut txn = self.yrs_doc().transact_mut_with("system");
        if !is_empty_cell(&txn, story_id) {
            return false;
        }
        self.create_empty_cell(&mut txn, story_id);
        true
    }

    /// Writes the story of an empty cell in `txn`: one empty paragraph keyed
    /// [`empty_cell_para_key`].
    pub(crate) fn create_empty_cell(
        &self,
        txn: &mut TransactionMut<'_>,
        story_id: &str,
    ) -> TextRef {
        let pilcrow = write_empty_cell(txn, story_id);
        let key = empty_cell_para_key(story_id);
        IdAllocator::new(self, &*txn).bind(txn, &pilcrow, &key, ParagraphIdOrigin::Authored);
        story_ref(txn, story_id).expect("the empty cell story was just written")
    }
}

#[cfg(test)]
mod tests {
    use super::cell_parent;

    #[test]
    fn cell_parents_follow_the_cell_id_shape() {
        assert_eq!(cell_parent("body:t0:r1c2"), Some("body"));
        assert_eq!(cell_parent("body:t0:r0c2~9701"), Some("body"));
        assert_eq!(cell_parent("body:t0:r0c0:t1:r3c4"), Some("body:t0:r0c0"));
        assert_eq!(cell_parent("body"), None);
        assert_eq!(cell_parent("body:t0"), None);
        assert_eq!(cell_parent("body:sdt0"), None);
        assert_eq!(cell_parent("body:tx:r0c0"), None);
        assert_eq!(cell_parent("body:t0:r0c0~"), None);
    }
}
