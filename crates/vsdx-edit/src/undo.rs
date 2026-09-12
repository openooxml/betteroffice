use std::collections::HashSet;
use std::sync::Arc;

use yrs::sync::time::Clock;
use yrs::{Doc, Origin, ReadTxn, Transact};

use crate::{EditError, EditResult, META, PAGE_ORDER, PAGES, SHEETS, STORIES};

const CAPTURE_TIMEOUT_MS: u64 = 500;

pub struct DiagramUndoManager {
    inner: yrs::undo::UndoManager<()>,
}

impl DiagramUndoManager {
    pub(crate) fn new(doc: &Doc, client_id: u64) -> EditResult<Self> {
        let (order, mut roots) = {
            let txn = doc.transact();
            let order = txn
                .get_array(PAGE_ORDER)
                .ok_or_else(|| EditError::InvalidState("missing page order".to_owned()))?;
            let roots = vec![
                txn.get_map(META),
                txn.get_map(PAGES),
                txn.get_map(SHEETS),
                txn.get_map(STORIES),
            ];
            (order, roots)
        };
        if roots.iter().any(Option::is_none) {
            return Err(EditError::InvalidState("missing diagram root".to_owned()));
        }
        let options = yrs::undo::Options {
            capture_timeout_millis: CAPTURE_TIMEOUT_MS,
            tracked_origins: HashSet::from([Origin::from(client_id)]),
            capture_transaction: None,
            timestamp: default_clock(),
            init_undo_stack: Vec::new(),
            init_redo_stack: Vec::new(),
        };
        let mut inner = yrs::undo::UndoManager::with_options(options);
        inner.expand_scope(doc, &order);
        for root in roots.drain(..).flatten() {
            inner.expand_scope(doc, &root);
        }
        Ok(Self { inner })
    }
    pub fn undo(&mut self) -> bool {
        self.inner.undo_blocking()
    }
    pub fn redo(&mut self) -> bool {
        self.inner.redo_blocking()
    }
    pub fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }
    pub fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }
    pub fn add_undo_barrier(&mut self) {
        self.inner.reset();
    }
}

fn default_clock() -> Arc<dyn Clock> {
    #[cfg(not(all(target_family = "wasm", target_os = "unknown")))]
    {
        Arc::new(|| {
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|elapsed| elapsed.as_millis() as u64)
                .unwrap_or_default()
        })
    }
    #[cfg(all(target_family = "wasm", target_os = "unknown"))]
    {
        use std::sync::atomic::{AtomicU64, Ordering};
        let ticks = AtomicU64::new(0);
        Arc::new(move || ticks.fetch_add(CAPTURE_TIMEOUT_MS + 1, Ordering::Relaxed))
    }
}
