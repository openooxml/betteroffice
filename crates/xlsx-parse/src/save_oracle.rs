use std::cell::Cell;

thread_local! {
    static LEGACY_SAVE_PATH: Cell<Option<usize>> = const { Cell::new(None) };
}

struct LegacySavePath(Option<usize>);

impl Drop for LegacySavePath {
    fn drop(&mut self) {
        LEGACY_SAVE_PATH.with(|state| {
            let count = state.get().unwrap_or_default();
            state.set(self.0.map(|previous| previous + count));
        });
    }
}

#[doc(hidden)]
pub fn with_legacy_save_path<R>(f: impl FnOnce() -> R) -> (R, usize) {
    let guard = LegacySavePath(LEGACY_SAVE_PATH.with(|state| state.replace(Some(0))));
    let result = f();
    let count = LEGACY_SAVE_PATH.with(|state| state.get().expect("legacy save scope"));
    drop(guard);
    (result, count)
}

pub(crate) fn use_legacy_save_path() -> bool {
    LEGACY_SAVE_PATH.with(|state| match state.get() {
        Some(count) => {
            state.set(Some(count + 1));
            true
        }
        None => false,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn nested_scopes_restore_and_count_legacy_dispatches() {
        assert!(!use_legacy_save_path());
        let (result, count) = with_legacy_save_path(|| {
            assert!(use_legacy_save_path());
            let (result, count) = with_legacy_save_path(|| {
                assert!(use_legacy_save_path());
                42
            });
            assert_eq!((result, count), (42, 1));
            assert!(use_legacy_save_path());
            result
        });
        assert_eq!((result, count), (42, 3));
        assert!(!use_legacy_save_path());
    }

    #[test]
    fn legacy_selection_is_thread_local() {
        let (_, count) = with_legacy_save_path(|| {
            std::thread::spawn(|| assert!(!use_legacy_save_path()))
                .join()
                .unwrap();
            assert!(use_legacy_save_path());
        });
        assert_eq!(count, 1);
        assert!(!use_legacy_save_path());
    }

    #[test]
    fn panics_restore_the_previous_scope() {
        let (_, count) = with_legacy_save_path(|| {
            assert!(use_legacy_save_path());
            let result = std::panic::catch_unwind(|| {
                with_legacy_save_path(|| {
                    assert!(use_legacy_save_path());
                    panic!("legacy save panic");
                });
            });
            assert!(result.is_err());
            assert!(use_legacy_save_path());
        });
        assert_eq!(count, 3);
        assert!(!use_legacy_save_path());
        let result = std::panic::catch_unwind(|| {
            with_legacy_save_path(|| panic!("legacy save panic"));
        });
        assert!(result.is_err());
        assert!(!use_legacy_save_path());
    }
}
