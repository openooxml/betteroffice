use vsdx_parse::{CellLocator, MutationGesture};

use crate::{Expr, ParseLimits, parse};

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum MutationOutcome {
    Allowed {
        target: CellLocator,
        formula: String,
    },
    Refused {
        reason: String,
    },
    Unsupported {
        reason: String,
    },
}

pub trait MutationContext {
    fn current_formula(&self, locator: &CellLocator) -> Result<Option<String>, String>;
    fn resolve_reference(&self, from: &CellLocator, reference: &str)
    -> Result<CellLocator, String>;
    fn lock_enabled(&self, locator: &CellLocator, lock: &str) -> Result<bool, String>;
}

/// Applies documented UI mutation interception before a ShapeSheet formula is written.
pub fn decide(
    context: &impl MutationContext,
    locator: CellLocator,
    gesture: MutationGesture,
    formula: String,
    limits: &ParseLimits,
) -> MutationOutcome {
    const MAX_SETATREF_HOPS: usize = 10;
    let Some(lock) = lock_for(gesture) else {
        return MutationOutcome::Unsupported {
            reason: "structural edits are deferred to phase 5b-2".to_owned(),
        };
    };
    let mut current = locator;
    let mut visited = vec![current.clone()];
    for hops in 0..=MAX_SETATREF_HOPS {
        match context.lock_enabled(&current, lock) {
            Ok(true) => {
                return MutationOutcome::Refused {
                    reason: format!("{lock} protects this {} gesture", gesture_name(gesture)),
                };
            }
            Ok(false) => {}
            Err(reason) => return MutationOutcome::Unsupported { reason },
        }
        let existing = match context.current_formula(&current) {
            Ok(value) => value,
            Err(reason) => return MutationOutcome::Unsupported { reason },
        };
        let Some(existing) = existing else {
            return MutationOutcome::Allowed {
                target: current,
                formula,
            };
        };
        let expression = match parse(existing.trim_start_matches('='), limits) {
            Ok(expression) => expression,
            Err(error) => {
                return MutationOutcome::Unsupported {
                    reason: format!("cannot inspect existing formula: {}", error.message),
                };
            }
        };
        if contains_call(&expression, "GUARD") {
            return MutationOutcome::Refused {
                reason: "GUARD protects the requested cell".to_owned(),
            };
        }
        if contains_call(&expression, "SETATREFEXPR") || contains_call(&expression, "SETATREFEVAL")
        {
            return MutationOutcome::Unsupported {
                reason: "SETATREFEXPR/SETATREFEVAL transformations are not implemented".to_owned(),
            };
        }
        let Expr::Call(name, arguments) = expression else {
            if contains_call(&expression, "SETATREF") {
                return MutationOutcome::Unsupported {
                    reason: "SETATREF is only supported as the complete cell formula".to_owned(),
                };
            }
            return MutationOutcome::Allowed {
                target: current,
                formula,
            };
        };
        if !name.eq_ignore_ascii_case("SETATREF") {
            if contains_call(&Expr::Call(name, arguments), "SETATREF") {
                return MutationOutcome::Unsupported {
                    reason: "SETATREF is only supported as the complete cell formula".to_owned(),
                };
            }
            return MutationOutcome::Allowed {
                target: current,
                formula,
            };
        }
        let Some(Expr::Reference(reference)) = arguments.first() else {
            return MutationOutcome::Unsupported {
                reason: "SETATREF requires a cell-reference first argument".to_owned(),
            };
        };
        if arguments.len() != 1 {
            return MutationOutcome::Unsupported {
                reason: "SETATREF set_expression handling is not implemented".to_owned(),
            };
        }
        if hops == MAX_SETATREF_HOPS {
            return MutationOutcome::Unsupported {
                reason: "SETATREF redirect depth exceeds 10 hops".to_owned(),
            };
        }
        let to = match context.resolve_reference(&current, reference) {
            Ok(to) => to,
            Err(reason) => return MutationOutcome::Unsupported { reason },
        };
        if visited.contains(&to) {
            return MutationOutcome::Unsupported {
                reason: "SETATREF redirect cycle detected".to_owned(),
            };
        }
        visited.push(to.clone());
        current = to;
    }
    unreachable!("redirect loop has a bounded return path")
}

fn contains_call(expression: &Expr, wanted: &str) -> bool {
    match expression {
        Expr::Call(name, arguments) => {
            name.eq_ignore_ascii_case(wanted)
                || arguments
                    .iter()
                    .any(|argument| contains_call(argument, wanted))
        }
        Expr::Unary(expression) => contains_call(expression, wanted),
        Expr::Binary(left, _, right) => contains_call(left, wanted) || contains_call(right, wanted),
        Expr::Number(_, _) | Expr::String(_) | Expr::Reference(_) => false,
    }
}

fn lock_for(gesture: MutationGesture) -> Option<&'static str> {
    match gesture {
        MutationGesture::CellEdit => Some(""),
        MutationGesture::MoveX => Some("LockMoveX"),
        MutationGesture::MoveY => Some("LockMoveY"),
        MutationGesture::ResizeWidth => Some("LockWidth"),
        MutationGesture::ResizeHeight => Some("LockHeight"),
        MutationGesture::ResizeAspect => Some("LockAspect"),
        MutationGesture::TextEdit => Some("LockTextEdit"),
        MutationGesture::Format => Some("LockFormat"),
        MutationGesture::Delete => Some("LockDelete"),
    }
}

fn gesture_name(gesture: MutationGesture) -> &'static str {
    match gesture {
        MutationGesture::CellEdit => "cell-edit",
        MutationGesture::MoveX | MutationGesture::MoveY => "move",
        MutationGesture::ResizeWidth
        | MutationGesture::ResizeHeight
        | MutationGesture::ResizeAspect => "resize",
        MutationGesture::TextEdit => "text-edit",
        MutationGesture::Format => "format",
        MutationGesture::Delete => "delete",
    }
}

#[cfg(test)]
mod tests {
    use std::collections::{BTreeMap, BTreeSet};

    use super::*;

    struct Context {
        formulas: BTreeMap<String, String>,
        locks: BTreeSet<String>,
    }

    impl MutationContext for Context {
        fn current_formula(&self, locator: &CellLocator) -> Result<Option<String>, String> {
            Ok(self.formulas.get(&locator.cell_name).cloned())
        }

        fn resolve_reference(
            &self,
            from: &CellLocator,
            reference: &str,
        ) -> Result<CellLocator, String> {
            Ok(CellLocator {
                cell_name: reference.to_owned(),
                ..from.clone()
            })
        }

        fn lock_enabled(&self, _locator: &CellLocator, lock: &str) -> Result<bool, String> {
            Ok(self.locks.contains(lock))
        }
    }

    fn locator() -> CellLocator {
        CellLocator {
            sheet: vsdx_parse::CellSheet::Page(1),
            shape_id: Some(1),
            section: None,
            row: None,
            cell_name: "Width".to_owned(),
        }
    }

    #[test]
    fn reports_unsupported_setatref_transformations() {
        let context = Context {
            formulas: [(
                "Width".to_owned(),
                "SETATREF(Target, SETATREFEXPR())".to_owned(),
            )]
            .into_iter()
            .collect(),
            locks: BTreeSet::new(),
        };
        assert!(matches!(
            decide(&context, locator(), MutationGesture::ResizeWidth, "2".to_owned(), &ParseLimits::default()),
            MutationOutcome::Unsupported { reason } if reason.contains("SETATREFEXPR")
        ));
    }

    #[test]
    fn each_lock_refuses_only_its_matching_gesture() {
        for (lock, gesture) in [
            ("LockWidth", MutationGesture::ResizeWidth),
            ("LockHeight", MutationGesture::ResizeHeight),
            ("LockMoveX", MutationGesture::MoveX),
            ("LockMoveY", MutationGesture::MoveY),
            ("LockAspect", MutationGesture::ResizeAspect),
            ("LockTextEdit", MutationGesture::TextEdit),
            ("LockFormat", MutationGesture::Format),
        ] {
            let context = Context {
                formulas: BTreeMap::new(),
                locks: [lock.to_owned()].into_iter().collect(),
            };
            assert!(matches!(
                decide(
                    &context,
                    locator(),
                    gesture,
                    "2".to_owned(),
                    &ParseLimits::default()
                ),
                MutationOutcome::Refused { .. }
            ));
            assert!(matches!(
                decide(
                    &context,
                    locator(),
                    MutationGesture::CellEdit,
                    "2".to_owned(),
                    &ParseLimits::default()
                ),
                MutationOutcome::Allowed { .. }
            ));
        }
    }

    #[test]
    fn rejects_setatref_anywhere_except_a_single_root_redirect() {
        for formula in [
            "SETATREF(Target)+1",
            "SETATREF(Target)+SETATREF(Other)",
            "IF(1, SETATREF(Target), 0)",
        ] {
            let context = Context {
                formulas: [("Width".to_owned(), formula.to_owned())]
                    .into_iter()
                    .collect(),
                locks: BTreeSet::new(),
            };
            assert!(matches!(
                decide(
                    &context,
                    locator(),
                    MutationGesture::ResizeWidth,
                    "2".to_owned(),
                    &ParseLimits::default()
                ),
                MutationOutcome::Unsupported { .. }
            ));
        }
    }

    #[test]
    fn rechecks_policy_at_each_redirect_target() {
        let context = Context {
            formulas: [
                ("Width".to_owned(), "SETATREF(Target)".to_owned()),
                ("Target".to_owned(), "GUARD(1)".to_owned()),
            ]
            .into_iter()
            .collect(),
            locks: BTreeSet::new(),
        };
        assert!(matches!(
            decide(
                &context,
                locator(),
                MutationGesture::ResizeWidth,
                "2".to_owned(),
                &ParseLimits::default()
            ),
            MutationOutcome::Refused { .. }
        ));
    }

    #[test]
    fn resolves_redirect_chains_and_rejects_cycles_and_excessive_depth() {
        let context = Context {
            formulas: [
                ("Width".to_owned(), "SETATREF(A)".to_owned()),
                ("A".to_owned(), "SETATREF(B)".to_owned()),
                ("B".to_owned(), "3".to_owned()),
            ]
            .into_iter()
            .collect(),
            locks: BTreeSet::new(),
        };
        assert!(matches!(
            decide(&context, locator(), MutationGesture::ResizeWidth, "2".to_owned(), &ParseLimits::default()),
            MutationOutcome::Allowed { target, .. } if target.cell_name == "B"
        ));

        let cycle = Context {
            formulas: [
                ("Width".to_owned(), "SETATREF(A)".to_owned()),
                ("A".to_owned(), "SETATREF(Width)".to_owned()),
            ]
            .into_iter()
            .collect(),
            locks: BTreeSet::new(),
        };
        assert!(matches!(
            decide(
                &cycle,
                locator(),
                MutationGesture::ResizeWidth,
                "2".to_owned(),
                &ParseLimits::default()
            ),
            MutationOutcome::Unsupported { .. }
        ));

        let mut formulas = BTreeMap::new();
        formulas.insert("Width".to_owned(), "SETATREF(A0)".to_owned());
        for index in 0..10 {
            formulas.insert(format!("A{index}"), format!("SETATREF(A{})", index + 1));
        }
        formulas.insert("A10".to_owned(), "3".to_owned());
        let too_deep = Context {
            formulas,
            locks: BTreeSet::new(),
        };
        assert!(matches!(
            decide(
                &too_deep,
                locator(),
                MutationGesture::ResizeWidth,
                "2".to_owned(),
                &ParseLimits::default()
            ),
            MutationOutcome::Unsupported { .. }
        ));
    }
}
