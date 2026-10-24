use xlsx_model::workbook::{DEFINED_NAME_ENTRY_BYTES, MAX_DEFINED_NAME_BYTES, defined_names_bytes};
use xlsx_model::{Cell, CellRange, CellRef, DefinedName, Sheet, SheetId, Table, Workbook};
use xlsx_ops::{Op, OpError, apply, apply_ops};

fn set(name: &str, formula: &str, scope: Option<SheetId>, hidden: Option<bool>) -> Op {
    Op::SetDefinedName {
        name: name.into(),
        refers_to: formula.into(),
        scope,
        hidden,
    }
}

#[test]
fn defined_name_upsert_preserves_position_and_inverse_restores_exact_vector() {
    let mut book = Workbook::default();
    book.sheets.push(Sheet::new("Data"));
    book.defined_names.push(DefinedName {
        name: "Rate".into(),
        formula: "1".into(),
        local_sheet: None,
        hidden: true,
    });
    let original = book.clone();
    let inverse = apply_ops(
        &mut book,
        &[
            set("Other", "2", None, None),
            set("rate", "3", None, None),
            set("Rate", "4", Some(SheetId(0)), Some(false)),
        ],
    )
    .unwrap();
    assert_eq!(book.defined_names[0].name, "Rate");
    assert_eq!(book.defined_names[0].formula, "3");
    assert!(book.defined_names[0].hidden);
    assert_eq!(book.defined_names.len(), 3);
    apply_ops(&mut book, &inverse).unwrap();
    assert_eq!(book, original);
    let before = book.clone();
    apply(
        &mut book,
        &Op::DeleteDefinedName {
            name: "missing".into(),
            scope: None,
        },
    )
    .unwrap();
    assert_eq!(book, before);
}

#[test]
fn defined_name_validation_rejects_table_conflicts_duplicates_and_invalid_formulas_atomically() {
    let mut book = Workbook::default();
    book.sheets.push(Sheet::new("Data"));
    book.tables.push(Table {
        name: "Items".into(),
        sheet: SheetId(0),
        range: CellRange::new(CellRef::new(0, 0), CellRef::new(1, 0)),
        header_rows: 1,
        totals_rows: 0,
        columns: vec!["Value".into()],
    });
    for op in [
        set("items", "1", None, None),
        set("bad name", "1", None, None),
        set("r1c1", "1", None, None),
        set("Valid", "1+", None, None),
        set("Valid", &"1".repeat(32769), None, None),
        set("Valid", "1", Some(SheetId(1)), None),
    ] {
        let before = book.clone();
        assert!(apply_ops(&mut book, &[set("First", "1", None, None), op]).is_err());
        assert_eq!(book, before);
    }
    book.defined_names = vec![
        DefinedName {
            name: "Rate".into(),
            formula: "1".into(),
            local_sheet: None,
            hidden: false,
        },
        DefinedName {
            name: "RATE".into(),
            formula: "2".into(),
            local_sheet: None,
            hidden: false,
        },
    ];
    let before = book.clone();
    assert!(apply(&mut book, &set("Rate", "3", None, None)).is_err());
    assert!(
        apply(
            &mut book,
            &Op::DeleteDefinedName {
                name: "rate".into(),
                scope: None
            }
        )
        .is_err()
    );
    assert_eq!(book, before);
}

#[test]
fn defined_name_updates_validate_target_names_without_revalidating_other_imports() {
    let mut book = Workbook::default();
    book.sheets.push(Sheet::new("Data"));
    book.defined_names.push(DefinedName {
        name: "bad name".into(),
        formula: "1".into(),
        local_sheet: None,
        hidden: false,
    });
    apply(&mut book, &set("Valid", "2", None, None)).unwrap();
    let before = book.clone();
    assert!(apply(&mut book, &set("bad name", "3", None, None)).is_err());
    assert_eq!(book, before);
}

fn numbered(count: usize, formula: &str) -> Vec<DefinedName> {
    (0..count)
        .map(|index| DefinedName {
            name: format!("Name_{index}"),
            formula: formula.into(),
            local_sheet: None,
            hidden: index % 2 == 0,
        })
        .collect()
}

#[test]
fn a_table_past_the_old_count_cap_takes_additions_deletions_and_undo() {
    let mut book = Workbook::default();
    book.sheets.push(Sheet::new("Data"));
    book.defined_names = numbered(66_391, "#REF!");
    let original = book.clone();
    let inverse = apply_ops(
        &mut book,
        &[
            set("Added", "2", None, None),
            set("Scoped", "3", Some(SheetId(0)), None),
            set("name_7", "4", None, Some(true)),
            Op::DeleteDefinedName {
                name: "Name_8".into(),
                scope: None,
            },
        ],
    )
    .unwrap();
    assert_eq!(book.defined_names.len(), 66_392);
    assert_eq!(book.defined_names[7].formula, "4");
    apply_ops(&mut book, &inverse).unwrap();
    assert_eq!(book, original);
}

#[test]
fn defined_name_byte_bound_refuses_growth_atomically_but_allows_updates() {
    let mut book = Workbook::default();
    book.sheets.push(Sheet::new("Data"));
    book.defined_names = numbered(31, &"1".repeat(1024 * 1024));
    let fits = set("Last", "2", None, None);
    let room = MAX_DEFINED_NAME_BYTES - defined_names_bytes(&book.defined_names);
    let pad = room - (DEFINED_NAME_ENTRY_BYTES + "Last".len() + 1);
    book.defined_names.push(DefinedName {
        name: "Pad".into(),
        formula: "1".repeat(pad - DEFINED_NAME_ENTRY_BYTES - "Pad".len()),
        local_sheet: None,
        hidden: false,
    });
    let original = book.clone();
    assert_eq!(
        apply_ops(&mut book, &[fits.clone(), set("Over", "3", None, None)]),
        Err(OpError::TooManyDefinedNames)
    );
    assert_eq!(book, original);
    let inverse = apply(&mut book, &fits).unwrap();
    assert_eq!(
        defined_names_bytes(&book.defined_names),
        MAX_DEFINED_NAME_BYTES
    );
    let full = book.clone();
    for op in [
        set("Over", "3", Some(SheetId(0)), None),
        set("Last", "22", None, None),
    ] {
        assert_eq!(apply(&mut book, &op), Err(OpError::TooManyDefinedNames));
        assert_eq!(book, full);
    }
    apply(&mut book, &set("last", "4", None, Some(true))).unwrap();
    assert_eq!(book.defined_names.last().unwrap().formula, "4");
    assert!(book.defined_names.last().unwrap().hidden);
    apply_ops(&mut book, &inverse.0).unwrap();
    assert_eq!(book, original);
}

#[test]
fn defined_name_byte_bound_refuses_oversized_replacements_atomically() {
    let mut book = Workbook::default();
    let original = book.clone();
    let defined_names = numbered(2, &"1".repeat(MAX_DEFINED_NAME_BYTES / 2));
    assert_eq!(
        apply(&mut book, &Op::SetDefinedNames { defined_names }),
        Err(OpError::TooManyDefinedNames)
    );
    assert_eq!(book, original);
}

#[test]
fn defined_name_replacements_over_the_bound_are_refused_even_when_they_shrink() {
    let mut book = Workbook {
        defined_names: numbered(2, &"1".repeat(MAX_DEFINED_NAME_BYTES / 2)),
        ..Workbook::default()
    };
    let original = book.clone();
    let defined_names = numbered(2, &"1".repeat(MAX_DEFINED_NAME_BYTES / 2 - 1));
    assert_eq!(
        apply(&mut book, &Op::SetDefinedNames { defined_names }),
        Err(OpError::TooManyDefinedNames)
    );
    assert_eq!(book, original);
}

fn full_table(tails: Vec<DefinedName>) -> Vec<DefinedName> {
    let text = |len: usize| format!("\"{}\"", "a".repeat(len - 2));
    let mut names = numbered(31, &text(1024 * 1024));
    let room = MAX_DEFINED_NAME_BYTES - defined_names_bytes(&names) - defined_names_bytes(&tails);
    names.push(DefinedName {
        name: "Pad".into(),
        formula: text(room - DEFINED_NAME_ENTRY_BYTES - "Pad".len()),
        local_sheet: None,
        hidden: false,
    });
    names.extend(tails);
    assert_eq!(defined_names_bytes(&names), MAX_DEFINED_NAME_BYTES);
    names
}

#[test]
fn structural_edits_and_sheet_renames_refuse_to_grow_a_full_name_table_atomically() {
    let mut book = Workbook::default();
    book.sheets.push(Sheet::new("Data"));
    let mut other = Sheet::new("Other");
    other.set_cell(
        CellRef::new(0, 0),
        Cell {
            formula: Some("Data!A1".into()),
            ..Cell::default()
        },
    );
    book.sheets.push(other);
    book.defined_names = full_table(vec![DefinedName {
        name: "Ref".into(),
        formula: "Data!$A$9".into(),
        local_sheet: None,
        hidden: false,
    }]);
    let original = book.clone();
    for op in [
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Longer".into(),
        },
    ] {
        assert_eq!(apply(&mut book, &op), Err(OpError::TooManyDefinedNames));
        assert_eq!(book, original);
    }
    let inverse = apply(
        &mut book,
        &Op::RenameSheet {
            sheet: SheetId(0),
            name: "Dat".into(),
        },
    )
    .unwrap();
    assert!(defined_names_bytes(&book.defined_names) < MAX_DEFINED_NAME_BYTES);
    apply_ops(&mut book, &inverse.0).unwrap();
    assert_eq!(book, original);
}

fn sheet_refs(refs: &[(&str, &str)]) -> Vec<DefinedName> {
    refs.iter()
        .map(|(name, formula)| DefinedName {
            name: (*name).into(),
            formula: (*formula).into(),
            local_sheet: None,
            hidden: false,
        })
        .collect()
}

#[test]
fn undoing_a_rename_that_shrank_a_full_name_table_restores_it_exactly() {
    let mut book = Workbook::default();
    book.sheets.push(Sheet::new("Data"));
    book.defined_names = full_table(sheet_refs(&[("Ref", "Data!$A$1"), ("Other", "D!$A$1")]));
    let original = book.clone();
    let inverse = apply(
        &mut book,
        &Op::RenameSheet {
            sheet: SheetId(0),
            name: "D".into(),
        },
    )
    .unwrap();
    assert!(defined_names_bytes(&book.defined_names) < MAX_DEFINED_NAME_BYTES);
    let renamed = book.clone();
    let redo = apply_ops(&mut book, &inverse.0).unwrap();
    assert_eq!(book, original);
    apply_ops(&mut book, &redo).unwrap();
    assert_eq!(book, renamed);
}

#[test]
fn undoing_a_rename_keeps_names_that_already_named_the_new_sheet_text() {
    let mut book = Workbook::default();
    book.sheets.push(Sheet::new("Data"));
    book.defined_names = sheet_refs(&[("Other", "D!$A$1")]);
    let original = book.clone();
    let inverse = apply(
        &mut book,
        &Op::RenameSheet {
            sheet: SheetId(0),
            name: "D".into(),
        },
    )
    .unwrap();
    let renamed = book.clone();
    let redo = apply_ops(&mut book, &inverse.0).unwrap();
    assert_eq!(book, original);
    apply_ops(&mut book, &redo).unwrap();
    assert_eq!(book, renamed);
}
