use xlsx_model::{CellRange, CellRef, DefinedName, Sheet, SheetId, Table, Workbook};
use xlsx_ops::{Op, apply, apply_ops};

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
