use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRef, CellValue, DefinedName, DefinedNamesRequest, Op, Sheet,
    SheetId, Workbook, WorkbookModel,
};
use xlsx_model::CellProvider;
use xlsx_model::workbook::{DEFINED_NAME_ENTRY_BYTES, MAX_DEFINED_NAME_BYTES, defined_names_bytes};

const NAME_COUNT: usize = 66_391;

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn names() -> String {
    (0..NAME_COUNT)
        .map(|index| match index {
            40_000 => r#"<definedName name="Rate">Data!$A$1</definedName>"#.to_owned(),
            65_999 => {
                r#"<definedName name="_Tail" hidden="1">Data!$A$1*3</definedName>"#.to_owned()
            }
            _ if index % 2 == 0 => {
                format!(r#"<definedName name="_junk{index}" hidden="1">#REF!</definedName>"#)
            }
            _ => format!(
                r#"<definedName name="Area{index}" localSheetId="0" comment="a &amp; b">Data!$B${}</definedName>"#,
                index % 1000 + 1
            ),
        })
        .collect()
}

fn package(names: &str) -> Vec<u8> {
    let mut model = WorkbookModel::default();
    let mut sheet = Sheet::new("Data");
    sheet.set_cell(
        cell("A1"),
        Cell {
            value: number(5.0),
            ..Cell::default()
        },
    );
    for (address, formula) in [("C1", "Rate*2"), ("C2", "_Tail+1")] {
        sheet.set_cell(
            cell(address),
            Cell {
                value: number(0.0),
                formula: Some(formula.into()),
                style: None,
            },
        );
    }
    model.sheets.push(sheet);
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let workbook = &mut parts
        .iter_mut()
        .find(|(path, _)| path == "xl/workbook.xml")
        .unwrap()
        .1;
    *workbook = std::str::from_utf8(workbook)
        .unwrap()
        .replace(
            "</workbook>",
            &format!("<definedNames>{names}</definedNames></workbook>"),
        )
        .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_workbook_xml(saved: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(saved).unwrap();
    let (_, bytes) = parts
        .iter()
        .find(|(path, _)| path == "xl/workbook.xml")
        .unwrap();
    String::from_utf8(bytes.clone()).unwrap()
}

#[test]
fn a_name_table_past_the_old_count_cap_opens_evaluates_and_saves_verbatim() {
    let names = names();
    let mut workbook = Workbook::open(&package(&names)).unwrap();
    assert_eq!(workbook.model().defined_names.len(), NAME_COUNT);

    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(workbook.model().value(SheetId(0), cell("C1")), number(10.0));
    assert_eq!(workbook.model().value(SheetId(0), cell("C2")), number(16.0));

    workbook
        .edit_cell(SheetId(0), cell("A1"), "7", CalculationOptions::default())
        .unwrap();
    assert_eq!(workbook.model().value(SheetId(0), cell("C1")), number(14.0));
    assert_eq!(workbook.model().value(SheetId(0), cell("C2")), number(22.0));

    let saved = workbook.save().unwrap();
    let container = format!("<definedNames>{names}</definedNames>");
    assert!(saved_workbook_xml(&saved).contains(&container));
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().defined_names,
        workbook.model().defined_names
    );
}

#[test]
fn a_model_whose_name_table_is_over_the_byte_bound_is_refused() {
    let formula = "1".repeat(xlsx_calc::lexer::MAX_FORMULA_BYTES);
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    let count = xlsx_model::workbook::MAX_DEFINED_NAME_BYTES / formula.len() + 1;
    model.defined_names = (0..count)
        .map(|index| xlsx_model::DefinedName {
            name: format!("N_{index}"),
            formula: formula.clone(),
            local_sheet: None,
            hidden: false,
        })
        .collect();
    let failure = match Workbook::from_model(model) {
        Ok(_) => panic!("an over-bound name table must be refused"),
        Err(failure) => failure.to_string(),
    };
    assert!(failure.contains("defined name table"), "{failure}");
}

#[test]
fn undoing_a_sheet_rename_restores_the_name_table_exactly() {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    model.defined_names = [("Ref", "Data!$A$1"), ("Other", "D!$A$1")]
        .into_iter()
        .map(|(name, formula)| xlsx_model::DefinedName {
            name: name.into(),
            formula: formula.into(),
            local_sheet: None,
            hidden: false,
        })
        .collect();
    let original = model.defined_names.clone();
    let mut workbook = Workbook::from_model(model).unwrap();
    workbook
        .apply_ops(
            vec![xlsx_ops::Op::RenameSheet {
                sheet: SheetId(0),
                name: "D".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let renamed = workbook.model().defined_names.clone();
    assert_eq!(renamed[0].formula, "D!$A$1");
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model().defined_names, original);
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model().defined_names, renamed);
}

fn mixed_names(count: usize) -> String {
    (0..count)
        .map(|index| match index % 10 {
            0..=5 => format!(
                r#"<definedName name="Name{index}">Data!$A${}</definedName>"#,
                index % 100 + 1
            ),
            6 | 7 => format!(
                r#"<definedName name="Local{index}" localSheetId="0">Data!$A$1:$A${}</definedName>"#,
                index % 100 + 1
            ),
            8 => format!(
                r#"<definedName name="Name{}" localSheetId="0">Data!$A$2</definedName>"#,
                index - 8
            ),
            _ => format!(
                r#"<definedName name="Ref{index}">Name{}*2+SUM(Name{},1)</definedName>"#,
                index - 9,
                index - 8
            ),
        })
        .collect()
}

fn recommit_ops(workbook: &Workbook) -> Vec<Op> {
    let mut ops = Vec::new();
    let mut offset = 0;
    loop {
        let page = workbook
            .list_defined_names(&DefinedNamesRequest {
                offset: Some(offset),
                limit: Some(100),
                ..Default::default()
            })
            .unwrap();
        ops.extend(page.items.into_iter().map(|info| Op::SetDefinedName {
            name: info.name,
            refers_to: info.refers_to,
            scope: info.scope.map(SheetId),
            hidden: None,
        }));
        match page.next_offset {
            Some(next) if next > offset => offset = next,
            _ => return ops,
        }
    }
}

fn set_name(name: &str, refers_to: &str, scope: Option<u32>, hidden: Option<bool>) -> Op {
    Op::SetDefinedName {
        name: name.into(),
        refers_to: refers_to.into(),
        scope: scope.map(SheetId),
        hidden,
    }
}

#[test]
fn recommitting_every_name_in_one_batch_keeps_the_version_and_the_saved_bytes() {
    let bytes = package(&mixed_names(300));
    let mut workbook = Workbook::open(&bytes).unwrap();
    let version = workbook.version();
    let ops = recommit_ops(&workbook);
    assert_eq!(ops.len(), 300);

    let result = workbook
        .apply_ops(ops, CalculationOptions::default())
        .unwrap();

    assert!(!result.applied);
    assert_eq!(workbook.version(), version);
    assert!(!workbook.can_undo());
    assert_eq!(workbook.save().unwrap(), bytes);
}

#[test]
fn a_batch_of_name_edits_matches_the_same_edits_one_at_a_time_and_undoes_exactly() {
    let bytes = package(&mixed_names(60));
    let ops = vec![
        set_name("Name1", "Data!$A$7", None, None),
        set_name("name3", "Data!$A$8", None, Some(true)),
        set_name("Local6", "Data!$A$1:$A$9", Some(0), Some(true)),
        set_name("Fresh", "Name1+1", None, None),
        Op::DeleteDefinedName {
            name: "Name2".into(),
            scope: None,
        },
        set_name("Name2", "Data!$A$3", None, None),
        set_name("Name0", "Data!$A$2", Some(0), None),
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Facts".into(),
        },
        set_name("Later", "Facts!$A$5", Some(0), None),
        set_name("Name4", "Facts!$A$6", None, None),
        set_name("Fresh", "Name1+2", None, None),
    ];
    let mut batched = Workbook::open(&bytes).unwrap();
    let original = batched.model().defined_names.clone();
    batched
        .apply_ops(ops.clone(), CalculationOptions::default())
        .unwrap();
    let mut stepped = Workbook::open(&bytes).unwrap();
    for op in ops {
        stepped
            .apply_ops(vec![op], CalculationOptions::default())
            .unwrap();
    }

    let edited = batched.model().defined_names.clone();
    assert_eq!(edited, stepped.model().defined_names);
    assert_eq!(
        batched.model().value(SheetId(0), cell("C1")),
        stepped.model().value(SheetId(0), cell("C1"))
    );
    assert_eq!(batched.save().unwrap(), stepped.save().unwrap());
    assert_eq!(batched.history_state().undo_depth, 1);

    batched.undo(CalculationOptions::default()).unwrap();
    assert_eq!(batched.model().defined_names, original);
    assert_eq!(batched.model().sheets[0].name, "Data");
    batched.redo(CalculationOptions::default()).unwrap();
    assert_eq!(batched.model().defined_names, edited);
}

#[test]
fn a_batch_still_refuses_an_ambiguous_name_and_changes_nothing() {
    let filter = r#"<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$A$2</definedName>"#;
    let bytes = package(&format!("{}{filter}{filter}", mixed_names(40)));
    for ambiguous_at in [0, 2, 3] {
        let mut workbook = Workbook::open(&bytes).unwrap();
        let version = workbook.version();
        let names = workbook.model().defined_names.clone();
        let mut ops = vec![
            set_name("Name0", "Data!$A$1", None, None),
            set_name("Name1", "Data!$A$9", None, None),
            set_name("Fresh", "Data!$A$3", None, None),
        ];
        ops.insert(
            ambiguous_at,
            set_name("_XLNM._FilterDatabase", "Data!$A$1:$A$2", Some(0), None),
        );

        let failure = workbook
            .apply_ops(ops, CalculationOptions::default())
            .unwrap_err()
            .to_string();

        println!("EQ ambiguous-at-{ambiguous_at} {failure}");
        assert!(
            failure.contains("ambiguous defined name in this scope"),
            "{failure}"
        );
        assert_eq!(workbook.version(), version);
        assert_eq!(workbook.model().defined_names, names);
        assert_eq!(workbook.save().unwrap(), bytes);
    }
}

fn digest(bytes: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    let hash = Sha256::digest(bytes);
    let hex: String = hash.iter().map(|byte| format!("{byte:02x}")).collect();
    format!("{}:{hex}", bytes.len())
}

fn names_digest(names: &[DefinedName]) -> String {
    digest(format!("{names:?}").as_bytes())
}

/// A book whose name table charges exactly `room` bytes less than the bound.
fn nearly_full_package(room: usize) -> Vec<u8> {
    let target = MAX_DEFINED_NAME_BYTES - room;
    let bytes = {
        let mut model = WorkbookModel::default();
        let mut sheet = Sheet::new("Data");
        sheet.set_cell(
            cell("A1"),
            Cell {
                value: number(5.0),
                ..Cell::default()
            },
        );
        sheet.set_cell(
            cell("C1"),
            Cell {
                value: number(0.0),
                formula: Some("Name0*2".into()),
                style: None,
            },
        );
        model.sheets.push(sheet);
        let defined = |name: String, formula: String| DefinedName {
            name,
            formula,
            local_sheet: None,
            hidden: false,
        };
        model.defined_names = vec![
            defined("Name0".into(), "Data!$A$1".into()),
            defined("Name1".into(), "Data!$A$2".into()),
        ];
        let bulk = "0".repeat(30_000);
        let pad_entry = DEFINED_NAME_ENTRY_BYTES + "Pad".len();
        for index in 0.. {
            let next = defined(format!("Bulk{index}"), bulk.clone());
            if defined_names_bytes(&model.defined_names) + next.retained_bytes() + pad_entry + 1
                > target
            {
                break;
            }
            model.defined_names.push(next);
        }
        let pad = target - defined_names_bytes(&model.defined_names) - pad_entry;
        model
            .defined_names
            .push(defined("Pad".into(), "0".repeat(pad)));
        assert_eq!(defined_names_bytes(&model.defined_names), target);
        ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap()
    };
    let opened = Workbook::open(&bytes).unwrap();
    assert_eq!(defined_names_bytes(&opened.model().defined_names), target);
    bytes
}

/// Fails naming each package part where two saves differ.
fn assert_same_save(label: &str, left: &[u8], right: &[u8]) {
    if left == right {
        return;
    }
    let (left, right) = (
        ooxml_opc::unzip_parts(left).unwrap(),
        ooxml_opc::unzip_parts(right).unwrap(),
    );
    let mut differing = Vec::new();
    for (path, bytes) in &left {
        let other = right
            .iter()
            .find(|(other, _)| other == path)
            .map(|(_, b)| b);
        if other != Some(bytes) {
            differing.push(format!(
                "{path}:\n  left  {}\n  right {}",
                String::from_utf8_lossy(bytes),
                other.map_or("<missing>".into(), |b| String::from_utf8_lossy(b))
            ));
        }
    }
    panic!("{label}: saves differ\n{}", differing.join("\n"));
}

#[test]
fn a_run_refused_at_its_third_op_by_the_byte_bound_leaves_the_book_as_it_was() {
    let fresh = DEFINED_NAME_ENTRY_BYTES + "Fresh".len() + 1;
    let over = DEFINED_NAME_ENTRY_BYTES + "Over".len() + 1;
    let bytes = nearly_full_package(1 + fresh + over - 1);
    let mut workbook = Workbook::open(&bytes).unwrap();
    let version = workbook.version();
    let names = workbook.model().defined_names.clone();
    let ops = vec![
        set_name("Name0", "Data!$A$12", None, None),
        set_name("Fresh", "1", None, None),
        set_name("Over", "2", None, None),
        set_name("Name1", "Data!$A$3", None, None),
    ];

    let failure = workbook
        .apply_ops(ops, CalculationOptions::default())
        .unwrap_err()
        .to_string();

    println!("EQ byte-bound-third {failure}");
    assert_eq!(failure, "defined name table exceeded cap");
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.model().defined_names, names);
    assert!(!workbook.can_undo());
    assert_eq!(workbook.save().unwrap(), bytes);
}

#[test]
fn undo_and_redo_of_name_runs_around_a_cell_edit_restore_the_name_table_exactly() {
    let bytes = package(&mixed_names(60));
    let mut workbook = Workbook::open(&bytes).unwrap();
    let names_before = workbook.model().defined_names.clone();
    let saved_before = workbook.save().unwrap();
    let ops = vec![
        set_name("Name1", "Data!$A$7", None, Some(true)),
        set_name("Fresh", "Name1+1", None, None),
        Op::DeleteDefinedName {
            name: "Name2".into(),
            scope: None,
        },
        set_name("Local6", "Data!$A$1:$A$4", Some(0), None),
        Op::SetCell {
            sheet: SheetId(0),
            at: cell("A1"),
            cell: Cell {
                value: number(7.0),
                ..Cell::default()
            }
            .into(),
        },
        set_name("Name3", "Data!$A$9", None, None),
        set_name("fresh", "Name3+Data!$A$1", None, Some(true)),
        Op::DeleteDefinedName {
            name: "Name4".into(),
            scope: None,
        },
        set_name("Later", "Data!$A$5", Some(0), None),
    ];

    workbook
        .apply_ops(ops, CalculationOptions::default())
        .unwrap();
    let names_after = workbook.model().defined_names.clone();
    let saved_after = workbook.save().unwrap();
    assert_eq!(workbook.history_state().undo_depth, 1);
    assert_ne!(names_after, names_before);

    workbook.undo(CalculationOptions::default()).unwrap();
    let names_undone = workbook.model().defined_names.clone();
    let saved_undone = workbook.save().unwrap();
    workbook.redo(CalculationOptions::default()).unwrap();
    let names_redone = workbook.model().defined_names.clone();
    let saved_redone = workbook.save().unwrap();

    for (label, names, saved) in [
        ("before", &names_before, &saved_before),
        ("after", &names_after, &saved_after),
        ("undone", &names_undone, &saved_undone),
        ("redone", &names_redone, &saved_redone),
    ] {
        println!(
            "EQ mixed-{label} names={} save={}",
            names_digest(names),
            digest(saved)
        );
    }
    let out = std::path::Path::new(env!("CARGO_TARGET_TMPDIR"));
    std::fs::write(out.join("mixed-saved-before.xlsx"), &saved_before).unwrap();
    std::fs::write(out.join("mixed-saved-undone.xlsx"), &saved_undone).unwrap();
    assert_eq!(names_undone, names_before);
    assert_eq!(names_redone, names_after);
    assert_same_save("redone", &saved_redone, &saved_after);
    assert_eq!(
        digest(&saved_undone),
        "2013:845bca35a4d85902b3522022bd0e20571e18c58452c1908a7a965e4451263a54",
        "the save after undo differs from the one captured from main's per-op inverses"
    );
}
