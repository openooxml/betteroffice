use betteroffice_xlsx::{
    CalculationOptions, CellAddress, CellRef, EditRequest, Error, Op, ProposalEditInput,
    ProposalRequest, Sheet, SheetId, Workbook, WorkbookModel,
};

fn cell(value: &str) -> CellRef {
    CellRef::parse_a1(value).unwrap()
}

fn model() -> WorkbookModel {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    model
}

fn stage(workbook: &mut Workbook, input: &str) -> String {
    workbook
        .propose(
            ProposalRequest {
                agent_id: "audit-agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("C1"),
                    input: input.into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap()
        .id
}

#[test]
fn dependency_drift_requires_review_of_the_refreshed_preview() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "10", CalculationOptions::default())
        .unwrap();
    let id = stage(&mut workbook, "=A1*2");
    assert_eq!(workbook.proposals()[0].edits[0].new_text, "20");
    workbook
        .edit_cell(SheetId(0), cell("A1"), "99", CalculationOptions::default())
        .unwrap();
    assert_eq!(workbook.proposals()[0].edits[0].new_text, "20");
    assert!(matches!(
        workbook.accept_proposal(&id, false, CalculationOptions::default()),
        Err(betteroffice_xlsx::Error::StaleProposal(_))
    ));
    assert_eq!(workbook.cell(SheetId(0), cell("C1")).unwrap().input, "");
    assert_eq!(workbook.proposals()[0].edits[0].new_text, "198");
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    assert_eq!(
        workbook
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("C1"))
            .unwrap()
            .value,
        betteroffice_xlsx::CellValue::Number { value: 198.0 }
    );
}

#[test]
fn unrelated_remote_edits_preserve_pending_local_proposals() {
    let mut local = Workbook::from_model_collaborative(model(), 901).unwrap();
    let mut remote = Workbook::from_model_collaborative(model(), 902).unwrap();
    stage(&mut local, "42");
    let update = local
        .encode_diff_v1(&remote.encode_state_vector_v1())
        .unwrap();
    remote
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert!(remote.proposals().is_empty());
    assert_eq!(local.proposals().len(), 1);
    remote
        .edit_cell(
            SheetId(0),
            cell("Z99"),
            "unrelated",
            CalculationOptions::default(),
        )
        .unwrap();
    let update = remote
        .encode_diff_v1(&local.encode_state_vector_v1())
        .unwrap();
    local
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert_eq!(local.proposals().len(), 1);
    local
        .accept_proposal("p1", false, CalculationOptions::default())
        .unwrap();
    assert_eq!(local.cell(SheetId(0), cell("C1")).unwrap().input, "42");
}

#[test]
fn collaborative_acceptance_is_one_undo_step() {
    let mut workbook = Workbook::from_model_collaborative(model(), 903).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "10", CalculationOptions::default())
        .unwrap();
    let id = stage(&mut workbook, "42");
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.cell(SheetId(0), cell("C1")).unwrap().input, "");
    assert_eq!(workbook.cell(SheetId(0), cell("A1")).unwrap().input, "10");
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.cell(SheetId(0), cell("C1")).unwrap().input, "42");
}

#[test]
fn acceptance_undo_and_redo_preserve_a_peer_edit_and_converge() {
    let mut local = Workbook::from_model_collaborative(model(), 911).unwrap();
    let mut peer = Workbook::from_model_collaborative(model(), 912).unwrap();
    let id = stage(&mut local, "42");
    local
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    let update = local
        .encode_diff_v1(&peer.encode_state_vector_v1())
        .unwrap();
    peer.apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    peer.edit_cell(
        SheetId(0),
        cell("Z99"),
        "keep",
        CalculationOptions::default(),
    )
    .unwrap();
    let update = peer
        .encode_diff_v1(&local.encode_state_vector_v1())
        .unwrap();
    local
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    local.undo(CalculationOptions::default()).unwrap();
    assert_eq!(local.cell(SheetId(0), cell("C1")).unwrap().input, "");
    assert_eq!(local.cell(SheetId(0), cell("Z99")).unwrap().input, "keep");
    let update = local
        .encode_diff_v1(&peer.encode_state_vector_v1())
        .unwrap();
    peer.apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert_eq!(peer.cell(SheetId(0), cell("C1")).unwrap().input, "");
    local.redo(CalculationOptions::default()).unwrap();
    let update = local
        .encode_diff_v1(&peer.encode_state_vector_v1())
        .unwrap();
    peer.apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert_eq!(peer.cell(SheetId(0), cell("C1")).unwrap().input, "42");
    assert_eq!(peer.cell(SheetId(0), cell("Z99")).unwrap().input, "keep");
}

#[test]
fn a_batch_that_moves_a_dependency_leaves_the_proposal_to_its_review() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "10", CalculationOptions::default())
        .unwrap();
    let id = stage(&mut workbook, "=A1*2");
    let request: EditRequest = serde_json::from_value(serde_json::json!({
        "expectVersion": workbook.version(),
        "steps": [{
            "op": "setCellInputs",
            "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": "A1" } },
            "inputs": [["50"]],
        }],
    }))
    .unwrap();
    workbook.apply_edits(&request).unwrap().unwrap();
    assert_eq!(workbook.proposals().len(), 1);
    assert_eq!(workbook.proposals()[0].edits[0].new_text, "20");
    assert!(matches!(
        workbook.accept_proposal(&id, false, CalculationOptions::default()),
        Err(betteroffice_xlsx::Error::StaleProposal(_))
    ));
    assert_eq!(workbook.proposals()[0].edits[0].new_text, "100");
    let version = workbook.version();
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    assert_ne!(workbook.version(), version);
    assert_eq!(
        workbook.cell(SheetId(0), cell("C1")).unwrap().input,
        "=A1*2"
    );
}

#[test]
fn standalone_acceptance_commits_the_authority_with_the_model() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    let id = stage(&mut workbook, "42");
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    let mut replica = Workbook::from_model_collaborative(model(), 931).unwrap();
    replica
        .apply_update_v1(
            &workbook.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(replica.cell(SheetId(0), cell("C1")).unwrap().input, "42");
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.cell(SheetId(0), cell("C1")).unwrap().input, "");
    let mut undone = Workbook::from_model_collaborative(model(), 932).unwrap();
    undone
        .apply_update_v1(
            &workbook.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(undone.cell(SheetId(0), cell("C1")).unwrap().input, "");
}

fn apply(workbook: &mut Workbook, op: Op) {
    workbook
        .apply_ops(vec![op], CalculationOptions::default())
        .unwrap();
}

fn assert_stale(workbook: &mut Workbook, id: &str, target: &str, force: bool) {
    match workbook.accept_proposal(id, force, CalculationOptions::default()) {
        Err(Error::StaleProposal(targets)) => assert_eq!(
            targets,
            vec![CellAddress {
                sheet: SheetId(0),
                cell: cell(target),
            }]
        ),
        result => panic!("expected stale target, got {result:?}"),
    }
    assert_eq!(workbook.proposals().len(), 1);
}

#[test]
fn inserted_rows_move_the_target_before_state_and_formula_input() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "10", CalculationOptions::default())
        .unwrap();
    workbook
        .edit_cell(SheetId(0), cell("C1"), "=A1", CalculationOptions::default())
        .unwrap();
    let id = stage(&mut workbook, "=A1*2");
    apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 2,
        },
    );
    let edit = &workbook.proposals()[0].edits[0];
    assert_eq!(edit.a1, "C3");
    assert_eq!(edit.input, "=A3*2");
    assert_eq!(edit.old_state.formula.as_deref(), Some("A3"));
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    assert_eq!(
        workbook.cell(SheetId(0), cell("C3")).unwrap().input,
        "=A3*2"
    );
}

#[test]
fn deleted_rows_are_stale_even_when_the_target_was_empty_and_force_refuses() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    let id = stage(&mut workbook, "42");
    apply(
        &mut workbook,
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    assert_stale(&mut workbook, &id, "C1", false);
    assert_stale(&mut workbook, &id, "C1", true);
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.proposals()[0].edits[0].a1, "C1");
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_stale(&mut workbook, &id, "C1", false);
    workbook.undo(CalculationOptions::default()).unwrap();
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    assert_eq!(workbook.cell(SheetId(0), cell("C1")).unwrap().input, "42");
}

#[test]
fn column_deletions_shift_or_stale_targets_and_undo_restores_them() {
    for at in [0, 2] {
        let mut workbook = Workbook::from_model(model()).unwrap();
        let id = stage(&mut workbook, "42");
        apply(
            &mut workbook,
            Op::DeleteCols {
                sheet: SheetId(0),
                at,
                count: 1,
            },
        );
        if at == 2 {
            assert_stale(&mut workbook, &id, "C1", false);
            assert_stale(&mut workbook, &id, "C1", true);
        } else {
            assert_eq!(workbook.proposals()[0].edits[0].a1, "B1");
        }
        workbook.undo(CalculationOptions::default()).unwrap();
        assert_eq!(workbook.proposals()[0].edits[0].a1, "C1");
        workbook
            .accept_proposal(&id, false, CalculationOptions::default())
            .unwrap();
    }
}

#[test]
fn removing_a_sheet_stales_only_its_edits_and_undo_restores_their_identity() {
    let mut source = model();
    source.sheets.push(Sheet::new("Other"));
    let mut workbook = Workbook::from_model(source).unwrap();
    let id = stage(&mut workbook, "42");
    let other = workbook
        .propose(
            ProposalRequest {
                agent_id: "other".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(1),
                    cell: cell("C1"),
                    input: "99".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap()
        .id;
    apply(&mut workbook, Op::RemoveSheet { index: 0 });
    assert_eq!(workbook.proposals().len(), 2);
    assert!(matches!(
        workbook.accept_proposal(&id, false, CalculationOptions::default()),
        Err(Error::StaleProposal(_))
    ));
    assert!(matches!(
        workbook.accept_proposal(&id, true, CalculationOptions::default()),
        Err(Error::StaleProposal(_))
    ));
    assert_eq!(workbook.proposals()[1].edits[0].sheet, 0);
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.proposals()[1].edits[0].sheet, 1);
    workbook
        .accept_proposal(&other, false, CalculationOptions::default())
        .unwrap();
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    assert_eq!(workbook.cell(SheetId(0), cell("C1")).unwrap().input, "42");
    assert_eq!(workbook.cell(SheetId(1), cell("C1")).unwrap().input, "99");
}

#[test]
fn sheet_rename_move_and_add_keep_proposals_attached_to_their_sheet() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    let id = stage(&mut workbook, "=Data!A1+42");
    apply(
        &mut workbook,
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Renamed".into(),
        },
    );
    apply(
        &mut workbook,
        Op::AddSheet {
            index: 0,
            name: "Other".into(),
        },
    );
    assert_eq!(workbook.proposals()[0].edits[0].sheet, 1);
    apply(&mut workbook, Op::MoveSheet { from: 1, to: 0 });
    assert_eq!(workbook.proposals()[0].edits[0].sheet, 0);
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    assert_eq!(
        workbook.cell(SheetId(0), cell("C1")).unwrap().input,
        "=Renamed!A1+42"
    );
}

#[test]
fn undo_redo_rebases_proposals_staged_after_an_insert_without_resurrecting_rejections() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "1", CalculationOptions::default())
        .unwrap();
    apply(
        &mut workbook,
        Op::InsertCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    let id = stage(&mut workbook, "42");
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.proposals()[0].edits[0].a1, "B1");
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.proposals()[0].edits[0].a1, "C1");
    workbook.reject_proposal(&id);
    workbook.undo(CalculationOptions::default()).unwrap();
    assert!(workbook.proposals().is_empty());
}

#[test]
fn collaborative_history_and_rejected_remote_structure_keep_proposals() {
    let mut local = Workbook::from_model_collaborative(model(), 951).unwrap();
    local
        .edit_cell(SheetId(0), cell("Z1"), "1", CalculationOptions::default())
        .unwrap();
    let id = stage(&mut local, "42");
    local.undo(CalculationOptions::default()).unwrap();
    assert_eq!(local.proposals().len(), 1);
    local.redo(CalculationOptions::default()).unwrap();
    assert_eq!(local.proposals().len(), 1);
    let mut remote = Workbook::from_model(model()).unwrap();
    remote
        .edit_cell(SheetId(0), cell("A1"), "1", CalculationOptions::default())
        .unwrap();
    apply(
        &mut remote,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    assert!(
        local
            .apply_update_v1(
                &remote.encode_state_as_update_v1(),
                CalculationOptions::default()
            )
            .is_err()
    );
    assert_eq!(local.proposals()[0].edits[0].a1, "C1");
    local
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
}

#[test]
fn recovery_keeps_deleted_targets_stale_and_preserves_their_undo() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    let id = stage(&mut workbook, "42");
    apply(
        &mut workbook,
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    let checkpoint = workbook
        .recovery_checkpoint_with_history(3, None, true)
        .unwrap();
    let (mut restored, _, _, _) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_stale(&mut restored, &id, "C1", false);
    restored.undo(CalculationOptions::default()).unwrap();
    restored
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    assert_eq!(restored.cell(SheetId(0), cell("C1")).unwrap().input, "42");
}

#[test]
fn deleting_a_referenced_cell_rewrites_the_input_and_undo_restores_it_exactly() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "10", CalculationOptions::default())
        .unwrap();
    let id = stage(&mut workbook, "=SUM($A$1:$B$1)");
    apply(
        &mut workbook,
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 0,
            count: 2,
        },
    );
    assert_eq!(workbook.proposals()[0].edits[0].a1, "A1");
    assert!(workbook.proposals()[0].edits[0].input.contains("#REF!"));
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.proposals()[0].edits[0].input, "=SUM($A$1:$B$1)");
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
}

#[test]
fn structural_rebasing_preserves_an_existing_stale_before_state() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    let id = stage(&mut workbook, "42");
    workbook
        .edit_cell(SheetId(0), cell("C1"), "99", CalculationOptions::default())
        .unwrap();
    apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    assert_stale(&mut workbook, &id, "C2", false);
    workbook
        .accept_proposal(&id, true, CalculationOptions::default())
        .unwrap();
    assert_eq!(workbook.cell(SheetId(0), cell("C2")).unwrap().input, "42");
}

#[test]
fn formula_like_plain_text_inputs_are_not_rewritten() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    apply(
        &mut workbook,
        Op::SetRangeNumberFormat {
            sheet: SheetId(0),
            range: betteroffice_xlsx::CellRange::new(cell("C1"), cell("C1")),
            format: betteroffice_xlsx::NumberFormatMutation::PlainText,
        },
    );
    let id = stage(&mut workbook, "=A1");
    apply(
        &mut workbook,
        Op::InsertCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    assert_eq!(workbook.proposals()[0].edits[0].input, "=A1");
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    assert_eq!(workbook.cell(SheetId(0), cell("D1")).unwrap().input, "=A1");
}

#[test]
fn structural_edits_rewrite_qualified_proposal_formulas_on_other_sheets() {
    let mut source = model();
    source.sheets.push(Sheet::new("Other"));
    let mut workbook = Workbook::from_model(source).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "10", CalculationOptions::default())
        .unwrap();
    let id = workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(1),
                    cell: cell("C1"),
                    input: "=Data!A1*2".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap()
        .id;
    apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    let edit = &workbook.proposals()[0].edits[0];
    assert_eq!(edit.a1, "C1");
    assert_eq!(edit.input, "=Data!A2*2");
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    assert_eq!(
        workbook.cell(SheetId(1), cell("C1")).unwrap().input,
        "=Data!A2*2"
    );
}

#[test]
fn formulas_that_cannot_be_rewritten_become_stale_and_undo_restores_them() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    let id = stage(&mut workbook, "=SUM(Table1[Amount])");
    apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    assert_stale(&mut workbook, &id, "C2", false);
    assert_stale(&mut workbook, &id, "C2", true);
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.proposals()[0].edits[0].stale, None);
    assert_eq!(
        workbook.proposals()[0].edits[0].input,
        "=SUM(Table1[Amount])"
    );
}

#[test]
fn hyperlinks_and_style_restores_keep_pending_proposals() {
    let mut source = model();
    source.styles.cell_xfs.push(xlsx_model::Xf::default());
    source.sheets[0].row_styles.insert(2, 0);
    source.sheets[0].col_styles.push(xlsx_model::ColStyle {
        first: 3,
        last: 3,
        xf: 0,
    });
    let mut workbook = Workbook::from_model(source).unwrap();
    let id = stage(&mut workbook, "42");
    let proposal = workbook.proposals()[0].clone();
    for op in [
        Op::SetHyperlinks {
            sheet: SheetId(0),
            hyperlinks: vec![betteroffice_xlsx::Hyperlink {
                range: betteroffice_xlsx::CellRange::new(cell("A1"), cell("A1")),
                external_target: Some("https://example.com".into()),
                location: None,
                tooltip: None,
                display: None,
            }],
        },
        Op::InsertRows {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(0),
            at: 3,
            count: 1,
        },
    ] {
        apply(&mut workbook, op);
        assert_eq!(workbook.proposals(), std::slice::from_ref(&proposal));
        workbook.undo(CalculationOptions::default()).unwrap();
        assert_eq!(workbook.proposals(), std::slice::from_ref(&proposal));
    }
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
}

#[test]
fn redo_restores_a_proposal_staged_inside_a_newly_inserted_column() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("Z1"), "1", CalculationOptions::default())
        .unwrap();
    apply(
        &mut workbook,
        Op::InsertCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
    );
    let id = stage(&mut workbook, "42");
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_stale(&mut workbook, &id, "C1", false);
    workbook.redo(CalculationOptions::default()).unwrap();
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    assert_eq!(workbook.cell(SheetId(0), cell("C1")).unwrap().input, "42");
}

#[test]
fn invalid_formula_text_remains_literal_through_structural_edits() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    let id = stage(&mut workbook, "=SUM(");
    apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    );
    assert_eq!(workbook.proposals()[0].edits[0].input, "=SUM(");
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    assert_eq!(
        workbook.cell(SheetId(0), cell("C2")).unwrap().input,
        "=SUM("
    );
}

#[test]
fn moving_sheets_preserves_reviewed_dependency_ghosts() {
    let mut source = model();
    source
        .sheets
        .extend([Sheet::new("First"), Sheet::new("Second")]);
    let mut workbook = Workbook::from_model(source).unwrap();
    for sheet in [SheetId(1), SheetId(2)] {
        workbook
            .edit_cell(sheet, cell("A1"), "=Data!C1", CalculationOptions::default())
            .unwrap();
    }
    let id = stage(&mut workbook, "42");
    apply(&mut workbook, Op::MoveSheet { from: 2, to: 1 });
    workbook
        .accept_proposal(&id, false, CalculationOptions::default())
        .unwrap();
    assert_eq!(workbook.cell(SheetId(0), cell("C1")).unwrap().input, "42");
}
