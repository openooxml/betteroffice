use betteroffice_xlsx::{
    CalculationOptions, CellRef, ProposalEditInput, ProposalRequest, Sheet, SheetId, Workbook,
    WorkbookModel,
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
