use anyhow::Result;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{ClientID, Update};

use crate::collaboration::{CollaborationClient, TransportEvent};
use crate::collaboration_protocol::{
    ProtocolMessage, decode_messages, encode_empty_awareness, encode_sync_step_1_with_fingerprint,
    encode_sync_step_2_with_fingerprint, encode_update_with_fingerprint, split_fingerprint,
};
use crate::document::DocumentView;

pub fn handle_transport_event(
    document: &mut DocumentView,
    collaboration: &mut CollaborationClient,
    event: TransportEvent,
) -> Result<bool> {
    collaboration.apply_transport_event(&event);
    match event {
        TransportEvent::Connected => {
            collaboration.send(encode_sync_step_1_with_fingerprint(
                &document.collaboration_state_vector()?,
                &document.collaboration_fingerprint()?,
            )?)?;
            Ok(false)
        }
        TransportEvent::Binary(frame) => apply_frame(document, collaboration, &frame),
        TransportEvent::Connecting
        | TransportEvent::PeerCount(_)
        | TransportEvent::Reconnecting { .. }
        | TransportEvent::Failed(_) => Ok(false),
    }
}

pub fn apply_frame(
    document: &mut DocumentView,
    collaboration: &mut CollaborationClient,
    frame: &[u8],
) -> Result<bool> {
    let messages = match decode_messages(frame) {
        Ok(messages) => messages,
        Err(error) => {
            collaboration.report_protocol_error(error);
            return Ok(false);
        }
    };
    let fingerprint = document.collaboration_fingerprint()?;
    for message in &messages {
        let payload = match message {
            ProtocolMessage::SyncStep1(payload)
            | ProtocolMessage::SyncStep2(payload)
            | ProtocolMessage::Update(payload) => payload,
            ProtocolMessage::Auth(reason) => {
                collaboration.deny(reason.clone());
                return Ok(false);
            }
            ProtocolMessage::Awareness(_) | ProtocolMessage::QueryAwareness => continue,
        };
        if split_fingerprint(payload)
            .1
            .is_some_and(|remote| remote != fingerprint)
        {
            collaboration.deny("document fingerprint mismatch; refusing room".to_owned());
            return Ok(false);
        }
    }

    let local_client_id = ClientID::new(collaboration.client_id());
    let mut updates = Vec::new();
    let mut saw_sync_step_2 = false;
    for message in &messages {
        let payload = match message {
            ProtocolMessage::SyncStep2(payload) => {
                saw_sync_step_2 = true;
                payload
            }
            ProtocolMessage::Update(payload) => payload,
            _ => continue,
        };
        let (payload, _) = split_fingerprint(payload);
        let update = match Update::decode_v1(payload) {
            Ok(update) => update,
            Err(error) => {
                collaboration.report_protocol_error(format!("invalid yrs update: {error}"));
                return Ok(false);
            }
        };
        if update
            .state_vector_lower()
            .contains_client(&local_client_id)
        {
            collaboration.report_protocol_error(format!(
                "incoming update is authored by local client id {}",
                collaboration.client_id()
            ));
            return Ok(false);
        }
        updates.push(update);
    }

    let mut sync_step_1 = None;
    let mut query_awareness = false;
    for message in messages {
        match message {
            ProtocolMessage::SyncStep1(state_vector) if sync_step_1.is_none() => {
                sync_step_1 = Some(state_vector);
            }
            ProtocolMessage::QueryAwareness => query_awareness = true,
            ProtocolMessage::SyncStep1(_)
            | ProtocolMessage::SyncStep2(_)
            | ProtocolMessage::Update(_)
            | ProtocolMessage::Auth(_) => {}
            ProtocolMessage::Awareness(_) => {}
        }
    }
    if let Some(state_vector) = sync_step_1 {
        let (state_vector, _) = split_fingerprint(&state_vector);
        collaboration.send(encode_sync_step_2_with_fingerprint(
            &document.collaboration_encode_diff(state_vector)?,
            &fingerprint,
        )?)?;
    }
    if query_awareness {
        collaboration.send(encode_empty_awareness())?;
    }

    let repainted = if updates.is_empty() {
        false
    } else {
        let update = Update::merge_updates(updates).encode_v1();
        match document.collaboration_apply_remote_update(&update) {
            Ok(repainted) => repainted,
            Err(error) => {
                collaboration.report_protocol_error(format!("invalid remote update: {error:#}"));
                return Ok(false);
            }
        }
    };
    if saw_sync_step_2 {
        collaboration.mark_synced();
    }
    Ok(repainted)
}

pub fn forward_local_updates(document: &DocumentView, collaboration: &mut CollaborationClient) {
    if !collaboration.is_connected() {
        return;
    }
    let fingerprint = match document.collaboration_fingerprint() {
        Ok(fingerprint) => fingerprint,
        Err(error) => {
            collaboration.report_protocol_error(error);
            return;
        }
    };
    for update in document.collaboration_drain_local_updates() {
        let result = encode_update_with_fingerprint(&update, &fingerprint)
            .map_err(anyhow::Error::from)
            .and_then(|frame| collaboration.send(frame));
        if let Err(error) = result {
            collaboration.report_protocol_error(error);
            break;
        }
    }
}

#[cfg(test)]
mod tests {
    use std::path::Path;

    use betteroffice_xlsx::{CalculationOptions, CellRef, Op, SheetId, Workbook};
    use tokio::sync::mpsc::Receiver;
    use yrs::{Any, Doc, Map, ReadTxn, StateVector, Transact, Update, merge_updates_v1};

    use super::*;
    use crate::collaboration::{CollaborationConfig, TransportCommand};
    use crate::collaboration_protocol::{
        MAX_MESSAGES_PER_FRAME, ProtocolMessage, decode_messages,
        encode_sync_step_1_with_fingerprint, encode_update_with_fingerprint,
    };
    use crate::document::{
        ReferenceDocument, load_collaborative_docx, load_collaborative_pptx,
        load_collaborative_xlsx,
    };
    use crate::editing::TextLoc;
    use crate::test_fixtures;

    fn detached(client_id: u64) -> (CollaborationClient, Receiver<TransportCommand>) {
        CollaborationClient::detached(CollaborationConfig::for_test("test", client_id))
    }

    fn next_frame(commands: &mut Receiver<TransportCommand>) -> Vec<u8> {
        match commands.try_recv().unwrap() {
            TransportCommand::Send(frame) => frame,
            TransportCommand::Shutdown => panic!("collaboration client shut down"),
        }
    }

    fn showcase_path() -> std::path::PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx")
    }

    fn presentation_path() -> std::path::PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx")
    }

    #[test]
    fn mismatched_document_fingerprint_refuses_the_room_before_sync() {
        let left_bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Left</w:t></w:r></w:p>"#,
        );
        let right_bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="22222222"><w:r><w:t>Right</w:t></w:r></w:p>"#,
        );
        let left_path = test_fixtures::write_docx("fingerprint-left", &left_bytes);
        let right_path = test_fixtures::write_docx("fingerprint-right", &right_bytes);
        let mut left = load_collaborative_docx(&left_path, 8_192, 101).unwrap();
        let right = load_collaborative_docx(&right_path, 8_192, 202).unwrap();
        let before = left.docx_canonical_checksum().unwrap();
        let frame = encode_sync_step_1_with_fingerprint(
            &right.docx_state_vector().unwrap(),
            &right.docx_fingerprint().unwrap(),
        )
        .unwrap();
        let (mut collaboration, _commands) = detached(101);

        assert!(!apply_frame(&mut left, &mut collaboration, &frame).unwrap());
        assert!(collaboration.status_line().contains("fingerprint mismatch"));
        assert_eq!(left.docx_canonical_checksum().unwrap(), before);
        std::fs::remove_file(left_path).unwrap();
        std::fs::remove_file(right_path).unwrap();
    }

    #[test]
    fn saved_copy_keeps_a_distinct_room_fingerprint() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Base line</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("fingerprint-original", &bytes);
        let output = source.with_extension("saved-copy.docx");
        let mut original = load_collaborative_docx(&source, 8_192, 101).unwrap();
        original
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 0,
                },
                false,
                false,
            )
            .unwrap();
        assert!(original.docx_insert_text("LEFT ").unwrap());
        original.save_docx_to(&output).unwrap();
        let copy = load_collaborative_docx(&output, 8_192, 202).unwrap();
        assert_ne!(
            original.docx_fingerprint().unwrap(),
            copy.docx_fingerprint().unwrap()
        );
        let frame = encode_sync_step_1_with_fingerprint(
            &copy.docx_state_vector().unwrap(),
            &copy.docx_fingerprint().unwrap(),
        )
        .unwrap();
        let (mut collaboration, _commands) = detached(101);
        assert!(!apply_frame(&mut original, &mut collaboration, &frame).unwrap());
        assert!(collaboration.status_line().contains("fingerprint mismatch"));
        std::fs::remove_file(source).unwrap();
        std::fs::remove_file(output).unwrap();
    }

    #[test]
    fn update_authored_by_the_local_client_id_is_rejected() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Base</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("client-id-collision", &bytes);
        let mut local = load_collaborative_docx(&source, 8_192, 101).unwrap();
        let mut impostor = load_collaborative_docx(&source, 8_192, 101).unwrap();
        impostor
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 4,
                },
                false,
                false,
            )
            .unwrap();
        assert!(impostor.docx_insert_text(" hostile").unwrap());
        let update = merge_updates_v1(impostor.docx_drain_local_updates()).unwrap();
        let frame =
            encode_update_with_fingerprint(&update, &local.docx_fingerprint().unwrap()).unwrap();
        let before = local.docx_canonical_checksum().unwrap();
        let (mut collaboration, _commands) = detached(101);

        assert!(!apply_frame(&mut local, &mut collaboration, &frame).unwrap());
        assert!(
            collaboration
                .status_line()
                .contains("authored by local client id 101")
        );
        assert_eq!(local.docx_canonical_checksum().unwrap(), before);
        std::fs::remove_file(source).unwrap();
    }

    #[test]
    fn one_frame_of_updates_causes_one_relayout() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Base</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("coalesced-updates", &bytes);
        let mut document = load_collaborative_docx(&source, 8_192, 101).unwrap();
        let fingerprint = document.docx_fingerprint().unwrap();
        let message = encode_update_with_fingerprint(Update::EMPTY_V1, &fingerprint).unwrap();
        let frame = message.repeat(MAX_MESSAGES_PER_FRAME);
        let before = document.docx_relayout_count();
        let (mut collaboration, _commands) = detached(101);

        assert!(apply_frame(&mut document, &mut collaboration, &frame).unwrap());
        assert_eq!(document.docx_relayout_count(), before + 1);
        std::fs::remove_file(source).unwrap();
    }

    #[test]
    fn repeated_awareness_queries_receive_one_response_per_frame() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Base</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("awareness-coalescing", &bytes);
        let mut document = load_collaborative_docx(&source, 8_192, 101).unwrap();
        let (mut collaboration, mut commands) = detached(101);
        let frame = vec![3; MAX_MESSAGES_PER_FRAME];

        assert!(!apply_frame(&mut document, &mut collaboration, &frame).unwrap());
        assert!(matches!(
            decode_messages(&next_frame(&mut commands))
                .unwrap()
                .as_slice(),
            [ProtocolMessage::Awareness(_)]
        ));
        assert!(commands.try_recv().is_err());
        std::fs::remove_file(source).unwrap();
    }

    #[test]
    fn mismatched_xlsx_fingerprint_refuses_the_room_before_sync() {
        let showcase = showcase_path();
        let sample = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/sample.xlsx");
        let mut local = load_collaborative_xlsx(&showcase, 0, 8_192, 101).unwrap();
        let remote = load_collaborative_xlsx(&sample, 0, 8_192, 202).unwrap();
        let before = local.collaboration_canonical_checksum().unwrap();
        let frame = encode_sync_step_1_with_fingerprint(
            &remote.collaboration_state_vector().unwrap(),
            &remote.collaboration_fingerprint().unwrap(),
        )
        .unwrap();
        let (mut collaboration, _commands) = detached(101);

        assert!(!apply_frame(&mut local, &mut collaboration, &frame).unwrap());
        assert!(collaboration.status_line().contains("fingerprint mismatch"));
        assert_eq!(local.collaboration_canonical_checksum().unwrap(), before);
    }

    #[test]
    fn xlsx_update_authored_by_the_local_client_id_is_rejected() {
        let source = showcase_path();
        let mut local = load_collaborative_xlsx(&source, 0, 8_192, 101).unwrap();
        let mut impostor = load_collaborative_xlsx(&source, 0, 8_192, 101).unwrap();
        impostor.xlsx_select_cell(CellRef::parse_a1("B5").unwrap());
        impostor.xlsx_begin_edit(Some("impostor")).unwrap();
        impostor.xlsx_commit(None).unwrap();
        let update = merge_updates_v1(impostor.collaboration_drain_local_updates()).unwrap();
        let frame =
            encode_update_with_fingerprint(&update, &local.collaboration_fingerprint().unwrap())
                .unwrap();
        let before = local.collaboration_canonical_checksum().unwrap();
        let (mut collaboration, _commands) = detached(101);

        assert!(!apply_frame(&mut local, &mut collaboration, &frame).unwrap());
        assert!(
            collaboration
                .status_line()
                .contains("authored by local client id 101")
        );
        assert_eq!(local.collaboration_canonical_checksum().unwrap(), before);
    }

    #[test]
    fn one_xlsx_frame_of_updates_causes_one_repaint() {
        let source = showcase_path();
        let mut local = load_collaborative_xlsx(&source, 0, 8_192, 101).unwrap();
        let mut peer = load_collaborative_xlsx(&source, 0, 8_192, 202).unwrap();
        for (cell, value) in [("B5", "first"), ("C5", "second")] {
            peer.xlsx_select_cell(CellRef::parse_a1(cell).unwrap());
            peer.xlsx_begin_edit(Some(value)).unwrap();
            peer.xlsx_commit(None).unwrap();
        }
        let fingerprint = local.collaboration_fingerprint().unwrap();
        let frame = peer
            .collaboration_drain_local_updates()
            .into_iter()
            .flat_map(|update| encode_update_with_fingerprint(&update, &fingerprint).unwrap())
            .collect::<Vec<_>>();
        let before = local.collaboration_relayout_count();
        let (mut collaboration, _commands) = detached(101);

        assert!(apply_frame(&mut local, &mut collaboration, &frame).unwrap());
        assert_eq!(local.collaboration_relayout_count(), before + 1);
    }

    #[test]
    fn hostile_structural_xlsx_frame_is_rolled_back() {
        let source = showcase_path();
        let bytes = std::fs::read(&source).unwrap();
        let mut local = load_collaborative_xlsx(&source, 0, 8_192, 101).unwrap();
        let mut hostile = Workbook::open(&bytes).unwrap();
        hostile
            .apply_ops(
                vec![Op::RenameSheet {
                    sheet: SheetId(0),
                    name: "Hostile".to_owned(),
                }],
                CalculationOptions::default(),
            )
            .unwrap();
        let update = hostile
            .encode_diff_v1(&local.collaboration_state_vector().unwrap())
            .unwrap();
        let frame =
            encode_update_with_fingerprint(&update, &local.collaboration_fingerprint().unwrap())
                .unwrap();
        let before = local.collaboration_canonical_checksum().unwrap();
        let (mut collaboration, _commands) = detached(101);

        assert!(!apply_frame(&mut local, &mut collaboration, &frame).unwrap());
        assert!(
            collaboration
                .status_line()
                .contains("invalid remote update")
        );
        assert_eq!(local.collaboration_canonical_checksum().unwrap(), before);
    }

    #[test]
    fn mismatched_pptx_fingerprint_refuses_the_room_before_sync() {
        let local_path = presentation_path();
        let remote_path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../crates/pptx-parse/tests/fixtures/chart-deck.pptx");
        let mut local = load_collaborative_pptx(&local_path, 8_192, 101).unwrap();
        let remote = load_collaborative_pptx(&remote_path, 8_192, 202).unwrap();
        let before = local.collaboration_canonical_checksum().unwrap();
        let frame = encode_sync_step_1_with_fingerprint(
            &remote.collaboration_state_vector().unwrap(),
            &remote.collaboration_fingerprint().unwrap(),
        )
        .unwrap();
        let (mut collaboration, _commands) = detached(101);

        assert!(!apply_frame(&mut local, &mut collaboration, &frame).unwrap());
        assert!(collaboration.status_line().contains("fingerprint mismatch"));
        assert_eq!(local.collaboration_canonical_checksum().unwrap(), before);
    }

    #[test]
    fn one_pptx_frame_repaints_the_changed_slide_once() {
        let source = presentation_path();
        let mut local = load_collaborative_pptx(&source, 8_192, 101).unwrap();
        let mut peer = load_collaborative_pptx(&source, 8_192, 202).unwrap();
        let ReferenceDocument::Pptx(reference) = &peer.reference else {
            unreachable!();
        };
        let story = reference
            .editor
            .snapshot()
            .unwrap()
            .slides
            .into_iter()
            .flat_map(|slide| slide.shapes)
            .flat_map(|shape| shape.text_stories)
            .find(|story| story.length > 1)
            .unwrap();
        assert!(
            peer.pptx_select_story_position(&story.id, story.length - 1)
                .unwrap()
        );
        assert!(peer.pptx_insert_text(" first").unwrap());
        assert!(peer.pptx_insert_text(" second").unwrap());
        let fingerprint = local.collaboration_fingerprint().unwrap();
        let frame = peer
            .collaboration_drain_local_updates()
            .into_iter()
            .flat_map(|update| encode_update_with_fingerprint(&update, &fingerprint).unwrap())
            .collect::<Vec<_>>();
        let before = local.collaboration_relayout_count();
        let (mut collaboration, _commands) = detached(101);

        assert!(apply_frame(&mut local, &mut collaboration, &frame).unwrap());
        assert_eq!(local.collaboration_relayout_count(), before + 1);
        let ReferenceDocument::Pptx(reference) = &local.reference else {
            unreachable!();
        };
        assert!(
            reference
                .editor
                .story(&story.id)
                .unwrap()
                .plain_text()
                .contains(" first second")
        );
    }

    #[test]
    fn hostile_structural_pptx_frame_is_rolled_back() {
        let source = presentation_path();
        let mut local = load_collaborative_pptx(&source, 8_192, 101).unwrap();
        let baseline = local
            .collaboration_encode_diff(&StateVector::default().encode_v1())
            .unwrap();
        let hostile = Doc::with_client_id(202);
        hostile
            .transact_mut()
            .apply_update(Update::decode_v1(&baseline).unwrap())
            .unwrap();
        let state_vector = hostile.transact().state_vector();
        {
            let mut transaction = hostile.transact_mut();
            let meta = transaction.get_map("pptx:meta").unwrap();
            meta.insert(&mut transaction, "schemaVersion", Any::Number(99.0));
        }
        let update = hostile.transact().encode_diff_v1(&state_vector);
        let frame =
            encode_update_with_fingerprint(&update, &local.collaboration_fingerprint().unwrap())
                .unwrap();
        let before = local.collaboration_canonical_checksum().unwrap();
        let (mut collaboration, _commands) = detached(101);

        assert!(!apply_frame(&mut local, &mut collaboration, &frame).unwrap());
        assert!(
            collaboration
                .status_line()
                .contains("invalid remote update")
        );
        assert_eq!(local.collaboration_canonical_checksum().unwrap(), before);
    }
}
