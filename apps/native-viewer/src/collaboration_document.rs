use anyhow::Result;

use crate::collaboration::{CollaborationClient, TransportEvent};
use crate::collaboration_protocol::{
    ProtocolMessage, decode_messages, encode_empty_awareness, encode_sync_step_1,
    encode_sync_step_2, encode_update,
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
            collaboration.send(encode_sync_step_1(&document.docx_state_vector()?)?)?;
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
    let mut repainted = false;
    for message in messages {
        match message {
            ProtocolMessage::SyncStep1(state_vector) => collaboration.send(encode_sync_step_2(
                &document.docx_encode_diff(&state_vector)?,
            )?)?,
            ProtocolMessage::SyncStep2(update) => {
                repainted |= document.docx_apply_remote_update(&update)?;
                collaboration.mark_synced();
            }
            ProtocolMessage::Update(update) => {
                repainted |= document.docx_apply_remote_update(&update)?;
            }
            ProtocolMessage::QueryAwareness => {
                collaboration.send(encode_empty_awareness())?;
            }
            ProtocolMessage::Auth(reason) => collaboration.deny(reason),
            ProtocolMessage::Awareness(_) => {}
        }
    }
    Ok(repainted)
}

pub fn forward_local_updates(document: &DocumentView, collaboration: &mut CollaborationClient) {
    if !collaboration.is_connected() {
        return;
    }
    for update in document.docx_drain_local_updates() {
        let result = encode_update(&update)
            .map_err(anyhow::Error::from)
            .and_then(|frame| collaboration.send(frame));
        if let Err(error) = result {
            collaboration.report_protocol_error(error);
            break;
        }
    }
}
