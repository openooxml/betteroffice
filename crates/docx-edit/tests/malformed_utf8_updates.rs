use docx_edit::{EditError, EditingDoc};

const HOSTILE: [u8; 18] = [
    0x01, 0x01, 0xca, 0x01, 0x00, 0xc4, 0x01, 0x0a, 0x01, 0x0b, 0x06, 0x20, 0x74, 0x68, 0x65,
    0xe5, 0x65, 0x00,
];

fn seeded(client_id: u64, text: &str) -> EditingDoc {
    let doc = EditingDoc::new(client_id);
    doc.create_story("body", text, "Normal", "left").unwrap();
    doc
}

fn corrupt_utf8(update: &[u8], text: &str) -> Vec<u8> {
    let needle = text.as_bytes();
    let at = update
        .windows(needle.len())
        .position(|window| window == needle)
        .expect("update carries the story text");
    let lead = text.find('é').unwrap();
    let mut bytes = update.to_vec();
    bytes[at + lead + 1] = b'A';
    assert!(std::str::from_utf8(&bytes[at..at + needle.len()]).is_err());
    bytes
}

#[test]
fn invalid_utf8_in_a_peer_update_is_refused() {
    let doc = seeded(11, "Safe");
    let before = doc.encode_state_as_update_v1();
    assert!(matches!(
        doc.apply_update_v1(&HOSTILE),
        Err(EditError::InvalidUpdate(_))
    ));
    assert!(matches!(
        doc.apply_update_chunks_v1(&[HOSTILE]),
        Err(EditError::InvalidUpdate(_))
    ));
    assert_eq!(doc.encode_state_as_update_v1(), before);
    assert!(!doc.holds_pending_update());
}

#[test]
fn a_real_update_with_broken_utf8_text_is_refused() {
    let text = "héllo wörld";
    let writer = seeded(21, text);
    let update = writer.encode_state_as_update_v1();
    let hostile = corrupt_utf8(&update, text);

    let reader = seeded(22, "Safe");
    let before = reader.encode_state_as_update_v1();
    assert!(matches!(
        reader.apply_update_v1(&hostile),
        Err(EditError::InvalidUpdate(_))
    ));
    assert_eq!(reader.encode_state_as_update_v1(), before);

    reader.apply_update_v1(&update).unwrap();
    assert_ne!(reader.encode_state_as_update_v1(), before);
}
