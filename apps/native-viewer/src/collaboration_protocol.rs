use std::error::Error;
use std::fmt;

include!(concat!(env!("OUT_DIR"), "/collaboration_limits.rs"));

const TOP_LEVEL_SYNC: u64 = 0;
const TOP_LEVEL_AWARENESS: u64 = 1;
const TOP_LEVEL_AUTH: u64 = 2;
const TOP_LEVEL_QUERY_AWARENESS: u64 = 3;
const SYNC_STEP_1: u64 = 0;
const SYNC_STEP_2: u64 = 1;
const SYNC_UPDATE: u64 = 2;
const AUTH_PERMISSION_DENIED: u64 = 0;
const MAX_VAR_UINT: u64 = 9_007_199_254_740_991;
const FINGERPRINT_TAG: &[u8] = b"\0betteroffice-document-fingerprint-v1\0";

pub type DocumentFingerprint = [u8; 32];

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ProtocolMessage {
    SyncStep1(Vec<u8>),
    SyncStep2(Vec<u8>),
    Update(Vec<u8>),
    Awareness(Vec<u8>),
    Auth(String),
    QueryAwareness,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ProtocolError(String);

impl ProtocolError {
    fn new(message: impl Into<String>) -> Self {
        Self(message.into())
    }
}

impl fmt::Display for ProtocolError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl Error for ProtocolError {}

struct Decoder<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> Decoder<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, offset: 0 }
    }

    fn done(&self) -> bool {
        self.offset == self.bytes.len()
    }

    fn read_var_uint(&mut self) -> Result<u64, ProtocolError> {
        let mut value = 0_u64;
        let mut multiplier = 1_u64;
        let mut count = 0_usize;
        loop {
            let byte = *self
                .bytes
                .get(self.offset)
                .ok_or_else(|| ProtocolError::new("Truncated varUint"))?;
            self.offset += 1;
            let digit = u64::from(byte & 0x7f);
            if digit > (MAX_VAR_UINT - value) / multiplier {
                return Err(ProtocolError::new(
                    "varUint exceeds Number.MAX_SAFE_INTEGER",
                ));
            }
            value += digit * multiplier;
            count += 1;
            if byte & 0x80 == 0 {
                if count > 1 && digit == 0 {
                    return Err(ProtocolError::new("Non-canonical varUint"));
                }
                return Ok(value);
            }
            if count >= 8 {
                return Err(ProtocolError::new(
                    "varUint exceeds Number.MAX_SAFE_INTEGER",
                ));
            }
            multiplier *= 128;
        }
    }

    fn read_bytes(&mut self) -> Result<Vec<u8>, ProtocolError> {
        let length = usize::try_from(self.read_var_uint()?)
            .map_err(|_| ProtocolError::new("Byte array length is unsupported"))?;
        let end = self
            .offset
            .checked_add(length)
            .filter(|end| *end <= self.bytes.len())
            .ok_or_else(|| ProtocolError::new("Truncated varUint8Array"))?;
        let value = self.bytes[self.offset..end].to_vec();
        self.offset = end;
        Ok(value)
    }

    fn read_string(&mut self) -> Result<String, ProtocolError> {
        String::from_utf8(self.read_bytes()?)
            .map_err(|error| ProtocolError::new(format!("Invalid UTF-8 string: {error}")))
    }
}

fn encode_var_uint(mut value: u64, output: &mut Vec<u8>) -> Result<(), ProtocolError> {
    if value > MAX_VAR_UINT {
        return Err(ProtocolError::new(
            "varUint exceeds Number.MAX_SAFE_INTEGER",
        ));
    }
    while value >= 128 {
        output.push((value as u8 & 0x7f) | 0x80);
        value /= 128;
    }
    output.push(value as u8);
    Ok(())
}

fn encode_sync(subtype: u64, payload: &[u8]) -> Result<Vec<u8>, ProtocolError> {
    if payload.len() > MAX_FRAME_BYTES {
        return Err(ProtocolError::new(format!(
            "Frame exceeds {MAX_FRAME_BYTES} bytes"
        )));
    }
    let mut frame = Vec::with_capacity(payload.len().saturating_add(12));
    encode_var_uint(TOP_LEVEL_SYNC, &mut frame)?;
    encode_var_uint(subtype, &mut frame)?;
    encode_var_uint(payload.len() as u64, &mut frame)?;
    frame.extend_from_slice(payload);
    if frame.len() > MAX_FRAME_BYTES {
        return Err(ProtocolError::new(format!(
            "Frame exceeds {MAX_FRAME_BYTES} bytes"
        )));
    }
    Ok(frame)
}

fn finish_frame(frame: Vec<u8>) -> Result<Vec<u8>, ProtocolError> {
    if frame.len() > MAX_FRAME_BYTES {
        return Err(ProtocolError::new(format!(
            "Frame exceeds {MAX_FRAME_BYTES} bytes"
        )));
    }
    Ok(frame)
}

pub fn encode_message(message: &ProtocolMessage) -> Result<Vec<u8>, ProtocolError> {
    match message {
        ProtocolMessage::SyncStep1(payload) => encode_sync(SYNC_STEP_1, payload),
        ProtocolMessage::SyncStep2(payload) => encode_sync(SYNC_STEP_2, payload),
        ProtocolMessage::Update(payload) => encode_sync(SYNC_UPDATE, payload),
        ProtocolMessage::Awareness(payload) => {
            let mut frame = Vec::with_capacity(payload.len().saturating_add(10));
            encode_var_uint(TOP_LEVEL_AWARENESS, &mut frame)?;
            encode_var_uint(payload.len() as u64, &mut frame)?;
            frame.extend_from_slice(payload);
            finish_frame(frame)
        }
        ProtocolMessage::Auth(reason) => {
            let bytes = reason.as_bytes();
            let mut frame = Vec::with_capacity(bytes.len().saturating_add(11));
            encode_var_uint(TOP_LEVEL_AUTH, &mut frame)?;
            encode_var_uint(AUTH_PERMISSION_DENIED, &mut frame)?;
            encode_var_uint(bytes.len() as u64, &mut frame)?;
            frame.extend_from_slice(bytes);
            finish_frame(frame)
        }
        ProtocolMessage::QueryAwareness => {
            let mut frame = Vec::with_capacity(1);
            encode_var_uint(TOP_LEVEL_QUERY_AWARENESS, &mut frame)?;
            finish_frame(frame)
        }
    }
}

#[cfg(test)]
pub fn encode_sync_step_1(state_vector: &[u8]) -> Result<Vec<u8>, ProtocolError> {
    encode_sync(SYNC_STEP_1, state_vector)
}

pub fn encode_sync_step_1_with_fingerprint(
    state_vector: &[u8],
    fingerprint: &DocumentFingerprint,
) -> Result<Vec<u8>, ProtocolError> {
    encode_sync(
        SYNC_STEP_1,
        &fingerprinted_payload(state_vector, fingerprint),
    )
}

#[cfg(test)]
pub fn encode_sync_step_2(update: &[u8]) -> Result<Vec<u8>, ProtocolError> {
    encode_sync(SYNC_STEP_2, update)
}

pub fn encode_sync_step_2_with_fingerprint(
    update: &[u8],
    fingerprint: &DocumentFingerprint,
) -> Result<Vec<u8>, ProtocolError> {
    encode_sync(SYNC_STEP_2, &fingerprinted_payload(update, fingerprint))
}

#[cfg(test)]
pub fn encode_update(update: &[u8]) -> Result<Vec<u8>, ProtocolError> {
    encode_sync(SYNC_UPDATE, update)
}

pub fn encode_update_with_fingerprint(
    update: &[u8],
    fingerprint: &DocumentFingerprint,
) -> Result<Vec<u8>, ProtocolError> {
    encode_sync(SYNC_UPDATE, &fingerprinted_payload(update, fingerprint))
}

pub fn split_fingerprint(payload: &[u8]) -> (&[u8], Option<DocumentFingerprint>) {
    let suffix_len = FINGERPRINT_TAG.len() + 32;
    let Some(tag_start) = payload.len().checked_sub(suffix_len) else {
        return (payload, None);
    };
    if &payload[tag_start..tag_start + FINGERPRINT_TAG.len()] != FINGERPRINT_TAG {
        return (payload, None);
    }
    let mut fingerprint = [0_u8; 32];
    fingerprint.copy_from_slice(&payload[tag_start + FINGERPRINT_TAG.len()..]);
    (&payload[..tag_start], Some(fingerprint))
}

fn fingerprinted_payload(payload: &[u8], fingerprint: &DocumentFingerprint) -> Vec<u8> {
    let mut output = Vec::with_capacity(payload.len() + FINGERPRINT_TAG.len() + fingerprint.len());
    output.extend_from_slice(payload);
    output.extend_from_slice(FINGERPRINT_TAG);
    output.extend_from_slice(fingerprint);
    output
}

pub fn encode_empty_awareness() -> Vec<u8> {
    encode_message(&ProtocolMessage::Awareness(vec![0])).unwrap()
}

pub fn decode_messages(frame: &[u8]) -> Result<Vec<ProtocolMessage>, ProtocolError> {
    if frame.is_empty() {
        return Err(ProtocolError::new("Frame is empty"));
    }
    if frame.len() > MAX_FRAME_BYTES {
        return Err(ProtocolError::new(format!(
            "Frame exceeds {MAX_FRAME_BYTES} bytes"
        )));
    }
    let mut decoder = Decoder::new(frame);
    let mut messages = Vec::new();
    while !decoder.done() {
        if messages.len() >= MAX_MESSAGES_PER_FRAME {
            return Err(ProtocolError::new(format!(
                "Frame exceeds {MAX_MESSAGES_PER_FRAME} messages"
            )));
        }
        match decoder.read_var_uint()? {
            TOP_LEVEL_SYNC => {
                let subtype = decoder.read_var_uint()?;
                let payload = decoder.read_bytes()?;
                messages.push(match subtype {
                    SYNC_STEP_1 => ProtocolMessage::SyncStep1(payload),
                    SYNC_STEP_2 => ProtocolMessage::SyncStep2(payload),
                    SYNC_UPDATE => ProtocolMessage::Update(payload),
                    _ => {
                        return Err(ProtocolError::new(format!(
                            "Unknown sync message type {subtype}"
                        )));
                    }
                });
            }
            TOP_LEVEL_AWARENESS => {
                messages.push(ProtocolMessage::Awareness(decoder.read_bytes()?));
            }
            TOP_LEVEL_AUTH => {
                let subtype = decoder.read_var_uint()?;
                if subtype != AUTH_PERMISSION_DENIED {
                    return Err(ProtocolError::new(format!(
                        "Unknown auth message type {subtype}"
                    )));
                }
                messages.push(ProtocolMessage::Auth(decoder.read_string()?));
            }
            TOP_LEVEL_QUERY_AWARENESS => messages.push(ProtocolMessage::QueryAwareness),
            message_type => {
                return Err(ProtocolError::new(format!(
                    "Unknown top-level message type {message_type}"
                )));
            }
        }
    }
    Ok(messages)
}

#[cfg(test)]
mod tests {
    use std::io::Write;
    use std::panic;
    use std::path::Path;
    use std::process::{Command, Stdio};

    use serde_json::{Value, json};
    use yrs::updates::decoder::Decode;
    use yrs::updates::encoder::Encode;
    use yrs::{StateVector, Update};

    use super::*;

    #[test]
    fn every_message_kind_matches_the_typescript_codec_both_directions() {
        let payload = vec![7; 130];
        let messages = [
            ProtocolMessage::SyncStep1(vec![1, 2, 3]),
            ProtocolMessage::SyncStep2(Vec::new()),
            ProtocolMessage::Update(payload),
            ProtocolMessage::Awareness(vec![0]),
            ProtocolMessage::Auth("denied 雪".to_owned()),
            ProtocolMessage::QueryAwareness,
        ];
        let rust_frames = messages
            .iter()
            .map(|message| encode_message(message).unwrap())
            .collect::<Vec<_>>();
        let request = json!({ "rustFrames": rust_frames });
        let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
        let script = manifest.join("tests/collaboration_protocol_conformance.ts");
        let mut child = Command::new("bun")
            .arg("run")
            .arg(script)
            .current_dir(manifest)
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::piped())
            .spawn()
            .expect("start TypeScript collaboration codec");
        serde_json::to_writer(child.stdin.as_mut().unwrap(), &request).unwrap();
        child.stdin.as_mut().unwrap().flush().unwrap();
        drop(child.stdin.take());
        let output = child.wait_with_output().unwrap();
        assert!(
            output.status.success(),
            "TypeScript collaboration codec failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        let response: Value = serde_json::from_slice(&output.stdout).unwrap();
        let expected = messages.iter().map(message_json).collect::<Vec<_>>();
        assert_eq!(response["decodedRustFrames"], json!(expected));

        let typescript_frames = response["typescriptFrames"].as_array().unwrap();
        assert_eq!(typescript_frames.len(), messages.len());
        for (frame, expected) in typescript_frames.iter().zip(messages) {
            let bytes: Vec<u8> = serde_json::from_value(frame.clone()).unwrap();
            assert_eq!(decode_messages(&bytes).unwrap(), [expected]);
        }
    }

    fn message_json(message: &ProtocolMessage) -> Value {
        match message {
            ProtocolMessage::SyncStep1(payload) => {
                json!({ "type": "sync-step-1", "stateVector": payload })
            }
            ProtocolMessage::SyncStep2(payload) => {
                json!({ "type": "sync-step-2", "update": payload })
            }
            ProtocolMessage::Update(payload) => json!({ "type": "update", "update": payload }),
            ProtocolMessage::Awareness(payload) => {
                json!({ "type": "awareness", "update": payload })
            }
            ProtocolMessage::Auth(reason) => json!({ "type": "auth", "reason": reason }),
            ProtocolMessage::QueryAwareness => json!({ "type": "query-awareness" }),
        }
    }

    #[test]
    fn decodes_every_top_level_message_and_concatenation() {
        let frame = [
            &[0, 0, 1, 7][..],
            &[0, 1, 1, 8],
            &[0, 2, 1, 9],
            &[1, 1, 0],
            &[2, 0, 2, b'o', b'k'],
            &[3],
        ]
        .concat();
        assert_eq!(
            decode_messages(&frame).unwrap(),
            [
                ProtocolMessage::SyncStep1(vec![7]),
                ProtocolMessage::SyncStep2(vec![8]),
                ProtocolMessage::Update(vec![9]),
                ProtocolMessage::Awareness(vec![0]),
                ProtocolMessage::Auth("ok".to_owned()),
                ProtocolMessage::QueryAwareness,
            ]
        );
    }

    #[test]
    fn malformed_and_oversized_frames_never_panic() {
        let mut cases = vec![
            vec![],
            vec![0x80],
            vec![0x80, 0],
            vec![0xff; 8],
            vec![0, 3, 0],
            vec![0, 2, 4, 1],
            vec![1, 3, 0],
            vec![2, 1, 0],
            vec![2, 0, 1, 0xff],
            vec![4],
            vec![3; MAX_MESSAGES_PER_FRAME + 1],
        ];
        cases.push(vec![0; MAX_FRAME_BYTES + 1]);
        for frame in cases {
            let result = panic::catch_unwind(|| decode_messages(&frame));
            assert!(result.is_ok());
            assert!(result.unwrap().is_err());
        }
    }

    #[test]
    fn every_truncation_is_rejected_without_panicking() {
        let valid = encode_update(&vec![5; 300]).unwrap();
        for end in 0..valid.len() {
            let result = panic::catch_unwind(|| decode_messages(&valid[..end]));
            assert!(result.is_ok());
            assert!(result.unwrap().is_err());
        }
    }

    #[test]
    fn encoder_rejects_frames_above_the_shared_limit() {
        let payload = vec![0; MAX_FRAME_BYTES];
        assert!(encode_update(&payload).is_err());
    }

    #[test]
    fn fingerprint_suffix_round_trips_and_remains_yrs_compatible() {
        let fingerprint = [0x5a; 32];
        let state_vector = StateVector::default().encode_v1();
        let frame = encode_sync_step_1_with_fingerprint(&state_vector, &fingerprint).unwrap();
        let messages = decode_messages(&frame).unwrap();
        let [ProtocolMessage::SyncStep1(payload)] = messages.as_slice() else {
            panic!("expected sync step 1");
        };
        let (decoded, remote_fingerprint) = split_fingerprint(payload);
        assert_eq!(decoded, state_vector);
        assert_eq!(remote_fingerprint, Some(fingerprint));
        assert!(StateVector::decode_v1(payload).is_ok());

        let frame = encode_update_with_fingerprint(Update::EMPTY_V1, &fingerprint).unwrap();
        let messages = decode_messages(&frame).unwrap();
        let [ProtocolMessage::Update(payload)] = messages.as_slice() else {
            panic!("expected update");
        };
        assert_eq!(
            split_fingerprint(payload),
            (Update::EMPTY_V1, Some(fingerprint))
        );
        assert!(Update::decode_v1(payload).is_ok());
    }
}
