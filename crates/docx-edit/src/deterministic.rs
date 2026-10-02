use std::collections::BTreeMap;

use serde_json::{Map, Number, Value};
use yrs::any::{F64_MAX_SAFE_INTEGER, F64_MIN_SAFE_INTEGER};
use yrs::encoding::write::Write;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::{Encode, Encoder, EncoderV1};
use yrs::{Any, ClientID, ID, ReadTxn, StateVector, Update};

pub(crate) struct DeterministicEncoderV1(EncoderV1);

impl DeterministicEncoderV1 {
    fn new() -> Self {
        Self(EncoderV1::new())
    }
}

impl Write for DeterministicEncoderV1 {
    fn write_all(&mut self, buf: &[u8]) {
        self.0.write_all(buf);
    }

    fn write_u8(&mut self, value: u8) {
        self.0.write_u8(value);
    }
}

impl Encoder for DeterministicEncoderV1 {
    fn to_vec(self) -> Vec<u8> {
        self.0.to_vec()
    }

    fn reset_ds_cur_val(&mut self) {
        self.0.reset_ds_cur_val();
    }

    fn write_ds_clock(&mut self, clock: u32) {
        self.0.write_ds_clock(clock);
    }

    fn write_ds_len(&mut self, len: u32) {
        self.0.write_ds_len(len);
    }

    fn write_left_id(&mut self, id: &ID) {
        self.0.write_left_id(id);
    }

    fn write_right_id(&mut self, id: &ID) {
        self.0.write_right_id(id);
    }

    fn write_client(&mut self, client: ClientID) {
        self.0.write_client(client);
    }

    fn write_info(&mut self, info: u8) {
        self.0.write_info(info);
    }

    fn write_parent_info(&mut self, is_y_key: bool) {
        self.0.write_parent_info(is_y_key);
    }

    fn write_type_ref(&mut self, info: u8) {
        self.0.write_type_ref(info);
    }

    fn write_len(&mut self, len: u32) {
        self.0.write_len(len);
    }

    fn write_any(&mut self, any: &Any) {
        encode_any(any, self);
    }

    fn write_json(&mut self, any: &Any) {
        self.write_string(&serde_json::to_string(&json_value(any)).unwrap());
    }

    fn write_key(&mut self, string: &str) {
        self.0.write_key(string);
    }
}

#[derive(Default)]
struct CountingEncoderV1 {
    len: usize,
}

impl Write for CountingEncoderV1 {
    fn write_all(&mut self, buf: &[u8]) {
        self.len += buf.len();
    }

    fn write_u8(&mut self, _: u8) {
        self.len += 1;
    }
}

impl Encoder for CountingEncoderV1 {
    fn to_vec(self) -> Vec<u8> {
        unreachable!("a counting encoder has no byte buffer")
    }

    fn reset_ds_cur_val(&mut self) {}

    fn write_ds_clock(&mut self, clock: u32) {
        self.write_var(clock);
    }

    fn write_ds_len(&mut self, len: u32) {
        self.write_var(len);
    }

    fn write_left_id(&mut self, id: &ID) {
        self.write_var(id.client.get());
        self.write_var(id.clock);
    }

    fn write_right_id(&mut self, id: &ID) {
        self.write_left_id(id);
    }

    fn write_client(&mut self, client: ClientID) {
        self.write_var(client.get());
    }

    fn write_info(&mut self, info: u8) {
        self.write_u8(info);
    }

    fn write_parent_info(&mut self, is_y_key: bool) {
        self.write_var(u32::from(is_y_key));
    }

    fn write_type_ref(&mut self, info: u8) {
        self.write_u8(info);
    }

    fn write_len(&mut self, len: u32) {
        self.write_var(len);
    }

    fn write_any(&mut self, any: &Any) {
        encode_any(any, self);
    }

    fn write_json(&mut self, any: &Any) {
        self.write_string(&serde_json::to_string(&json_value(any)).unwrap());
    }

    fn write_key(&mut self, key: &str) {
        self.write_string(key);
    }
}

pub(crate) fn encoded_state_len_v1<T: ReadTxn>(txn: &T) -> Option<usize> {
    let store = txn.store();
    if store.pending_update().is_some() || store.pending_ds().is_some() {
        return None;
    }
    let mut encoder = CountingEncoderV1::default();
    txn.encode_state_as_update(&StateVector::default(), &mut encoder);
    Some(encoder.len)
}

pub(crate) fn encode_state_as_update_v1<T: ReadTxn>(
    txn: &T,
    state_vector: &StateVector,
) -> Vec<u8> {
    let integrated = encode_state(txn, state_vector);
    let store = txn.store();
    if store.pending_update().is_none() && store.pending_ds().is_none() {
        return integrated;
    }

    let mut updates = vec![Update::decode_v1(&integrated).unwrap()];
    if let Some(pending) = store.pending_update() {
        let mut encoder = DeterministicEncoderV1::new();
        pending.update.encode(&mut encoder);
        updates.push(Update::decode_v1(&encoder.to_vec()).unwrap());
    }
    if let Some(pending) = store.pending_ds() {
        let mut encoder = DeterministicEncoderV1::new();
        encoder.write_var(0_u32);
        pending.encode(&mut encoder);
        updates.push(Update::decode_v1(&encoder.to_vec()).unwrap());
    }

    let merged = Update::merge_updates(updates);
    let mut encoder = DeterministicEncoderV1::new();
    merged.encode(&mut encoder);
    encoder.to_vec()
}

pub(crate) fn encode_diff_v1<T: ReadTxn>(txn: &T, state_vector: &StateVector) -> Vec<u8> {
    let mut encoder = DeterministicEncoderV1::new();
    txn.encode_diff(state_vector, &mut encoder);
    encoder.to_vec()
}

fn encode_state<T: ReadTxn>(txn: &T, state_vector: &StateVector) -> Vec<u8> {
    let mut encoder = DeterministicEncoderV1::new();
    txn.encode_state_as_update(state_vector, &mut encoder);
    encoder.to_vec()
}

fn encode_any<W: Write>(any: &Any, encoder: &mut W) {
    match any {
        Any::Undefined => encoder.write_u8(127),
        Any::Null => encoder.write_u8(126),
        Any::Bool(value) => encoder.write_u8(if *value { 120 } else { 121 }),
        Any::String(value) => {
            encoder.write_u8(119);
            encoder.write_string(value);
        }
        Any::Number(value) => {
            let truncated = value.trunc();
            if truncated == *value
                && (F64_MIN_SAFE_INTEGER..=F64_MAX_SAFE_INTEGER).contains(&truncated)
            {
                encoder.write_u8(125);
                encoder.write_var(truncated as i64);
            } else if ((*value as f32) as f64) == *value {
                encoder.write_u8(124);
                encoder.write_f32(*value as f32);
            } else {
                encoder.write_u8(123);
                encoder.write_f64(*value);
            }
        }
        Any::BigInt(value) => {
            encoder.write_u8(122);
            encoder.write_i64(*value);
        }
        Any::Array(values) => {
            encoder.write_u8(117);
            encoder.write_var(values.len() as u64);
            for value in values.iter() {
                encode_any(value, encoder);
            }
        }
        Any::Map(values) => {
            encoder.write_u8(118);
            encoder.write_var(values.len() as u64);
            let sorted: BTreeMap<_, _> = values.iter().collect();
            for (key, value) in sorted {
                encoder.write_string(key);
                encode_any(value, encoder);
            }
        }
        Any::Buffer(value) => {
            encoder.write_u8(116);
            encoder.write_buf(value);
        }
    }
}

fn json_value(any: &Any) -> Value {
    match any {
        Any::Null | Any::Undefined => Value::Null,
        Any::Bool(value) => Value::Bool(*value),
        Any::Number(value) if *value as i64 as f64 == *value => {
            Value::Number(Number::from(*value as i64))
        }
        Any::Number(value) => Number::from_f64(*value)
            .map(Value::Number)
            .unwrap_or(Value::Null),
        Any::BigInt(value) => Value::Number(Number::from(*value)),
        Any::String(value) => Value::String(value.to_string()),
        Any::Buffer(value) => Value::Array(
            value
                .iter()
                .map(|byte| Value::Number(Number::from(*byte)))
                .collect(),
        ),
        Any::Array(values) => Value::Array(values.iter().map(json_value).collect()),
        Any::Map(values) => {
            let sorted: BTreeMap<_, _> = values
                .iter()
                .map(|(key, value)| (key.clone(), json_value(value)))
                .collect();
            Value::Object(sorted.into_iter().collect::<Map<_, _>>())
        }
    }
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use yrs::{Text, Transact};

    use super::*;
    use crate::{
        EditCtx, EditingDoc, FormatPolicy, InlineFormatDelta, Patch, Position, StoryRange,
    };

    fn assert_len(doc: &EditingDoc) {
        let txn = doc.yrs_doc().transact();
        assert_eq!(
            encoded_state_len_v1(&txn),
            Some(encode_state_as_update_v1(&txn, &StateVector::default()).len()),
        );
    }

    #[test]
    fn state_lengths_cover_text_attributes_comments_and_embeds() {
        let doc = EditingDoc::new(901);
        assert_len(&doc);
        doc.create_story("body", "Alpha 😀 beta", "Normal", "left")
            .unwrap();
        assert_len(&doc);
        let ctx = EditCtx::local("Ann", "2026-09-24T12:00:00Z");
        doc.insert_text(
            &ctx,
            Position::new("body", 6),
            "text ",
            FormatPolicy::Inherit,
        )
        .unwrap();
        assert_len(&doc);
        doc.delete_range(&ctx, StoryRange::new("body", 0, 2))
            .unwrap();
        assert_len(&doc);
        doc.format_range(
            &ctx,
            StoryRange::new("body", 0, 3),
            &InlineFormatDelta {
                bold: Patch::Set(true),
                italic: Patch::Set(true),
                ..Default::default()
            },
        )
        .unwrap();
        assert_len(&doc);
        doc.replace_range(
            &ctx.clone().suggesting(),
            StoryRange::new("body", 3, 5),
            "🚀",
        )
        .unwrap();
        assert_len(&doc);
        doc.add_comment(
            &[StoryRange::new("body", 0, 3)],
            "Ann",
            "2026-09-24T12:00:00Z",
            Any::from("remark"),
        )
        .unwrap();
        assert_len(&doc);
        doc.insert_embed(&ctx, Position::new("body", 0), "break", Vec::new())
            .unwrap();
        assert_len(&doc);
        {
            let mut txn = doc.yrs_doc().transact_mut();
            let story = crate::story_ref(&txn, "body").unwrap();
            story.insert_embed(
                &mut txn,
                1,
                Any::Map(Arc::new(std::collections::HashMap::from([
                    ("_kind".to_owned(), Any::from("inlineSdt")),
                    ("tag".to_owned(), Any::from("control")),
                    (
                        "content".to_owned(),
                        Any::Array(Arc::from([Any::from("value")])),
                    ),
                ]))),
            );
        }
        assert_len(&doc);
    }

    #[test]
    fn counting_encoder_matches_every_v1_field() {
        fn fields<E: Encoder>(encoder: &mut E) {
            encoder.reset_ds_cur_val();
            encoder.write_ds_clock(128);
            encoder.write_ds_len(u32::MAX);
            encoder.write_left_id(&ID::new(ClientID::new(129), 16_384));
            encoder.write_right_id(&ID::new(ClientID::new(1), u32::MAX));
            encoder.write_client(ClientID::new(256));
            encoder.write_info(255);
            encoder.write_parent_info(true);
            encoder.write_parent_info(false);
            encoder.write_type_ref(7);
            encoder.write_len(16_384);
            encoder.write_key("key 😀");
            for value in [
                Any::Null,
                Any::Undefined,
                Any::Bool(true),
                Any::Bool(false),
                Any::from("text 😀"),
                Any::Number(-129.0),
                Any::Number(0.5),
                Any::Number(0.1),
                Any::BigInt(i64::MIN),
                Any::Buffer(Arc::from([0, 127, 255])),
                Any::Array(Arc::from([Any::Bool(true), Any::from("child")])),
                Any::Map(Arc::new(std::collections::HashMap::from([
                    ("z".to_owned(), Any::Number(1.0)),
                    (
                        "a".to_owned(),
                        Any::Array(Arc::from([Any::Null, Any::from("😀")])),
                    ),
                ]))),
            ] {
                encoder.write_any(&value);
                encoder.write_json(&value);
            }
        }
        let mut actual = DeterministicEncoderV1::new();
        fields(&mut actual);
        let mut counted = CountingEncoderV1::default();
        fields(&mut counted);
        assert_eq!(counted.len, actual.to_vec().len());
    }

    #[test]
    fn pending_updates_and_deletions_have_no_proven_length() {
        let peer = EditingDoc::new(902);
        peer.create_story("body", "abc", "Normal", "left").unwrap();
        let known = peer.encode_state_vector_v1();
        peer.insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 1),
            "x",
            FormatPolicy::Plain,
        )
        .unwrap();
        let pending = EditingDoc::new(903);
        pending
            .apply_update_v1(&peer.encode_diff_v1(&known).unwrap())
            .unwrap();
        assert!(
            pending
                .yrs_doc()
                .transact()
                .store()
                .pending_update()
                .is_some()
        );
        assert_eq!(encoded_state_len_v1(&pending.yrs_doc().transact()), None);

        let known = peer.encode_state_vector_v1();
        peer.delete_range(&EditCtx::local("", ""), StoryRange::new("body", 0, 2))
            .unwrap();
        let pending = EditingDoc::new(904);
        pending
            .apply_update_v1(&peer.encode_diff_v1(&known).unwrap())
            .unwrap();
        assert!(pending.yrs_doc().transact().store().pending_ds().is_some());
        assert_eq!(encoded_state_len_v1(&pending.yrs_doc().transact()), None);
    }

    #[test]
    fn empty_update_preserves_integrated_state() {
        let doc = EditingDoc::new(905);
        doc.create_story("body", "abc", "Normal", "left").unwrap();
        let base = doc.encode_state_as_update_v1();
        let vector = doc.encode_state_vector_v1();
        let epoch = doc.committed_epoch();
        doc.yrs_doc()
            .transact_mut_with(crate::batch::HOST_ORIGIN)
            .apply_update(Update::new())
            .unwrap();
        assert_eq!(base, doc.encode_state_as_update_v1());
        assert_eq!(vector, doc.encode_state_vector_v1());
        assert_eq!(epoch, doc.committed_epoch());
    }
}
