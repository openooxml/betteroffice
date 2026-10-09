use std::collections::BTreeMap;

use serde_json::{Map, Number, Value};
use yrs::any::{F64_MAX_SAFE_INTEGER, F64_MIN_SAFE_INTEGER};
use yrs::block::{BLOCK_GC_REF_NUMBER, BLOCK_ITEM_DELETED_REF_NUMBER, HAS_PARENT_SUB};
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
    info: u8,
    unstable_json: bool,
    root_sequence: bool,
    tombstone_len: u64,
    deleted_len: u64,
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
        self.deleted_len += u64::from(len);
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
        self.info = info;
        self.write_u8(info);
    }

    fn write_parent_info(&mut self, is_y_key: bool) {
        self.root_sequence |= is_y_key && self.info & HAS_PARENT_SUB == 0;
        self.write_var(u32::from(is_y_key));
    }

    fn write_type_ref(&mut self, info: u8) {
        self.write_u8(info);
    }

    fn write_len(&mut self, len: u32) {
        if self.info == BLOCK_GC_REF_NUMBER || self.info & 0x1F == BLOCK_ITEM_DELETED_REF_NUMBER {
            self.tombstone_len += u64::from(len);
        }
        self.write_var(len);
    }

    fn write_any(&mut self, any: &Any) {
        encode_any(any, self);
    }

    fn write_json(&mut self, any: &Any) {
        let json = serde_json::to_string(&json_value(any)).unwrap();
        self.write_string(&json);
        if !self.unstable_json && !json_stable_fast(any) && !json_stable(any, &json) {
            self.unstable_json = true;
        }
    }

    fn write_key(&mut self, key: &str) {
        self.write_string(key);
    }
}

/// Returns None for pending data, unstable JSON values, root sequence content, or deleted content
/// a fork would collect.
pub(crate) fn fork_state_len_v1<T: ReadTxn>(txn: &T) -> Option<usize> {
    let store = txn.store();
    if store.pending_update().is_some() || store.pending_ds().is_some() {
        return None;
    }
    let mut encoder = CountingEncoderV1::default();
    txn.encode_state_as_update(&StateVector::default(), &mut encoder);
    (!encoder.unstable_json
        && !encoder.root_sequence
        && encoder.tombstone_len == encoder.deleted_len)
        .then_some(encoder.len)
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

fn json_stable_fast(any: &Any) -> bool {
    json_stable_fast_at(any, 0)
}

fn json_stable_fast_at(any: &Any, depth: usize) -> bool {
    match any {
        Any::Null | Any::Bool(_) => true,
        Any::String(value) => value
            .chars()
            .all(|ch| !ch.is_control() && ch != '"' && ch != '\\'),
        Any::Number(value) => {
            value.abs() < 2_147_483_648.0
                && value.trunc() == *value
                && (*value != 0.0 || !value.is_sign_negative())
        }
        Any::Array(values) => {
            depth < 32 && values.iter().all(|v| json_stable_fast_at(v, depth + 1))
        }
        Any::Map(values) => {
            depth < 32 && values.values().all(|v| json_stable_fast_at(v, depth + 1))
        }
        _ => false,
    }
}

fn json_stable(any: &Any, json: &str) -> bool {
    let Ok(round_trip) = Any::from_json(json) else {
        return false;
    };
    let mut original = Vec::new();
    let mut decoded = Vec::new();
    encode_any(any, &mut original);
    encode_any(&round_trip, &mut decoded);
    original == decoded
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

    use yrs::{Map, Text, Transact};

    use super::*;
    use crate::{
        EditCtx, EditingDoc, FormatPolicy, InlineFormatDelta, Patch, Position, StoryRange,
    };

    fn assert_len(doc: &EditingDoc) {
        let txn = doc.yrs_doc().transact();
        assert_eq!(
            fork_state_len_v1(&txn),
            Some(encode_state_as_update_v1(&txn, &StateVector::default()).len()),
        );
    }

    fn assert_json_stability(value: &Any, stable: bool) {
        let json = serde_json::to_string(&json_value(value)).unwrap();
        assert_eq!(json_stable(value, &json), stable, "{value:?}");
        let mut encoder = CountingEncoderV1::default();
        encoder.write_json(value);
        assert_eq!(!encoder.unstable_json, stable, "{value:?}");
    }

    #[test]
    fn json_stability_matches_deterministic_any_bytes() {
        for (value, stable) in [
            (Any::Null, true),
            (Any::Bool(true), true),
            (Any::Bool(false), true),
            (Any::from("text 😀"), true),
            (Any::from("\"quoted\" \\ slash\n\t\u{0}"), true),
            (Any::Number(0.5), true),
            (Any::Number(12.0), true),
            (Any::Number(0.0), true),
            (Any::Number(-0.0), true),
            (Any::Number(2_147_483_647.0), true),
            (Any::Number(-2_147_483_647.0), true),
            (Any::Number(2_147_483_648.0), true),
            (Any::Number(-2_147_483_648.0), true),
            (Any::Number(f64::NAN), false),
            (Any::Number(f64::INFINITY), false),
            (Any::Number(f64::NEG_INFINITY), false),
            (Any::Number((1_u64 << 60) as f64), false),
            (Any::BigInt(7), false),
            (Any::Buffer(Arc::from([0, 127, 255])), false),
            (Any::Undefined, false),
        ] {
            assert_json_stability(&value, stable);
        }
        for (member, stable) in [
            (Any::from("text 😀"), true),
            (Any::Number(12.0), true),
            (Any::Number(0.5), true),
            (Any::from("\"quoted\" \\ slash\n"), true),
            (Any::BigInt(7), false),
            (Any::Buffer(Arc::from([1, 2])), false),
            (Any::Undefined, false),
            (Any::Number(f64::NAN), false),
            (Any::Number(f64::INFINITY), false),
        ] {
            let array = Any::Array(Arc::from([Any::Null, member]));
            let map = Any::Map(Arc::new(std::collections::HashMap::from([
                ("z".to_owned(), Any::Bool(true)),
                ("a".to_owned(), array.clone()),
            ])));
            assert_json_stability(&array, stable);
            assert_json_stability(&map, stable);
            let nested = Any::Array(Arc::from([map]));
            assert_json_stability(&nested, stable);
        }
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
    fn unstable_format_values_have_no_fork_length_in_any_story_or_paragraph() {
        for value in [
            Any::Number(f64::INFINITY),
            Any::BigInt(7),
            Any::Buffer(Arc::from([1, 2, 3])),
        ] {
            for (story_id, index) in [("body", 0), ("body", 3), ("header", 0)] {
                let doc = EditingDoc::new(901);
                doc.seed_story(
                    "body",
                    &["AB", "CD"].map(|text| crate::SeedParagraph {
                        text: text.to_owned(),
                        p_style: "Normal".to_owned(),
                        alignment: "left".to_owned(),
                    }),
                )
                .unwrap();
                doc.create_story("header", "Header", "Normal", "left")
                    .unwrap();
                assert_len(&doc);
                {
                    let mut txn = doc.yrs_doc().transact_mut();
                    crate::story_ref(&txn, story_id).unwrap().format(
                        &mut txn,
                        index,
                        1,
                        std::collections::HashMap::from([("opaque".into(), value.clone())]),
                    );
                }
                assert_eq!(fork_state_len_v1(&doc.yrs_doc().transact()), None);
            }
        }
    }

    #[test]
    fn unstable_embeds_have_no_fork_length() {
        for value in [
            Any::Number(f64::INFINITY),
            Any::BigInt(7),
            Any::Buffer(Arc::from([1, 2, 3])),
        ] {
            let doc = EditingDoc::new(901);
            doc.create_story("body", "AB", "Normal", "left").unwrap();
            doc.create_story("header", "Header", "Normal", "left")
                .unwrap();
            {
                let mut txn = doc.yrs_doc().transact_mut();
                crate::story_ref(&txn, "header")
                    .unwrap()
                    .insert_embed(&mut txn, 0, value);
            }
            assert_eq!(fork_state_len_v1(&doc.yrs_doc().transact()), None);
        }
    }

    #[test]
    fn non_json_map_values_preserve_exact_fork_lengths() {
        let doc = EditingDoc::new(901);
        {
            let mut txn = doc.yrs_doc().transact_mut();
            let comments = txn.get_map(crate::COMMENTS).unwrap();
            for (key, value) in [
                ("number", Any::Number(f64::INFINITY)),
                ("integer", Any::BigInt(7)),
                ("buffer", Any::Buffer(Arc::from([1, 2, 3]))),
            ] {
                comments.insert(&mut txn, key, value);
            }
        }
        assert_len(&doc);
    }

    #[test]
    fn declared_root_sequences_have_no_fork_length_even_after_deletion() {
        for formatted in [false, true] {
            let doc = EditingDoc::new(901);
            doc.create_story("body", "AB", "Normal", "left").unwrap();
            let text = doc.yrs_doc().get_or_insert_text(crate::COMMENTS);
            assert_len(&doc);
            if formatted {
                text.insert_with_attributes(
                    &mut doc.yrs_doc().transact_mut(),
                    0,
                    "root text",
                    std::collections::HashMap::from([("bold".into(), Any::Bool(true))]),
                );
            } else {
                text.insert(&mut doc.yrs_doc().transact_mut(), 0, "root text");
            }
            assert_eq!(fork_state_len_v1(&doc.yrs_doc().transact()), None);
            text.remove_range(&mut doc.yrs_doc().transact_mut(), 0, 9);
            let txn = doc.yrs_doc().transact();
            assert_eq!(text.len(&txn), 0);
            assert_eq!(fork_state_len_v1(&txn), None);
        }
    }

    #[test]
    fn retained_deleted_content_has_no_fork_length() {
        let ctx = EditCtx::local("Ann", "2026-09-24T12:00:00Z");
        for tracked in [false, true] {
            let doc = EditingDoc::new(901);
            doc.create_story("body", "ABX", "Normal", "left").unwrap();
            let history = tracked.then(|| doc.undo_manager());
            doc.delete_range(&ctx, StoryRange::new("body", 2, 3))
                .unwrap();
            if tracked {
                assert_eq!(history.unwrap().undo_depth(), 1);
                assert_eq!(fork_state_len_v1(&doc.yrs_doc().transact()), None);
            } else {
                assert_len(&doc);
            }
        }
    }

    #[test]
    fn deeply_nested_json_takes_the_exact_proof() {
        let nest =
            |depth: usize| (0..depth).fold(Any::Null, |inner, _| Any::Array(Arc::from([inner])));
        assert!(json_stable_fast(&nest(32)));
        for (depth, stable) in [(33, true), (128, false)] {
            let value = nest(depth);
            assert!(!json_stable_fast(&value));
            assert_json_stability(&value, stable);
        }
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
        assert_eq!(fork_state_len_v1(&pending.yrs_doc().transact()), None);

        let known = peer.encode_state_vector_v1();
        peer.delete_range(&EditCtx::local("", ""), StoryRange::new("body", 0, 2))
            .unwrap();
        let pending = EditingDoc::new(904);
        pending
            .apply_update_v1(&peer.encode_diff_v1(&known).unwrap())
            .unwrap();
        assert!(pending.yrs_doc().transact().store().pending_ds().is_some());
        assert_eq!(fork_state_len_v1(&pending.yrs_doc().transact()), None);
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
