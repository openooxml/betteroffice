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
