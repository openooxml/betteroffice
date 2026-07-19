//! Streaming typed-page serialization for the FrameDelta encoder.
//!
//! Rebuilt pages previously detoured through `serde_json::to_value` — one
//! fully allocated `Value` tree per page, then four independent walks
//! (structural hash, visual hash, string collection, wire emission). The
//! three serializers here run directly over the typed [`DisplayPage`] via
//! serde, so a rebuilt page costs three allocation-free passes:
//!
//! - [`hash_page`] — structural + visual fingerprints in one pass;
//! - [`collect_page_strings`] — string-table population (upsert pages only);
//! - [`encode_page`] — the typed value stream (upsert pages only).
//!
//! Wire compatibility: the emitted value stream uses the same opcodes,
//! container framing, and glyph compaction as the `Value`-based encoder. The
//! only difference is object field ORDER (declaration order instead of
//! `serde_json::Map`'s sorted order), which the browser decoder is agnostic
//! to — it materializes fields one by one and rejects duplicates only.
//! Numeric opcode selection mirrors the old `as_i64`-first introspection:
//! unsigned values within `i64` range emit `VALUE_I64`. Non-finite floats
//! emit `VALUE_NULL`, exactly as `serde_json::to_value` mapped them.
//!
//! Fingerprints hash the same logical structure as before minus the key
//! sorting; they only ever compare against fingerprints produced by the same
//! session, so the definition change is invisible outside this module.

use std::collections::{BTreeSet, HashMap};
use std::fmt;

use docx_layout::display_list::DisplayPage;
use serde::Serialize;
use serde::ser::{self, Serializer};

use super::{
    GLYPH_BIDI_LEVEL, GLYPH_LOGICAL_ORDER, VALUE_ARRAY, VALUE_F64, VALUE_FALSE, VALUE_GLYPH_ARRAY,
    VALUE_I64, VALUE_NULL, VALUE_OBJECT, VALUE_STRING, VALUE_TRUE, VALUE_U64, checked_u32,
    hash_write, patch_u32, string_id, write_f64, write_i64, write_u32, write_u64,
};

pub(super) struct PageHashes {
    pub fingerprint: u64,
    pub visual_fingerprint: u64,
}

/// Structural + visual fingerprints in one streaming pass.
pub(super) fn hash_page(page: &DisplayPage) -> Result<PageHashes, String> {
    let mut state = HashState {
        fingerprint: super::FNV_OFFSET,
        visual: super::FNV_OFFSET,
    };
    page.serialize(HashSer {
        state: &mut state,
        fp_on: true,
        vfp_on: true,
        root: true,
        slot: Slot::None,
    })
    .map_err(|error| format!("hash display page: {error}"))?;
    Ok(PageHashes {
        fingerprint: state.fingerprint,
        visual_fingerprint: state.visual,
    })
}

/// Collect every string the wire encoding of `page` will reference.
pub(super) fn collect_page_strings(
    page: &DisplayPage,
    strings: &mut BTreeSet<String>,
) -> Result<(), String> {
    page.serialize(StrSer { strings })
        .map_err(|error| format!("collect display page strings: {error}"))
}

/// Emit the typed value stream for `page` using the prepared string table.
pub(super) fn encode_page(
    page: &DisplayPage,
    ids: &HashMap<&str, u32>,
    out: &mut Vec<u8>,
) -> Result<(), String> {
    page.serialize(EmitSer { ids, out })
        .map_err(|error| format!("encode display page: {error}"))
}

// ---------------------------------------------------------------------------
// shared plumbing
// ---------------------------------------------------------------------------

#[derive(Debug)]
pub(super) struct SerError(String);

impl fmt::Display for SerError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for SerError {}

impl ser::Error for SerError {
    fn custom<T: fmt::Display>(msg: T) -> Self {
        SerError(msg.to_string())
    }
}

fn unsupported<T>(what: &str) -> Result<T, SerError> {
    Err(SerError(format!(
        "display page serialization does not use {what}"
    )))
}

/// Key classifications the exclusion / compaction rules depend on.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Slot {
    None,
    Glyphs,
    InlineSdtWidget,
}

fn classify(key: &str) -> Slot {
    match key {
        "glyphs" => Slot::Glyphs,
        "inlineSdtWidget" => Slot::InlineSdtWidget,
        _ => Slot::None,
    }
}

fn is_position_key(key: &str) -> bool {
    matches!(
        key,
        "docStart" | "docEnd" | "fragmentDocStart" | "fragmentDocEnd"
    )
}

/// Captures one dynamic map key into a reusable buffer (Value-object entries;
/// typed structs pass `&'static str` keys directly and never hit this).
struct KeySer<'a>(&'a mut String);

impl Serializer for KeySer<'_> {
    type Ok = ();
    type Error = SerError;
    type SerializeSeq = ser::Impossible<(), SerError>;
    type SerializeTuple = ser::Impossible<(), SerError>;
    type SerializeTupleStruct = ser::Impossible<(), SerError>;
    type SerializeTupleVariant = ser::Impossible<(), SerError>;
    type SerializeMap = ser::Impossible<(), SerError>;
    type SerializeStruct = ser::Impossible<(), SerError>;
    type SerializeStructVariant = ser::Impossible<(), SerError>;

    fn serialize_str(self, value: &str) -> Result<(), SerError> {
        self.0.clear();
        self.0.push_str(value);
        Ok(())
    }

    fn serialize_char(self, value: char) -> Result<(), SerError> {
        self.0.clear();
        self.0.push(value);
        Ok(())
    }

    fn serialize_bool(self, _: bool) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_i8(self, _: i8) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_i16(self, _: i16) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_i32(self, _: i32) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_i64(self, _: i64) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_u8(self, _: u8) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_u16(self, _: u16) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_u32(self, _: u32) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_u64(self, _: u64) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_f32(self, _: f32) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_f64(self, _: f64) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_bytes(self, _: &[u8]) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_none(self) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<(), SerError> {
        value.serialize(self)
    }
    fn serialize_unit(self) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_unit_struct(self, _: &'static str) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_unit_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
    ) -> Result<(), SerError> {
        self.serialize_str(variant)
    }
    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        value.serialize(self)
    }
    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: &T,
    ) -> Result<(), SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_seq(self, _: Option<usize>) -> Result<Self::SerializeSeq, SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_tuple(self, _: usize) -> Result<Self::SerializeTuple, SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_tuple_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleStruct, SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleVariant, SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_map(self, _: Option<usize>) -> Result<Self::SerializeMap, SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStruct, SerError> {
        unsupported("non-string map keys")
    }
    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStructVariant, SerError> {
        unsupported("non-string map keys")
    }
}

// ---------------------------------------------------------------------------
// pass 1: fingerprints
// ---------------------------------------------------------------------------

struct HashState {
    fingerprint: u64,
    visual: u64,
}

impl HashState {
    fn write(&mut self, fp_on: bool, vfp_on: bool, bytes: &[u8]) {
        if fp_on {
            hash_write(&mut self.fingerprint, bytes);
        }
        if vfp_on {
            hash_write(&mut self.visual, bytes);
        }
    }
}

struct HashSer<'a> {
    state: &'a mut HashState,
    fp_on: bool,
    vfp_on: bool,
    /// True only for the page's outermost container: its direct fields apply
    /// the root-level `pageIndex` exclusion.
    root: bool,
    /// The key this value sits under (drives the `inlineSdtWidget.pos` rule).
    slot: Slot,
}

impl HashSer<'_> {
    fn write(&mut self, bytes: &[u8]) {
        self.state.write(self.fp_on, self.vfp_on, bytes);
    }
}

struct HashContainer<'a> {
    state: &'a mut HashState,
    fp_on: bool,
    vfp_on: bool,
    fields_root: bool,
    slot: Slot,
    key_buf: String,
}

impl HashContainer<'_> {
    fn field<T: Serialize + ?Sized>(&mut self, key: &str, value: &T) -> Result<(), SerError> {
        let fp_skip = self.fields_root && key == "pageIndex";
        let vfp_skip =
            fp_skip || is_position_key(key) || (self.slot == Slot::InlineSdtWidget && key == "pos");
        let fp_on = self.fp_on && !fp_skip;
        let vfp_on = self.vfp_on && !vfp_skip;
        if !fp_on && !vfp_on {
            return Ok(());
        }
        self.state.write(fp_on, vfp_on, key.as_bytes());
        value.serialize(HashSer {
            state: &mut *self.state,
            fp_on,
            vfp_on,
            root: false,
            slot: classify(key),
        })
    }

    fn element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        value.serialize(HashSer {
            state: &mut *self.state,
            fp_on: self.fp_on,
            vfp_on: self.vfp_on,
            root: false,
            // array elements inherit the array's slot (glyph objects sit
            // under "glyphs"), matching the Value-walk's parent_key rule
            slot: self.slot,
        })
    }
}

impl<'a> Serializer for HashSer<'a> {
    type Ok = ();
    type Error = SerError;
    type SerializeSeq = HashContainer<'a>;
    type SerializeTuple = Self::SerializeSeq;
    type SerializeTupleStruct = Self::SerializeSeq;
    type SerializeTupleVariant = Self::SerializeSeq;
    type SerializeMap = Self::SerializeSeq;
    type SerializeStruct = Self::SerializeSeq;
    type SerializeStructVariant = Self::SerializeSeq;

    fn serialize_bool(mut self, value: bool) -> Result<(), SerError> {
        self.write(&[if value { VALUE_TRUE } else { VALUE_FALSE }]);
        Ok(())
    }

    fn serialize_i8(self, value: i8) -> Result<(), SerError> {
        self.serialize_i64(value.into())
    }
    fn serialize_i16(self, value: i16) -> Result<(), SerError> {
        self.serialize_i64(value.into())
    }
    fn serialize_i32(self, value: i32) -> Result<(), SerError> {
        self.serialize_i64(value.into())
    }
    fn serialize_i64(mut self, value: i64) -> Result<(), SerError> {
        self.write(&[VALUE_I64]);
        self.write(&value.to_le_bytes());
        Ok(())
    }
    fn serialize_u8(self, value: u8) -> Result<(), SerError> {
        self.serialize_u64(value.into())
    }
    fn serialize_u16(self, value: u16) -> Result<(), SerError> {
        self.serialize_u64(value.into())
    }
    fn serialize_u32(self, value: u32) -> Result<(), SerError> {
        self.serialize_u64(value.into())
    }
    fn serialize_u64(mut self, value: u64) -> Result<(), SerError> {
        // mirrors the old Value introspection: as_i64() first
        if let Ok(signed) = i64::try_from(value) {
            self.write(&[VALUE_I64]);
            self.write(&signed.to_le_bytes());
        } else {
            self.write(&[VALUE_U64]);
            self.write(&value.to_le_bytes());
        }
        Ok(())
    }
    fn serialize_f32(self, value: f32) -> Result<(), SerError> {
        self.serialize_f64(value.into())
    }
    fn serialize_f64(mut self, value: f64) -> Result<(), SerError> {
        // serde_json::to_value maps non-finite floats to null
        if value.is_finite() {
            self.write(&[VALUE_F64]);
            self.write(&value.to_bits().to_le_bytes());
        } else {
            self.write(&[VALUE_NULL]);
        }
        Ok(())
    }

    fn serialize_char(self, value: char) -> Result<(), SerError> {
        self.serialize_str(value.encode_utf8(&mut [0; 4]))
    }
    fn serialize_str(mut self, value: &str) -> Result<(), SerError> {
        self.write(&[VALUE_STRING]);
        self.write(value.as_bytes());
        Ok(())
    }
    fn serialize_bytes(self, value: &[u8]) -> Result<(), SerError> {
        let mut seq = self.serialize_seq(Some(value.len()))?;
        for byte in value {
            ser::SerializeSeq::serialize_element(&mut seq, byte)?;
        }
        ser::SerializeSeq::end(seq)
    }

    fn serialize_none(mut self) -> Result<(), SerError> {
        self.write(&[VALUE_NULL]);
        Ok(())
    }
    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<(), SerError> {
        value.serialize(self)
    }
    fn serialize_unit(mut self) -> Result<(), SerError> {
        self.write(&[VALUE_NULL]);
        Ok(())
    }
    fn serialize_unit_struct(self, _: &'static str) -> Result<(), SerError> {
        self.serialize_unit()
    }
    fn serialize_unit_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
    ) -> Result<(), SerError> {
        self.serialize_str(variant)
    }
    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        value.serialize(self)
    }
    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        mut self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        // serde_json shape: {variant: value}
        self.write(&[VALUE_OBJECT]);
        self.state
            .write(self.fp_on, self.vfp_on, variant.as_bytes());
        value.serialize(HashSer {
            state: self.state,
            fp_on: self.fp_on,
            vfp_on: self.vfp_on,
            root: false,
            slot: Slot::None,
        })
    }

    fn serialize_seq(mut self, len: Option<usize>) -> Result<Self::SerializeSeq, SerError> {
        let len = len.ok_or_else(|| SerError("unsized sequence in display page".to_owned()))?;
        self.write(&[VALUE_ARRAY]);
        self.write(&(len as u64).to_le_bytes());
        Ok(HashContainer {
            state: self.state,
            fp_on: self.fp_on,
            vfp_on: self.vfp_on,
            fields_root: false,
            slot: self.slot,
            key_buf: String::new(),
        })
    }
    fn serialize_tuple(self, len: usize) -> Result<Self::SerializeTuple, SerError> {
        self.serialize_seq(Some(len))
    }
    fn serialize_tuple_struct(
        self,
        _: &'static str,
        len: usize,
    ) -> Result<Self::SerializeTupleStruct, SerError> {
        self.serialize_seq(Some(len))
    }
    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleVariant, SerError> {
        unsupported("tuple variants")
    }
    fn serialize_map(mut self, _: Option<usize>) -> Result<Self::SerializeMap, SerError> {
        self.write(&[VALUE_OBJECT]);
        Ok(HashContainer {
            state: self.state,
            fp_on: self.fp_on,
            vfp_on: self.vfp_on,
            fields_root: self.root,
            slot: self.slot,
            key_buf: String::new(),
        })
    }
    fn serialize_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStruct, SerError> {
        self.serialize_map(None)
    }
    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStructVariant, SerError> {
        unsupported("struct variants")
    }
}

impl ser::SerializeSeq for HashContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        self.element(value)
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

impl ser::SerializeTuple for HashContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        self.element(value)
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

impl ser::SerializeTupleStruct for HashContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        self.element(value)
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

impl ser::SerializeTupleVariant for HashContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        self.element(value)
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

impl ser::SerializeMap for HashContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), SerError> {
        let mut buf = std::mem::take(&mut self.key_buf);
        key.serialize(KeySer(&mut buf))?;
        self.key_buf = buf;
        Ok(())
    }
    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        let key = std::mem::take(&mut self.key_buf);
        let result = self.field(&key, value);
        self.key_buf = key;
        result
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

impl ser::SerializeStruct for HashContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        self.field(key, value)
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

impl ser::SerializeStructVariant for HashContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        self.field(key, value)
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// pass 2: string-table collection (upsert pages only)
// ---------------------------------------------------------------------------

struct StrSer<'a> {
    strings: &'a mut BTreeSet<String>,
}

struct StrContainer<'a> {
    strings: &'a mut BTreeSet<String>,
    key_buf: String,
}

impl StrContainer<'_> {
    fn field<T: Serialize + ?Sized>(&mut self, key: &str, value: &T) -> Result<(), SerError> {
        if !self.strings.contains(key) {
            self.strings.insert(key.to_owned());
        }
        // compact glyph arrays carry no strings on the wire
        if key == "glyphs" && probe_glyphs(value).is_some() {
            return Ok(());
        }
        value.serialize(StrSer {
            strings: &mut *self.strings,
        })
    }
}

impl<'a> Serializer for StrSer<'a> {
    type Ok = ();
    type Error = SerError;
    type SerializeSeq = StrContainer<'a>;
    type SerializeTuple = Self::SerializeSeq;
    type SerializeTupleStruct = Self::SerializeSeq;
    type SerializeTupleVariant = Self::SerializeSeq;
    type SerializeMap = Self::SerializeSeq;
    type SerializeStruct = Self::SerializeSeq;
    type SerializeStructVariant = Self::SerializeSeq;

    fn serialize_bool(self, _: bool) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_i8(self, _: i8) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_i16(self, _: i16) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_i32(self, _: i32) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_i64(self, _: i64) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_u8(self, _: u8) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_u16(self, _: u16) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_u32(self, _: u32) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_u64(self, _: u64) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_f32(self, _: f32) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_f64(self, _: f64) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_char(self, value: char) -> Result<(), SerError> {
        self.serialize_str(value.encode_utf8(&mut [0; 4]))
    }
    fn serialize_str(self, value: &str) -> Result<(), SerError> {
        if !self.strings.contains(value) {
            self.strings.insert(value.to_owned());
        }
        Ok(())
    }
    fn serialize_bytes(self, _: &[u8]) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_none(self) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<(), SerError> {
        value.serialize(self)
    }
    fn serialize_unit(self) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_unit_struct(self, _: &'static str) -> Result<(), SerError> {
        Ok(())
    }
    fn serialize_unit_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
    ) -> Result<(), SerError> {
        self.serialize_str(variant)
    }
    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        value.serialize(self)
    }
    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        if !self.strings.contains(variant) {
            self.strings.insert(variant.to_owned());
        }
        value.serialize(StrSer {
            strings: self.strings,
        })
    }
    fn serialize_seq(self, _: Option<usize>) -> Result<Self::SerializeSeq, SerError> {
        Ok(StrContainer {
            strings: self.strings,
            key_buf: String::new(),
        })
    }
    fn serialize_tuple(self, len: usize) -> Result<Self::SerializeTuple, SerError> {
        self.serialize_seq(Some(len))
    }
    fn serialize_tuple_struct(
        self,
        _: &'static str,
        len: usize,
    ) -> Result<Self::SerializeTupleStruct, SerError> {
        self.serialize_seq(Some(len))
    }
    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleVariant, SerError> {
        unsupported("tuple variants")
    }
    fn serialize_map(self, _: Option<usize>) -> Result<Self::SerializeMap, SerError> {
        Ok(StrContainer {
            strings: self.strings,
            key_buf: String::new(),
        })
    }
    fn serialize_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStruct, SerError> {
        self.serialize_map(None)
    }
    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStructVariant, SerError> {
        unsupported("struct variants")
    }
}

impl ser::SerializeSeq for StrContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        value.serialize(StrSer {
            strings: &mut *self.strings,
        })
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

impl ser::SerializeTuple for StrContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        ser::SerializeSeq::serialize_element(self, value)
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

impl ser::SerializeTupleStruct for StrContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        ser::SerializeSeq::serialize_element(self, value)
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

impl ser::SerializeTupleVariant for StrContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        ser::SerializeSeq::serialize_element(self, value)
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

impl ser::SerializeMap for StrContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), SerError> {
        let mut buf = std::mem::take(&mut self.key_buf);
        key.serialize(KeySer(&mut buf))?;
        self.key_buf = buf;
        Ok(())
    }
    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        let key = std::mem::take(&mut self.key_buf);
        let result = self.field(&key, value);
        self.key_buf = key;
        result
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

impl ser::SerializeStruct for StrContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        self.field(key, value)
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

impl ser::SerializeStructVariant for StrContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        self.field(key, value)
    }
    fn end(self) -> Result<(), SerError> {
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// pass 3: wire emission (upsert pages only)
// ---------------------------------------------------------------------------

struct EmitSer<'a> {
    ids: &'a HashMap<&'a str, u32>,
    out: &'a mut Vec<u8>,
}

struct EmitContainer<'a> {
    ids: &'a HashMap<&'a str, u32>,
    out: &'a mut Vec<u8>,
    length_at: usize,
    count_at: usize,
    payload_at: usize,
    count: u32,
    key_buf: String,
    /// Object entries as `(key id, byte offset of the entry's key id)`, so a
    /// duplicate key — a named struct field colliding with a `serde(flatten)`
    /// member — can be compacted to the last write, exactly like
    /// `serde_json::Map` insertion did on the old `to_value` path. The browser
    /// decoder rejects duplicate object keys, so this is a correctness rule,
    /// not cosmetics.
    entries: Vec<(u32, usize)>,
}

impl<'a> EmitContainer<'a> {
    fn open(opcode: u8, ids: &'a HashMap<&'a str, u32>, out: &'a mut Vec<u8>) -> Self {
        out.push(opcode);
        let length_at = out.len();
        write_u32(out, 0);
        let count_at = out.len();
        write_u32(out, 0);
        let payload_at = out.len();
        EmitContainer {
            ids,
            out,
            length_at,
            count_at,
            payload_at,
            count: 0,
            key_buf: String::new(),
            entries: Vec::new(),
        }
    }

    fn close(mut self) -> Result<(), SerError> {
        self.compact_duplicate_keys();
        let payload_len = self.out.len() - self.payload_at;
        patch_u32(
            self.out,
            self.length_at,
            checked_u32(payload_len, "container payload length").map_err(SerError)?,
        );
        patch_u32(self.out, self.count_at, self.count);
        Ok(())
    }

    /// Drop every non-final occurrence of a repeated object key (last write
    /// wins). Duplicate-free objects — the overwhelming majority — pay one
    /// linear scan and nothing else; typed value payloads only carry relative
    /// lengths, so splicing bytes out never invalidates inner containers.
    fn compact_duplicate_keys(&mut self) {
        let has_duplicate = self.entries.iter().enumerate().any(|(index, (key, _))| {
            self.entries[index + 1..]
                .iter()
                .any(|(other, _)| other == key)
        });
        if !has_duplicate {
            return;
        }
        let mut retained = Vec::with_capacity(self.out.len() - self.payload_at);
        let mut kept = 0_u32;
        for (index, (key, start)) in self.entries.iter().enumerate() {
            let end = self
                .entries
                .get(index + 1)
                .map_or(self.out.len(), |(_, next_start)| *next_start);
            let superseded = self.entries[index + 1..]
                .iter()
                .any(|(other, _)| other == key);
            if superseded {
                continue;
            }
            retained.extend_from_slice(&self.out[*start..end]);
            kept += 1;
        }
        self.out.truncate(self.payload_at);
        self.out.extend_from_slice(&retained);
        self.count = kept;
    }

    fn element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        self.count = self
            .count
            .checked_add(1)
            .ok_or_else(|| SerError("container element count exceeds u32".to_owned()))?;
        value.serialize(EmitSer {
            ids: self.ids,
            out: &mut *self.out,
        })
    }

    fn field<T: Serialize + ?Sized>(&mut self, key: &str, value: &T) -> Result<(), SerError> {
        self.count = self
            .count
            .checked_add(1)
            .ok_or_else(|| SerError("object field count exceeds u32".to_owned()))?;
        let key_id = string_id(self.ids, key).map_err(SerError)?;
        self.entries.push((key_id, self.out.len()));
        write_u32(self.out, key_id);
        if key == "glyphs"
            && let Some(glyphs) = probe_glyphs(value)
        {
            return emit_glyph_array(&glyphs, self.out);
        }
        value.serialize(EmitSer {
            ids: self.ids,
            out: &mut *self.out,
        })
    }
}

impl<'a> Serializer for EmitSer<'a> {
    type Ok = ();
    type Error = SerError;
    type SerializeSeq = EmitContainer<'a>;
    type SerializeTuple = Self::SerializeSeq;
    type SerializeTupleStruct = Self::SerializeSeq;
    type SerializeTupleVariant = Self::SerializeSeq;
    type SerializeMap = Self::SerializeSeq;
    type SerializeStruct = Self::SerializeSeq;
    type SerializeStructVariant = Self::SerializeSeq;

    fn serialize_bool(self, value: bool) -> Result<(), SerError> {
        self.out.push(if value { VALUE_TRUE } else { VALUE_FALSE });
        Ok(())
    }
    fn serialize_i8(self, value: i8) -> Result<(), SerError> {
        self.serialize_i64(value.into())
    }
    fn serialize_i16(self, value: i16) -> Result<(), SerError> {
        self.serialize_i64(value.into())
    }
    fn serialize_i32(self, value: i32) -> Result<(), SerError> {
        self.serialize_i64(value.into())
    }
    fn serialize_i64(self, value: i64) -> Result<(), SerError> {
        self.out.push(VALUE_I64);
        write_i64(self.out, value);
        Ok(())
    }
    fn serialize_u8(self, value: u8) -> Result<(), SerError> {
        self.serialize_u64(value.into())
    }
    fn serialize_u16(self, value: u16) -> Result<(), SerError> {
        self.serialize_u64(value.into())
    }
    fn serialize_u32(self, value: u32) -> Result<(), SerError> {
        self.serialize_u64(value.into())
    }
    fn serialize_u64(self, value: u64) -> Result<(), SerError> {
        // mirrors the old Value introspection: as_i64() first
        if let Ok(signed) = i64::try_from(value) {
            self.out.push(VALUE_I64);
            write_i64(self.out, signed);
        } else {
            self.out.push(VALUE_U64);
            write_u64(self.out, value);
        }
        Ok(())
    }
    fn serialize_f32(self, value: f32) -> Result<(), SerError> {
        self.serialize_f64(value.into())
    }
    fn serialize_f64(self, value: f64) -> Result<(), SerError> {
        // serde_json::to_value maps non-finite floats to null
        if value.is_finite() {
            self.out.push(VALUE_F64);
            write_f64(self.out, value);
        } else {
            self.out.push(VALUE_NULL);
        }
        Ok(())
    }
    fn serialize_char(self, value: char) -> Result<(), SerError> {
        let mut buf = [0; 4];
        let text: &str = value.encode_utf8(&mut buf);
        self.out.push(VALUE_STRING);
        write_u32(self.out, string_id(self.ids, text).map_err(SerError)?);
        Ok(())
    }
    fn serialize_str(self, value: &str) -> Result<(), SerError> {
        self.out.push(VALUE_STRING);
        write_u32(self.out, string_id(self.ids, value).map_err(SerError)?);
        Ok(())
    }
    fn serialize_bytes(self, value: &[u8]) -> Result<(), SerError> {
        let mut seq = self.serialize_seq(Some(value.len()))?;
        for byte in value {
            ser::SerializeSeq::serialize_element(&mut seq, byte)?;
        }
        ser::SerializeSeq::end(seq)
    }
    fn serialize_none(self) -> Result<(), SerError> {
        self.out.push(VALUE_NULL);
        Ok(())
    }
    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<(), SerError> {
        value.serialize(self)
    }
    fn serialize_unit(self) -> Result<(), SerError> {
        self.out.push(VALUE_NULL);
        Ok(())
    }
    fn serialize_unit_struct(self, _: &'static str) -> Result<(), SerError> {
        self.serialize_unit()
    }
    fn serialize_unit_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
    ) -> Result<(), SerError> {
        self.serialize_str(variant)
    }
    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        value.serialize(self)
    }
    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        // serde_json shape: {variant: value}
        let mut container = EmitContainer::open(VALUE_OBJECT, self.ids, self.out);
        container.field(variant, value)?;
        container.close()
    }
    fn serialize_seq(self, _: Option<usize>) -> Result<Self::SerializeSeq, SerError> {
        Ok(EmitContainer::open(VALUE_ARRAY, self.ids, self.out))
    }
    fn serialize_tuple(self, len: usize) -> Result<Self::SerializeTuple, SerError> {
        self.serialize_seq(Some(len))
    }
    fn serialize_tuple_struct(
        self,
        _: &'static str,
        len: usize,
    ) -> Result<Self::SerializeTupleStruct, SerError> {
        self.serialize_seq(Some(len))
    }
    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleVariant, SerError> {
        unsupported("tuple variants")
    }
    fn serialize_map(self, _: Option<usize>) -> Result<Self::SerializeMap, SerError> {
        Ok(EmitContainer::open(VALUE_OBJECT, self.ids, self.out))
    }
    fn serialize_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStruct, SerError> {
        self.serialize_map(None)
    }
    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStructVariant, SerError> {
        unsupported("struct variants")
    }
}

impl ser::SerializeSeq for EmitContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        self.element(value)
    }
    fn end(self) -> Result<(), SerError> {
        self.close()
    }
}

impl ser::SerializeTuple for EmitContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        self.element(value)
    }
    fn end(self) -> Result<(), SerError> {
        self.close()
    }
}

impl ser::SerializeTupleStruct for EmitContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        self.element(value)
    }
    fn end(self) -> Result<(), SerError> {
        self.close()
    }
}

impl ser::SerializeTupleVariant for EmitContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        self.element(value)
    }
    fn end(self) -> Result<(), SerError> {
        self.close()
    }
}

impl ser::SerializeMap for EmitContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), SerError> {
        let mut buf = std::mem::take(&mut self.key_buf);
        key.serialize(KeySer(&mut buf))?;
        self.key_buf = buf;
        Ok(())
    }
    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        let key = std::mem::take(&mut self.key_buf);
        let result = self.field(&key, value);
        self.key_buf = key;
        result
    }
    fn end(self) -> Result<(), SerError> {
        self.close()
    }
}

impl ser::SerializeStruct for EmitContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        self.field(key, value)
    }
    fn end(self) -> Result<(), SerError> {
        self.close()
    }
}

impl ser::SerializeStructVariant for EmitContainer<'_> {
    type Ok = ();
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        self.field(key, value)
    }
    fn end(self) -> Result<(), SerError> {
        self.close()
    }
}

// ---------------------------------------------------------------------------
// compact glyph probe (shared by the strings and emit passes)
// ---------------------------------------------------------------------------

/// One glyph eligible for the fixed-field wire payload. Range checks beyond
/// eligibility (id/cluster into u32) happen at emission, exactly like the
/// `Value`-based encoder: an out-of-range id is a hard error, not a fallback.
struct CompactGlyph {
    id: u64,
    x: f64,
    y: f64,
    cluster: u64,
    advance: f64,
    logical_order: Option<u64>,
    bidi_level: Option<u64>,
}

/// Serialize `value` through the eligibility probe. `None` means the value is
/// not a compact-eligible glyph array and must encode generically.
fn probe_glyphs<T: Serialize + ?Sized>(value: &T) -> Option<Vec<CompactGlyph>> {
    value.serialize(GlyphSeqProbe).ok()
}

fn emit_glyph_array(glyphs: &[CompactGlyph], out: &mut Vec<u8>) -> Result<(), SerError> {
    out.push(VALUE_GLYPH_ARRAY);
    let length_at = out.len();
    write_u32(out, 0);
    write_u32(
        out,
        checked_u32(glyphs.len(), "glyph array count").map_err(SerError)?,
    );
    let payload_at = out.len();
    for glyph in glyphs {
        write_u32(
            out,
            u32::try_from(glyph.id).map_err(|_| SerError("glyph id exceeds u32".to_owned()))?,
        );
        write_f64(out, glyph.x);
        write_f64(out, glyph.y);
        write_u32(
            out,
            u32::try_from(glyph.cluster)
                .map_err(|_| SerError("glyph cluster exceeds u32".to_owned()))?,
        );
        write_f64(out, glyph.advance);
        let flags = if glyph.logical_order.is_some() {
            GLYPH_LOGICAL_ORDER
        } else {
            0
        } | if glyph.bidi_level.is_some() {
            GLYPH_BIDI_LEVEL
        } else {
            0
        };
        out.push(flags);
        if let Some(value) = glyph.logical_order {
            write_u64(out, value);
        }
        if let Some(value) = glyph.bidi_level {
            out.push(value as u8);
        }
    }
    let payload_len = out.len() - payload_at;
    patch_u32(
        out,
        length_at,
        checked_u32(payload_len, "glyph array payload length").map_err(SerError)?,
    );
    Ok(())
}

fn ineligible<T>() -> Result<T, SerError> {
    Err(SerError("glyph array is not compact-eligible".to_owned()))
}

/// Numeric capture for one glyph field.
#[derive(Clone, Copy)]
enum GlyphNum {
    Int(i64),
    Uint(u64),
    Float(f64),
}

impl GlyphNum {
    /// `Value::as_u64` equivalence.
    fn as_u64(self) -> Option<u64> {
        match self {
            GlyphNum::Uint(value) => Some(value),
            GlyphNum::Int(value) => u64::try_from(value).ok(),
            GlyphNum::Float(_) => None,
        }
    }

    /// `Value::as_f64` equivalence.
    fn as_f64(self) -> Option<f64> {
        match self {
            GlyphNum::Uint(value) => Some(value as f64),
            GlyphNum::Int(value) => Some(value as f64),
            GlyphNum::Float(value) => value.is_finite().then_some(value),
        }
    }
}

struct GlyphNumProbe;

impl Serializer for GlyphNumProbe {
    type Ok = GlyphNum;
    type Error = SerError;
    type SerializeSeq = ser::Impossible<GlyphNum, SerError>;
    type SerializeTuple = ser::Impossible<GlyphNum, SerError>;
    type SerializeTupleStruct = ser::Impossible<GlyphNum, SerError>;
    type SerializeTupleVariant = ser::Impossible<GlyphNum, SerError>;
    type SerializeMap = ser::Impossible<GlyphNum, SerError>;
    type SerializeStruct = ser::Impossible<GlyphNum, SerError>;
    type SerializeStructVariant = ser::Impossible<GlyphNum, SerError>;

    fn serialize_i8(self, value: i8) -> Result<GlyphNum, SerError> {
        Ok(GlyphNum::Int(value.into()))
    }
    fn serialize_i16(self, value: i16) -> Result<GlyphNum, SerError> {
        Ok(GlyphNum::Int(value.into()))
    }
    fn serialize_i32(self, value: i32) -> Result<GlyphNum, SerError> {
        Ok(GlyphNum::Int(value.into()))
    }
    fn serialize_i64(self, value: i64) -> Result<GlyphNum, SerError> {
        Ok(GlyphNum::Int(value))
    }
    fn serialize_u8(self, value: u8) -> Result<GlyphNum, SerError> {
        Ok(GlyphNum::Uint(value.into()))
    }
    fn serialize_u16(self, value: u16) -> Result<GlyphNum, SerError> {
        Ok(GlyphNum::Uint(value.into()))
    }
    fn serialize_u32(self, value: u32) -> Result<GlyphNum, SerError> {
        Ok(GlyphNum::Uint(value.into()))
    }
    fn serialize_u64(self, value: u64) -> Result<GlyphNum, SerError> {
        Ok(GlyphNum::Uint(value))
    }
    fn serialize_f32(self, value: f32) -> Result<GlyphNum, SerError> {
        Ok(GlyphNum::Float(value.into()))
    }
    fn serialize_f64(self, value: f64) -> Result<GlyphNum, SerError> {
        Ok(GlyphNum::Float(value))
    }
    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<GlyphNum, SerError> {
        value.serialize(self)
    }
    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        value: &T,
    ) -> Result<GlyphNum, SerError> {
        value.serialize(self)
    }

    fn serialize_bool(self, _: bool) -> Result<GlyphNum, SerError> {
        ineligible()
    }
    fn serialize_char(self, _: char) -> Result<GlyphNum, SerError> {
        ineligible()
    }
    fn serialize_str(self, _: &str) -> Result<GlyphNum, SerError> {
        ineligible()
    }
    fn serialize_bytes(self, _: &[u8]) -> Result<GlyphNum, SerError> {
        ineligible()
    }
    fn serialize_none(self) -> Result<GlyphNum, SerError> {
        ineligible()
    }
    fn serialize_unit(self) -> Result<GlyphNum, SerError> {
        ineligible()
    }
    fn serialize_unit_struct(self, _: &'static str) -> Result<GlyphNum, SerError> {
        ineligible()
    }
    fn serialize_unit_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
    ) -> Result<GlyphNum, SerError> {
        ineligible()
    }
    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: &T,
    ) -> Result<GlyphNum, SerError> {
        ineligible()
    }
    fn serialize_seq(self, _: Option<usize>) -> Result<Self::SerializeSeq, SerError> {
        ineligible()
    }
    fn serialize_tuple(self, _: usize) -> Result<Self::SerializeTuple, SerError> {
        ineligible()
    }
    fn serialize_tuple_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleStruct, SerError> {
        ineligible()
    }
    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleVariant, SerError> {
        ineligible()
    }
    fn serialize_map(self, _: Option<usize>) -> Result<Self::SerializeMap, SerError> {
        ineligible()
    }
    fn serialize_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStruct, SerError> {
        ineligible()
    }
    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStructVariant, SerError> {
        ineligible()
    }
}

struct GlyphSeqProbe;

impl Serializer for GlyphSeqProbe {
    type Ok = Vec<CompactGlyph>;
    type Error = SerError;
    type SerializeSeq = GlyphSeqBuilder;
    type SerializeTuple = ser::Impossible<Vec<CompactGlyph>, SerError>;
    type SerializeTupleStruct = ser::Impossible<Vec<CompactGlyph>, SerError>;
    type SerializeTupleVariant = ser::Impossible<Vec<CompactGlyph>, SerError>;
    type SerializeMap = ser::Impossible<Vec<CompactGlyph>, SerError>;
    type SerializeStruct = ser::Impossible<Vec<CompactGlyph>, SerError>;
    type SerializeStructVariant = ser::Impossible<Vec<CompactGlyph>, SerError>;

    fn serialize_seq(self, len: Option<usize>) -> Result<Self::SerializeSeq, SerError> {
        Ok(GlyphSeqBuilder {
            glyphs: Vec::with_capacity(len.unwrap_or(0)),
        })
    }
    fn serialize_some<T: Serialize + ?Sized>(
        self,
        value: &T,
    ) -> Result<Vec<CompactGlyph>, SerError> {
        value.serialize(self)
    }
    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        value: &T,
    ) -> Result<Vec<CompactGlyph>, SerError> {
        value.serialize(self)
    }

    fn serialize_bool(self, _: bool) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_i8(self, _: i8) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_i16(self, _: i16) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_i32(self, _: i32) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_i64(self, _: i64) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_u8(self, _: u8) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_u16(self, _: u16) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_u32(self, _: u32) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_u64(self, _: u64) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_f32(self, _: f32) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_f64(self, _: f64) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_char(self, _: char) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_str(self, _: &str) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_bytes(self, _: &[u8]) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_none(self) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_unit(self) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_unit_struct(self, _: &'static str) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_unit_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
    ) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: &T,
    ) -> Result<Vec<CompactGlyph>, SerError> {
        ineligible()
    }
    fn serialize_tuple(self, _: usize) -> Result<Self::SerializeTuple, SerError> {
        ineligible()
    }
    fn serialize_tuple_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleStruct, SerError> {
        ineligible()
    }
    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleVariant, SerError> {
        ineligible()
    }
    fn serialize_map(self, _: Option<usize>) -> Result<Self::SerializeMap, SerError> {
        ineligible()
    }
    fn serialize_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStruct, SerError> {
        ineligible()
    }
    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStructVariant, SerError> {
        ineligible()
    }
}

struct GlyphSeqBuilder {
    glyphs: Vec<CompactGlyph>,
}

impl ser::SerializeSeq for GlyphSeqBuilder {
    type Ok = Vec<CompactGlyph>;
    type Error = SerError;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        self.glyphs.push(value.serialize(GlyphElemProbe)?);
        Ok(())
    }
    fn end(self) -> Result<Vec<CompactGlyph>, SerError> {
        Ok(self.glyphs)
    }
}

const GLYPH_SEEN_ID: u8 = 1 << 0;
const GLYPH_SEEN_X: u8 = 1 << 1;
const GLYPH_SEEN_Y: u8 = 1 << 2;
const GLYPH_SEEN_CLUSTER: u8 = 1 << 3;
const GLYPH_SEEN_ADVANCE: u8 = 1 << 4;
const GLYPH_SEEN_LOGICAL_ORDER: u8 = 1 << 5;
const GLYPH_SEEN_BIDI_LEVEL: u8 = 1 << 6;
const GLYPH_SEEN_MANDATORY: u8 =
    GLYPH_SEEN_ID | GLYPH_SEEN_X | GLYPH_SEEN_Y | GLYPH_SEEN_CLUSTER | GLYPH_SEEN_ADVANCE;

struct GlyphElemProbe;

impl Serializer for GlyphElemProbe {
    type Ok = CompactGlyph;
    type Error = SerError;
    type SerializeSeq = ser::Impossible<CompactGlyph, SerError>;
    type SerializeTuple = ser::Impossible<CompactGlyph, SerError>;
    type SerializeTupleStruct = ser::Impossible<CompactGlyph, SerError>;
    type SerializeTupleVariant = ser::Impossible<CompactGlyph, SerError>;
    type SerializeMap = GlyphFieldProbe;
    type SerializeStruct = GlyphFieldProbe;
    type SerializeStructVariant = ser::Impossible<CompactGlyph, SerError>;

    fn serialize_map(self, _: Option<usize>) -> Result<Self::SerializeMap, SerError> {
        Ok(GlyphFieldProbe::default())
    }
    fn serialize_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStruct, SerError> {
        Ok(GlyphFieldProbe::default())
    }
    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<CompactGlyph, SerError> {
        value.serialize(self)
    }
    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        value: &T,
    ) -> Result<CompactGlyph, SerError> {
        value.serialize(self)
    }

    fn serialize_bool(self, _: bool) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_i8(self, _: i8) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_i16(self, _: i16) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_i32(self, _: i32) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_i64(self, _: i64) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_u8(self, _: u8) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_u16(self, _: u16) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_u32(self, _: u32) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_u64(self, _: u64) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_f32(self, _: f32) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_f64(self, _: f64) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_char(self, _: char) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_str(self, _: &str) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_bytes(self, _: &[u8]) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_none(self) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_unit(self) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_unit_struct(self, _: &'static str) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_unit_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
    ) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: &T,
    ) -> Result<CompactGlyph, SerError> {
        ineligible()
    }
    fn serialize_seq(self, _: Option<usize>) -> Result<Self::SerializeSeq, SerError> {
        ineligible()
    }
    fn serialize_tuple(self, _: usize) -> Result<Self::SerializeTuple, SerError> {
        ineligible()
    }
    fn serialize_tuple_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleStruct, SerError> {
        ineligible()
    }
    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleVariant, SerError> {
        ineligible()
    }
    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStructVariant, SerError> {
        ineligible()
    }
}

#[derive(Default)]
struct GlyphFieldProbe {
    seen: u8,
    id: u64,
    x: f64,
    y: f64,
    cluster: u64,
    advance: f64,
    logical_order: Option<u64>,
    bidi_level: Option<u64>,
    key_buf: String,
}

impl GlyphFieldProbe {
    fn field<T: Serialize + ?Sized>(&mut self, key: &str, value: &T) -> Result<(), SerError> {
        let (bit, number) = match key {
            "id" => (GLYPH_SEEN_ID, value.serialize(GlyphNumProbe)?),
            "x" => (GLYPH_SEEN_X, value.serialize(GlyphNumProbe)?),
            "y" => (GLYPH_SEEN_Y, value.serialize(GlyphNumProbe)?),
            "cluster" => (GLYPH_SEEN_CLUSTER, value.serialize(GlyphNumProbe)?),
            "advance" => (GLYPH_SEEN_ADVANCE, value.serialize(GlyphNumProbe)?),
            "logicalOrder" => (GLYPH_SEEN_LOGICAL_ORDER, value.serialize(GlyphNumProbe)?),
            "bidiLevel" => (GLYPH_SEEN_BIDI_LEVEL, value.serialize(GlyphNumProbe)?),
            _ => return ineligible(),
        };
        if self.seen & bit != 0 {
            return ineligible();
        }
        self.seen |= bit;
        match bit {
            GLYPH_SEEN_ID => self.id = number.as_u64().map_or_else(ineligible, Ok)?,
            GLYPH_SEEN_X => self.x = number.as_f64().map_or_else(ineligible, Ok)?,
            GLYPH_SEEN_Y => self.y = number.as_f64().map_or_else(ineligible, Ok)?,
            GLYPH_SEEN_CLUSTER => self.cluster = number.as_u64().map_or_else(ineligible, Ok)?,
            GLYPH_SEEN_ADVANCE => self.advance = number.as_f64().map_or_else(ineligible, Ok)?,
            GLYPH_SEEN_LOGICAL_ORDER => {
                self.logical_order = Some(number.as_u64().map_or_else(ineligible, Ok)?);
            }
            _ => {
                let level = number.as_u64().map_or_else(ineligible, Ok)?;
                if level > u8::MAX as u64 {
                    return ineligible();
                }
                self.bidi_level = Some(level);
            }
        }
        Ok(())
    }

    fn finish(self) -> Result<CompactGlyph, SerError> {
        if self.seen & GLYPH_SEEN_MANDATORY != GLYPH_SEEN_MANDATORY {
            return ineligible();
        }
        Ok(CompactGlyph {
            id: self.id,
            x: self.x,
            y: self.y,
            cluster: self.cluster,
            advance: self.advance,
            logical_order: self.logical_order,
            bidi_level: self.bidi_level,
        })
    }
}

impl ser::SerializeMap for GlyphFieldProbe {
    type Ok = CompactGlyph;
    type Error = SerError;
    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), SerError> {
        let mut buf = std::mem::take(&mut self.key_buf);
        key.serialize(KeySer(&mut buf))?;
        self.key_buf = buf;
        Ok(())
    }
    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), SerError> {
        let key = std::mem::take(&mut self.key_buf);
        let result = self.field(&key, value);
        self.key_buf = key;
        result
    }
    fn end(self) -> Result<CompactGlyph, SerError> {
        self.finish()
    }
}

impl ser::SerializeStruct for GlyphFieldProbe {
    type Ok = CompactGlyph;
    type Error = SerError;
    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), SerError> {
        self.field(key, value)
    }
    fn end(self) -> Result<CompactGlyph, SerError> {
        self.finish()
    }
}
