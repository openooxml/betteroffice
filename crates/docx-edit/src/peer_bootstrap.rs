use std::collections::{HashMap, HashSet};
use std::fmt;
use std::hash::Hash;
use std::sync::Arc;

use docx_parse::media::{MediaDescriptor, MediaTable};
use ooxml_opc::{PackageBytes, RetainedPackage};
use serde::{Deserialize, Deserializer, Serialize, Serializer};
use serde_json::Value;
use sha2::{Digest, Sha256};
use yrs::updates::decoder::Decode;
use yrs::{ReadTxn, Transact, Update};

use crate::EditingDoc;
use crate::identity::{SourceIndex, SourcePackage};
use crate::seed::SourceMetadata;

const MAGIC: [u8; 8] = *b"BOPEER\0\0";
const FORMAT_VERSION: u32 = 1;
const SHAPE: &str = "docx-edit.peer-metadata/unseeded-v1;source-serde-v1;identity-occurrences-v1;media-blobs-v1;yrs-v1;comments-first-load-rebase;no-pins-or-occurrence-safety";
const HEADER_LEN: usize = 8 + 4 + 32 + 8 + 8;

#[derive(Debug, Eq, PartialEq)]
pub enum PeerMetadataError {
    MissingSource,
    SourceRequired,
    NonEmptyDocument,
    InvalidState(String),
    BadMagic,
    UnsupportedVersion(u32),
    ShapeMismatch,
    Truncated,
    InvalidLength,
    Json(String),
    InvalidMetadata(String),
    SourceMismatch,
    Package(String),
}

impl fmt::Display for PeerMetadataError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::MissingSource => {
                f.write_str("peer metadata requires an indexed source and media")
            }
            Self::SourceRequired => f.write_str("peer bootstrap requires retained source bytes"),
            Self::NonEmptyDocument => f.write_str("peer bootstrap requires an empty document"),
            Self::InvalidState(error) => write!(f, "invalid peer state: {error}"),
            Self::BadMagic => f.write_str("invalid peer metadata magic"),
            Self::UnsupportedVersion(version) => {
                write!(f, "unsupported peer metadata version {version}")
            }
            Self::ShapeMismatch => f.write_str("peer metadata shape fingerprint mismatch"),
            Self::Truncated => f.write_str("truncated peer metadata"),
            Self::InvalidLength => f.write_str("invalid peer metadata section length"),
            Self::Json(error) => write!(f, "invalid peer metadata JSON: {error}"),
            Self::InvalidMetadata(error) => write!(f, "invalid peer metadata: {error}"),
            Self::SourceMismatch => f.write_str("peer metadata retained source mismatch"),
            Self::Package(error) => write!(f, "invalid peer retained package: {error}"),
        }
    }
}

impl std::error::Error for PeerMetadataError {}

impl PeerMetadataError {
    pub fn code(&self) -> &'static str {
        match self {
            Self::MissingSource => "missing-source",
            Self::SourceRequired => "source-required",
            Self::NonEmptyDocument => "non-empty-document",
            Self::InvalidState(_) => "invalid-state",
            Self::BadMagic => "bad-magic",
            Self::UnsupportedVersion(_) => "unsupported-version",
            Self::ShapeMismatch => "shape-mismatch",
            Self::Truncated => "truncated",
            Self::InvalidLength => "invalid-length",
            Self::Json(_) => "invalid-json",
            Self::InvalidMetadata(_) => "invalid-metadata",
            Self::SourceMismatch => "source-mismatch",
            Self::Package(_) => "invalid-package",
        }
    }
}

pub struct PeerBootstrap {
    metadata: PeerMetadata,
}

pub struct PeerBootstrapSource {
    pub source: PackageBytes,
    pub digest: String,
}

impl EditingDoc {
    pub fn encode_peer_metadata(&self) -> Result<Vec<u8>, PeerMetadataError> {
        encode(self)
    }

    fn require_empty_peer(&self) -> Result<(), PeerMetadataError> {
        let txn = self.yrs_doc().transact();
        if !txn.state_vector().is_empty() || txn.has_missing_updates() {
            return Err(PeerMetadataError::NonEmptyDocument);
        }
        Ok(())
    }

    pub fn prepare_peer_bootstrap(
        &self,
        state: &[u8],
        metadata: &[u8],
        source: Option<PackageBytes>,
    ) -> Result<PeerBootstrap, PeerMetadataError> {
        sections(metadata)?;
        self.require_empty_peer()?;
        let update = Update::decode_v1(state)
            .map_err(|error| PeerMetadataError::InvalidState(error.to_string()))?;
        let (source, digest) = match source {
            Some(source) => {
                let digest = crate::seed::package_digest(&source);
                (source, digest)
            }
            None => {
                let retained = self
                    .source
                    .lock()
                    .unwrap_or_else(|error| error.into_inner());
                match retained.as_ref() {
                    Some(SourcePackage::Ready(index)) => {
                        (index.bytes(), index.package_digest().to_owned())
                    }
                    Some(SourcePackage::Pending(bytes, digest)) => (
                        bytes.clone(),
                        digest
                            .clone()
                            .unwrap_or_else(|| crate::seed::package_digest(bytes)),
                    ),
                    None => return Err(PeerMetadataError::SourceRequired),
                }
            }
        };
        let metadata = decode(metadata, source, &digest)?;
        let validation = EditingDoc::new(self.client_id);
        validation
            .yrs_doc()
            .transact_mut()
            .apply_update(update)
            .map_err(|error| PeerMetadataError::InvalidState(error.to_string()))?;
        Ok(PeerBootstrap { metadata })
    }

    pub fn install_peer_bootstrap(
        &self,
        bootstrap: PeerBootstrap,
        entropy: u64,
    ) -> Result<PeerBootstrapSource, PeerMetadataError> {
        self.require_empty_peer()?;
        let PeerMetadata {
            mut source,
            index,
            media,
            retained_source,
            source_digest,
            comment_baseline: CommentBaseline::RebaseOnFirstLoad,
            media_sources,
        } = bootstrap.metadata;
        source.watch_comments(self);
        self.install_source(source, entropy);
        self.retain_source(SourcePackage::Ready(Arc::new(index)));
        self.install_media(media);
        self.set_media_sources(media_sources);
        Ok(PeerBootstrapSource {
            source: retained_source,
            digest: source_digest,
        })
    }
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
pub(crate) enum CommentBaseline {
    RebaseOnFirstLoad,
}

pub(crate) struct PeerMetadata {
    pub source: SourceMetadata,
    pub index: SourceIndex,
    pub media: MediaTable,
    pub retained_source: PackageBytes,
    pub source_digest: String,
    pub comment_baseline: CommentBaseline,
    pub media_sources: crate::media::MediaSources,
}

#[derive(Serialize)]
struct EncodeWire<'a> {
    source_digest: &'a str,
    source_length: u64,
    comment_baseline: CommentBaseline,
    source: &'a SourceMetadata,
    index: &'a SourceIndex,
    media: MediaWire,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct DecodeWire {
    source_digest: String,
    source_length: u64,
    comment_baseline: CommentBaseline,
    source: SourceMetadata,
    index: SourceIndex,
    media: MediaWire,
}

#[derive(Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
struct MediaWire {
    parts: Vec<MediaPartWire>,
    warnings: Vec<String>,
}

#[derive(Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
struct MediaPartWire {
    path: String,
    position: usize,
    size: u64,
    image: bool,
    mime_type: String,
    #[serde(deserialize_with = "required_option")]
    display: Option<Blob>,
}

#[derive(Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
struct Blob {
    offset: u64,
    length: u64,
}

fn valid_digest(digest: &str) -> bool {
    digest.len() == 64
        && digest
            .bytes()
            .all(|byte| matches!(byte, b'0'..=b'9' | b'a'..=b'f'))
}

fn shape_fingerprint() -> [u8; 32] {
    Sha256::digest(SHAPE.as_bytes()).into()
}

pub(crate) fn encode(document: &EditingDoc) -> Result<Vec<u8>, PeerMetadataError> {
    let source = document
        .source_metadata()
        .ok_or(PeerMetadataError::MissingSource)?;
    let index = {
        let retained = document
            .source
            .lock()
            .unwrap_or_else(|error| error.into_inner());
        match retained.as_ref() {
            Some(SourcePackage::Ready(index)) => Arc::clone(index),
            _ => return Err(PeerMetadataError::MissingSource),
        }
    };
    let media = document
        .media
        .lock()
        .unwrap_or_else(|error| error.into_inner())
        .clone()
        .ok_or(PeerMetadataError::MissingSource)?;
    let mut blobs = Vec::new();
    let parts = media
        .descriptors()
        .into_iter()
        .map(|part| {
            let display = part.display.map(|bytes| {
                let blob = Blob {
                    offset: blobs.len() as u64,
                    length: bytes.len() as u64,
                };
                blobs.extend_from_slice(&bytes);
                blob
            });
            MediaPartWire {
                path: part.path,
                position: part.position,
                size: part.size,
                image: part.image,
                mime_type: part.mime_type,
                display,
            }
        })
        .collect();
    let wire = EncodeWire {
        source_digest: index.package_digest(),
        source_length: index.bytes().len() as u64,
        comment_baseline: CommentBaseline::RebaseOnFirstLoad,
        source: &source,
        index: &index,
        media: MediaWire {
            parts,
            warnings: media.warnings().to_vec(),
        },
    };
    let value =
        serde_json::to_value(wire).map_err(|error| PeerMetadataError::Json(error.to_string()))?;
    let json = serde_json::to_vec(&canonical(value))
        .map_err(|error| PeerMetadataError::Json(error.to_string()))?;
    frame(&json, &blobs)
}

fn canonical(value: Value) -> Value {
    match value {
        Value::Object(object) => {
            let mut entries: Vec<_> = object.into_iter().collect();
            entries.sort_by(|a, b| a.0.cmp(&b.0));
            Value::Object(
                entries
                    .into_iter()
                    .map(|(key, value)| (key, canonical(value)))
                    .collect(),
            )
        }
        Value::Array(values) => Value::Array(values.into_iter().map(canonical).collect()),
        value => value,
    }
}

fn frame(json: &[u8], blobs: &[u8]) -> Result<Vec<u8>, PeerMetadataError> {
    let length = HEADER_LEN
        .checked_add(json.len())
        .and_then(|length| length.checked_add(blobs.len()))
        .ok_or(PeerMetadataError::InvalidLength)?;
    let mut bytes = Vec::with_capacity(length);
    bytes.extend_from_slice(&MAGIC);
    bytes.extend_from_slice(&FORMAT_VERSION.to_le_bytes());
    bytes.extend_from_slice(&shape_fingerprint());
    bytes.extend_from_slice(&(json.len() as u64).to_le_bytes());
    bytes.extend_from_slice(&(blobs.len() as u64).to_le_bytes());
    bytes.extend_from_slice(json);
    bytes.extend_from_slice(blobs);
    Ok(bytes)
}

fn sections(bytes: &[u8]) -> Result<(&[u8], &[u8]), PeerMetadataError> {
    if bytes.len() < HEADER_LEN {
        return Err(PeerMetadataError::Truncated);
    }
    if bytes[..8] != MAGIC {
        return Err(PeerMetadataError::BadMagic);
    }
    let mut version = [0; 4];
    version.copy_from_slice(&bytes[8..12]);
    let version = u32::from_le_bytes(version);
    if version != FORMAT_VERSION {
        return Err(PeerMetadataError::UnsupportedVersion(version));
    }
    if bytes[12..44] != shape_fingerprint() {
        return Err(PeerMetadataError::ShapeMismatch);
    }
    let length_at = |offset| {
        let mut length = [0; 8];
        length.copy_from_slice(&bytes[offset..offset + 8]);
        usize::try_from(u64::from_le_bytes(length)).map_err(|_| PeerMetadataError::InvalidLength)
    };
    let json_end = HEADER_LEN
        .checked_add(length_at(44)?)
        .ok_or(PeerMetadataError::InvalidLength)?;
    let end = json_end
        .checked_add(length_at(52)?)
        .ok_or(PeerMetadataError::InvalidLength)?;
    if end > bytes.len() {
        return Err(PeerMetadataError::Truncated);
    }
    if end != bytes.len() {
        return Err(PeerMetadataError::InvalidLength);
    }
    Ok((&bytes[HEADER_LEN..json_end], &bytes[json_end..end]))
}

pub(crate) fn decode(
    bytes: &[u8],
    retained_source: PackageBytes,
    known_digest: &str,
) -> Result<PeerMetadata, PeerMetadataError> {
    let (json, blobs) = sections(bytes)?;
    let mut wire: DecodeWire =
        serde_json::from_slice(json).map_err(|error| PeerMetadataError::Json(error.to_string()))?;
    if wire.source_digest != known_digest
        || wire.index.package_digest() != known_digest
        || wire.source_length != retained_source.len() as u64
        || !valid_digest(known_digest)
    {
        return Err(PeerMetadataError::SourceMismatch);
    }
    wire.source
        .rebuild_peer_metadata()
        .map_err(PeerMetadataError::InvalidMetadata)?;
    wire.index
        .attach_peer_source(retained_source.clone())
        .map_err(PeerMetadataError::InvalidMetadata)?;
    let mut next_offset = 0usize;
    let mut descriptors = Vec::with_capacity(wire.media.parts.len());
    for part in wire.media.parts {
        let display = if let Some(blob) = part.display {
            let offset =
                usize::try_from(blob.offset).map_err(|_| PeerMetadataError::InvalidLength)?;
            let length =
                usize::try_from(blob.length).map_err(|_| PeerMetadataError::InvalidLength)?;
            let end = offset
                .checked_add(length)
                .ok_or(PeerMetadataError::InvalidLength)?;
            if offset != next_offset {
                return Err(PeerMetadataError::InvalidLength);
            }
            let data = blobs
                .get(offset..end)
                .ok_or(PeerMetadataError::Truncated)?
                .to_vec();
            next_offset = end;
            Some(data)
        } else {
            None
        };
        descriptors.push(MediaDescriptor {
            path: part.path,
            position: part.position,
            size: part.size,
            image: part.image,
            mime_type: part.mime_type,
            display,
        });
    }
    if next_offset != blobs.len() {
        return Err(PeerMetadataError::InvalidLength);
    }
    let package =
        RetainedPackage::from_bytes(retained_source.clone()).map_err(PeerMetadataError::Package)?;
    let media = MediaTable::from_descriptors(package, descriptors, wire.media.warnings)
        .map_err(PeerMetadataError::InvalidMetadata)?;
    Ok(PeerMetadata {
        source: wire.source,
        index: wire.index,
        media,
        retained_source,
        source_digest: wire.source_digest,
        comment_baseline: wire.comment_baseline,
        media_sources: crate::media::MediaSources::default(),
    })
}

pub(crate) fn required_option<'de, D, T>(deserializer: D) -> Result<Option<T>, D::Error>
where
    D: Deserializer<'de>,
    T: Deserialize<'de>,
{
    Option::deserialize(deserializer)
}

pub(crate) mod sorted_map {
    use super::*;

    pub fn serialize<S, K, V>(map: &HashMap<K, V>, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
        K: Ord + Serialize,
        V: Serialize,
    {
        let mut entries: Vec<_> = map.iter().collect();
        entries.sort_by(|a, b| a.0.cmp(b.0));
        entries.serialize(serializer)
    }

    pub fn deserialize<'de, D, K, V>(deserializer: D) -> Result<HashMap<K, V>, D::Error>
    where
        D: Deserializer<'de>,
        K: Eq + Hash + Deserialize<'de>,
        V: Deserialize<'de>,
    {
        let entries = Vec::<(K, V)>::deserialize(deserializer)?;
        let mut map = HashMap::with_capacity(entries.len());
        for (key, value) in entries {
            if map.insert(key, value).is_some() {
                return Err(serde::de::Error::custom("duplicate metadata map key"));
            }
        }
        Ok(map)
    }
}

pub(crate) mod sorted_set {
    use super::*;

    pub fn serialize<S, T>(set: &HashSet<T>, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
        T: Ord + Serialize,
    {
        let mut values: Vec<_> = set.iter().collect();
        values.sort();
        values.serialize(serializer)
    }

    pub fn deserialize<'de, D, T>(deserializer: D) -> Result<HashSet<T>, D::Error>
    where
        D: Deserializer<'de>,
        T: Eq + Hash + Deserialize<'de>,
    {
        let values = Vec::<T>::deserialize(deserializer)?;
        let mut set = HashSet::with_capacity(values.len());
        for value in values {
            if !set.insert(value) {
                return Err(serde::de::Error::custom("duplicate metadata set value"));
            }
        }
        Ok(set)
    }
}

pub(crate) mod optional_map {
    use super::*;

    pub fn serialize<S, K, V>(map: &Option<HashMap<K, V>>, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
        K: Ord + Serialize,
        V: Serialize,
    {
        let entries = map.as_ref().map(|map| {
            let mut entries: Vec<_> = map.iter().collect();
            entries.sort_by(|a, b| a.0.cmp(b.0));
            entries
        });
        entries.serialize(serializer)
    }

    pub fn deserialize<'de, D, K, V>(deserializer: D) -> Result<Option<HashMap<K, V>>, D::Error>
    where
        D: Deserializer<'de>,
        K: Eq + Hash + Deserialize<'de>,
        V: Deserialize<'de>,
    {
        Option::<Vec<(K, V)>>::deserialize(deserializer)?
            .map(|entries| {
                let mut map = HashMap::with_capacity(entries.len());
                for (key, value) in entries {
                    if map.insert(key, value).is_some() {
                        return Err(serde::de::Error::custom("duplicate metadata map key"));
                    }
                }
                Ok(map)
            })
            .transpose()
    }
}

pub(crate) mod numbering {
    use super::*;

    pub fn serialize<S: Serializer>(
        numbering: &Arc<docx_parse::NumberingMap>,
        serializer: S,
    ) -> Result<S::Ok, S::Error> {
        numbering.definitions.serialize(serializer)
    }

    pub fn deserialize<'de, D: Deserializer<'de>>(
        deserializer: D,
    ) -> Result<Arc<docx_parse::NumberingMap>, D::Error> {
        Ok(Arc::new(docx_parse::NumberingMap {
            definitions: Deserialize::deserialize(deserializer)?,
        }))
    }
}

pub(crate) mod occurrences {
    use std::ops::Range;

    use docx_parse::paragraph_identity::ParagraphOccurrence;

    use super::*;

    #[derive(Deserialize, Serialize)]
    #[serde(deny_unknown_fields)]
    struct WireOccurrence {
        ordinal: u32,
        tag: Range<usize>,
        #[serde(deserialize_with = "required_option")]
        para_id: Option<String>,
        #[serde(deserialize_with = "required_option")]
        item_id: Option<String>,
    }

    pub fn serialize<S: Serializer>(
        occurrences: &[ParagraphOccurrence],
        serializer: S,
    ) -> Result<S::Ok, S::Error> {
        let occurrences: Vec<_> = occurrences
            .iter()
            .map(|occurrence| WireOccurrence {
                ordinal: occurrence.ordinal,
                tag: occurrence.tag.clone(),
                para_id: occurrence.para_id.clone(),
                item_id: occurrence.item_id.clone(),
            })
            .collect();
        occurrences.serialize(serializer)
    }

    pub fn deserialize<'de, D: Deserializer<'de>>(
        deserializer: D,
    ) -> Result<Vec<ParagraphOccurrence>, D::Error> {
        Ok(Vec::<WireOccurrence>::deserialize(deserializer)?
            .into_iter()
            .map(|occurrence| ParagraphOccurrence {
                ordinal: occurrence.ordinal,
                tag: occurrence.tag,
                para_id: occurrence.para_id,
                item_id: occurrence.item_id,
            })
            .collect())
    }
}

pub(crate) mod steps {
    use super::*;
    use crate::structured::source::Step;

    #[derive(Deserialize, Serialize)]
    enum WireStep {
        Body,
        Footnote(String),
        Endnote(String),
        Comment(String),
        Block(usize),
        Row(usize),
        Cell(usize),
        Content,
    }

    pub fn serialize<S: Serializer>(steps: &[Step], serializer: S) -> Result<S::Ok, S::Error> {
        let steps = steps
            .iter()
            .map(|step| {
                Ok(match step {
                    Step::Body => WireStep::Body,
                    Step::Note("footnote", id) => WireStep::Footnote(id.clone()),
                    Step::Note("endnote", id) => WireStep::Endnote(id.clone()),
                    Step::Note(_, _) => {
                        return Err(serde::ser::Error::custom("unknown source note locator"));
                    }
                    Step::Comment(id) => WireStep::Comment(id.clone()),
                    Step::Block(index) => WireStep::Block(*index),
                    Step::Row(index) => WireStep::Row(*index),
                    Step::Cell(index) => WireStep::Cell(*index),
                    Step::Content => WireStep::Content,
                })
            })
            .collect::<Result<Vec<_>, S::Error>>()?;
        steps.serialize(serializer)
    }

    pub fn deserialize<'de, D: Deserializer<'de>>(deserializer: D) -> Result<Vec<Step>, D::Error> {
        Ok(Vec::<WireStep>::deserialize(deserializer)?
            .into_iter()
            .map(|step| match step {
                WireStep::Body => Step::Body,
                WireStep::Footnote(id) => Step::Note("footnote", id),
                WireStep::Endnote(id) => Step::Note("endnote", id),
                WireStep::Comment(id) => Step::Comment(id),
                WireStep::Block(index) => Step::Block(index),
                WireStep::Row(index) => Step::Row(index),
                WireStep::Cell(index) => Step::Cell(index),
                WireStep::Content => Step::Content,
            })
            .collect())
    }
}

#[cfg(test)]
mod tests {
    use std::path::{Path, PathBuf};

    use serde_json::json;
    use yrs::block::{BLOCK_ITEM_ANY_REF_NUMBER, HAS_PARENT_SUB};
    use yrs::encoding::write::Write;
    use yrs::updates::encoder::{Encoder, EncoderV1};
    use yrs::{Any, Map, MapPrelim, MapRef, ReadTxn, Transact};

    use super::*;
    use crate::seed::{self, SeedMedia, fixture};
    use crate::structured::source::ReadSource;

    fn synthetic_package() -> PackageBytes {
        ooxml_opc::rezip_parts(&fixture::principal_parts())
            .unwrap()
            .into()
    }

    #[test]
    fn peer_bootstrap_installs_normalized_metadata_and_shares_source() {
        let source = synthetic_package();
        let digest = seed::package_digest(&source);
        let worker = worker(&source, &digest, 7).unwrap();
        let metadata = worker.encode_peer_metadata().unwrap();
        let state = worker.encode_state_as_update_v1();
        let peer = EditingDoc::new(19);
        let version = peer.version();
        let prepared = peer
            .prepare_peer_bootstrap(&state, &metadata, Some(source.clone()))
            .unwrap();
        assert!(peer.source_metadata().is_none());
        assert_eq!(peer.version(), version);
        let retained = peer.install_peer_bootstrap(prepared, 17).unwrap();
        assert_ne!(peer.version(), version);
        assert_eq!(retained.digest, digest);
        assert_eq!(retained.source.as_ptr(), source.as_ptr());
        let index = peer.source_index().unwrap();
        assert_eq!(index.bytes().as_ptr(), source.as_ptr());
        assert!(!index.seed_states_initialized());
        let read = peer.source_metadata().unwrap();
        assert_peer_runtime(read.read());
        assert!(peer.media_sources().is_empty());
        let expected = decode(&metadata, source.clone(), &digest).unwrap();
        assert_fields(
            &canonical(serde_json::to_value(read.as_ref()).unwrap()),
            &canonical(serde_json::to_value(&expected.source).unwrap()),
            "installed.source",
        );
        assert_eq!(
            peer.media.lock().unwrap().as_ref().unwrap().descriptors(),
            expected.media.descriptors(),
        );
        peer.apply_update_v1(&state).unwrap();
        assert_eq!(peer.encode_state_as_update_v1(), state);
        assert_eq!(peer.paragraph_identities(), worker.paragraph_identities());
        assert!(!read.read().comment_writes.snapshot().is_empty());
    }

    #[test]
    fn peer_bootstrap_rejections_leave_native_session_untouched() {
        let source = synthetic_package();
        let digest = seed::package_digest(&source);
        let worker = worker(&source, &digest, 7).unwrap();
        let metadata = worker.encode_peer_metadata().unwrap();
        let state = worker.encode_state_as_update_v1();
        let mut cases = Vec::new();
        for (offset, code) in [
            (0, "bad-magic"),
            (8, "unsupported-version"),
            (12, "shape-mismatch"),
        ] {
            let mut bad = metadata.clone();
            bad[offset] ^= 1;
            cases.push((bad, code));
        }
        cases.push((metadata[..HEADER_LEN - 1].to_vec(), "truncated"));
        cases.push((metadata[..metadata.len() - 1].to_vec(), "truncated"));
        let mut bad = metadata.clone();
        bad.push(0);
        cases.push((bad, "invalid-length"));
        let mut bad = metadata.clone();
        bad[44..52].copy_from_slice(&u64::MAX.to_le_bytes());
        cases.push((bad, "invalid-length"));
        let mut bad = metadata.clone();
        bad[HEADER_LEN] = b'!';
        cases.push((bad, "invalid-json"));
        let (json, blobs) = sections(&metadata).unwrap();
        let wire: Value = serde_json::from_slice(json).unwrap();
        let mut bad = wire.clone();
        bad["source"].as_object_mut().unwrap().remove("read");
        cases.push((
            frame(&serde_json::to_vec(&bad).unwrap(), blobs).unwrap(),
            "invalid-json",
        ));
        let mut bad = wire.clone();
        bad["index"]["parts"][0]["uri"] = json!("");
        cases.push((
            frame(&serde_json::to_vec(&bad).unwrap(), blobs).unwrap(),
            "invalid-metadata",
        ));
        let mut bad = wire.clone();
        bad["media"]["parts"][0]["display"] = json!({"offset": 1, "length": 0});
        cases.push((
            frame(&serde_json::to_vec(&bad).unwrap(), blobs).unwrap(),
            "invalid-length",
        ));
        let mut bad = wire;
        bad["source_digest"] = json!("0".repeat(64));
        cases.push((
            frame(&serde_json::to_vec(&bad).unwrap(), blobs).unwrap(),
            "source-mismatch",
        ));
        for (bad, code) in cases {
            let peer = EditingDoc::new(19);
            let before = peer.encode_state_as_update_v1();
            let version = peer.version();
            let error = peer
                .prepare_peer_bootstrap(&state, &bad, Some(source.clone()))
                .err()
                .unwrap();
            assert_eq!(error.code(), code);
            assert_eq!(peer.encode_state_as_update_v1(), before);
            assert_eq!(peer.version(), version);
            assert!(peer.source_metadata().is_none());
            assert!(peer.source.lock().unwrap().is_none());
            assert!(peer.media.lock().unwrap().is_none());
            let prepared = peer
                .prepare_peer_bootstrap(&state, &metadata, Some(source.clone()))
                .unwrap();
            peer.install_peer_bootstrap(prepared, 17).unwrap();
            peer.apply_update_v1(&state).unwrap();
            assert_eq!(peer.paragraph_identities(), worker.paragraph_identities());
        }
    }

    fn table_package(horizontal: bool, vertical: bool) -> PackageBytes {
        let cell = |properties: &str, content: &str| {
            format!("<w:tc><w:tcPr>{properties}</w:tcPr>{content}</w:tc>")
        };
        let paragraph = |text: &str| format!("<w:p><w:r><w:t>{text}</w:t></w:r></w:p>");
        let nested = concat!(
            "<w:tbl><w:tblGrid><w:gridCol w:w=\"900\"/></w:tblGrid><w:tr><w:tc>",
            "<w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        );
        let mut rows = String::new();
        for row in 0..3 {
            let span = if horizontal {
                "<w:gridSpan w:val=\"2\"/>"
            } else {
                ""
            };
            let merge = if vertical {
                if row == 0 {
                    "<w:vMerge w:val=\"restart\"/>"
                } else {
                    "<w:vMerge/>"
                }
            } else {
                ""
            };
            let properties = format!("{span}{merge}");
            let content = if vertical && row > 0 {
                "<w:p/>".to_owned()
            } else if row == 0 {
                format!("{}{nested}", paragraph("Left"))
            } else {
                paragraph("Left")
            };
            rows.push_str(&format!(
                "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/><w:gridAfter w:val=\"1\"/></w:trPr>{}{}</w:tr>",
                cell(&properties, &content), cell("", &paragraph("Right")),
            ));
        }
        let mut parts = fixture::principal_parts();
        let body = parts
            .iter_mut()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap();
        body.1 = format!(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:tbl><w:tblGrid><w:gridCol w:w=\"900\"/><w:gridCol w:w=\"900\"/><w:gridCol w:w=\"900\"/><w:gridCol w:w=\"900\"/><w:gridCol w:w=\"900\"/></w:tblGrid>{rows}</w:tbl><w:sectPr/></w:body></w:document>"
        ).into_bytes();
        ooxml_opc::rezip_parts(&parts).unwrap().into()
    }

    fn assert_invalid_metadata_is_atomic(
        worker: &EditingDoc,
        source: &PackageBytes,
        metadata: &[u8],
        bad: &Value,
        expected: &str,
    ) {
        let (_, blobs) = sections(metadata).unwrap();
        let bad = frame(&serde_json::to_vec(bad).unwrap(), blobs).unwrap();
        let state = worker.encode_state_as_update_v1();
        let peer = EditingDoc::new(19);
        let before = peer.encode_state_as_update_v1();
        let version = peer.version();
        let error = peer
            .prepare_peer_bootstrap(&state, &bad, Some(source.clone()))
            .err()
            .unwrap();
        assert_eq!(
            error,
            PeerMetadataError::InvalidMetadata(expected.to_owned())
        );
        assert_eq!(peer.encode_state_as_update_v1(), before);
        assert_eq!(peer.version(), version);
        assert!(peer.source_metadata().is_none());
        assert!(peer.source.lock().unwrap().is_none());
        assert!(peer.media.lock().unwrap().is_none());
        assert!(peer.media_sources().is_empty());
        let prepared = peer
            .prepare_peer_bootstrap(&state, metadata, Some(source.clone()))
            .unwrap();
        peer.install_peer_bootstrap(prepared, 17).unwrap();
        peer.apply_update_v1(&state).unwrap();
        assert_eq!(peer.paragraph_identities(), worker.paragraph_identities());
    }

    fn continuation_table_package() -> PackageBytes {
        let cell = |properties: &str, text: &str| {
            format!(
                "<w:tc><w:tcPr>{properties}</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"
            )
        };
        let table = |rows: &str| {
            format!(
                "<w:tbl><w:tblGrid><w:gridCol w:w=\"900\"/><w:gridCol w:w=\"900\"/><w:gridCol w:w=\"900\"/></w:tblGrid>{rows}</w:tbl>"
            )
        };
        let mut rows = String::new();
        for (properties, text) in [
            ("<w:vMerge w:val=\"continue\"/>", "Orphan"),
            ("<w:vMerge w:val=\"restart\"/>", "Restart"),
            (
                "<w:gridSpan w:val=\"2\"/><w:vMerge w:val=\"continue\"/>",
                "Mismatched",
            ),
        ] {
            rows.push_str(&format!(
                "<w:tr>{}{}</w:tr>",
                cell(properties, text),
                cell("", "Right"),
            ));
        }
        let mut parts = fixture::principal_parts();
        let body = parts
            .iter_mut()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap();
        body.1 = format!(
            "<w:document {}><w:body>{}<w:sectPr/></w:body></w:document>",
            fixture::NS,
            table(&rows),
        )
        .into_bytes();
        let comments = parts
            .iter_mut()
            .find(|(name, _)| name == "word/comments.xml")
            .unwrap();
        comments.1 = format!(
            "<w:comments {}><w:comment w:id=\"1\">{}</w:comment></w:comments>",
            fixture::NS,
            table(&format!(
                "<w:tr>{}</w:tr>",
                cell("<w:vMerge w:val=\"continue\"/>", "Comment orphan"),
            )),
        )
        .into_bytes();
        ooxml_opc::rezip_parts(&parts).unwrap().into()
    }

    #[test]
    fn peer_metadata_table_row_count_fits_export_positions_and_spans() {
        use crate::structured::source::TableLayout;

        assert!(TableLayout::validate_peer_row_count(u32::MAX as usize).is_ok());
        if let Ok(rows) = usize::try_from(u64::from(u32::MAX) + 1) {
            assert_eq!(
                TableLayout::validate_peer_row_count(rows),
                Err("too many table rows".to_owned())
            );
        }
    }

    #[test]
    fn peer_bootstrap_rejects_invalid_table_layouts_atomically() {
        let source = table_package(false, false);
        let worker = worker(&source, &seed::package_digest(&source), 7).unwrap();
        let metadata = worker.encode_peer_metadata().unwrap();
        let (json, _) = sections(&metadata).unwrap();
        let wire: Value = serde_json::from_slice(json).unwrap();
        let tables = wire["source"]["read"]["provenance"]["tables"]
            .as_array()
            .unwrap();
        let table_index = tables
            .iter()
            .position(|entry| entry[0] == "body:t0")
            .unwrap();
        for (case, expected) in [
            ("same-row-origin", "invalid table cell column"),
            ("zero-span", "invalid table cell span"),
            ("column-overflow", "table cell column overflow"),
            ("trailing-overflow", "table row column overflow"),
            ("overlap", "invalid table cell column"),
            ("unordered", "invalid table cell column"),
            ("gap", "invalid table cell column"),
            ("leading-gap", "invalid table cell column"),
            ("grid-extent", "invalid table grid extent"),
            ("empty-row-extent", "invalid table grid extent"),
        ] {
            let mut bad = wire.clone();
            let layout = &mut bad["source"]["read"]["provenance"]["tables"][table_index][1];
            match case {
                "same-row-origin" => {
                    layout["rows"][0]["cells"][0]["merge"] = json!("Restart");
                    layout["rows"][0]["cells"][1]["merge"] = json!("Continue");
                    layout["rows"][0]["cells"][1]["column"] =
                        layout["rows"][0]["cells"][0]["column"].clone();
                    layout["rows"][0]["cells"][1]["span"] =
                        layout["rows"][0]["cells"][0]["span"].clone();
                }
                "zero-span" => layout["rows"][0]["cells"][0]["span"] = json!(0),
                "column-overflow" => {
                    layout["rows"][0]["grid_before"] = json!(u32::MAX);
                    layout["rows"][0]["cells"][0]["column"] = json!(u32::MAX);
                }
                "trailing-overflow" => layout["rows"][0]["grid_after"] = json!(u32::MAX),
                "overlap" => layout["rows"][0]["cells"][0]["span"] = json!(2),
                "unordered" => layout["rows"][0]["cells"]
                    .as_array_mut()
                    .unwrap()
                    .swap(0, 1),
                "gap" => layout["rows"][0]["cells"][1]["column"] = json!(3),
                "leading-gap" => layout["rows"][0]["grid_before"] = json!(0),
                "grid-extent" => layout["grid_columns"] = json!(3),
                "empty-row-extent" => {
                    layout["rows"][0]["cells"] = json!([]);
                    layout["rows"][0]["grid_before"] = json!(6);
                }
                _ => unreachable!(),
            }
            assert_invalid_metadata_is_atomic(&worker, &source, &metadata, &bad, expected);
        }
    }

    #[test]
    fn peer_bootstrap_unmatched_continuation_metadata_decodes_and_exports() {
        let source = table_package(false, false);
        let digest = seed::package_digest(&source);
        let worker = worker(&source, &digest, 7).unwrap();
        let metadata = worker.encode_peer_metadata().unwrap();
        let state = worker.encode_state_as_update_v1();
        let (json, blobs) = sections(&metadata).unwrap();
        let wire: Value = serde_json::from_slice(json).unwrap();
        let tables = wire["source"]["read"]["provenance"]["tables"]
            .as_array()
            .unwrap();
        let table_index = tables
            .iter()
            .position(|entry| entry[0] == "body:t0")
            .unwrap();
        for (case, row_index) in [
            ("first-row-continuation", 0),
            ("no-restart", 1),
            ("different-span", 1),
            ("different-column", 1),
            ("interrupted-merge", 2),
            ("orphan-continuation", 0),
        ] {
            let mut accepted = wire.clone();
            if case == "orphan-continuation" {
                accepted["source"]["read"]["comments"][0]["body"] = json!([{
                    "type": "table", "columnWidths": [900, 900],
                    "rows": [{ "cells": [
                        { "formatting": { "vMerge": "continue" }, "content": [] },
                        { "content": [] }
                    ] }]
                }]);
            } else {
                let layout =
                    &mut accepted["source"]["read"]["provenance"]["tables"][table_index][1];
                match case {
                    "first-row-continuation" => {
                        layout["rows"][0]["cells"][0]["merge"] = json!("Continue")
                    }
                    "no-restart" => layout["rows"][1]["cells"][0]["merge"] = json!("Continue"),
                    "different-span" => {
                        layout["rows"][0]["cells"][0]["merge"] = json!("Restart");
                        layout["rows"][1]["cells"][0]["merge"] = json!("Continue");
                        layout["rows"][1]["cells"][0]["span"] = json!(2);
                        layout["rows"][1]["cells"][1]["column"] = json!(3);
                    }
                    "different-column" => {
                        layout["rows"][0]["cells"][0]["merge"] = json!("Restart");
                        layout["rows"][1]["cells"][0]["merge"] = json!("Continue");
                        layout["rows"][1]["grid_before"] = json!(0);
                        layout["rows"][1]["cells"][0]["column"] = json!(0);
                        layout["rows"][1]["cells"][1]["column"] = json!(1);
                    }
                    "interrupted-merge" => {
                        layout["rows"][0]["cells"][0]["merge"] = json!("Restart");
                        layout["rows"][2]["cells"][0]["merge"] = json!("Continue");
                    }
                    _ => unreachable!(),
                }
            }
            let accepted = frame(&serde_json::to_vec(&accepted).unwrap(), blobs).unwrap();
            assert!(decode(&accepted, source.clone(), &digest).is_ok(), "{case}");
            let peer = EditingDoc::new(19);
            let prepared = peer
                .prepare_peer_bootstrap(&state, &accepted, Some(source.clone()))
                .unwrap_or_else(|error| panic!("{case}: {error}"));
            peer.install_peer_bootstrap(prepared, 17).unwrap();
            peer.apply_update_v1(&state).unwrap();
            for view in [
                crate::structured::RevisionView::Accepted,
                crate::structured::RevisionView::Original,
                crate::structured::RevisionView::Markup,
            ] {
                let mut options = crate::structured::ExportOptions::new(view);
                if case == "orphan-continuation" {
                    options.stories = Some(vec![crate::structured::StorySelection::Comments]);
                }
                let actual = peer
                    .export_structured(&options)
                    .unwrap_or_else(|error| panic!("{case}: {error:?}"))
                    .content;
                let crate::structured::BlockKind::Table { table } =
                    &actual.stories[0].blocks[0].content
                else {
                    panic!("{case}: expected table");
                };
                let cell = &table.rows[row_index].cells[0];
                assert_eq!(cell.row_span, 1, "{case}");
                assert_eq!(
                    cell.vertical_merge,
                    crate::structured::VerticalMerge::None,
                    "{case}"
                );
                assert_eq!(cell.merge_origin, None, "{case}");
                assert!(
                    actual.diagnostics.iter().any(|diagnostic| {
                        diagnostic.code == crate::structured::DiagnosticCode::UnsupportedContent
                    }),
                    "{case}"
                );
            }
        }
    }

    #[test]
    fn peer_bootstrap_merged_and_nested_tables_export_like_source_open() {
        for (horizontal, vertical) in [(false, false), (true, false), (false, true), (true, true)] {
            let source = table_package(horizontal, vertical);
            let digest = seed::package_digest(&source);
            let worker = worker(&source, &digest, 7).unwrap();
            let state = worker.encode_state_as_update_v1();
            let metadata = worker.encode_peer_metadata().unwrap();
            let (envelope, parts, media) =
                seed::parse_docx_package_with_media(source.clone(), digest.clone()).unwrap();
            let (mut expected_source, index, _) =
                seed::replica_source(envelope, parts, source.clone(), digest).unwrap();
            let baseline = EditingDoc::new(19);
            expected_source.watch_comments(&baseline);
            baseline.install_source(expected_source, 17);
            baseline.retain_source(SourcePackage::Ready(Arc::new(index)));
            baseline.install_media(media);
            baseline.apply_update_v1(&state).unwrap();
            let peer = EditingDoc::new(19);
            let prepared = peer
                .prepare_peer_bootstrap(&state, &metadata, Some(source))
                .unwrap();
            peer.install_peer_bootstrap(prepared, 17).unwrap();
            peer.apply_update_v1(&state).unwrap();
            let read = peer.source_metadata().unwrap();
            let layout = &read.read().provenance.tables["body:t0"];
            assert_eq!(layout.rows.len(), 3);
            assert_eq!(layout.rows[0].cells[0].span, if horizontal { 2 } else { 1 });
            assert!(read.read().provenance.tables.len() >= 2);
            for view in [
                crate::structured::RevisionView::Accepted,
                crate::structured::RevisionView::Original,
                crate::structured::RevisionView::Markup,
            ] {
                let options = crate::structured::ExportOptions::new(view);
                let actual = peer.export_structured(&options).unwrap().content;
                assert_eq!(
                    actual,
                    baseline.export_structured(&options).unwrap().content
                );
                let crate::structured::BlockKind::Table { table } =
                    &actual.stories[0].blocks[0].content
                else {
                    panic!("expected table");
                };
                assert_eq!(
                    table.rows[0].cells[0].row_span,
                    if vertical { 3 } else { 1 }
                );
                assert_eq!(
                    table.rows[0].cells[0].grid_span,
                    if horizontal { 2 } else { 1 }
                );
                if vertical {
                    assert_eq!(table.rows[1].cells[0].row_span, 0);
                    assert_eq!(
                        table.rows[2].cells[0].merge_origin,
                        Some(crate::structured::CellPosition { row: 0, column: 1 })
                    );
                }
            }
        }
    }

    #[test]
    fn peer_bootstrap_source_unmatched_continuations_export_like_source_open() {
        let source = continuation_table_package();
        let digest = seed::package_digest(&source);
        let worker = worker(&source, &digest, 7).unwrap();
        let read = worker.source_metadata().unwrap();
        let layout = &read.read().provenance.tables["body:t0"];
        assert_eq!(layout.rows.len(), 3);
        assert!(matches!(
            layout.rows[0].cells[0].merge,
            crate::structured::source::SourceMerge::Continue
        ));
        assert!(matches!(
            layout.rows[1].cells[0].merge,
            crate::structured::source::SourceMerge::Restart
        ));
        assert!(matches!(
            layout.rows[2].cells[0].merge,
            crate::structured::source::SourceMerge::Continue
        ));
        assert_eq!(layout.rows[1].cells[0].span, 1);
        assert_eq!(layout.rows[2].cells[0].span, 2);
        assert!(!read.read().seeded_comments.contains("1"));
        let comment = read
            .read()
            .comments
            .iter()
            .find(|comment| comment.id == "1")
            .unwrap();
        assert_eq!(
            comment.body[0]["rows"][0]["cells"][0]["formatting"]["vMerge"],
            json!("continue")
        );
        let state = worker.encode_state_as_update_v1();
        let metadata = worker.encode_peer_metadata().unwrap();
        let (envelope, parts, media) =
            seed::parse_docx_package_with_media(source.clone(), digest.clone()).unwrap();
        let (mut expected_source, index, _) =
            seed::replica_source(envelope, parts, source.clone(), digest).unwrap();
        let baseline = EditingDoc::new(19);
        expected_source.watch_comments(&baseline);
        baseline.install_source(expected_source, 17);
        baseline.retain_source(SourcePackage::Ready(Arc::new(index)));
        baseline.install_media(media);
        baseline.apply_update_v1(&state).unwrap();
        let peer = EditingDoc::new(19);
        let prepared = peer
            .prepare_peer_bootstrap(&state, &metadata, Some(source))
            .unwrap();
        peer.install_peer_bootstrap(prepared, 17).unwrap();
        peer.apply_update_v1(&state).unwrap();
        for view in [
            crate::structured::RevisionView::Accepted,
            crate::structured::RevisionView::Original,
            crate::structured::RevisionView::Markup,
        ] {
            for comments in [false, true] {
                let mut options = crate::structured::ExportOptions::new(view);
                if comments {
                    options.stories = Some(vec![crate::structured::StorySelection::Comments]);
                }
                let actual = peer.export_structured(&options).unwrap().content;
                assert_eq!(
                    actual,
                    baseline.export_structured(&options).unwrap().content
                );
                let crate::structured::BlockKind::Table { table } =
                    &actual.stories[0].blocks[0].content
                else {
                    panic!("expected continuation table");
                };
                assert_eq!(table.rows.len(), if comments { 1 } else { 3 });
                let unmatched_rows: &[usize] = if comments { &[0] } else { &[0, 2] };
                for &row_index in unmatched_rows {
                    let cell = &table.rows[row_index].cells[0];
                    assert_eq!(cell.row_span, 1);
                    assert_eq!(cell.vertical_merge, crate::structured::VerticalMerge::None);
                    assert_eq!(cell.merge_origin, None);
                }
                if !comments {
                    assert_eq!(table.rows[2].cells[0].grid_span, 2);
                    assert_eq!(table.rows[1].cells[0].row_span, 1);
                    assert_eq!(
                        table.rows[1].cells[0].vertical_merge,
                        crate::structured::VerticalMerge::Restart
                    );
                }
                assert_eq!(
                    actual
                        .diagnostics
                        .iter()
                        .filter(|diagnostic| {
                            diagnostic.code == crate::structured::DiagnosticCode::UnsupportedContent
                                && diagnostic.message.contains("continues a vertical merge")
                        })
                        .count(),
                    if comments { 1 } else { 2 },
                );
            }
        }
    }

    #[test]
    fn peer_bootstrap_rejects_invalid_control_metadata_atomically() {
        let source = synthetic_package();
        let worker = worker(&source, &seed::package_digest(&source), 7).unwrap();
        let metadata = worker.encode_peer_metadata().unwrap();
        let (json, _) = sections(&metadata).unwrap();
        let wire: Value = serde_json::from_slice(json).unwrap();
        let anchor = serde_json::to_value(crate::structured::Anchor::Control {
            story: "body".to_owned(),
            control_id: "source-control".to_owned(),
        })
        .unwrap();
        for (case, expected) in [
            ("anchor-count", "invalid unrepresented control count"),
            ("raw-count-overflow", "unlocated control count overflow"),
            ("comment-count-overflow", "unlocated control count overflow"),
            (
                "parent-story",
                "source control parent belongs to another container",
            ),
            (
                "parent-block",
                "source control parent belongs to another container",
            ),
        ] {
            let mut bad = wire.clone();
            let read = &mut bad["source"]["read"];
            match case {
                "anchor-count" => {
                    read["unrepresented_controls"] = json!(1);
                    read["unrepresented_anchors"] = json!([anchor.clone(), anchor.clone()]);
                }
                "raw-count-overflow" => {
                    read["control_safety"] = Value::Null;
                    read["unlocated_controls"] = json!(usize::MAX);
                    read["provenance"]["raw_sources"] = json!([{
                        "story": "body", "index": 0, "steps": [], "xml": "<w:sdtContent/>"
                    }]);
                }
                "comment-count-overflow" => {
                    read["control_safety"] = Value::Null;
                    read["unlocated_controls"] = json!(usize::MAX);
                    read["provenance"]["raw_sources"] = json!([]);
                    read["comments"] = json!([{
                        "id": "1", "author": null, "date": null, "parent_id": null,
                        "done": false, "body": [{ "type": "blockSdt" }]
                    }]);
                }
                "parent-story" | "parent-block" => {
                    let parent = json!({
                        "story": "body", "raw_block": 0, "anchor": anchor.clone(),
                        "parent": null, "block": true,
                        "properties": { "sdtType": "plainText", "lock": "contentLocked" }
                    });
                    let mut child = parent.clone();
                    child["parent"] = json!(0);
                    if case == "parent-story" {
                        child["story"] = json!("hf:rId1");
                    } else {
                        child["raw_block"] = json!(1);
                    }
                    read["source_controls"] = json!([parent, child]);
                }
                _ => unreachable!(),
            }
            assert_invalid_metadata_is_atomic(&worker, &source, &metadata, &bad, expected);
        }
    }

    #[test]
    fn peer_bootstrap_normalizes_retained_table_counts() {
        let source = synthetic_package();
        let worker = worker(&source, &seed::package_digest(&source), 7).unwrap();
        let metadata = worker.encode_peer_metadata().unwrap();
        let (json, blobs) = sections(&metadata).unwrap();
        let wire: Value = serde_json::from_slice(json).unwrap();
        let state = worker.encode_state_as_update_v1();
        for case in [
            "zero-span",
            "negative-span",
            "fractional-span",
            "oversized-span",
            "width-overflow",
            "omission-overflow",
            "leading-count",
            "trailing-count",
            "nested-table",
            "separator-table",
        ] {
            let mut table = json!({
                "type": "table", "columnWidths": [900, 900],
                "rows": [{ "cells": [{ "content": [] }, { "content": [] }] }]
            });
            match case {
                "zero-span" => table["rows"][0]["cells"][0]["formatting"] = json!({"gridSpan": 0}),
                "negative-span" => {
                    table["rows"][0]["cells"][0]["formatting"] = json!({"gridSpan": -1})
                }
                "fractional-span" => {
                    table["rows"][0]["cells"][0]["formatting"] = json!({"gridSpan": 1.5})
                }
                "oversized-span" | "nested-table" | "separator-table" => {
                    table["rows"][0]["cells"][0]["formatting"] = json!({"gridSpan": 1e300});
                }
                "width-overflow" => {
                    table["rows"][0]["cells"][0]["formatting"] = json!({"gridSpan": u32::MAX});
                }
                "omission-overflow" => {
                    table["rows"][0]["cells"] = json!([]);
                    table["rows"][0]["formatting"] =
                        json!({"gridBefore": u32::MAX, "gridAfter": 1});
                }
                "leading-count" => table["rows"][0]["formatting"] = json!({"gridBefore": 1e300}),
                "trailing-count" => table["rows"][0]["formatting"] = json!({"gridAfter": 1e300}),
                _ => unreachable!(),
            }
            let mut bad = wire.clone();
            if case == "separator-table" {
                bad["source"]["read"]["note_separator_paragraphs"] = json!([["footnote", [table]]]);
            } else {
                let body = if case == "nested-table" {
                    json!([{ "type": "blockSdt", "content": [table] }])
                } else {
                    json!([table])
                };
                bad["source"]["read"]["comments"] = json!([{
                    "id": "peer-table", "author": null, "date": null, "parent_id": null,
                    "done": false, "body": body
                }]);
            }
            if matches!(case, "oversized-span" | "nested-table" | "separator-table")
                || (case == "width-overflow" && usize::BITS == 32)
            {
                assert_invalid_metadata_is_atomic(
                    &worker,
                    &source,
                    &metadata,
                    &bad,
                    "source table column overflow",
                );
                continue;
            }
            let accepted = frame(&serde_json::to_vec(&bad).unwrap(), blobs).unwrap();
            let peer = EditingDoc::new(19);
            let prepared = peer
                .prepare_peer_bootstrap(&state, &accepted, Some(source.clone()))
                .unwrap_or_else(|error| panic!("{case}: {error}"));
            peer.install_peer_bootstrap(prepared, 17).unwrap();
            peer.apply_update_v1(&state).unwrap();
            assert!(peer.source_metadata().is_some(), "{case}");
        }
    }

    #[test]
    fn peer_bootstrap_source_comment_tables_export_like_source_open() {
        let mut parts = fixture::principal_parts();
        let comments = parts
            .iter_mut()
            .find(|(name, _)| name == "word/comments.xml")
            .unwrap();
        comments.1 = format!(
            "<w:comments {}><w:comment w:id=\"1\"><w:tbl><w:tblGrid><w:gridCol w:w=\"900\"/><w:gridCol w:w=\"900\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val=\"2\"/><w:vMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>Comment cell</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:gridSpan w:val=\"2\"/><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl></w:comment></w:comments>",
            fixture::NS
        ).into_bytes();
        let source: PackageBytes = ooxml_opc::rezip_parts(&parts).unwrap().into();
        let digest = seed::package_digest(&source);
        let worker = worker(&source, &digest, 7).unwrap();
        let metadata = worker.encode_peer_metadata().unwrap();
        let state = worker.encode_state_as_update_v1();
        let (envelope, parts, media) =
            seed::parse_docx_package_with_media(source.clone(), digest.clone()).unwrap();
        let (mut expected_source, index, _) =
            seed::replica_source(envelope, parts, source.clone(), digest).unwrap();
        let baseline = EditingDoc::new(19);
        expected_source.watch_comments(&baseline);
        baseline.install_source(expected_source, 17);
        baseline.retain_source(SourcePackage::Ready(Arc::new(index)));
        baseline.install_media(media);
        baseline.apply_update_v1(&state).unwrap();
        let peer = EditingDoc::new(19);
        let prepared = peer
            .prepare_peer_bootstrap(&state, &metadata, Some(source))
            .unwrap();
        peer.install_peer_bootstrap(prepared, 17).unwrap();
        peer.apply_update_v1(&state).unwrap();
        let mut options =
            crate::structured::ExportOptions::new(crate::structured::RevisionView::Accepted);
        options.stories = Some(vec![crate::structured::StorySelection::Comments]);
        let actual = peer.export_structured(&options).unwrap().content;
        assert_eq!(
            actual,
            baseline.export_structured(&options).unwrap().content
        );
        let crate::structured::BlockKind::Table { table } = &actual.stories[0].blocks[0].content
        else {
            panic!("expected comment table");
        };
        assert_eq!(table.rows[0].cells[0].grid_span, 2);
        assert_eq!(table.rows[0].cells[0].row_span, 2);
    }

    #[test]
    fn peer_bootstrap_rejects_inconsistent_identity_views_atomically() {
        let source = table_package(false, false);
        let worker = worker(&source, &seed::package_digest(&source), 7).unwrap();
        let metadata = worker.encode_peer_metadata().unwrap();
        let (json, _) = sections(&metadata).unwrap();
        let wire: Value = serde_json::from_slice(json).unwrap();
        let parts = wire["index"]["parts"].as_array().unwrap();
        let part_index = parts
            .iter()
            .position(|part| part["uri"] == "/word/document.xml")
            .unwrap();
        let views = parts[part_index]["backed"].as_array().unwrap();
        assert!(views.len() >= 2);
        let key = views[0][1][0].as_str().unwrap();
        let seeded = wire["index"]["seeded"].as_array().unwrap();
        let seed_index = seeded.iter().position(|entry| entry[0] == key).unwrap();
        for (case, expected) in [
            ("empty-views", "invalid source paragraph views"),
            ("duplicate-view", "inconsistent source paragraph view"),
            ("duplicate-root", "inconsistent source paragraph view"),
            ("different-part", "inconsistent source paragraph view"),
            ("different-ordinal", "inconsistent source paragraph view"),
            (
                "missing-reciprocal-view",
                "seeded paragraph missing its source view",
            ),
        ] {
            let mut bad = wire.clone();
            match case {
                "empty-views" => bad["index"]["parts"][part_index]["backed"][0][1] = json!([]),
                "duplicate-view" => {
                    bad["index"]["parts"][part_index]["backed"][0][1] = json!([key, key]);
                }
                "duplicate-root" => {
                    let mut duplicate = seeded[seed_index].clone();
                    duplicate[0] = json!("duplicate-view-key");
                    bad["index"]["seeded"]
                        .as_array_mut()
                        .unwrap()
                        .push(duplicate);
                    bad["index"]["parts"][part_index]["backed"][0][1] =
                        json!([key, "duplicate-view-key"]);
                }
                "different-part" => {
                    let other = (part_index + 1) % parts.len();
                    assert_ne!(other, part_index);
                    bad["index"]["seeded"][seed_index][1]["part"] = json!(other);
                }
                "different-ordinal" => {
                    bad["index"]["seeded"][seed_index][1]["ordinal"] = views[1][0].clone();
                }
                "missing-reciprocal-view" => {
                    bad["index"]["parts"][part_index]["backed"]
                        .as_array_mut()
                        .unwrap()
                        .remove(0);
                }
                _ => unreachable!(),
            }
            assert_invalid_metadata_is_atomic(&worker, &source, &metadata, &bad, expected);
        }
    }

    #[test]
    fn peer_bootstrap_requires_matching_source_and_decodable_state() {
        let source = synthetic_package();
        let digest = seed::package_digest(&source);
        let worker = worker(&source, &digest, 7).unwrap();
        let metadata = worker.encode_peer_metadata().unwrap();
        let state = worker.encode_state_as_update_v1();
        let peer = EditingDoc::new(19);
        let version = peer.version();
        assert_eq!(
            peer.prepare_peer_bootstrap(&state, &metadata, None)
                .err()
                .unwrap(),
            PeerMetadataError::SourceRequired
        );
        let mut wrong = source.to_vec();
        wrong[0] ^= 1;
        assert_eq!(
            peer.prepare_peer_bootstrap(&state, &metadata, Some(wrong.into()))
                .err()
                .unwrap(),
            PeerMetadataError::SourceMismatch
        );
        assert!(matches!(
            peer.prepare_peer_bootstrap(&[], &metadata, Some(source.clone()))
                .err()
                .unwrap(),
            PeerMetadataError::InvalidState(_)
        ));
        assert_eq!(peer.version(), version);
        assert!(peer.source_metadata().is_none());
        let prepared = peer
            .prepare_peer_bootstrap(&state, &metadata, Some(source.clone()))
            .unwrap();
        let retained = peer.install_peer_bootstrap(prepared, 17).unwrap();
        let prepared = peer
            .prepare_peer_bootstrap(&state, &metadata, None)
            .unwrap();
        let reused = peer.install_peer_bootstrap(prepared, 18).unwrap();
        assert_eq!(retained.source.as_ptr(), reused.source.as_ptr());
        peer.apply_update_v1(&state).unwrap();
        assert_eq!(
            peer.prepare_peer_bootstrap(&state, &metadata, None)
                .err()
                .unwrap(),
            PeerMetadataError::NonEmptyDocument
        );
        let pending_source = EditingDoc::new(23);
        pending_source.retain_source_docx(Arc::<[u8]>::from(source.as_ref()));
        let prepared = pending_source
            .prepare_peer_bootstrap(&state, &metadata, None)
            .unwrap();
        let retained = pending_source.install_peer_bootstrap(prepared, 17).unwrap();
        assert_eq!(retained.source.as_ref(), source.as_ref());
        assert_eq!(
            retained.source.as_ptr(),
            pending_source.source_index().unwrap().bytes().as_ptr()
        );
    }

    #[test]
    fn peer_bootstrap_rechecks_empty_document_at_installation() {
        let source = synthetic_package();
        let digest = seed::package_digest(&source);
        let worker = worker(&source, &digest, 7).unwrap();
        let metadata = worker.encode_peer_metadata().unwrap();
        let state = worker.encode_state_as_update_v1();
        for deleted in [false, true] {
            let peer = EditingDoc::new(19);
            let prepared = peer
                .prepare_peer_bootstrap(&state, &metadata, Some(source.clone()))
                .unwrap();
            let map = peer.yrs_doc().get_or_insert_map("other-root");
            map.insert(&mut peer.yrs_doc().transact_mut(), "value", true);
            if deleted {
                map.remove(&mut peer.yrs_doc().transact_mut(), "value");
            }
            let before = peer.encode_state_as_update_v1();
            let version = peer.version();
            assert_eq!(
                peer.install_peer_bootstrap(prepared, 17).err().unwrap(),
                PeerMetadataError::NonEmptyDocument
            );
            assert_eq!(
                peer.prepare_peer_bootstrap(&state, &metadata, Some(source.clone()))
                    .err()
                    .unwrap(),
                PeerMetadataError::NonEmptyDocument
            );
            assert_eq!(peer.encode_state_as_update_v1(), before);
            assert_eq!(peer.version(), version);
            assert!(peer.source_metadata().is_none());
            assert!(peer.media.lock().unwrap().is_none());
        }
        let map = worker.yrs_doc().get_or_insert_map("other-root").insert(
            &mut worker.yrs_doc().transact_mut(),
            "parent",
            MapPrelim::default(),
        );
        let vector = worker.encode_state_vector_v1();
        map.insert(&mut worker.yrs_doc().transact_mut(), "value", true);
        let pending = EditingDoc::new(19);
        pending
            .apply_update_v1(&worker.encode_diff_v1(&vector).unwrap())
            .unwrap();
        assert!(pending.yrs_doc().transact().state_vector().is_empty());
        assert!(pending.yrs_doc().transact().has_missing_updates());
        let before = pending.encode_state_as_update_v1();
        let version = pending.version();
        assert_eq!(
            pending
                .prepare_peer_bootstrap(&state, &metadata, Some(source.clone()))
                .err()
                .unwrap(),
            PeerMetadataError::NonEmptyDocument
        );
        assert_eq!(pending.encode_state_as_update_v1(), before);
        assert_eq!(pending.version(), version);
        let vector = worker.encode_state_vector_v1();
        map.remove(&mut worker.yrs_doc().transact_mut(), "value");
        let pending_delete = EditingDoc::new(19);
        pending_delete
            .apply_update_v1(&worker.encode_diff_v1(&vector).unwrap())
            .unwrap();
        assert!(
            pending_delete
                .yrs_doc()
                .transact()
                .state_vector()
                .is_empty()
        );
        assert!(pending_delete.yrs_doc().transact().has_missing_updates());
        let before = pending_delete.encode_state_as_update_v1();
        let version = pending_delete.version();
        assert_eq!(
            pending_delete
                .prepare_peer_bootstrap(&state, &metadata, Some(source))
                .err()
                .unwrap(),
            PeerMetadataError::NonEmptyDocument
        );
        assert_eq!(pending_delete.encode_state_as_update_v1(), before);
        assert_eq!(pending_delete.version(), version);
    }

    #[test]
    fn peer_bootstrap_rejects_decodable_state_with_invalid_parent_atomically() {
        let source = synthetic_package();
        let digest = seed::package_digest(&source);
        let worker = worker(&source, &digest, 7).unwrap();
        let metadata = worker.encode_peer_metadata().unwrap();
        let mut encoder = EncoderV1::new();
        encoder.write_var(1u32);
        encoder.write_var(2u32);
        encoder.write_client(yrs::ClientID::new(7));
        encoder.write_var(0u32);
        encoder.write_info(HAS_PARENT_SUB | BLOCK_ITEM_ANY_REF_NUMBER);
        encoder.write_parent_info(true);
        encoder.write_string("other-root");
        encoder.write_string("parent");
        encoder.write_len(1);
        encoder.write_any(&Any::Bool(true));
        encoder.write_info(HAS_PARENT_SUB | BLOCK_ITEM_ANY_REF_NUMBER);
        encoder.write_parent_info(false);
        encoder.write_left_id(&yrs::ID::new(yrs::ClientID::new(7), 0));
        encoder.write_string("child");
        encoder.write_len(1);
        encoder.write_any(&Any::Bool(true));
        encoder.write_var(0u32);
        let invalid = encoder.to_vec();
        assert!(Update::decode_v1(&invalid).is_ok());
        let peer = EditingDoc::new(19);
        let before = peer.encode_state_as_update_v1();
        let version = peer.version();
        assert!(matches!(
            peer.prepare_peer_bootstrap(&invalid, &metadata, Some(source.clone()))
                .err()
                .unwrap(),
            PeerMetadataError::InvalidState(_)
        ));
        assert_eq!(peer.encode_state_as_update_v1(), before);
        assert_eq!(peer.version(), version);
        assert!(peer.source_metadata().is_none());
        assert!(peer.source.lock().unwrap().is_none());
        assert!(peer.media.lock().unwrap().is_none());
        let state = worker.encode_state_as_update_v1();
        let prepared = peer
            .prepare_peer_bootstrap(&state, &metadata, Some(source))
            .unwrap();
        peer.install_peer_bootstrap(prepared, 17).unwrap();
        peer.apply_update_v1(&state).unwrap();
        assert_eq!(peer.paragraph_identities(), worker.paragraph_identities());
    }

    #[test]
    fn peer_bootstrap_invalid_package_is_atomic() {
        let source = synthetic_package();
        let digest = seed::package_digest(&source);
        let worker = worker(&source, &digest, 7).unwrap();
        let state = worker.encode_state_as_update_v1();
        let metadata = worker.encode_peer_metadata().unwrap();
        let (json, blobs) = sections(&metadata).unwrap();
        let mut wire: Value = serde_json::from_slice(json).unwrap();
        let invalid: PackageBytes = vec![0; source.len()].into();
        let invalid_digest = seed::package_digest(&invalid);
        wire["source_digest"] = json!(invalid_digest);
        wire["index"]["package_sha256"] = wire["source_digest"].clone();
        let bad = frame(&serde_json::to_vec(&wire).unwrap(), blobs).unwrap();
        let peer = EditingDoc::new(19);
        let before = peer.encode_state_as_update_v1();
        let version = peer.version();
        assert!(matches!(
            peer.prepare_peer_bootstrap(&state, &bad, Some(invalid))
                .err()
                .unwrap(),
            PeerMetadataError::Package(_)
        ));
        assert_eq!(peer.encode_state_as_update_v1(), before);
        assert_eq!(peer.version(), version);
        assert!(peer.source_metadata().is_none());
        let prepared = peer
            .prepare_peer_bootstrap(&state, &metadata, Some(source))
            .unwrap();
        peer.install_peer_bootstrap(prepared, 17).unwrap();
    }

    #[test]
    fn peer_bootstrap_full_corpus_installation_matches_replica_source() {
        let root = Path::new(env!("CARGO_MANIFEST_DIR"))
            .ancestors()
            .nth(2)
            .unwrap();
        let mut paths = Vec::new();
        for directory in [
            "crates/docx-edit/tests/fixtures",
            "crates/betteroffice-docx/tests/corpus/fixtures",
            "packages/docx/src/yrs/__fixtures__",
        ] {
            let start = paths.len();
            collect_docx(&root.join(directory), &mut paths);
            assert!(paths.len() > start, "empty corpus root {directory}");
        }
        paths.sort();
        let mut supported = 0;
        for path in paths {
            let source: PackageBytes = std::fs::read(&path).unwrap().into();
            let digest = seed::package_digest(&source);
            let (envelope, parts, media) =
                match seed::parse_docx_package_with_media(source.clone(), digest.clone()) {
                    Ok(parsed) => parsed,
                    Err(error) => {
                        eprintln!("unsupported {}: {error}", path.display());
                        continue;
                    }
                };
            let worker = worker(&source, &digest, 7).unwrap();
            let state = worker.encode_state_as_update_v1();
            let metadata = worker.encode_peer_metadata().unwrap();
            let (mut expected_source, index, _) =
                seed::replica_source(envelope, parts, source.clone(), digest).unwrap();
            let baseline = EditingDoc::new(19);
            expected_source.watch_comments(&baseline);
            baseline.install_source(expected_source, 17);
            baseline.retain_source(SourcePackage::Ready(Arc::new(index)));
            baseline.install_media(media);
            let peer = EditingDoc::new(19);
            let prepared = peer
                .prepare_peer_bootstrap(&state, &metadata, Some(source))
                .unwrap();
            peer.install_peer_bootstrap(prepared, 17).unwrap();
            assert_fields(
                &canonical(serde_json::to_value(peer.source_metadata().unwrap().as_ref()).unwrap()),
                &canonical(
                    serde_json::to_value(baseline.source_metadata().unwrap().as_ref()).unwrap(),
                ),
                &path.display().to_string(),
            );
            peer.apply_update_v1(&state).unwrap();
            baseline.apply_update_v1(&state).unwrap();
            assert_eq!(
                peer.encode_state_as_update_v1(),
                baseline.encode_state_as_update_v1(),
                "{}",
                path.display()
            );
            assert_eq!(
                peer.paragraph_identities(),
                baseline.paragraph_identities(),
                "{}",
                path.display()
            );
            supported += 1;
        }
        assert!(supported > 0);
    }

    fn worker(bytes: &PackageBytes, digest: &str, client: u64) -> Result<EditingDoc, String> {
        let (envelope, parts, media) =
            seed::parse_docx_package_with_media(bytes.clone(), digest.to_owned())?;
        let document = EditingDoc::new(client);
        seed::seed_parsed_docx(
            &document,
            envelope,
            parts,
            bytes.clone(),
            digest.to_owned(),
            SeedMedia::DataUrls {
                table: &media,
                layout_tokens: true,
            },
        )?;
        document.install_media(media);
        document.begin_opening(Some("peer-bootstrap-test"));
        Ok(document)
    }

    fn assert_fields(actual: &Value, expected: &Value, path: &str) {
        match (actual, expected) {
            (Value::Object(actual), Value::Object(expected)) => {
                assert_eq!(
                    actual.keys().collect::<Vec<_>>(),
                    expected.keys().collect::<Vec<_>>(),
                    "{path}"
                );
                for (key, value) in expected {
                    assert_fields(&actual[key], value, &format!("{path}.{key}"));
                }
            }
            (Value::Array(actual), Value::Array(expected)) => {
                assert_eq!(actual.len(), expected.len(), "{path}");
                for (index, (actual, expected)) in actual.iter().zip(expected).enumerate() {
                    assert_fields(actual, expected, &format!("{path}[{index}]"));
                }
            }
            _ => assert_eq!(actual, expected, "{path}"),
        }
    }

    fn assert_keys(value: &Value, expected: &[&str]) {
        use std::collections::BTreeSet;

        let actual: BTreeSet<_> = value
            .as_object()
            .unwrap()
            .keys()
            .map(String::as_str)
            .collect();
        assert_eq!(actual, expected.iter().copied().collect::<BTreeSet<_>>());
    }

    fn assert_source_fields(source: &SourceMetadata, index: &SourceIndex) {
        let source = serde_json::to_value(source).unwrap();
        assert_keys(&source, &["styles", "structure", "read"]);
        assert_keys(
            &source["styles"],
            &[
                "enabled",
                "styles",
                "doc_defaults",
                "default_paragraph",
                "default_table",
                "default_character",
                "table_paragraph_formatting",
            ],
        );
        assert_keys(&source["structure"], &["blocks", "run_revisions"]);
        assert_keys(
            &source["read"],
            &[
                "document_part",
                "stories",
                "footnote_separators",
                "endnote_separators",
                "note_separator_paragraphs",
                "comments",
                "seeded_comments",
                "raw_block_anchors",
                "comment_anchors",
                "final_section",
                "settings",
                "relationships",
                "numbering",
                "provenance",
                "warnings",
                "control_safety",
                "source_controls",
                "unlocated_controls",
                "unrepresented_controls",
                "unrepresented_anchors",
                "ambiguous_safety",
            ],
        );
        assert_keys(
            &source["read"]["provenance"],
            &[
                "raw_blocks",
                "block_order",
                "inline",
                "raw_sources",
                "tables",
                "relocated",
                "moves",
                "paragraph_sources",
            ],
        );
        let index = serde_json::to_value(index).unwrap();
        assert_keys(
            &index,
            &[
                "package_sha256",
                "occupied",
                "parts",
                "roots",
                "seeded",
                "comment_references",
            ],
        );
        for part in index["parts"].as_array().unwrap() {
            assert_keys(
                part,
                &[
                    "uri",
                    "sha256",
                    "as_written",
                    "kind",
                    "occurrences",
                    "backed",
                    "roots",
                ],
            );
        }
        for entry in index["seeded"].as_array().unwrap() {
            assert_keys(&entry[1], &["root", "part", "ordinal", "source_para_id"]);
        }
    }

    fn assert_peer_runtime(read: &ReadSource) {
        assert!(!read.pinned);
        assert!(read.provenance.by_story.is_empty());
        assert!(read.provenance.witnessed.is_empty());
        assert!(
            read.provenance
                .inline
                .iter()
                .all(|record| record.pin.position.is_none())
        );
        assert!(
            read.provenance
                .relocated
                .iter()
                .all(|record| record.pin.position.is_none())
        );
        assert!(read.embed_safety.get().is_none());
        assert!(read.comment_writes.snapshot().is_empty());
    }

    fn assert_round_trip(
        bytes: &PackageBytes,
        digest: &str,
        document: &EditingDoc,
        label: &str,
    ) -> PeerMetadata {
        assert!(
            document.source_metadata().unwrap().read().pinned,
            "{label}: worker profile"
        );
        let index = document.source_index().unwrap();
        assert!(!index.seed_states_initialized(), "{label}");
        let encoded = encode(document).unwrap();
        assert_eq!(encoded, encode(document).unwrap(), "{label}: determinism");
        assert!(
            !index.seed_states_initialized(),
            "{label}: export warmed save cache"
        );
        let decoded = decode(&encoded, bytes.clone(), digest).unwrap();
        let (envelope, parts, media) =
            seed::parse_docx_package_with_media(bytes.clone(), digest.to_owned()).unwrap();
        let (source, expected_index, _) =
            seed::replica_source(envelope, parts, bytes.clone(), digest.to_owned()).unwrap();
        assert_source_fields(&decoded.source, &decoded.index);
        assert_fields(
            &canonical(serde_json::to_value(&decoded.source).unwrap()),
            &canonical(serde_json::to_value(&source).unwrap()),
            &format!("{label}.source"),
        );
        assert_fields(
            &canonical(serde_json::to_value(&decoded.index).unwrap()),
            &canonical(serde_json::to_value(&expected_index).unwrap()),
            &format!("{label}.index"),
        );
        assert_peer_runtime(source.read());
        assert_peer_runtime(decoded.source.read());
        for story in &source.read().stories {
            assert_fields(
                &serde_json::to_value(decoded.source.read().story(&story.story).unwrap()).unwrap(),
                &serde_json::to_value(source.read().story(&story.story).unwrap()).unwrap(),
                &format!("{label}.story_index.{}", story.story),
            );
        }
        assert_eq!(
            decoded.media.descriptors(),
            media.descriptors(),
            "{label}: media descriptors"
        );
        assert_eq!(
            decoded.media.warnings(),
            media.warnings(),
            "{label}: media warnings"
        );
        assert_eq!(decoded.source_digest, digest, "{label}");
        assert_eq!(decoded.retained_source.as_ref(), bytes.as_ref(), "{label}");
        assert_eq!(
            decoded.index.bytes().as_ptr(),
            decoded.retained_source.as_ptr(),
            "{label}: shared source backing"
        );
        assert!(!decoded.index.seed_states_initialized());
        assert_eq!(decoded.comment_baseline, CommentBaseline::RebaseOnFirstLoad);
        assert!(decoded.media_sources.is_empty());
        assert_eq!(decoded.media_sources, crate::media::MediaSources::default());
        decoded
    }

    fn collect_docx(directory: &Path, paths: &mut Vec<PathBuf>) {
        for entry in std::fs::read_dir(directory).unwrap() {
            let entry = entry.unwrap();
            let path = entry.path();
            if entry.file_type().unwrap().is_dir() {
                collect_docx(&path, paths);
            } else if path
                .extension()
                .is_some_and(|extension| extension == "docx")
            {
                paths.push(path);
            }
        }
    }

    #[test]
    fn peer_metadata_full_corpus_round_trip() {
        let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
        let root = manifest.ancestors().nth(2).unwrap();
        let mut paths = Vec::new();
        for directory in [
            "crates/docx-edit/tests/fixtures",
            "crates/betteroffice-docx/tests/corpus/fixtures",
            "packages/docx/src/yrs/__fixtures__",
        ] {
            let start = paths.len();
            collect_docx(&root.join(directory), &mut paths);
            assert!(paths.len() > start, "empty corpus root {directory}");
        }
        paths.sort();
        let mut supported = 0;
        for path in paths {
            let bytes: PackageBytes = std::fs::read(&path).unwrap().into();
            let digest = seed::package_digest(&bytes);
            if let Err(error) = seed::parse_docx_package_with_media(bytes.clone(), digest.clone()) {
                eprintln!("unsupported {}: {error}", path.display());
                continue;
            }
            let label = path.display().to_string();
            let document =
                worker(&bytes, &digest, 7).unwrap_or_else(|error| panic!("{label}: {error}"));
            assert_round_trip(&bytes, &digest, &document, &label);
            let second = worker(&bytes, &digest, 19).unwrap();
            assert_eq!(
                encode(&document).unwrap(),
                encode(&second).unwrap(),
                "{label}: independent opens"
            );
            supported += 1;
        }
        assert!(supported > 0);
    }

    #[test]
    fn peer_metadata_deterministic_independent_opens() {
        let bytes = synthetic_package();
        let digest = seed::package_digest(&bytes);
        let first = worker(&bytes, &digest, 7).unwrap();
        let second = worker(&bytes, &digest, 19).unwrap();
        assert_eq!(encode(&first).unwrap(), encode(&first).unwrap());
        assert_eq!(encode(&first).unwrap(), encode(&second).unwrap());
        assert_round_trip(&bytes, &digest, &first, "synthetic");
    }

    #[test]
    fn peer_metadata_export_requires_resident_source_and_media() {
        assert_eq!(
            encode(&EditingDoc::new(7)),
            Err(PeerMetadataError::MissingSource)
        );
        let bytes = synthetic_package();
        let digest = seed::package_digest(&bytes);
        let document = worker(&bytes, &digest, 7).unwrap();
        *document.media.lock().unwrap() = None;
        assert_eq!(encode(&document), Err(PeerMetadataError::MissingSource));
        assert!(document.media.lock().unwrap().is_none());
        document.retain_source(SourcePackage::Pending(bytes, Some(digest)));
        assert_eq!(encode(&document), Err(PeerMetadataError::MissingSource));
        assert!(matches!(
            document.source.lock().unwrap().as_ref(),
            Some(SourcePackage::Pending(_, _))
        ));
    }

    #[test]
    fn peer_metadata_round_trip_without_body_paragraphs() {
        let mut parts = fixture::principal_parts();
        let body = parts
            .iter_mut()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap();
        body.1 = format!(
            "<w:document {}><w:body><w:sectPr/></w:body></w:document>",
            fixture::NS
        )
        .into_bytes();
        let bytes: PackageBytes = ooxml_opc::rezip_parts(&parts).unwrap().into();
        let digest = seed::package_digest(&bytes);
        let document = worker(&bytes, &digest, 7).unwrap();
        assert_round_trip(&bytes, &digest, &document, "no body paragraphs");
    }

    #[test]
    fn peer_metadata_preserves_occupied_duplicate_shared_and_companion_ids() {
        let mut parts = fixture::principal_parts();
        for (name, bytes) in &mut parts {
            if name == "word/document.xml" {
                *bytes = String::from_utf8(bytes.clone())
                    .unwrap()
                    .replace("00000002", "00000001")
                    .into_bytes();
            } else if name == "word/_rels/document.xml.rels" {
                *bytes = String::from_utf8(bytes.clone())
                    .unwrap()
                    .replace("Target=\"header2.xml\"", "Target=\"header1.xml\"")
                    .into_bytes();
            }
        }
        parts.push((
            "word/unused.xml".to_owned(),
            format!(
                "<w:document {}><w:body><w:p w14:paraId=\"7ABCDE01\"/></w:body></w:document>",
                fixture::NS,
            )
            .into_bytes(),
        ));
        parts.push(("word/commentsExtended.xml".to_owned(), br#"<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="30000001" w15:paraIdParent="7ABCDE02"/></w15:commentsEx>"#.to_vec()));
        parts.push(("word/commentsIds.xml".to_owned(), br#"<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w16cid:commentId w16cid:paraId="7ABCDE03"/></w16cid:commentsIds>"#.to_vec()));
        let bytes: PackageBytes = ooxml_opc::rezip_parts(&parts).unwrap().into();
        let digest = seed::package_digest(&bytes);
        let document = worker(&bytes, &digest, 7).unwrap();
        let peer = assert_round_trip(&bytes, &digest, &document, "identities");
        let index = serde_json::to_value(&peer.index).unwrap();
        assert!(
            index["occupied"]
                .as_array()
                .unwrap()
                .contains(&json!(0x7ABCDE01u32))
        );
        for id in [0x30000001u32, 0x7ABCDE02, 0x7ABCDE03] {
            assert!(
                index["comment_references"]
                    .as_array()
                    .unwrap()
                    .contains(&json!(id))
            );
        }
        let parts = index["parts"].as_array().unwrap();
        let body = parts.iter().find(|part| part["kind"] == "Body").unwrap();
        assert_eq!(
            body["occurrences"]
                .as_array()
                .unwrap()
                .iter()
                .filter(|occurrence| occurrence["para_id"] == "00000001")
                .count(),
            2
        );
        let header = parts
            .iter()
            .find(|part| part["uri"] == "/word/header1.xml")
            .unwrap();
        assert_eq!(header["roots"].as_array().unwrap().len(), 2);
        assert!(
            header["backed"]
                .as_array()
                .unwrap()
                .iter()
                .any(|entry| entry[1].as_array().unwrap().len() == 2)
        );
    }

    #[test]
    fn peer_metadata_preserves_safety_omissions_grids_and_separators() {
        let mut parts = fixture::principal_parts();
        let body = parts
            .iter_mut()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap();
        let control = |content: &str| {
            format!(
                "<w:sdt><w:sdtPr><w:id w:val=\"71\"/><w:tag w:val=\"same\"/><w:text/></w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>"
            )
        };
        let added = format!(
            "<w:p w14:paraId=\"00000071\">{}</w:p><w:p w14:paraId=\"00000072\">{}</w:p><bofx:raw>{}</bofx:raw>",
            control("<w:r><w:t>Safe</w:t></w:r>"),
            control("<w:fldSimple w:instr=\" PAGE \"><w:r><w:t>1</w:t></w:r></w:fldSimple>"),
            control("<w:p><w:r><w:t>Source only</w:t></w:r></w:p>"),
        );
        let uncertain = concat!(
            "<w:p xmlns:q=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w14:paraId=\"00000073\">",
            "<w:sdt><w:sdtPr><w:id w:val=\"73\"/><w:lock q:val=\"contentLocked\"/></w:sdtPr>",
            "<w:sdtContent><w:r><w:t>Uncertain policy</w:t></w:r></w:sdtContent></w:sdt></w:p>",
        );
        body.1 = String::from_utf8(body.1.clone())
            .unwrap()
            .replace("</w:body>", &format!("{added}{uncertain}</w:body>"))
            .into_bytes();
        let bytes: PackageBytes = ooxml_opc::rezip_parts(&parts).unwrap().into();
        let digest = seed::package_digest(&bytes);
        let document = worker(&bytes, &digest, 7).unwrap();
        let peer = assert_round_trip(&bytes, &digest, &document, "safety and provenance");
        let read = peer.source.read();
        assert!(!read.ambiguous_safety.is_empty());
        assert!(
            read.control_safety
                .as_ref()
                .unwrap()
                .values()
                .any(|safety| !safety.unsupported.is_empty())
        );
        assert!(
            read.control_safety
                .as_ref()
                .unwrap()
                .values()
                .any(|safety| safety.uncertain)
        );
        assert!(!read.source_controls.is_empty());
        assert!(!read.provenance.inline.is_empty());
        assert!(!read.provenance.raw_blocks.is_empty());
        assert!(!read.provenance.tables.is_empty());
        assert!(!read.provenance.relocated.is_empty());
        assert!(!read.provenance.moves.is_empty());
        assert!(read.footnote_separators > 0);
        assert!(read.endnote_separators > 0);
        assert!(!read.note_separator_paragraphs.is_empty());
        assert!(!read.comments.is_empty());
        assert!(!read.comment_anchors.is_empty());
    }

    #[test]
    fn peer_metadata_discards_worker_comment_history_including_placeholders() {
        let bytes = synthetic_package();
        let digest = seed::package_digest(&bytes);
        let document = worker(&bytes, &digest, 7).unwrap();
        let before = encode(&document).unwrap();
        let source = document.source_metadata().unwrap();
        {
            let mut txn = document.yrs_doc().transact_mut();
            let comments = txn.get_map(crate::COMMENTS).unwrap();
            let comment = comments.get(&txn, "1").unwrap().cast::<MapRef>().unwrap();
            comment.insert(&mut txn, "author", "Edited author");
            comment.insert(&mut txn, "done", true);
            let authored = comments.insert(&mut txn, "authored", MapPrelim::default());
            authored.insert(&mut txn, "author", "New author");
        }
        {
            let mut txn = document.yrs_doc().transact_mut();
            let comments = txn.get_map(crate::COMMENTS).unwrap();
            let comment = comments.get(&txn, "1").unwrap().cast::<MapRef>().unwrap();
            comment.insert(&mut txn, "author", "");
            comment.insert(&mut txn, "date", "");
            comment.insert(&mut txn, "parentId", Any::Null);
            comment.insert(&mut txn, "body", Any::Null);
            comment.insert(&mut txn, "done", false);
        }
        assert!(source.read().comment_writes.written("1", "author"));
        assert!(source.read().comment_writes.written("1", "done"));
        assert!(source.read().comment_writes.written("authored", "author"));
        assert_eq!(before, encode(&document).unwrap());
        let peer = decode(&before, bytes, &digest).unwrap();
        assert_eq!(peer.comment_baseline, CommentBaseline::RebaseOnFirstLoad);
        assert!(peer.source.read().comment_writes.snapshot().is_empty());
        assert!(!peer.source.read().comment_writes.written("1", "author"));
        assert!(!peer.source.read().comment_writes.written("1", "done"));
        assert!(
            !peer
                .source
                .read()
                .comment_writes
                .written("authored", "body")
        );
    }

    #[test]
    fn peer_metadata_raw_locators_round_trip() {
        use crate::structured::source::{RawSource, Step};

        let source = RawSource {
            story: "fn:1".to_owned(),
            index: 3,
            steps: vec![
                Step::Body,
                Step::Note("footnote", "1".to_owned()),
                Step::Note("endnote", "2".to_owned()),
                Step::Comment("3".to_owned()),
                Step::Block(4),
                Step::Row(5),
                Step::Cell(6),
                Step::Content,
            ],
            xml: "<opaque/>".to_owned(),
        };
        let json = serde_json::to_value(&source).unwrap();
        let decoded: RawSource = serde_json::from_value(json.clone()).unwrap();
        assert_fields(
            &serde_json::to_value(&decoded).unwrap(),
            &json,
            "raw locator",
        );
        assert!(matches!(&decoded.steps[1], Step::Note("footnote", id) if id == "1"));
        assert!(matches!(&decoded.steps[2], Step::Note("endnote", id) if id == "2"));
        let bad = RawSource {
            steps: vec![Step::Note("unknown", "1".to_owned())],
            ..source
        };
        assert!(serde_json::to_value(&bad).is_err());
    }

    #[test]
    fn peer_metadata_keeps_transcodes_in_binary_blobs() {
        let bytes = synthetic_package();
        let digest = seed::package_digest(&bytes);
        let document = worker(&bytes, &digest, 7).unwrap();
        let media = document.media_table().unwrap();
        let mut descriptors = media.descriptors();
        let display = b"<svg>peer bootstrap display bytes</svg>".to_vec();
        descriptors[0].display = Some(display.clone());
        descriptors[0].mime_type = "image/svg+xml".to_owned();
        document.install_media(
            MediaTable::from_descriptors(
                RetainedPackage::from_bytes(bytes.clone()).unwrap(),
                descriptors.clone(),
                vec!["existing transcode warning".to_owned()],
            )
            .unwrap(),
        );
        let encoded = encode(&document).unwrap();
        let (json, blobs) = sections(&encoded).unwrap();
        assert_eq!(blobs, display.as_slice());
        for offset in [HEADER_LEN + json.len(), encoded.len() - 1] {
            assert_eq!(
                error(&encoded[..offset], &bytes, &digest),
                PeerMetadataError::Truncated,
            );
        }
        assert!(
            !json
                .windows(display.len())
                .any(|window| window == display.as_slice())
        );
        assert!(
            !encoded
                .windows(bytes.len())
                .any(|window| window == bytes.as_ref())
        );
        let peer = decode(&encoded, bytes, &digest).unwrap();
        assert_eq!(peer.media.descriptors(), descriptors);
        assert_eq!(peer.media.bytes(0).unwrap().as_ref(), display.as_slice());
        assert_eq!(
            peer.media.warnings(),
            &["existing transcode warning".to_owned()]
        );
        assert!(peer.media_sources.is_empty());
    }

    fn error(bytes: &[u8], source: &PackageBytes, digest: &str) -> PeerMetadataError {
        match decode(bytes, source.clone(), digest) {
            Err(error) => error,
            Ok(_) => panic!("invalid metadata decoded"),
        }
    }

    #[test]
    fn peer_metadata_frame_errors_are_typed() {
        let source = synthetic_package();
        let digest = seed::package_digest(&source);
        let document = worker(&source, &digest, 7).unwrap();
        let bytes = encode(&document).unwrap();
        let mut bad = bytes.clone();
        bad[0] ^= 1;
        assert_eq!(error(&bad, &source, &digest), PeerMetadataError::BadMagic);
        let mut bad = bytes.clone();
        bad[8..12].copy_from_slice(&2u32.to_le_bytes());
        assert_eq!(
            error(&bad, &source, &digest),
            PeerMetadataError::UnsupportedVersion(2)
        );
        let mut bad = bytes.clone();
        bad[12] ^= 1;
        assert_eq!(
            error(&bad, &source, &digest),
            PeerMetadataError::ShapeMismatch
        );
        for offset in [
            0,
            1,
            7,
            8,
            11,
            12,
            43,
            44,
            51,
            52,
            HEADER_LEN - 1,
            HEADER_LEN,
            bytes.len() / 2,
            bytes.len() - 1,
        ] {
            assert_eq!(
                error(&bytes[..offset], &source, &digest),
                PeerMetadataError::Truncated,
                "offset {offset}"
            );
        }
        let mut bad = bytes.clone();
        bad[HEADER_LEN] = b'!';
        assert!(matches!(
            error(&bad, &source, &digest),
            PeerMetadataError::Json(_)
        ));
        let mut bad = bytes.clone();
        bad[44..52].copy_from_slice(&u64::MAX.to_le_bytes());
        assert_eq!(
            error(&bad, &source, &digest),
            PeerMetadataError::InvalidLength
        );
        let mut bad = bytes.clone();
        bad.push(0);
        assert_eq!(
            error(&bad, &source, &digest),
            PeerMetadataError::InvalidLength
        );
        assert_eq!(
            error(&bytes, &source, &"0".repeat(64)),
            PeerMetadataError::SourceMismatch
        );
        assert_eq!(
            error(&bytes, &Vec::new().into(), &digest),
            PeerMetadataError::SourceMismatch
        );
    }

    #[test]
    fn peer_metadata_rejects_invalid_metadata_and_blob_ranges() {
        let source = synthetic_package();
        let digest = seed::package_digest(&source);
        let document = worker(&source, &digest, 7).unwrap();
        let encoded = encode(&document).unwrap();
        let (json, blobs) = sections(&encoded).unwrap();
        let wire: Value = serde_json::from_slice(json).unwrap();
        let check = |value: &Value| {
            error(
                &frame(&serde_json::to_vec(value).unwrap(), blobs).unwrap(),
                &source,
                &digest,
            )
        };
        let mut bad = wire.clone();
        bad["source"]["read"]
            .as_object_mut()
            .unwrap()
            .remove("settings");
        assert!(matches!(check(&bad), PeerMetadataError::Json(_)));
        let mut bad = wire.clone();
        bad["source"]["read"]["pinned"] = json!(true);
        assert!(matches!(check(&bad), PeerMetadataError::Json(_)));
        let mut bad = wire.clone();
        bad["index"]["parts"][0]["uri"] = json!("");
        assert!(matches!(check(&bad), PeerMetadataError::InvalidMetadata(_)));
        let mut bad = wire.clone();
        bad["media"]["parts"][0]["mime_type"] = json!("unknown/type");
        assert!(matches!(check(&bad), PeerMetadataError::InvalidMetadata(_)));
        let mut bad = wire.clone();
        bad["media"]["parts"][0]["path"] = json!("word/media/wrong.png");
        assert!(matches!(check(&bad), PeerMetadataError::InvalidMetadata(_)));
        let mut bad = wire.clone();
        bad["media"]["parts"] = json!([]);
        assert!(matches!(check(&bad), PeerMetadataError::InvalidMetadata(_)));
        let mut bad = wire.clone();
        bad["comment_baseline"] = json!("CopyWorkerWrites");
        assert!(matches!(check(&bad), PeerMetadataError::Json(_)));
        let mut bad = wire.clone();
        bad["media"]["parts"][0]["display"] = json!({"offset": 1, "length": 0});
        assert_eq!(check(&bad), PeerMetadataError::InvalidLength);
        bad["media"]["parts"][0]["display"] = json!({"offset": 0, "length": 1});
        assert_eq!(check(&bad), PeerMetadataError::Truncated);
        let mut bad = wire;
        bad["media"]["parts"][0]["display"] = json!({"offset": 0, "length": u64::MAX});
        assert!(matches!(
            check(&bad),
            PeerMetadataError::InvalidLength | PeerMetadataError::Truncated
        ));
    }
}
