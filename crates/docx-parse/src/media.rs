//! Embedded media table and image-resolution aliases.

use std::borrow::Cow;
use std::collections::HashMap;
use std::sync::Arc;

use base64::Engine as _;
use indexmap::IndexMap;
use serde::ser::SerializeMap;
use serde::{Deserialize, Deserializer, Serialize, Serializer};

use crate::relationships::RelationshipMap;

/// Media files keyed by package path, with a lowercase index so
/// case-insensitive lookups stay map hits.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct MediaMap {
    entries: IndexMap<String, Arc<MediaFile>>,
    lower: HashMap<String, String>,
}

impl MediaMap {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn len(&self) -> usize {
        self.entries.len()
    }

    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    pub fn get(&self, key: &str) -> Option<&MediaFile> {
        self.entries.get(key).map(|file| &**file)
    }

    /// Exact hit first, then a case-insensitive retry through the index.
    pub fn get_case_insensitive(&self, key: &str) -> Option<&MediaFile> {
        self.get(key).or_else(|| {
            self.lower
                .get(key.to_ascii_lowercase().as_str())
                .and_then(|stored| self.entries.get(stored).map(|file| &**file))
        })
    }

    /// Bare (`media/x.png`) and `word/`-rooted targets resolve the same.
    pub(crate) fn find_target(&self, target: &str) -> Option<&MediaFile> {
        let trimmed = target.trim_start_matches('/');
        self.get_case_insensitive(trimmed).or_else(|| {
            (!trimmed.starts_with("word/"))
                .then(|| format!("word/{trimmed}"))
                .and_then(|candidate| self.get_case_insensitive(&candidate))
        })
    }

    pub fn iter(&self) -> impl Iterator<Item = (&String, &MediaFile)> {
        self.entries.iter().map(|(key, file)| (key, &**file))
    }

    pub fn keys(&self) -> impl Iterator<Item = &String> {
        self.entries.keys()
    }

    /// IndexMap-compatible insert; returns any replaced entry.
    pub fn insert(&mut self, key: String, file: MediaFile) -> Option<MediaFile> {
        self.lower
            .entry(key.to_ascii_lowercase())
            .or_insert_with(|| key.clone());
        self.entries
            .insert(key, Arc::new(file))
            .map(|old| (*old).clone())
    }

    fn insert_key(&mut self, key: String, file: Arc<MediaFile>) {
        self.lower
            .entry(key.to_ascii_lowercase())
            .or_insert_with(|| key.clone());
        self.entries.insert(key, file);
    }
}

impl std::ops::Deref for MediaMap {
    type Target = IndexMap<String, Arc<MediaFile>>;

    fn deref(&self) -> &Self::Target {
        &self.entries
    }
}

impl std::ops::Index<&str> for MediaMap {
    type Output = MediaFile;

    fn index(&self, key: &str) -> &MediaFile {
        self.entries[key].as_ref()
    }
}

fn as_ref_entry<'a>((key, file): (&'a String, &'a Arc<MediaFile>)) -> (&'a String, &'a MediaFile) {
    (key, file.as_ref())
}

fn unwrap_entry((key, file): (String, Arc<MediaFile>)) -> (String, MediaFile) {
    (
        key,
        Arc::try_unwrap(file).unwrap_or_else(|file| (*file).clone()),
    )
}

impl<'a> IntoIterator for &'a MediaMap {
    type Item = (&'a String, &'a MediaFile);
    type IntoIter = std::iter::Map<
        indexmap::map::Iter<'a, String, Arc<MediaFile>>,
        fn((&'a String, &'a Arc<MediaFile>)) -> (&'a String, &'a MediaFile),
    >;

    fn into_iter(self) -> Self::IntoIter {
        self.entries.iter().map(as_ref_entry)
    }
}

impl IntoIterator for MediaMap {
    type Item = (String, MediaFile);
    type IntoIter = std::iter::Map<
        indexmap::map::IntoIter<String, Arc<MediaFile>>,
        fn((String, Arc<MediaFile>)) -> (String, MediaFile),
    >;

    fn into_iter(self) -> Self::IntoIter {
        self.entries.into_iter().map(unwrap_entry)
    }
}

impl FromIterator<(String, MediaFile)> for MediaMap {
    fn from_iter<I: IntoIterator<Item = (String, MediaFile)>>(iter: I) -> Self {
        let mut map = Self::new();
        map.extend(iter);
        map
    }
}

impl Extend<(String, MediaFile)> for MediaMap {
    fn extend<I: IntoIterator<Item = (String, MediaFile)>>(&mut self, iter: I) {
        for (key, file) in iter {
            self.insert(key, file);
        }
    }
}

impl Serialize for MediaMap {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(self.entries.len()))?;
        for (key, file) in &self.entries {
            map.serialize_entry(key, &**file)?;
        }
        map.end()
    }
}

impl<'de> Deserialize<'de> for MediaMap {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        IndexMap::<String, MediaFile>::deserialize(deserializer).map(Self::from_iter)
    }
}

/// One media payload; `base64` is canonical and [`Self::data_url`] derives
/// the display form, so the bytes are stored once.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MediaFile {
    pub path: String,
    pub filename: Option<String>,
    pub mime_type: String,
    pub base64: String,
}

impl MediaFile {
    /// `data:<mime>;base64,<payload>` derived from `base64`.
    pub fn data_url(&self) -> String {
        format!("data:{};base64,{}", self.mime_type, self.base64)
    }
}

impl Serialize for MediaFile {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(4 + usize::from(self.filename.is_some())))?;
        map.serialize_entry("path", &self.path)?;
        if let Some(filename) = &self.filename {
            map.serialize_entry("filename", filename)?;
        }
        map.serialize_entry("mimeType", &self.mime_type)?;
        map.serialize_entry("base64", &self.base64)?;
        map.serialize_entry("dataUrl", &self.data_url())?;
        map.end()
    }
}

impl<'de> Deserialize<'de> for MediaFile {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        #[derive(Deserialize)]
        #[serde(rename_all = "camelCase")]
        struct WireMediaFile {
            path: String,
            filename: Option<String>,
            mime_type: String,
            #[serde(default)]
            base64: String,
            data_url: String,
        }
        let wire = WireMediaFile::deserialize(deserializer)?;
        // A payload serialized only as a data URL keeps its base64 half.
        let base64 = if wire.base64.is_empty() {
            wire.data_url
                .split_once(',')
                .filter(|(head, _)| head.starts_with("data:") && head.ends_with(";base64"))
                .map_or_else(String::new, |(_, payload)| payload.to_owned())
        } else {
            wire.base64
        };
        Ok(Self {
            path: wire.path,
            filename: wire.filename,
            mime_type: wire.mime_type,
            base64,
        })
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ResolvedImageData {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub src: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mime_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub filename: Option<String>,
}

pub fn build_media_map(parts: &[(String, Vec<u8>)]) -> MediaMap {
    build_media_map_with_warnings(parts).0
}

/// Media map plus one warning per part that could not be transcoded for display.
pub fn build_media_map_with_warnings(parts: &[(String, Vec<u8>)]) -> (MediaMap, Vec<String>) {
    let mut media = MediaMap::new();
    let mut warnings = Vec::new();
    for (path, data) in parts {
        if !path.to_ascii_lowercase().starts_with("word/media/") {
            continue;
        }
        let filename = path.rsplit('/').next().unwrap_or(path).to_owned();
        let (data, mime_type, warning) = display_form(data, media_mime_type(path), path);
        warnings.extend(warning);
        let mime_type = mime_type.to_owned();
        let base64 = base64::engine::general_purpose::STANDARD.encode(&data);
        let file = Arc::new(MediaFile {
            path: path.clone(),
            filename: Some(filename),
            base64,
            mime_type,
        });
        media.insert_key(path.clone(), Arc::clone(&file));
        if let Some(normalized) = path.strip_prefix("word/") {
            media.insert_key(normalized.to_owned(), file);
        }
    }
    (media, warnings)
}

pub fn resolve_image_data(
    relationship_id: &str,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
) -> ResolvedImageData {
    if relationship_id.is_empty() {
        return ResolvedImageData::default();
    }
    let Some(relationship) = relationships.and_then(|map| map.get(relationship_id)) else {
        return ResolvedImageData::default();
    };
    if relationship.target.is_empty() {
        return ResolvedImageData::default();
    }
    let target = &relationship.target;
    let filename = target.rsplit('/').next().map(str::to_owned);
    if let Some(file) = media.and_then(|media| media.find_target(target)) {
        return ResolvedImageData {
            src: Some(if file.base64.is_empty() {
                file.base64.clone()
            } else {
                file.data_url()
            }),
            mime_type: Some(file.mime_type.clone()),
            filename,
        };
    }
    ResolvedImageData {
        src: None,
        mime_type: Some(media_mime_type(target).to_owned()),
        filename,
    }
}

/// Browsers have no TIFF decoder, so the display copy carries a PNG transcode.
/// Save reads the untouched package part, so the original bytes still round-trip.
/// An encoding the decoder does not support keeps the TIFF source — decoders that
/// do handle it still render — and reports why the transcode was skipped.
#[cfg(feature = "tiff")]
fn display_form<'a>(
    data: &'a [u8],
    mime_type: &'static str,
    path: &str,
) -> (Cow<'a, [u8]>, &'static str, Option<String>) {
    if !is_tiff(data) {
        return (Cow::Borrowed(data), mime_type, None);
    }
    match ooxml_drawingml::media::decode_tiff_png(data) {
        Ok(png) => (Cow::Owned(png), "image/png", None),
        Err(error) => (
            Cow::Borrowed(data),
            mime_type,
            Some(format!(
                "TIFF image {path} could not be decoded for display: {error}"
            )),
        ),
    }
}

#[cfg(not(feature = "tiff"))]
fn display_form<'a>(
    data: &'a [u8],
    mime_type: &'static str,
    _path: &str,
) -> (Cow<'a, [u8]>, &'static str, Option<String>) {
    (Cow::Borrowed(data), mime_type, None)
}

#[cfg(feature = "tiff")]
fn is_tiff(data: &[u8]) -> bool {
    matches!(data.first_chunk::<4>(), Some(b"II\x2a\x00" | b"MM\x00\x2a"))
}

pub fn media_mime_type(path: &str) -> &'static str {
    match path
        .rsplit('.')
        .next()
        .unwrap_or_default()
        .to_ascii_lowercase()
        .as_str()
    {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        "webp" => "image/webp",
        "svg" => "image/svg+xml",
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf",
        _ => "application/octet-stream",
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::relationships::{Relationship, TargetMode};

    #[test]
    fn aliases_bytes_and_resolves_paths_case_insensitively() {
        let parts = vec![("word/media/IMAGE.PNG".to_owned(), vec![0, 255, 16])];
        let media = build_media_map(&parts);
        assert_eq!(
            media.keys().collect::<Vec<_>>(),
            ["word/media/IMAGE.PNG", "media/IMAGE.PNG"]
        );
        let relationships = RelationshipMap::from([(
            "rId1".into(),
            Relationship {
                id: "rId1".into(),
                relationship_type: "image".into(),
                target: "media/image.png".into(),
                target_mode: Some(TargetMode::External),
            },
        )]);
        let resolved = resolve_image_data("rId1", Some(&relationships), Some(&media));
        // TargetMode is deliberately not used as a fetch signal. Resolution
        // only consults already embedded package bytes and performs no I/O.
        assert_eq!(resolved.mime_type.as_deref(), Some("image/png"));
        assert_eq!(resolved.filename.as_deref(), Some("image.png"));
        assert_eq!(resolved.src.as_deref(), Some("data:image/png;base64,AP8Q"));
    }

    #[test]
    fn missing_media_returns_only_extension_metadata() {
        let relationships = RelationshipMap::from([(
            "rId2".into(),
            Relationship {
                id: "rId2".into(),
                relationship_type: "image".into(),
                target: "../outside.WMF".into(),
                target_mode: None,
            },
        )]);
        assert_eq!(
            resolve_image_data("rId2", Some(&relationships), None),
            ResolvedImageData {
                src: None,
                mime_type: Some("image/x-wmf".into()),
                filename: Some("outside.WMF".into()),
            }
        );
    }
}
