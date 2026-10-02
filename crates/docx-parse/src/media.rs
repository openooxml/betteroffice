//! Embedded media table and image-resolution aliases.

use std::borrow::Cow;
use std::sync::Arc;

use base64::Engine as _;
use indexmap::IndexMap;
use ooxml_opc::{MAX_TOTAL_UNCOMPRESSED_BYTES, RetainedPackage};
use serde::{Deserialize, Serialize};

use crate::relationships::RelationshipMap;

pub type MediaMap = IndexMap<String, Arc<MediaFile>>;

/// How a [`MediaTable`] token starts: `media:{n}` names its `n`th part.
pub const MEDIA_TOKEN_PREFIX: &str = "media:";

/// The token naming a [`MediaTable`]'s `index`th part.
pub fn media_token(index: usize) -> String {
    format!("{MEDIA_TOKEN_PREFIX}{index}")
}

/// The part index `token` names, when it is a `media:{n}` token in its one
/// spelling.
pub fn media_token_index(token: &str) -> Option<usize> {
    let digits = token.strip_prefix(MEDIA_TOKEN_PREFIX)?;
    let canonical = !digits.is_empty()
        && digits.bytes().all(|byte| byte.is_ascii_digit())
        && (digits == "0" || !digits.starts_with('0'));
    canonical.then(|| digits.parse().ok()).flatten()
}

/// Bytes read to tell whether [`display_form`] replaces a part.
const SNIFFED_BYTES: usize = 64;

/// The `word/media/` parts of a retained package in archive order, the
/// `n`th of which a `media:{n}` token names. An image part stays compressed
/// in the package until read, and its `data:` URL is built only when asked
/// for. Parts browsers cannot decode are transcoded when the table is built,
/// with the warnings [`build_media_map_with_warnings`] reports.
#[derive(Clone, Debug)]
pub struct MediaTable {
    package: RetainedPackage,
    parts: Arc<[MediaPart]>,
    warnings: Arc<[String]>,
}

#[derive(Debug)]
struct MediaPart {
    path: String,
    position: usize,
    size: u64,
    /// An image, which no parser reads as markup and so need not inflate
    /// with the package.
    image: bool,
    mime_type: &'static str,
    display: Option<Vec<u8>>,
}

impl MediaTable {
    /// The media of `package`. The images it leaves compressed must fit the
    /// container budget by their declared sizes.
    pub fn new(package: RetainedPackage) -> Result<Self, String> {
        let scan = MediaScan::new(package, MAX_TOTAL_UNCOMPRESSED_BYTES)?;
        let total = scan
            .parts
            .iter()
            .filter(|part| part.transcode && !part.metafile)
            .fold(scan.images, |total, part| total.saturating_add(part.size));
        if total > MAX_TOTAL_UNCOMPRESSED_BYTES {
            return Err(budget_exceeded());
        }
        let mut parts = scan
            .parts
            .iter()
            .filter(|part| part.transcode && !part.metafile)
            .map(|part| {
                scan.package
                    .read(part.position)
                    .map(|bytes| (part.path.clone(), bytes))
            })
            .collect::<Result<Vec<_>, String>>()?;
        scan.finish(&mut parts)
    }

    /// Whether the package's part at `path` is an image the table reads
    /// itself, which inflating the package may skip.
    pub fn keeps_compressed(&self, path: &str) -> bool {
        self.parts
            .iter()
            .any(|part| part.image && part.path == path)
    }

    /// Whether the images left compressed, by their declared sizes, and the
    /// `inflated` bytes of the package's other parts fit the container budget.
    pub fn check_budget(&self, inflated: u64) -> Result<(), String> {
        let total = self
            .parts
            .iter()
            .filter(|part| part.image)
            .fold(inflated, |total, part| total.saturating_add(part.size));
        if total > MAX_TOTAL_UNCOMPRESSED_BYTES {
            return Err(budget_exceeded());
        }
        Ok(())
    }

    pub fn len(&self) -> usize {
        self.parts.len()
    }

    pub fn is_empty(&self) -> bool {
        self.parts.is_empty()
    }

    /// The package path of the `index`th part.
    pub fn path(&self, index: usize) -> Option<&str> {
        self.parts.get(index).map(|part| part.path.as_str())
    }

    /// The media type of the bytes [`MediaTable::bytes`] returns.
    pub fn mime_type(&self, index: usize) -> Option<&'static str> {
        self.parts.get(index).map(|part| part.mime_type)
    }

    /// The `index`th part as displayed: its transcode, or its inflated bytes.
    pub fn bytes(&self, index: usize) -> Result<Cow<'_, [u8]>, String> {
        let part = self
            .parts
            .get(index)
            .ok_or_else(|| format!("no media part {index}"))?;
        match &part.display {
            Some(display) => Ok(Cow::Borrowed(display)),
            None => self.package.read(part.position).map(Cow::Owned),
        }
    }

    /// The `data:` URL [`build_media_map`] gives the `index`th part.
    pub fn data_url(&self, index: usize) -> Result<String, String> {
        let bytes = self.bytes(index)?;
        let mime_type = self.parts[index].mime_type;
        let encoded = base64::encoded_len(bytes.len(), true).unwrap_or(0);
        let mut url = String::with_capacity("data:;base64,".len() + mime_type.len() + encoded);
        url.push_str("data:");
        url.push_str(mime_type);
        url.push_str(";base64,");
        base64::engine::general_purpose::STANDARD.encode_string(&bytes, &mut url);
        Ok(url)
    }

    /// The `data:` URL a `media:{n}` token stands for, or `None` for any other
    /// string or a part that cannot be read.
    pub fn resolve(&self, token: &str) -> Option<String> {
        media_token_index(token)
            .filter(|index| *index < self.len())
            .and_then(|index| self.data_url(index).ok())
    }

    /// One warning per part that could not be transcoded for display.
    pub fn warnings(&self) -> &[String] {
        &self.warnings
    }

    /// The media map image resolution reads, each part's token standing
    /// where [`build_media_map`] puts its `data:` URL.
    pub fn media_map(&self) -> MediaMap {
        let mut media = MediaMap::new();
        for (index, part) in self.parts.iter().enumerate() {
            let file = Arc::new(MediaFile {
                path: part.path.clone(),
                filename: Some(
                    part.path
                        .rsplit('/')
                        .next()
                        .unwrap_or(&part.path)
                        .to_owned(),
                ),
                mime_type: part.mime_type.to_owned(),
                base64: String::new(),
                data_url: media_token(index),
            });
            media.insert(part.path.clone(), Arc::clone(&file));
            if let Some(normalized) = part.path.strip_prefix("word/") {
                media.insert(normalized.to_owned(), file);
            }
        }
        media
    }
}

/// Media metadata and the budget reserved for images kept compressed.
pub(crate) struct MediaScan {
    package: RetainedPackage,
    parts: Vec<ScannedPart>,
    budget: u64,
    images: u64,
}

struct ScannedPart {
    path: String,
    position: usize,
    size: u64,
    image: bool,
    transcode: bool,
    metafile: bool,
}

impl MediaScan {
    /// Fails when compressed images' declared sizes exceed `budget`.
    pub(crate) fn new(package: RetainedPackage, budget: u64) -> Result<Self, String> {
        let budget = budget.min(MAX_TOTAL_UNCOMPRESSED_BYTES);
        let mut parts = Vec::new();
        let mut images = 0_u64;
        for (position, (path, size)) in package.parts().enumerate() {
            if !is_media_path(path) {
                continue;
            }
            let mime_type = media_mime_type(path);
            let prefix = match package.read_prefix(position, SNIFFED_BYTES) {
                Ok(prefix) => prefix,
                Err(_) if metafile_transcode(&[], mime_type) => Vec::new(),
                Err(error) => return Err(error),
            };
            let metafile = metafile_transcode(&prefix, mime_type);
            let image = is_image(&prefix) || metafile;
            let transcode = transcodes(&prefix, mime_type);
            if image && (!transcode || metafile) {
                images = images.saturating_add(size);
                if images > budget {
                    return Err(format!("inflated size exceeds {budget} bytes"));
                }
            }
            parts.push(ScannedPart {
                path: path.to_owned(),
                position,
                size,
                image,
                transcode,
                metafile,
            });
        }
        Ok(Self {
            package,
            parts,
            budget,
            images,
        })
    }

    /// Whether the image at `path` stays compressed during extraction.
    pub(crate) fn keeps_compressed(&self, path: &str) -> bool {
        self.parts
            .iter()
            .any(|part| part.image && (!part.transcode || part.metafile) && part.path == path)
    }

    /// What the images left compressed leave of the budget.
    pub(crate) fn remaining_budget(&self) -> u64 {
        self.budget - self.images
    }

    /// Moves transcode inputs from the bounded extraction into the table.
    pub(crate) fn finish(
        self,
        inflated: &mut Vec<(String, Vec<u8>)>,
    ) -> Result<MediaTable, String> {
        let package = self.package;
        let mut warnings = Vec::new();
        let mut budget = DisplayBudget::default();
        let parts = self
            .parts
            .into_iter()
            .map(|part| {
                let mime_type = media_mime_type(&part.path);
                let data = if part.metafile {
                    match package.read(part.position) {
                        Ok(data) => Some(data),
                        Err(error) => {
                            warnings.push(format!(
                                "Metafile image {} could not be read for display: {error}",
                                part.path
                            ));
                            None
                        }
                    }
                } else if part.transcode {
                    let index = inflated
                        .iter()
                        .position(|(path, _)| path == &part.path)
                        .ok_or_else(|| format!("missing inflated media part {}", part.path))?;
                    Some(inflated.remove(index).1)
                } else {
                    None
                };
                let (mime_type, display) = if let Some(data) = data {
                    let (display, mime_type, warning) =
                        display_form(&data, mime_type, &part.path, &mut budget);
                    warnings.extend(warning);
                    let display = match display {
                        Cow::Owned(display) => display,
                        Cow::Borrowed(_) => data,
                    };
                    (mime_type, Some(display))
                } else {
                    (mime_type, None)
                };
                Ok(MediaPart {
                    path: part.path,
                    position: part.position,
                    size: part.size,
                    image: part.image,
                    mime_type,
                    display,
                })
            })
            .collect::<Result<Vec<_>, String>>()?;
        Ok(MediaTable {
            package,
            parts: parts.into(),
            warnings: warnings.into(),
        })
    }
}

fn budget_exceeded() -> String {
    format!("inflated size exceeds {MAX_TOTAL_UNCOMPRESSED_BYTES} bytes")
}

/// Whether `prefix` starts an image format a media part may hold.
fn is_image(prefix: &[u8]) -> bool {
    prefix.starts_with(b"\x89PNG\r\n\x1a\n")
        || prefix.starts_with(&[0xff, 0xd8, 0xff])
        || prefix.starts_with(b"GIF8")
        || prefix.starts_with(b"BM")
        || (prefix.starts_with(b"RIFF") && prefix.get(8..12) == Some(b"WEBP"))
        || matches!(
            prefix.first_chunk::<4>(),
            Some(b"II\x2a\x00" | b"MM\x00\x2a")
        )
        || prefix.starts_with(b"II\xbc")
        || (prefix.starts_with(&[1, 0, 0, 0]) && prefix.get(40..44) == Some(b" EMF"))
        || prefix.starts_with(&[0xd7, 0xcd, 0xc6, 0x9a])
}

fn is_media_path(path: &str) -> bool {
    path.to_ascii_lowercase().starts_with("word/media/")
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MediaFile {
    pub path: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub filename: Option<String>,
    pub mime_type: String,
    pub base64: String,
    pub data_url: String,
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
    let mut budget = DisplayBudget::default();
    for (path, data) in parts {
        if !is_media_path(path) {
            continue;
        }
        let filename = path.rsplit('/').next().unwrap_or(path).to_owned();
        let (data, mime_type, warning) =
            display_form(data, media_mime_type(path), path, &mut budget);
        warnings.extend(warning);
        let mime_type = mime_type.to_owned();
        let base64 = base64::engine::general_purpose::STANDARD.encode(&data);
        let file = Arc::new(MediaFile {
            path: path.clone(),
            filename: Some(filename),
            mime_type: mime_type.clone(),
            data_url: format!("data:{mime_type};base64,{base64}"),
            base64,
        });
        media.insert(path.clone(), Arc::clone(&file));
        if let Some(normalized) = path.strip_prefix("word/") {
            media.insert(normalized.to_owned(), file);
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
    let normalized = normalize_media_path(target);
    let filename = target.rsplit('/').next().map(str::to_owned);
    if let Some(media) = media {
        for candidate in [
            normalized,
            target.trim_start_matches('/').to_owned(),
            format!("word/{}", target.trim_start_matches('/')),
        ] {
            if let Some(file) = find_case_insensitive(media, &candidate) {
                return ResolvedImageData {
                    src: Some(if file.data_url.is_empty() {
                        file.base64.clone()
                    } else {
                        file.data_url.clone()
                    }),
                    mime_type: Some(file.mime_type.clone()),
                    filename,
                };
            }
        }
    }
    ResolvedImageData {
        src: None,
        mime_type: Some(media_mime_type(target).to_owned()),
        filename,
    }
}

/// Whether [`display_form`] replaces a part whose bytes start with `prefix`.
fn transcodes(prefix: &[u8], mime_type: &str) -> bool {
    #[cfg(feature = "tiff")]
    if is_tiff(prefix) {
        return true;
    }
    metafile_transcode(prefix, mime_type)
}

fn metafile_transcode(prefix: &[u8], mime_type: &str) -> bool {
    #[cfg(feature = "tiff")]
    if is_tiff(prefix) {
        return false;
    }
    #[cfg(feature = "metafile")]
    if ooxml_metafile::is_metafile(prefix)
        || (matches!(mime_type, "image/x-emf" | "image/x-wmf") && !is_browser_image(prefix))
    {
        return true;
    }
    let _ = (prefix, mime_type);
    false
}

/// The copy of a media part the renderer paints: formats browsers cannot
/// decode get a transcode, with a warning when that fails. Save reads the
/// untouched package part, so the original bytes still round-trip.
fn display_form<'a>(
    data: &'a [u8],
    mime_type: &'static str,
    path: &str,
    budget: &mut DisplayBudget,
) -> (Cow<'a, [u8]>, &'static str, Option<String>) {
    #[cfg(feature = "tiff")]
    if is_tiff(data) {
        return tiff_display_form(data, mime_type, path);
    }
    #[cfg(feature = "metafile")]
    if ooxml_metafile::is_metafile(data)
        || (matches!(mime_type, "image/x-emf" | "image/x-wmf") && !is_browser_image(data))
    {
        return metafile_display_form(data, mime_type, path, budget);
    }
    let _ = (path, &budget);
    (Cow::Borrowed(data), mime_type, None)
}

/// One document's cumulative metafile display allowances.
#[derive(Debug)]
#[cfg_attr(not(feature = "metafile"), allow(dead_code))]
struct DisplayBudget {
    metafile_bytes: usize,
    svg_bytes: usize,
    #[cfg(feature = "metafile")]
    replay: ooxml_metafile::ReplayBudget,
}

impl Default for DisplayBudget {
    fn default() -> Self {
        Self {
            metafile_bytes: 64 * 1024 * 1024,
            svg_bytes: 64 * 1024 * 1024,
            #[cfg(feature = "metafile")]
            replay: {
                let picture = ooxml_metafile::ReplayBudget::default();
                ooxml_metafile::ReplayBudget {
                    work: 4 * picture.work,
                    pixels: 4 * picture.pixels,
                }
            },
        }
    }
}

/// An encoding the decoder does not support keeps the TIFF source, so
/// decoders that do handle it still render, and reports why.
#[cfg(feature = "tiff")]
fn tiff_display_form<'a>(
    data: &'a [u8],
    mime_type: &'static str,
    path: &str,
) -> (Cow<'a, [u8]>, &'static str, Option<String>) {
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

/// Largest metafile part transcoded for display.
#[cfg(feature = "metafile")]
const MAX_METAFILE_BYTES: usize = 16 * 1024 * 1024;

/// Browsers have no EMF or WMF decoder, so the display copy is an SVG replay.
/// A metafile the replay refuses shows a neutral placeholder instead of
/// nothing, and ink it drew without is reported.
#[cfg(feature = "metafile")]
fn metafile_display_form<'a>(
    data: &'a [u8],
    mime_type: &str,
    path: &str,
    budget: &mut DisplayBudget,
) -> (Cow<'a, [u8]>, &'static str, Option<String>) {
    let kind = if ooxml_metafile::is_wmf(data)
        || (!ooxml_metafile::is_metafile(data) && mime_type == "image/x-wmf")
    {
        "WMF"
    } else {
        "EMF"
    };
    let placeholder = |why: String| {
        let (width, height) = ooxml_metafile::picture_size(data).unwrap_or((96.0, 96.0));
        (
            Cow::Owned(ooxml_metafile::placeholder_svg(width, height).into_bytes()),
            "image/svg+xml",
            Some(format!(
                "{kind} image {path} could not be converted for display: {why}"
            )),
        )
    };
    if data.len() > MAX_METAFILE_BYTES {
        return placeholder("it exceeds the display size limit".to_owned());
    }
    if data.len() > budget.metafile_bytes {
        return placeholder("the document's pictures exceed the display size limit".to_owned());
    }
    if budget.svg_bytes == 0 {
        return placeholder("the document's pictures exceed the display size limit".to_owned());
    }
    if budget.replay.work == 0 {
        return placeholder("the document's pictures exceed the replay limits".to_owned());
    }
    budget.metafile_bytes -= data.len();
    match ooxml_metafile::to_svg_with_limits(data, &mut budget.replay, &mut budget.svg_bytes) {
        Ok(svg) => {
            let warning = (!svg.omissions.is_empty()).then(|| {
                let omitted: Vec<String> = svg
                    .omissions
                    .iter()
                    .map(|omission| format!("{} ({})", omission.what, omission.count))
                    .collect();
                format!(
                    "{kind} image {path} is displayed without {}",
                    omitted.join(", ")
                )
            });
            (
                Cow::Owned(svg.markup.into_bytes()),
                "image/svg+xml",
                warning,
            )
        }
        Err(refusal) => placeholder(refusal.to_string()),
    }
}

/// Whether `data` starts like a raster format or markup (SVG) browsers
/// decode, whatever its part name says.
#[cfg(feature = "metafile")]
fn is_browser_image(data: &[u8]) -> bool {
    let text = data.strip_prefix(b"\xEF\xBB\xBF").unwrap_or(data);
    text.trim_ascii_start().starts_with(b"<")
        || data.starts_with(b"\x89PNG")
        || data.starts_with(&[0xFF, 0xD8, 0xFF])
        || data.starts_with(b"GIF8")
        || data.starts_with(b"BM")
        || (data.starts_with(b"RIFF") && data.get(8..12) == Some(b"WEBP"))
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

fn normalize_media_path(path: &str) -> String {
    let path = path.trim_start_matches('/');
    if path.starts_with("media/") {
        format!("word/{path}")
    } else if path.starts_with("word/") {
        path.to_owned()
    } else {
        format!("word/{path}")
    }
}

fn find_case_insensitive<'a>(media: &'a MediaMap, path: &str) -> Option<&'a MediaFile> {
    media
        .iter()
        .find(|(candidate, _)| candidate.eq_ignore_ascii_case(path))
        .map(|(_, file)| &**file)
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
    fn alias_keys_share_one_allocation() {
        let parts = vec![("word/media/a.png".to_owned(), vec![1, 2, 3])];
        let media = build_media_map(&parts);
        assert!(Arc::ptr_eq(
            &media["word/media/a.png"],
            &media["media/a.png"]
        ));
    }

    #[cfg(any(feature = "metafile", feature = "tiff"))]
    fn corrupt_image_package(path: &str, prefix: &[u8], broken_deflate: bool) -> (Arc<[u8]>, u64) {
        let mut image = vec![0u8; 128];
        image[..prefix.len()].copy_from_slice(prefix);
        let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml"><w:body><w:p><w:r><w:pict><v:shape id="Picture 1" style="width:10pt;height:10pt"><v:imagedata r:id="rId1"/></v:shape></w:pict></w:r></w:p></w:body></w:document>"#.to_vec();
        let relationships = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{}"/></Relationships>"#,
            path.strip_prefix("word/").unwrap()
        ).into_bytes();
        let total = (image.len() + document.len() + relationships.len()) as u64;
        let mut bytes = ooxml_opc::rezip_parts(&[
            ("word/document.xml".to_owned(), document),
            ("word/_rels/document.xml.rels".to_owned(), relationships),
            (path.to_owned(), image),
        ])
        .unwrap();
        let central = bytes
            .windows(4)
            .enumerate()
            .find_map(|(at, signature)| {
                (signature == b"PK\x01\x02"
                    && bytes.get(at + 46..at + 46 + path.len()) == Some(path.as_bytes()))
                .then_some(at)
            })
            .unwrap();
        let local =
            u32::from_le_bytes(bytes[central + 42..central + 46].try_into().unwrap()) as usize;
        if broken_deflate {
            let name =
                u16::from_le_bytes(bytes[local + 26..local + 28].try_into().unwrap()) as usize;
            let extra =
                u16::from_le_bytes(bytes[local + 28..local + 30].try_into().unwrap()) as usize;
            bytes[local + 30 + name + extra] = 0x07;
        } else {
            bytes[central + 16] ^= 1;
            bytes[local + 14] ^= 1;
        }
        (bytes.into(), total)
    }

    #[cfg(feature = "metafile")]
    #[test]
    fn corrupt_metafile_entries_open_with_lazy_broken_pictures_and_warnings() {
        use crate::s9::{
            S9ParseOptions, media_table_parts_within, parse_docx_s9_wire_with_media_table,
        };

        let mut emf_prefix = vec![0u8; 44];
        emf_prefix[..4].copy_from_slice(&1u32.to_le_bytes());
        emf_prefix[40..44].copy_from_slice(b" EMF");
        for (path, prefix) in [
            ("word/media/image1.emf", emf_prefix.as_slice()),
            ("word/media/image1.wmf", &[0xd7, 0xcd, 0xc6, 0x9a][..]),
        ] {
            for broken_deflate in [false, true] {
                let (bytes, total) = corrupt_image_package(path, prefix, broken_deflate);
                let package = RetainedPackage::new(Arc::clone(&bytes)).unwrap();
                let table = MediaTable::new(package).unwrap();
                assert_eq!(table.len(), 1);
                assert_eq!(table.path(0), Some(path));
                assert_eq!(table.mime_type(0), Some(media_mime_type(path)));
                assert!(table.parts[0].display.is_none());
                assert!(table.keeps_compressed(path));
                assert!(table.bytes(0).is_err());
                assert_eq!(table.resolve("media:0"), None);
                assert_eq!(table.warnings().len(), 1);
                assert!(table.warnings()[0].contains(path));
                assert!(
                    table
                        .check_budget(MAX_TOTAL_UNCOMPRESSED_BYTES - 128)
                        .is_ok()
                );
                assert!(
                    table
                        .check_budget(MAX_TOTAL_UNCOMPRESSED_BYTES - 127)
                        .is_err()
                );
                let (parts, bounded) = media_table_parts_within(&bytes, total).unwrap();
                assert_eq!(parts.len(), 2);
                assert_eq!(bounded.warnings(), table.warnings());
                assert!(media_table_parts_within(&bytes, total - 1).is_err());
                let (wire, _, parsed) = parse_docx_s9_wire_with_media_table(
                    bytes,
                    S9ParseOptions::default(),
                    &crate::ParseLimits::default(),
                )
                .unwrap();
                assert_eq!(wire.document.warnings.as_deref(), Some(parsed.warnings()));
                let json = serde_json::to_string(&wire).unwrap();
                assert!(json.contains(r#""type":"image""#));
                assert!(json.contains(r#""src":"media:0""#));
                assert!(parsed.bytes(0).is_err());
            }
        }
    }

    #[cfg(feature = "tiff")]
    #[test]
    fn corrupt_tiff_entries_still_fail_extraction() {
        for path in ["word/media/image1.tif", "word/media/image1.emf"] {
            let (bytes, _) = corrupt_image_package(path, b"II\x2a\x00", false);
            let package = RetainedPackage::new(Arc::clone(&bytes)).unwrap();
            assert!(MediaTable::new(package).is_err());
            assert!(crate::s9::media_table_parts(&bytes).is_err());
        }
    }

    #[cfg(feature = "metafile")]
    #[test]
    fn metafile_output_exhaustion_skips_later_pictures() {
        let mut data = vec![0u8; 88];
        for (at, value) in [
            (0, 1u32),
            (4, 88),
            (16, 100),
            (20, 100),
            (32, 2540),
            (36, 2540),
            (40, 0x464D_4520),
            (44, 0x0001_0000),
            (48, 132),
            (52, 3),
            (56, 1),
            (72, 96),
            (76, 96),
            (80, 25),
            (84, 25),
        ] {
            data[at..at + 4].copy_from_slice(&value.to_le_bytes());
        }
        data.extend(
            [43u32, 24, 0, 0, 100, 100, 14, 20, 0, 16, 20]
                .into_iter()
                .flat_map(u32::to_le_bytes),
        );
        let expected = ooxml_metafile::to_svg(&data).unwrap();
        let mut budget = DisplayBudget {
            svg_bytes: expected.markup.len() + expected.markup.len() / 2,
            ..DisplayBudget::default()
        };
        let (display, mime, warning) =
            display_form(&data, "image/x-emf", "word/media/first.emf", &mut budget);
        assert_eq!(mime, "image/svg+xml");
        assert_eq!(std::str::from_utf8(&display).unwrap(), expected.markup);
        assert!(warning.is_none());
        assert_eq!(budget.svg_bytes, expected.markup.len() / 2);
        let before = budget.replay;
        let (display, _, warning) =
            display_form(&data, "image/x-emf", "word/media/second.emf", &mut budget);
        assert!(warning.unwrap().contains("display size limit"));
        assert!(
            std::str::from_utf8(&display)
                .unwrap()
                .contains(r##"fill="#f1f3f4""##)
        );
        assert!(budget.replay.work < before.work);
        assert_eq!(budget.svg_bytes, 0);
        let before = (budget.replay, budget.metafile_bytes);
        for _ in 0..4 {
            let (display, _, warning) =
                display_form(&data, "image/x-emf", "word/media/later.emf", &mut budget);
            assert!(warning.unwrap().contains("document's pictures"));
            assert!(
                std::str::from_utf8(&display)
                    .unwrap()
                    .contains(r##"fill="#f1f3f4""##)
            );
            assert_eq!((budget.replay, budget.metafile_bytes), before);
        }
    }

    #[cfg(feature = "metafile")]
    #[test]
    fn metafile_parts_share_the_document_work_budget() {
        let data = include_bytes!("../../ooxml-metafile/tests/fixtures/shapes.emf");
        let mut replay = ooxml_metafile::ReplayBudget::default();
        let before = replay.work;
        ooxml_metafile::to_svg_with_budget(data, &mut replay).unwrap();
        let mut budget = DisplayBudget::default();
        budget.replay.work = 2 * (before - replay.work);
        for index in 0..3 {
            let (display, mime, warning) = display_form(
                data,
                "image/x-emf",
                &format!("word/media/image{index}.emf"),
                &mut budget,
            );
            assert_eq!(mime, "image/svg+xml");
            let svg = std::str::from_utf8(&display).unwrap();
            if index < 2 {
                assert!(warning.is_none());
                assert!(!svg.contains(r##"fill="#f1f3f4""##));
            } else {
                assert!(warning.unwrap().contains("replay limits"));
                assert!(svg.contains(r##"fill="#f1f3f4""##));
            }
        }
    }

    #[cfg(feature = "metafile")]
    #[test]
    fn vector_metafiles_render_after_the_document_pixel_budget_is_spent() {
        let bitmap = include_bytes!("../../ooxml-metafile/tests/fixtures/clip-bitmap.emf");
        let vector = include_bytes!("../../ooxml-metafile/tests/fixtures/shapes.emf");
        let mut replay = ooxml_metafile::ReplayBudget::default();
        let before = replay.pixels;
        let expected = ooxml_metafile::to_svg_with_budget(bitmap, &mut replay).unwrap();
        let mut budget = DisplayBudget::default();
        budget.replay.pixels = before - replay.pixels;
        assert!(budget.replay.pixels > 0);
        let (display, mime, _) =
            display_form(bitmap, "image/x-emf", "word/media/bitmap.emf", &mut budget);
        assert_eq!(mime, "image/svg+xml");
        assert_eq!(std::str::from_utf8(&display).unwrap(), expected.markup);
        assert_eq!(budget.replay.pixels, 0);
        assert!(budget.replay.work > 0);
        let (display, mime, warning) =
            display_form(vector, "image/x-emf", "word/media/vector.emf", &mut budget);
        assert_eq!(mime, "image/svg+xml");
        assert!(warning.is_none());
        let svg = std::str::from_utf8(&display).unwrap();
        assert!(svg.contains("<path"));
        assert!(!svg.contains(r##"fill="#f1f3f4""##));
        assert_eq!(budget.replay.pixels, 0);
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
