//! Embedded media table and image-resolution aliases.

use std::borrow::Cow;
use std::sync::Arc;

use base64::Engine as _;
use indexmap::IndexMap;
use serde::{Deserialize, Serialize};

use crate::relationships::RelationshipMap;

pub type MediaMap = IndexMap<String, Arc<MediaFile>>;

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
        if !path.to_ascii_lowercase().starts_with("word/media/") {
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
    if budget.replay.work == 0 {
        return placeholder("the document's pictures exceed the replay limits".to_owned());
    }
    budget.metafile_bytes -= data.len();
    match ooxml_metafile::to_svg_with_budget(data, &mut budget.replay) {
        Ok(svg) if svg.markup.len() <= budget.svg_bytes => {
            budget.svg_bytes -= svg.markup.len();
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
        Ok(_) => placeholder("the document's pictures exceed the display size limit".to_owned()),
        Err(refusal) => placeholder(refusal.to_string()),
    }
}

/// Whether `data` starts like a raster format browsers decode, whatever its
/// part name says.
#[cfg(feature = "metafile")]
fn is_browser_image(data: &[u8]) -> bool {
    data.starts_with(b"\x89PNG")
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
