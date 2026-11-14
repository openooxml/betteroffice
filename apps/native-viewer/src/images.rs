use std::cell::RefCell;
use std::collections::HashMap;

use anyhow::Result;
use base64::Engine as _;
use docx_parse::{
    RelationshipTarget, is_footer_relationship, is_header_relationship, is_image_relationship,
    parse_docx_relationship_parts, resolve_relationship_target,
};
use docx_raster::{ImageMap, ImageScope, MAX_DATA_URL_BYTES, NoteKind, scoped_image_key};
use vello::peniko::ImageData;

pub struct ImageRegistry {
    pub raw: ImageMap,
    decoded: HashMap<String, ImageData>,
    data_urls: RefCell<HashMap<String, Option<ImageData>>>,
}

impl ImageRegistry {
    pub fn load(docx: &[u8]) -> Result<Self> {
        let parts = ooxml_opc::unzip_parts(docx).map_err(anyhow::Error::msg)?;
        let package = parts
            .into_iter()
            .map(|(path, bytes)| (path.to_lowercase(), bytes))
            .collect::<HashMap<_, _>>();
        let relationships = parse_docx_relationship_parts(docx)?;
        let body_part = relationships.relationship_parts.iter().find(|part| {
            part.path
                .eq_ignore_ascii_case("word/_rels/document.xml.rels")
        });
        let mut header_footer_ids = HashMap::new();
        if let Some(body_part) = body_part {
            for (_, relationship) in &body_part.relationships {
                if !(is_header_relationship(relationship) || is_footer_relationship(relationship)) {
                    continue;
                }
                if let RelationshipTarget::Internal(target) =
                    resolve_relationship_target(&body_part.path, relationship)?
                {
                    header_footer_ids.insert(target.to_lowercase(), relationship.id.clone());
                }
            }
        }

        let mut raw = HashMap::new();
        let mut decoded = HashMap::new();
        for part in &relationships.relationship_parts {
            let Some(owner) = relationship_owner(&part.path) else {
                continue;
            };
            let owner_lower = owner.to_lowercase();
            for (_, relationship) in &part.relationships {
                if !is_image_relationship(relationship) {
                    continue;
                }
                let RelationshipTarget::Internal(target) =
                    resolve_relationship_target(&part.path, relationship)?
                else {
                    continue;
                };
                let Some(bytes) = package.get(&target.to_lowercase()) else {
                    continue;
                };
                let key = if owner_lower == "word/document.xml" {
                    scoped_image_key(ImageScope::Body, &relationship.id)
                } else if owner_lower == "word/footnotes.xml" {
                    scoped_image_key(ImageScope::Notes(NoteKind::Footnote), &relationship.id)
                } else if owner_lower == "word/endnotes.xml" {
                    scoped_image_key(ImageScope::Notes(NoteKind::Endnote), &relationship.id)
                } else if let Some(r_id) = header_footer_ids.get(&owner_lower) {
                    scoped_image_key(ImageScope::HeaderFooter(r_id), &relationship.id)
                } else {
                    continue;
                };
                raw.insert(key.clone(), bytes.clone());
                if let Some(image) = decode_image(bytes) {
                    decoded.insert(key, image);
                }
            }
        }
        Ok(Self {
            raw,
            decoded,
            data_urls: RefCell::new(HashMap::new()),
        })
    }

    pub fn get(&self, scope: ImageScope<'_>, rel_id: &str) -> Option<ImageData> {
        if rel_id.starts_with("data:") {
            if let Some(image) = self.data_urls.borrow().get(rel_id) {
                return image.clone();
            }
            let image = decode_data_url(rel_id);
            self.data_urls
                .borrow_mut()
                .insert(rel_id.to_owned(), image.clone());
            return image;
        }
        let key = scoped_image_key(scope, rel_id);
        self.decoded.get(&key).cloned()
    }
}

fn decode_data_url(source: &str) -> Option<ImageData> {
    let (metadata, payload) = source.split_once(',')?;
    if !metadata.ends_with(";base64") {
        return None;
    }
    let declared = payload.len() as u64 / 4 * 3;
    if declared > MAX_DATA_URL_BYTES {
        return None;
    }
    let bytes = base64::engine::general_purpose::STANDARD
        .decode(payload)
        .ok()?;
    decode_image(&bytes)
}

fn decode_image(bytes: &[u8]) -> Option<ImageData> {
    let mut unlimited = u64::MAX;
    crate::picture_decode::decode_bounded(bytes, true, &mut unlimited).ok()
}

fn relationship_owner(path: &str) -> Option<String> {
    let normalized = path.replace('\\', "/");
    if let Some((directory, file)) = normalized.rsplit_once("/_rels/") {
        return Some(format!("{directory}/{}", file.strip_suffix(".rels")?));
    }
    let file = normalized.strip_prefix("_rels/")?;
    Some(file.strip_suffix(".rels")?.to_owned())
}

#[cfg(test)]
mod tests {

    /// A 2x1 TIFF whose strip is a real JPEG of those pixels.
    fn jpeg_tiff() -> Vec<u8> {
        let mut jpeg = Vec::new();
        image::codecs::jpeg::JpegEncoder::new(&mut jpeg)
            .encode(
                &[255, 0, 0, 0, 255, 0],
                2,
                1,
                image::ExtendedColorType::Rgb8,
            )
            .unwrap();
        let entries: [(u16, u16, u32, u32); 9] = [
            (256, 3, 1, 2),
            (257, 3, 1, 1),
            (258, 3, 3, 122),
            (259, 3, 1, 7),
            (262, 3, 1, 2),
            (273, 4, 1, 128),
            (277, 3, 1, 3),
            (278, 3, 1, 1),
            (279, 4, 1, jpeg.len() as u32),
        ];
        let mut tiff = b"II*\0\x08\0\0\0".to_vec();
        tiff.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        for (tag, kind, count, value) in entries {
            tiff.extend_from_slice(&tag.to_le_bytes());
            tiff.extend_from_slice(&kind.to_le_bytes());
            tiff.extend_from_slice(&count.to_le_bytes());
            tiff.extend_from_slice(&value.to_le_bytes());
        }
        tiff.extend_from_slice(&0u32.to_le_bytes());
        for _ in 0..3 {
            tiff.extend_from_slice(&8u16.to_le_bytes());
        }
        tiff.extend_from_slice(&jpeg);
        tiff
    }

    #[test]
    fn a_tiff_whose_decode_its_size_does_not_bound_is_not_loaded() {
        let tiff = jpeg_tiff();
        assert!(image::load_from_memory(&tiff).is_ok());
        assert!(decode_image(&tiff).is_none());
    }

    use super::*;
    use crate::test_fixtures;

    #[test]
    fn resolves_relationship_owner_parts() {
        assert_eq!(
            relationship_owner("word/_rels/header1.xml.rels").as_deref(),
            Some("word/header1.xml")
        );
        assert_eq!(relationship_owner("_rels/.rels").as_deref(), Some(""));
    }

    #[test]
    fn decodes_and_caches_data_urls_with_raster_error_rules() {
        assert!(decode_data_url("data:image/png,AAAA").is_none());
        assert!(decode_data_url("data:image/png;base64,%").is_none());
        assert!(decode_data_url("data:image/png;base64,bm90IGFuIGltYWdl").is_none());

        let package = test_fixtures::image_docx(None);
        let png = test_fixtures::part(&package, "word/media/image1.png");
        let source = format!(
            "data:image/png;base64,{}",
            base64::engine::general_purpose::STANDARD.encode(png)
        );
        let registry = ImageRegistry::load(&package).unwrap();
        let first = registry.get(ImageScope::Body, &source).unwrap();
        let second = registry.get(ImageScope::Body, &source).unwrap();
        assert_eq!((first.width, first.height), (160, 80));
        assert_eq!((second.width, second.height), (160, 80));
        assert_eq!(registry.data_urls.borrow().len(), 1);
    }
}
