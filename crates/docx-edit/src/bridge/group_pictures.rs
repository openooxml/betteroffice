use std::collections::BTreeMap;
use std::sync::Arc;

use docx_parse::media::MediaTable;
use docx_parse::paragraph::{apply_shape_metadata, parse_alternate_content_branch};
use docx_parse::relationships::{RelationshipMap, parse_relationships};
use docx_parse::shape::{
    parse_shape_from_drawing, resolve_shape_fill_pictures, shape_for_rendering,
};
use docx_parse::xml::{ParseBudget, ParseLimits, XmlElement, parse_xml};
use ooxml_opc::{PackageBytes, RetainedPackage};
use serde_json::Value;

use crate::identity::SourceStory;

type StoryShapes = BTreeMap<String, BTreeMap<String, Vec<Value>>>;

#[derive(Clone, Debug, Default)]
pub struct GroupPictureSources {
    shapes: Arc<StoryShapes>,
    media: Option<Arc<MediaTable>>,
}

impl PartialEq for GroupPictureSources {
    fn eq(&self, other: &Self) -> bool {
        self.shapes == other.shapes
            && match (&self.media, &other.media) {
                (Some(left), Some(right)) => Arc::ptr_eq(left, right),
                (None, None) => true,
                _ => false,
            }
    }
}

pub(crate) fn source_shapes(
    bytes: PackageBytes,
    roots: &std::collections::HashMap<String, SourceStory>,
) -> GroupPictureSources {
    let mut sources = BTreeMap::new();
    let Ok(package) = RetainedPackage::from_bytes(bytes) else {
        return GroupPictureSources::default();
    };
    let Ok(table) = MediaTable::new(package.clone()) else {
        return GroupPictureSources::default();
    };
    let media = table.media_map();
    let parts: BTreeMap<_, _> = package
        .parts()
        .enumerate()
        .map(|(index, (path, _))| (path.to_owned(), index))
        .collect();
    let mut by_part: BTreeMap<&str, Vec<_>> = BTreeMap::new();
    for (story, source) in roots {
        by_part
            .entry(source.part_uri.trim_start_matches('/'))
            .or_default()
            .push((story, source));
    }
    let limits = ParseLimits::default();
    for (path, stories) in by_part {
        let mut budget = ParseBudget::new(&limits);
        let Some(bytes) = parts.get(path).and_then(|index| package.read(*index).ok()) else {
            continue;
        };
        let has = |needle: &[u8]| bytes.windows(needle.len()).any(|window| window == needle);
        if !(has(b":pic") || has(b"<pic") || has(b"txbx"))
            || !(has(b"wgp") || has(b"wpc") || has(b"grpSp") || has(b"spTree"))
        {
            continue;
        }
        let Ok(document) = parse_xml(&bytes, path, &mut budget) else {
            continue;
        };
        let Some(root) = document.root() else {
            continue;
        };
        let Some((directory, filename)) = path.rsplit_once('/') else {
            continue;
        };
        let rel_path = format!("{directory}/_rels/{filename}.rels");
        let relationships = parts
            .get(&rel_path)
            .and_then(|index| package.read(*index).ok())
            .and_then(|bytes| parse_relationships(&bytes, &rel_path, &mut budget).ok())
            .unwrap_or_default();
        for (story, source) in stories {
            let story_root = if let Some(id) = &source.item_id {
                let Some(item) = root
                    .child_elements()
                    .find(|item| item.attribute(Some("w"), "id") == Some(id))
                else {
                    continue;
                };
                item
            } else {
                root
            };
            let mut shapes = BTreeMap::new();
            collect(story_root, &relationships, &media, &mut shapes, &mut budget);
            if !shapes.is_empty() {
                sources.insert(story.clone(), shapes);
            }
        }
    }
    GroupPictureSources {
        shapes: Arc::new(sources),
        media: Some(Arc::new(table)),
    }
}

fn collect(
    node: &XmlElement,
    relationships: &RelationshipMap,
    media: &docx_parse::media::MediaMap,
    shapes: &mut BTreeMap<String, Vec<Value>>,
    budget: &mut ParseBudget<'_>,
) {
    if node.name == "w:drawing" {
        if let Some(mut baseline) = parse_shape_from_drawing(node)
            && baseline
                .scene
                .as_ref()
                .and_then(|scene| scene.root.as_ref())
                .is_some_and(has_picture_or_text)
            && let Some(mut projected) = shape_for_rendering(node)
        {
            resolve_shape_fill_pictures(&mut baseline, Some(relationships), Some(media));
            resolve_shape_fill_pictures(&mut projected, Some(relationships), Some(media));
            apply_shape_metadata(&mut baseline, node);
            shapes
                .entry(baseline.id.clone().unwrap_or_default())
                .or_default()
                .push(serde_json::json!({"baseline": baseline, "projected": projected}));
        }
        return;
    }
    if node.local_name() == "AlternateContent" {
        if let Ok(Some((branch, _))) =
            parse_alternate_content_branch(node, Some(relationships), budget, None, false)
        {
            collect(branch, relationships, media, shapes, budget);
        }
        return;
    }
    for child in node.child_elements() {
        collect(child, relationships, media, shapes, budget);
    }
}

fn has_picture_or_text(node: &docx_parse::shape::DrawingSceneNode) -> bool {
    node.kind.as_deref() == Some("picture")
        || node
            .shape
            .as_ref()
            .is_some_and(|shape| shape.text_body.is_some())
        || node
            .children
            .as_ref()
            .is_some_and(|children| children.iter().any(has_picture_or_text))
}

fn same_model(left: Option<&Value>, right: Option<&Value>) -> bool {
    match (left, right) {
        (Some(Value::Number(left)), Some(Value::Number(right))) => left
            .as_f64()
            .zip(right.as_f64())
            .is_some_and(|(left, right)| {
                left == right
                    || (left - right).abs() <= 2.0 * f64::EPSILON * left.abs().max(right.abs())
            }),
        (Some(Value::Array(left)), Some(Value::Array(right))) => {
            left.len() == right.len()
                && left
                    .iter()
                    .zip(right)
                    .all(|(left, right)| same_model(Some(left), Some(right)))
        }
        (Some(Value::Object(left)), Some(Value::Object(right))) => {
            left.len() == right.len()
                && left.iter().all(|(key, value)| {
                    if key == "src" && left.get("type").and_then(Value::as_str) == Some("image") {
                        right.contains_key(key)
                    } else if key == "textBody" {
                        same_text_body(value, right.get(key))
                    } else {
                        same_model(Some(value), right.get(key))
                    }
                })
        }
        _ => left == right,
    }
}

fn same_text_body(left: &Value, right: Option<&Value>) -> bool {
    let (Some(left), Some(right)) = (left.as_object(), right.and_then(Value::as_object)) else {
        return Some(left) == right;
    };
    left.keys()
        .filter(|key| key.as_str() != "content")
        .eq(right.keys().filter(|key| key.as_str() != "content"))
        && left
            .iter()
            .all(|(key, value)| key == "content" || same_model(Some(value), right.get(key)))
}

type TextBodyKey = (Option<String>, usize);

fn text_bodies(value: &Value, bodies: &mut BTreeMap<TextBodyKey, Value>) {
    if let Some(children) = value.as_array() {
        for child in children {
            text_bodies(child, bodies);
        }
    } else if value.is_object() {
        if let Some(body) = value.get("textBody") {
            let key = (
                value.get("id").and_then(Value::as_str).map(str::to_owned),
                bodies.len(),
            );
            bodies.insert(key, body.clone());
        }
        if let Some(children) = value.get("children") {
            text_bodies(children, bodies);
        }
    }
}

fn apply_text_bodies(
    value: &mut Value,
    bodies: &BTreeMap<TextBodyKey, Value>,
    ordinal: &mut usize,
) {
    if let Some(children) = value.as_array_mut() {
        for child in children {
            apply_text_bodies(child, bodies, ordinal);
        }
    } else if value.is_object() {
        if value.get("textBody").is_some() {
            let key = (
                value.get("id").and_then(Value::as_str).map(str::to_owned),
                *ordinal,
            );
            if let Some(body) = bodies.get(&key) {
                value["textBody"] = body.clone();
            }
            *ordinal += 1;
        }
        if let Some(children) = value.get_mut("children") {
            apply_text_bodies(children, bodies, ordinal);
        }
    }
}

pub(super) fn project(
    shape: &Value,
    story: &str,
    sources: &GroupPictureSources,
    media_tokens: bool,
) -> Option<Value> {
    let (_, shapes) = sources
        .shapes
        .iter()
        .filter(|(root, _)| {
            story == root.as_str()
                || story
                    .strip_prefix(root.as_str())
                    .is_some_and(|suffix| suffix.starts_with(':'))
        })
        .max_by_key(|(root, _)| root.len())?;
    let source = shapes
        .get(shape.get("id").and_then(Value::as_str).unwrap_or_default())?
        .iter()
        .find(|source| {
            let baseline = &source["baseline"];
            ["scene", "children", "size", "transform"]
                .iter()
                .all(|key| same_model(shape.get(*key), baseline.get(*key)))
        })?;
    let mut projected = shape.clone();
    projected["children"] = source["projected"]["children"].clone();
    let mut bodies = BTreeMap::new();
    text_bodies(&shape["children"], &mut bodies);
    let mut targets = BTreeMap::new();
    text_bodies(&projected["children"], &mut targets);
    if !bodies.keys().eq(targets.keys()) {
        return None;
    }
    apply_text_bodies(&mut projected["children"], &bodies, &mut 0);
    if !media_tokens && let Some(table) = &sources.media {
        resolve_data_urls(&mut projected["children"], table);
    }
    Some(projected)
}

pub(super) fn raw_projection(shape: &Value) -> Value {
    let mut shape = shape.clone();
    super::group_transforms::project(&mut shape);
    shape
}

fn resolve_data_urls(value: &mut Value, table: &MediaTable) {
    match value {
        Value::Object(object) => {
            if object.get("type").and_then(Value::as_str) == Some("image")
                && let Some(url) = object
                    .get("src")
                    .and_then(Value::as_str)
                    .and_then(|src| table.resolve(src))
            {
                object.insert("src".into(), Value::String(url));
            }
            for child in object.values_mut() {
                resolve_data_urls(child, table);
            }
        }
        Value::Array(array) => {
            for child in array {
                resolve_data_urls(child, table);
            }
        }
        _ => {}
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
    use crate::{EditingDoc, SegmentContent};
    use serde_json::json;

    fn package() -> Vec<u8> {
        ooxml_opc::rezip_parts(&[
            ("[Content_Types].xml".into(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
            ("_rels/.rels".into(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
            ("word/_rels/document.xml.rels".into(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/picture.png"/></Relationships>"#.to_vec()),
            ("word/document.xml".into(), br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p w14:paraId="11111111"><w:r><w:drawing><wp:inline><wp:extent cx="1000" cy="500"/><wp:docPr id="1" name="Group"/><a:graphic><a:graphicData><wpg:wgp><pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="Photo"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="100" cy="50"/></a:xfrm></pic:spPr></pic:pic></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#.to_vec()),
            ("word/media/picture.png".into(), vec![1, 2, 3, 4]),
        ]).unwrap()
    }

    fn rewrite_document(rewrite: impl FnOnce(&str) -> String) -> Vec<u8> {
        let mut parts = ooxml_opc::unzip_parts(&package()).unwrap();
        let (_, bytes) = parts
            .iter_mut()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap();
        *bytes = rewrite(std::str::from_utf8(bytes).unwrap()).into_bytes();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    fn seeded_shapes(doc: &EditingDoc) -> Vec<Value> {
        doc.story_segments("body")
            .unwrap()
            .iter()
            .filter_map(|segment| match &segment.content {
                SegmentContent::OtherEmbed { kind, payload } if kind == "shape" => {
                    let yrs::Any::String(shape) = payload.get("shapeJson")? else {
                        return None;
                    };
                    Some(serde_json::from_str(shape).unwrap())
                }
                _ => None,
            })
            .collect()
    }

    #[test]
    fn grouped_pictures_follow_render_env_media_mode() {
        let doc = EditingDoc::new(1);
        crate::seed::seed_from_docx(&doc, &package()).unwrap();
        for media_tokens in [false, true] {
            let env = RenderEnv {
                media_tokens,
                ..RenderEnv::default()
            };
            let blocks = yrs_doc_to_layout_blocks(&doc, "body", &env).unwrap();
            let blocks = serde_json::to_value(blocks).unwrap();
            let src = &blocks[0]["runs"][0]["inlineShape"]["children"][0]["fill"]["pictureSrc"];
            assert_eq!(
                src,
                if media_tokens {
                    "media:0"
                } else {
                    "data:image/png;base64,AQIDBA=="
                }
            );
        }
    }

    #[test]
    fn grouped_picture_baselines_include_seeded_metadata() {
        for metadata in [
            r#"<wp:docPr id="1" name="Group"><a:extLst><a:ext uri="decorative"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst></wp:docPr>"#,
            r#"<wp:docPr id="1" name="Group"/><a:cNvPr title="Title" descr="Description" hidden="1"/>"#,
        ] {
            let bytes =
                rewrite_document(|xml| xml.replace(r#"<wp:docPr id="1" name="Group"/>"#, metadata));
            let doc = EditingDoc::new(1);
            crate::seed::seed_from_docx(&doc, &bytes).unwrap();
            let source = doc.source_index().unwrap().group_picture_sources();
            let shape = seeded_shapes(&doc).remove(0);
            assert!(
                project(&shape, "body", &source, false).is_some(),
                "{metadata}"
            );
        }
    }

    #[test]
    fn alternate_content_projects_only_the_selected_branch() {
        for choice in [true, false] {
            let bytes = rewrite_document(|xml| {
                let start = xml.find("<w:drawing>").unwrap();
                let end = xml.find("</w:drawing>").unwrap() + "</w:drawing>".len();
                let drawing = &xml[start..end];
                let fallback = drawing.replace("Photo", "FallbackPhoto");
                let selected = if choice { drawing } else { "<w:drawing/>" };
                xml.replace(
                    drawing,
                    &format!(
                        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wpg">{selected}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#
                    ),
                )
            });
            let doc = EditingDoc::new(1);
            crate::seed::seed_from_docx(&doc, &bytes).unwrap();
            let shape = seeded_shapes(&doc).remove(0);
            let source = doc.source_index().unwrap().group_picture_sources();
            assert!(
                project(&shape, "body", &source, false).is_some(),
                "{choice}"
            );
        }
    }

    #[test]
    fn duplicate_docpr_ids_match_each_group_model() {
        let bytes = rewrite_document(|xml| {
            let start = xml.find("<w:p ").unwrap();
            let end = xml.find("</w:p>").unwrap() + "</w:p>".len();
            let paragraph = &xml[start..end];
            let second = paragraph
                .replace("11111111", "22222222")
                .replace("Photo", "SecondPhoto")
                .replace(r#"x="10""#, r#"x="30""#);
            xml.replace(paragraph, &format!("{paragraph}{second}"))
        });
        let doc = EditingDoc::new(1);
        crate::seed::seed_from_docx(&doc, &bytes).unwrap();
        let source = doc.source_index().unwrap().group_picture_sources();
        let shapes = seeded_shapes(&doc);
        assert_eq!(shapes.len(), 2);
        for (shape, x) in shapes.iter().zip([10, 30]) {
            let projected = project(shape, "body", &source, false).unwrap();
            assert_eq!(
                projected["children"][0]["offset"]["x"].as_f64(),
                Some(f64::from(x))
            );
        }
    }

    #[test]
    fn shared_note_parts_do_not_exhaust_later_story_budgets() {
        use crate::identity::SourceStoryKind;
        let mut parts = ooxml_opc::unzip_parts(&package()).unwrap();
        let document = String::from_utf8(
            parts
                .iter()
                .find(|(path, _)| path == "word/document.xml")
                .unwrap()
                .1
                .clone(),
        )
        .unwrap();
        let paragraph = &document
            [document.find("<w:p ").unwrap()..document.find("</w:p>").unwrap() + "</w:p>".len()];
        let relationships = parts
            .iter()
            .find(|(path, _)| path == "word/_rels/document.xml.rels")
            .unwrap()
            .1
            .clone();
        let mut roots = std::collections::HashMap::from([(
            "body".to_owned(),
            SourceStory {
                part_uri: "/word/document.xml".into(),
                kind: SourceStoryKind::Body,
                item_id: None,
            },
        )]);
        for (prefix, name, kind) in [
            ("fn", "footnote", SourceStoryKind::Footnote),
            ("en", "endnote", SourceStoryKind::Endnote),
        ] {
            let mut notes = format!("<w:{name}s>{}", " ".repeat(600 * 1024));
            for id in 1..=140 {
                notes.push_str(&format!(r#"<w:{name} w:id="{id}">{paragraph}</w:{name}>"#));
                roots.insert(
                    format!("{prefix}:{id}"),
                    SourceStory {
                        part_uri: format!("/word/{name}s.xml"),
                        kind,
                        item_id: Some(id.to_string()),
                    },
                );
            }
            notes.push_str(&format!("</w:{name}s>"));
            parts.push((format!("word/{name}s.xml"), notes.into_bytes()));
            parts.push((
                format!("word/_rels/{name}s.xml.rels"),
                relationships.clone(),
            ));
        }
        let sources = source_shapes(ooxml_opc::rezip_parts(&parts).unwrap().into(), &roots);
        assert_eq!(sources.shapes.len(), roots.len());
        for story in roots.keys() {
            assert!(
                project(&main_shape_json(), story, &sources, false).is_some(),
                "{story}"
            );
        }
    }

    #[test]
    fn edited_root_size_and_transform_reject_stale_projection() {
        let doc = EditingDoc::new(1);
        crate::seed::seed_from_docx(&doc, &package()).unwrap();
        let source = doc.source_index().unwrap().group_picture_sources();
        let baseline = seeded_shapes(&doc).remove(0);
        assert!(project(&baseline, "body", &source, false).is_some());
        for (field, edited) in [
            ("size", json!({"width": 2000, "height": 500})),
            ("transform", json!({"rotation": 90})),
        ] {
            let mut shape = baseline.clone();
            shape[field] = edited;
            assert!(project(&shape, "body", &source, false).is_none(), "{field}");
        }
    }

    const MAIN_SHAPE_JSON: &str = r#"{"type":"shape","shapeType":"rect","id":"1","name":"Group","size":{"width":1000,"height":500},"offset":{"x":0,"y":0},"transform":{"rotation":0},"children":[],"scene":{"version":1,"root":{"kind":"group","children":[{"kind":"picture","id":"2","name":"Photo","relationshipId":"rId7"}],"transform":{"offsetX":0,"offsetY":0,"rotation":0}}}}"#;

    fn main_shape_json() -> Value {
        serde_json::from_str(MAIN_SHAPE_JSON).unwrap()
    }

    #[test]
    fn restored_main_group_snapshot_renders_without_changing_save_fingerprint() {
        let bytes = package();
        let seeded = EditingDoc::new(1);
        crate::seed::seed_from_docx(&seeded, &bytes).unwrap();
        let segments = seeded.story_segments("body").unwrap();
        let payload = segments
            .iter()
            .find_map(|segment| match &segment.content {
                SegmentContent::OtherEmbed { kind, payload } if kind == "shape" => Some(payload),
                _ => None,
            })
            .unwrap();
        let yrs::Any::String(shape_json) = payload.get("shapeJson").unwrap() else {
            panic!("shape JSON");
        };
        assert_eq!(shape_json.as_ref(), MAIN_SHAPE_JSON);
        let restored = EditingDoc::new(2);
        restored
            .apply_update_v1(&seeded.encode_state_as_update_v1())
            .unwrap();
        restored.retain_source_docx(bytes);
        let before = restored.encode_state_as_update_v1();
        let blocks = yrs_doc_to_layout_blocks(&restored, "body", &RenderEnv::default()).unwrap();
        let blocks = serde_json::to_value(blocks).unwrap();
        let picture = &blocks[0]["runs"][0]["inlineShape"]["children"][0];
        assert_eq!(
            picture["fill"]["pictureSrc"],
            "data:image/png;base64,AQIDBA=="
        );
        assert_eq!(picture["fill"]["pictureRelId"], "rId7");
        assert_eq!(restored.encode_state_as_update_v1(), before);
        let plan = restored.paragraph_save_plan();
        assert!(
            plan.patched_parts
                .iter()
                .any(|(part, patches)| part == "word/document.xml" && patches.is_empty())
        );
        assert!(plan.spliced_parts.iter().all(|part| part.changed.is_empty()
            && part.inserted.is_empty()
            && part.removed.is_empty()));
    }

    #[test]
    fn projections_are_scoped_to_source_story_and_do_not_override_edited_children() {
        let baseline = main_shape_json();
        let mut first = baseline.clone();
        first["children"] = json!([{"fillPaint": {"picture": {"src": "media:0"}}}]);
        let mut second = first.clone();
        second["children"][0]["fillPaint"]["picture"]["src"] = json!("media:1");
        let sources = GroupPictureSources {
            shapes: Arc::new(BTreeMap::from([
                (
                    "body".into(),
                    BTreeMap::from([(
                        "1".into(),
                        vec![json!({"baseline": baseline, "projected": first})],
                    )]),
                ),
                (
                    "header:rId7".into(),
                    BTreeMap::from([(
                        "1".into(),
                        vec![json!({"baseline": baseline, "projected": second})],
                    )]),
                ),
            ])),
            ..GroupPictureSources::default()
        };
        assert_eq!(
            project(&baseline, "body:t0:r0c0", &sources, true).unwrap()["children"][0]["fillPaint"]
                ["picture"]["src"],
            "media:0"
        );
        assert_eq!(
            project(&baseline, "header:rId7", &sources, true).unwrap()["children"][0]["fillPaint"]
                ["picture"]["src"],
            "media:1"
        );
        let mut edited = baseline;
        edited["children"] = json!([{"shapeType": "ellipse"}]);
        assert!(project(&edited, "body", &sources, true).is_none());
    }
}
