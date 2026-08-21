use std::collections::{BTreeMap, BTreeSet, VecDeque};
use std::fs;
use std::path::Path;
use std::sync::{Arc, Mutex};

use anyhow::{Context, Result, bail};
use betteroffice_pptx::{
    CaretAnchor, DeckSnapshot, EditCtx, HitTestResult, ModelTextRun, ParagraphProperties,
    ParagraphSnapshot, PptxPackage, Presentation, Primitive, RenderedSlide, RunProperties,
    ShapeNode, ShapeSnapshot, StorySnapshot, TextBody, TextStyle, Transform, UpdateOrigin,
    UpdateSubscription,
};
use sha2::{Digest, Sha256};
use vello::kurbo::{Affine, Point};

use crate::editing::{DeleteDirection, MoveDirection};
use crate::pptx_scene::{PptxSceneResources, PptxSlideSummary};
use crate::scene_shared::PageScene;

const EDIT_AUTHOR: &str = "native-viewer";
const MAX_FIDELITY_DIFF_CELLS: usize = 4_000_000;

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct PptxTextHit {
    pub slide_index: usize,
    pub shape_id: String,
    pub story_id: String,
    pub position: u32,
    line_index: usize,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum PptxHit {
    Text(PptxTextHit),
    Other,
}

#[derive(Clone, Copy, Debug)]
pub struct PptxCaretGeometry {
    pub page_index: usize,
    pub x: f64,
    pub y: f64,
    pub height: f64,
    pub transform: Affine,
}

pub struct PptxEditChange {
    pub page_index: usize,
    pub page: PageScene,
}

pub enum PptxRemoteChange {
    Slides(Vec<PptxEditChange>),
    All(Vec<PageScene>),
}

#[derive(Clone, Debug)]
struct PptxCaret {
    slide_index: usize,
    shape_id: String,
    story_id: String,
    position: u32,
    line_index: usize,
}

pub struct PptxEditor {
    _local_update_subscription: Option<UpdateSubscription>,
    local_updates: Arc<Mutex<VecDeque<Vec<u8>>>>,
    presentation: Presentation,
    resources: PptxSceneResources,
    rendered: Vec<RenderedSlide>,
    summaries: Vec<PptxSlideSummary>,
    baseline: DeckSnapshot,
    saved_snapshot: DeckSnapshot,
    source: Vec<u8>,
    source_fingerprint: Option<[u8; 32]>,
    caret: Option<PptxCaret>,
    vertical_goal_x: Option<f32>,
    dirty: bool,
    local_dirty: bool,
    remote_dirty: bool,
    max_texture_dimension_2d: u32,
    #[cfg(test)]
    relayout_count: usize,
}

#[derive(Clone, Debug)]
struct StoryChange {
    baseline: StorySnapshot,
    current: StorySnapshot,
}

type EditedSlides = BTreeMap<String, BTreeMap<u32, Vec<StoryChange>>>;

impl PptxEditor {
    pub fn open(source: Vec<u8>, max_texture_dimension_2d: u32) -> Result<(Self, Vec<PageScene>)> {
        Self::open_internal(source, max_texture_dimension_2d, None)
    }

    pub fn open_collaborative(
        source: Vec<u8>,
        max_texture_dimension_2d: u32,
        client_id: u64,
    ) -> Result<(Self, Vec<PageScene>)> {
        Self::open_internal(source, max_texture_dimension_2d, Some(client_id))
    }

    fn open_internal(
        source: Vec<u8>,
        max_texture_dimension_2d: u32,
        collaboration_client_id: Option<u64>,
    ) -> Result<(Self, Vec<PageScene>)> {
        let mut presentation = match collaboration_client_id {
            Some(client_id) => Presentation::open_collaborative(&source, client_id),
            None => Presentation::open(&source),
        }?;
        let slide_count = presentation.slides().len();
        let model_slide_count = presentation.model().slides.len();
        if slide_count != model_slide_count {
            bail!(
                "PPTX model has {model_slide_count} slide references but {slide_count} parsed slides"
            );
        }
        let baseline = presentation.snapshot().context("snapshot opened PPTX")?;
        let resources = PptxSceneResources::new(&mut presentation)?;
        let mut rendered = Vec::with_capacity(slide_count);
        let mut summaries = Vec::with_capacity(slide_count);
        let mut pages = Vec::with_capacity(slide_count);
        for slide_index in 0..slide_count {
            let slide = presentation
                .render_slide(slide_index)
                .with_context(|| format!("layout PPTX slide {}", slide_index + 1))?;
            let (page, summary) =
                resources.translate(&slide.display_list, max_texture_dimension_2d)?;
            rendered.push(slide);
            summaries.push(summary);
            pages.push(page);
        }
        let source_fingerprint = collaboration_client_id.map(|_| Sha256::digest(&source).into());
        let local_updates = Arc::new(Mutex::new(VecDeque::new()));
        let local_update_subscription = if collaboration_client_id.is_some() {
            let observed = Arc::clone(&local_updates);
            Some(presentation.observe_update_v1(move |event| {
                if event.origin == UpdateOrigin::Local {
                    observed
                        .lock()
                        .unwrap_or_else(|poisoned| poisoned.into_inner())
                        .push_back(event.update);
                }
            })?)
        } else {
            None
        };
        Ok((
            Self {
                _local_update_subscription: local_update_subscription,
                local_updates,
                presentation,
                resources,
                rendered,
                summaries,
                saved_snapshot: baseline.clone(),
                baseline,
                source,
                source_fingerprint,
                caret: None,
                vertical_goal_x: None,
                dirty: false,
                local_dirty: false,
                remote_dirty: false,
                max_texture_dimension_2d,
                #[cfg(test)]
                relayout_count: 0,
            },
            pages,
        ))
    }

    pub fn slide_count(&self) -> usize {
        self.rendered.len()
    }

    pub fn summaries(&self) -> &[PptxSlideSummary] {
        &self.summaries
    }

    pub fn font_faces(&self) -> &[String] {
        &self.resources.font_faces
    }

    pub fn can_undo(&self) -> bool {
        self.presentation.can_undo()
    }

    pub fn can_redo(&self) -> bool {
        self.presentation.can_redo()
    }

    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    pub fn is_remote_only_dirty(&self) -> bool {
        self.dirty && self.remote_dirty && !self.local_dirty
    }

    pub fn state_vector(&self) -> Result<Vec<u8>> {
        self.source_fingerprint
            .context("collaboration is unavailable for this presentation")?;
        Ok(self.presentation.encode_state_vector_v1())
    }

    pub fn encode_diff(&self, state_vector: &[u8]) -> Result<Vec<u8>> {
        self.source_fingerprint
            .context("collaboration is unavailable for this presentation")?;
        Ok(self.presentation.encode_diff_v1(state_vector)?)
    }

    pub fn drain_local_updates(&self) -> Vec<Vec<u8>> {
        self.local_updates
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .drain(..)
            .collect()
    }

    pub fn source_fingerprint(&self) -> Result<[u8; 32]> {
        self.source_fingerprint
            .context("collaboration is unavailable for this presentation")
    }

    pub fn canonical_checksum(&self) -> Result<[u8; 32]> {
        let mut checksum = Sha256::new();
        checksum.update(b"betteroffice-native-pptx-deck-v1\0");
        checksum.update(self.presentation.save()?);
        Ok(checksum.finalize().into())
    }

    pub fn apply_remote_update(&mut self, update: &[u8]) -> Result<Option<PptxRemoteChange>> {
        self.source_fingerprint
            .context("collaboration is unavailable for this presentation")?;
        let before = self.presentation.snapshot()?;
        let caret = self.caret.clone();
        let caret_anchor = caret
            .as_ref()
            .map(|caret| {
                self.presentation
                    .anchor_caret(&caret.story_id, caret.position)
            })
            .transpose()?;
        let preferred_y = caret.as_ref().and_then(|caret| {
            caret_placement(
                self.rendered.get(caret.slide_index)?,
                caret,
                Some(caret.line_index),
            )
            .map(|placement| placement.y)
        });
        let after = self.presentation.apply_update_v1(update)?;
        if before == after {
            return Ok(None);
        }
        let rebuilt = before.slides.len() != after.slides.len()
            || before
                .slides
                .iter()
                .zip(&after.slides)
                .any(|(left, right)| left.id != right.id);
        let change = if rebuilt {
            PptxRemoteChange::All(self.relayout_all()?)
        } else {
            let mut changes = Vec::new();
            for (slide_index, (left, right)) in before.slides.iter().zip(&after.slides).enumerate()
            {
                if left != right {
                    changes.push(self.relayout_slide(slide_index)?);
                }
            }
            PptxRemoteChange::Slides(changes)
        };
        self.restore_caret(caret, caret_anchor, preferred_y, &after);
        self.remote_dirty = true;
        self.vertical_goal_x = None;
        self.update_dirty()?;
        Ok(Some(change))
    }

    pub(crate) fn snapshot(&self) -> Result<DeckSnapshot> {
        Ok(self.presentation.snapshot()?)
    }

    #[cfg(test)]
    pub fn relayout_count(&self) -> usize {
        self.relayout_count
    }

    pub fn has_caret(&self) -> bool {
        self.caret.is_some()
    }

    #[cfg(test)]
    pub fn caret_position(&self) -> Option<u32> {
        self.caret.as_ref().map(|caret| caret.position)
    }

    #[cfg(test)]
    pub fn caret_story_id(&self) -> Option<&str> {
        self.caret.as_ref().map(|caret| caret.story_id.as_str())
    }

    #[cfg(test)]
    pub fn story(&self, story_id: &str) -> Result<StorySnapshot> {
        Ok(self.presentation.story(story_id)?)
    }

    #[cfg(test)]
    pub fn rendered_slide(&self, slide_index: usize) -> Option<&RenderedSlide> {
        self.rendered.get(slide_index)
    }

    pub fn hit_test(&self, slide_index: usize, x: f32, y: f32) -> Option<PptxHit> {
        let rendered = self.rendered.get(slide_index)?;
        match rendered.hit_test(x, y) {
            Some(HitTestResult::Text {
                shape_id,
                story_id,
                position,
            }) if self.presentation.story(&story_id).is_ok() => {
                let text_box = text_box(rendered, &shape_id, &story_id)?;
                let point = text_box.transform.inverse() * Point::new(f64::from(x), f64::from(y));
                let line_index = nearest_line(text_box.lines, point.y as f32)?;
                Some(PptxHit::Text(PptxTextHit {
                    slide_index,
                    shape_id,
                    story_id,
                    position,
                    line_index,
                }))
            }
            Some(_) => Some(PptxHit::Other),
            None => None,
        }
    }

    pub fn select_hit(&mut self, hit: PptxTextHit) -> bool {
        if self.rendered.get(hit.slide_index).is_none() {
            return false;
        }
        self.caret = Some(PptxCaret {
            slide_index: hit.slide_index,
            shape_id: hit.shape_id,
            story_id: hit.story_id,
            position: hit.position,
            line_index: hit.line_index,
        });
        self.vertical_goal_x = None;
        true
    }

    pub(crate) fn select_story_position(&mut self, story_id: &str, position: u32) -> Result<bool> {
        self.presentation.anchor_caret(story_id, position)?;
        for (slide_index, rendered) in self.rendered.iter().enumerate() {
            for primitive in &rendered.display_list.primitives {
                let Primitive::TextBox {
                    shape_id: Some(shape_id),
                    story_id: Some(candidate_story),
                    lines,
                    ..
                } = primitive
                else {
                    continue;
                };
                if candidate_story != story_id {
                    continue;
                }
                if let Some(line_index) = lines.iter().position(|line| {
                    line.caret_stops
                        .iter()
                        .any(|stop| stop.position == position)
                }) {
                    self.caret = Some(PptxCaret {
                        slide_index,
                        shape_id: shape_id.clone(),
                        story_id: story_id.to_owned(),
                        position,
                        line_index,
                    });
                    self.vertical_goal_x = None;
                    return Ok(true);
                }
            }
        }
        Ok(false)
    }

    pub fn clear_caret(&mut self) -> bool {
        let changed = self.caret.take().is_some();
        self.vertical_goal_x = None;
        changed
    }

    pub fn caret_geometry(&self) -> Option<PptxCaretGeometry> {
        let caret = self.caret.as_ref()?;
        let placement = caret_placement(
            self.rendered.get(caret.slide_index)?,
            caret,
            Some(caret.line_index),
        )?;
        Some(PptxCaretGeometry {
            page_index: caret.slide_index,
            x: f64::from(placement.x),
            y: f64::from(placement.y),
            height: f64::from(placement.height),
            transform: placement.transform,
        })
    }

    pub fn move_caret(&mut self, direction: MoveDirection) -> bool {
        let Some(caret) = self.caret.clone() else {
            return false;
        };
        let Some(rendered) = self.rendered.get(caret.slide_index) else {
            return false;
        };
        let Some(text_box) = text_box(rendered, &caret.shape_id, &caret.story_id) else {
            return false;
        };
        let Some(current) = caret_placement(rendered, &caret, Some(caret.line_index)) else {
            return false;
        };
        let target = match direction {
            MoveDirection::Left | MoveDirection::Right => {
                self.vertical_goal_x = None;
                horizontal_target(
                    text_box.lines,
                    current.line_index,
                    caret.position,
                    matches!(direction, MoveDirection::Right),
                )
            }
            MoveDirection::Up | MoveDirection::Down => {
                let goal = *self.vertical_goal_x.get_or_insert(current.x);
                vertical_target(
                    text_box.lines,
                    current.line_index,
                    goal,
                    matches!(direction, MoveDirection::Down),
                )
            }
            MoveDirection::Home | MoveDirection::End => {
                self.vertical_goal_x = None;
                line_edge_target(
                    text_box.lines,
                    current.line_index,
                    matches!(direction, MoveDirection::End),
                )
            }
        };
        let Some((line_index, position)) = target else {
            return false;
        };
        let changed = line_index != caret.line_index || position != caret.position;
        if changed {
            self.caret = Some(PptxCaret {
                line_index,
                position,
                ..caret
            });
        }
        changed
    }

    pub fn insert_text(&mut self, text: &str) -> Result<Option<PptxEditChange>> {
        if text.is_empty() {
            return Ok(None);
        }
        let Some(caret) = self.caret.clone() else {
            return Ok(None);
        };
        let story = self.presentation.story(&caret.story_id)?;
        let style = style_at(&story, caret.position);
        let placement = caret_placement(
            &self.rendered[caret.slide_index],
            &caret,
            Some(caret.line_index),
        );
        self.presentation.insert_text(
            &edit_context(),
            &caret.story_id,
            caret.position,
            text,
            &style,
        )?;
        let position = caret.position + utf16_len(text);
        self.finish_edit(caret, position, placement.map(|value| value.y))
            .map(Some)
    }

    pub fn enter(&mut self) -> Result<Option<PptxEditChange>> {
        let Some(caret) = self.caret.clone() else {
            return Ok(None);
        };
        let placement = caret_placement(
            &self.rendered[caret.slide_index],
            &caret,
            Some(caret.line_index),
        );
        self.presentation.insert_paragraph_break(
            &edit_context(),
            &caret.story_id,
            caret.position,
        )?;
        self.finish_edit(
            caret.clone(),
            caret.position + 1,
            placement.map(|value| value.y),
        )
        .map(Some)
    }

    pub fn delete(&mut self, direction: DeleteDirection) -> Result<Option<PptxEditChange>> {
        let Some(caret) = self.caret.clone() else {
            return Ok(None);
        };
        let rendered = &self.rendered[caret.slide_index];
        let Some(text_box) = text_box(rendered, &caret.shape_id, &caret.story_id) else {
            return Ok(None);
        };
        let Some(current) = caret_placement(rendered, &caret, Some(caret.line_index)) else {
            return Ok(None);
        };
        let forward = matches!(direction, DeleteDirection::Forward);
        let Some((_, target)) =
            horizontal_target(text_box.lines, current.line_index, caret.position, forward)
        else {
            return Ok(None);
        };
        let (start, end, position) = if forward {
            (caret.position, target, caret.position)
        } else {
            (target, caret.position, target)
        };
        let story = self.presentation.story(&caret.story_id)?;
        if end == start + 1 && paragraph_breaks(&story).contains(&start) {
            self.presentation
                .delete_paragraph_break(&edit_context(), &caret.story_id, start)?;
        } else {
            self.presentation
                .delete_text(&edit_context(), &caret.story_id, start, end)?;
        }
        self.finish_edit(caret, position, Some(current.y)).map(Some)
    }

    pub fn undo(&mut self) -> Result<Option<Vec<PageScene>>> {
        if !self.presentation.undo() {
            return Ok(None);
        }
        self.local_dirty = true;
        self.clear_caret();
        self.relayout_all().map(Some)
    }

    pub fn redo(&mut self) -> Result<Option<Vec<PageScene>>> {
        if !self.presentation.redo() {
            return Ok(None);
        }
        self.local_dirty = true;
        self.clear_caret();
        self.relayout_all().map(Some)
    }

    pub fn save_to(&mut self, path: &Path) -> Result<()> {
        let current = self.presentation.snapshot()?;
        let edited = text_only_changes(&self.baseline, &current)?;
        let source_parts = package_parts(&self.source)?;
        reject_signatures(&source_parts)?;
        let saved = self.presentation.save()?;
        ensure_package_fidelity(&self.source, &saved, &edited)?;
        let reopened = Presentation::open(&saved)
            .context("cannot save PPTX safely: reopen the engine output")?;
        ensure_model_fidelity(&current, &reopened)?;
        ensure_text_metadata_fidelity(self.presentation.package(), reopened.package(), &edited)?;
        fs::write(path, saved).with_context(|| format!("write PPTX {}", path.display()))?;
        self.saved_snapshot = current;
        self.dirty = false;
        self.local_dirty = false;
        self.remote_dirty = false;
        Ok(())
    }

    pub fn recover_layout(&mut self) -> Result<Vec<PageScene>> {
        self.relayout_all()
    }

    fn finish_edit(
        &mut self,
        mut caret: PptxCaret,
        position: u32,
        preferred_y: Option<f32>,
    ) -> Result<PptxEditChange> {
        let slide_index = caret.slide_index;
        let rendered = self
            .presentation
            .render_slide(slide_index)
            .with_context(|| format!("relayout PPTX slide {}", slide_index + 1))?;
        let (page, summary) = self
            .resources
            .translate(&rendered.display_list, self.max_texture_dimension_2d)?;
        caret.position = position;
        caret.line_index = preferred_y
            .and_then(|y| placement_near_y(&rendered, &caret, y))
            .or_else(|| caret_placement(&rendered, &caret, None))
            .map_or(caret.line_index, |placement| placement.line_index);
        self.rendered[slide_index] = rendered;
        self.summaries[slide_index] = summary;
        self.caret = Some(caret);
        self.vertical_goal_x = None;
        self.local_dirty = true;
        self.update_dirty()?;
        Ok(PptxEditChange {
            page_index: slide_index,
            page,
        })
    }

    fn relayout_all(&mut self) -> Result<Vec<PageScene>> {
        let slide_count = self.presentation.snapshot()?.slides.len();
        let mut rendered_slides = Vec::with_capacity(slide_count);
        let mut summaries = Vec::with_capacity(slide_count);
        let mut pages = Vec::with_capacity(slide_count);
        for slide_index in 0..slide_count {
            let rendered = self
                .presentation
                .render_slide(slide_index)
                .with_context(|| format!("relayout PPTX slide {}", slide_index + 1))?;
            let (page, summary) = self
                .resources
                .translate(&rendered.display_list, self.max_texture_dimension_2d)?;
            rendered_slides.push(rendered);
            summaries.push(summary);
            pages.push(page);
            #[cfg(test)]
            {
                self.relayout_count += 1;
            }
        }
        self.rendered = rendered_slides;
        self.summaries = summaries;
        self.update_dirty()?;
        Ok(pages)
    }

    fn relayout_slide(&mut self, slide_index: usize) -> Result<PptxEditChange> {
        let rendered = self
            .presentation
            .render_slide(slide_index)
            .with_context(|| format!("relayout PPTX slide {}", slide_index + 1))?;
        let (page, summary) = self
            .resources
            .translate(&rendered.display_list, self.max_texture_dimension_2d)?;
        self.rendered[slide_index] = rendered;
        self.summaries[slide_index] = summary;
        #[cfg(test)]
        {
            self.relayout_count += 1;
        }
        Ok(PptxEditChange {
            page_index: slide_index,
            page,
        })
    }

    fn restore_caret(
        &mut self,
        caret: Option<PptxCaret>,
        anchor: Option<CaretAnchor>,
        preferred_y: Option<f32>,
        snapshot: &DeckSnapshot,
    ) {
        let (Some(mut caret), Some(anchor)) = (caret, anchor) else {
            return;
        };
        let Some(position) = self.presentation.resolve_caret_anchor(&anchor) else {
            self.caret = None;
            return;
        };
        let Some((slide_index, shape_id)) =
            story_location(snapshot, &caret.story_id, &caret.shape_id)
        else {
            self.caret = None;
            return;
        };
        caret.slide_index = slide_index;
        caret.shape_id = shape_id;
        caret.position = position;
        let placement = preferred_y
            .and_then(|y| placement_near_y(&self.rendered[slide_index], &caret, y))
            .or_else(|| caret_placement(&self.rendered[slide_index], &caret, None));
        let Some(placement) = placement else {
            self.caret = None;
            return;
        };
        caret.line_index = placement.line_index;
        self.caret = Some(caret);
    }

    fn update_dirty(&mut self) -> Result<()> {
        self.dirty = self.presentation.snapshot()? != self.saved_snapshot;
        Ok(())
    }
}

#[derive(Clone, Copy)]
struct TextBoxRef<'a> {
    lines: &'a [betteroffice_pptx::PositionedTextLine],
    transform: Affine,
}

#[derive(Clone, Copy)]
struct CaretPlacement {
    line_index: usize,
    x: f32,
    y: f32,
    height: f32,
    transform: Affine,
}

fn edit_context() -> EditCtx {
    EditCtx::local(EDIT_AUTHOR)
}

fn text_box<'a>(
    rendered: &'a RenderedSlide,
    shape_id: &str,
    story_id: &str,
) -> Option<TextBoxRef<'a>> {
    rendered
        .display_list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                shape_id: Some(candidate_shape),
                story_id: Some(candidate_story),
                x,
                y,
                w,
                h,
                lines,
                transform,
                ..
            } if candidate_shape == shape_id && candidate_story == story_id => Some(TextBoxRef {
                lines,
                transform: primitive_transform(*x, *y, *w, *h, *transform),
            }),
            _ => None,
        })
}

fn story_location(
    snapshot: &DeckSnapshot,
    story_id: &str,
    preferred_shape_id: &str,
) -> Option<(usize, String)> {
    snapshot
        .slides
        .iter()
        .enumerate()
        .find_map(|(slide_index, slide)| {
            shape_for_story(&slide.shapes, story_id, Some(preferred_shape_id))
                .or_else(|| shape_for_story(&slide.shapes, story_id, None))
                .map(|shape_id| (slide_index, shape_id.to_owned()))
        })
}

fn shape_for_story<'a>(
    shapes: &'a [ShapeSnapshot],
    story_id: &str,
    preferred_shape_id: Option<&str>,
) -> Option<&'a str> {
    for shape in shapes {
        if preferred_shape_id.is_none_or(|preferred| preferred == shape.id)
            && shape.text_stories.iter().any(|story| story.id == story_id)
        {
            return Some(&shape.id);
        }
        if let Some(shape_id) = shape_for_story(&shape.children, story_id, preferred_shape_id) {
            return Some(shape_id);
        }
    }
    None
}

fn primitive_transform(x: f32, y: f32, w: f32, h: f32, transform: Transform) -> Affine {
    let center = Point::new(f64::from(x + w / 2.0), f64::from(y + h / 2.0));
    let flip = Affine::translate(center.to_vec2())
        * Affine::scale_non_uniform(
            if transform.flip_h { -1.0 } else { 1.0 },
            if transform.flip_v { -1.0 } else { 1.0 },
        )
        * Affine::translate(-center.to_vec2());
    Affine::rotate_about(f64::from(transform.rotation_deg).to_radians(), center) * flip
}

fn nearest_line(lines: &[betteroffice_pptx::PositionedTextLine], y: f32) -> Option<usize> {
    lines
        .iter()
        .enumerate()
        .min_by(|(_, left), (_, right)| {
            distance_to_interval(y, left.y, left.y + left.height).total_cmp(&distance_to_interval(
                y,
                right.y,
                right.y + right.height,
            ))
        })
        .map(|(index, _)| index)
}

fn distance_to_interval(value: f32, start: f32, end: f32) -> f32 {
    if value < start {
        start - value
    } else if value > end {
        value - end
    } else {
        0.0
    }
}

fn caret_placement(
    rendered: &RenderedSlide,
    caret: &PptxCaret,
    preferred_line: Option<usize>,
) -> Option<CaretPlacement> {
    let text_box = text_box(rendered, &caret.shape_id, &caret.story_id)?;
    let line_index = preferred_line
        .filter(|index| {
            text_box.lines.get(*index).is_some_and(|line| {
                line.caret_stops
                    .iter()
                    .any(|stop| stop.position == caret.position)
            })
        })
        .or_else(|| {
            text_box.lines.iter().position(|line| {
                line.caret_stops
                    .iter()
                    .any(|stop| stop.position == caret.position)
            })
        })?;
    let line = &text_box.lines[line_index];
    let stop = line
        .caret_stops
        .iter()
        .find(|stop| stop.position == caret.position)?;
    Some(CaretPlacement {
        line_index,
        x: stop.x,
        y: line.y,
        height: line.height,
        transform: text_box.transform,
    })
}

fn placement_near_y(
    rendered: &RenderedSlide,
    caret: &PptxCaret,
    preferred_y: f32,
) -> Option<CaretPlacement> {
    let text_box = text_box(rendered, &caret.shape_id, &caret.story_id)?;
    let line_index = text_box
        .lines
        .iter()
        .enumerate()
        .filter(|(_, line)| {
            line.caret_stops
                .iter()
                .any(|stop| stop.position == caret.position)
        })
        .min_by(|(_, left), (_, right)| {
            (left.y - preferred_y)
                .abs()
                .total_cmp(&(right.y - preferred_y).abs())
        })
        .map(|(index, _)| index)?;
    caret_placement(rendered, caret, Some(line_index))
}

fn horizontal_target(
    lines: &[betteroffice_pptx::PositionedTextLine],
    line_index: usize,
    position: u32,
    forward: bool,
) -> Option<(usize, u32)> {
    let mut positions = BTreeSet::new();
    for line in lines {
        positions.extend(line.caret_stops.iter().map(|stop| stop.position));
    }
    let target = if forward {
        positions
            .range(position.saturating_add(1)..)
            .next()
            .copied()
    } else {
        positions.range(..position).next_back().copied()
    }?;
    let target_line = lines
        .iter()
        .enumerate()
        .filter(|(_, line)| line.caret_stops.iter().any(|stop| stop.position == target))
        .min_by_key(|(candidate, _)| candidate.abs_diff(line_index))
        .map(|(index, _)| index)?;
    Some((target_line, target))
}

fn vertical_target(
    lines: &[betteroffice_pptx::PositionedTextLine],
    line_index: usize,
    goal_x: f32,
    forward: bool,
) -> Option<(usize, u32)> {
    let target_line = if forward {
        line_index.checked_add(1)?
    } else {
        line_index.checked_sub(1)?
    };
    let line = lines.get(target_line)?;
    let stop = line
        .caret_stops
        .iter()
        .min_by(|left, right| (left.x - goal_x).abs().total_cmp(&(right.x - goal_x).abs()))?;
    Some((target_line, stop.position))
}

fn line_edge_target(
    lines: &[betteroffice_pptx::PositionedTextLine],
    line_index: usize,
    end: bool,
) -> Option<(usize, u32)> {
    let line = lines.get(line_index)?;
    let stop = if end {
        line.caret_stops.last()
    } else {
        line.caret_stops.first()
    }?;
    Some((line_index, stop.position))
}

fn style_at(story: &StorySnapshot, position: u32) -> TextStyle {
    let mut offset = 0;
    let mut preceding = None;
    for paragraph in &story.paragraphs {
        for run in &paragraph.runs {
            let end = offset + utf16_len(&run.text);
            if position >= offset && position < end {
                return run.style.clone();
            }
            if position == offset {
                return run.style.clone();
            }
            preceding = Some(run.style.clone());
            offset = end;
        }
        if position == offset {
            return preceding.unwrap_or_default();
        }
        offset += 1;
    }
    preceding.unwrap_or_default()
}

fn paragraph_breaks(story: &StorySnapshot) -> BTreeSet<u32> {
    let mut breaks = BTreeSet::new();
    let mut offset = 0;
    for (index, paragraph) in story.paragraphs.iter().enumerate() {
        offset += paragraph
            .runs
            .iter()
            .map(|run| utf16_len(&run.text))
            .sum::<u32>();
        if index + 1 < story.paragraphs.len() {
            breaks.insert(offset);
        }
        offset += 1;
    }
    breaks
}

fn utf16_len(text: &str) -> u32 {
    text.encode_utf16().count() as u32
}

fn text_only_changes(baseline: &DeckSnapshot, current: &DeckSnapshot) -> Result<EditedSlides> {
    if (baseline.width_emu, baseline.height_emu) != (current.width_emu, current.height_emu)
        || baseline.slides.len() != current.slides.len()
    {
        bail!("cannot save PPTX safely: the deck structure changed outside shape text");
    }
    let mut edited = BTreeMap::new();
    for (base_slide, slide) in baseline.slides.iter().zip(&current.slides) {
        if slide_shell(base_slide) != slide_shell(slide) {
            bail!("cannot save PPTX safely: slide structure changed outside shape text");
        }
        let part = slide
            .source_part_path
            .clone()
            .context("cannot save PPTX safely: a slide has no source part")?;
        let mut shapes = BTreeMap::new();
        compare_shapes(&base_slide.shapes, &slide.shapes, &mut shapes)?;
        if !shapes.is_empty() {
            edited.insert(part, shapes);
        }
    }
    Ok(edited)
}

fn slide_shell(slide: &betteroffice_pptx::SlideSnapshot) -> betteroffice_pptx::SlideSnapshot {
    let mut shell = slide.clone();
    shell.shapes.clear();
    shell
}

fn shape_shell(shape: &ShapeSnapshot) -> ShapeSnapshot {
    let mut shell = shape.clone();
    shell.text_stories.clear();
    shell.children.clear();
    shell
}

fn compare_shapes(
    baseline: &[ShapeSnapshot],
    current: &[ShapeSnapshot],
    edited: &mut BTreeMap<u32, Vec<StoryChange>>,
) -> Result<()> {
    if baseline.len() != current.len() {
        bail!("cannot save PPTX safely: shape structure changed outside shape text");
    }
    for (base, shape) in baseline.iter().zip(current) {
        if shape_shell(base) != shape_shell(shape) {
            bail!(
                "cannot save PPTX safely: shape {} geometry or non-text properties changed",
                shape.name
            );
        }
        if base.text_stories.len() != shape.text_stories.len()
            || base
                .text_stories
                .iter()
                .zip(&shape.text_stories)
                .any(|(left, right)| left.id != right.id)
        {
            bail!(
                "cannot save PPTX safely: shape {} text-story structure changed",
                shape.name
            );
        }
        if base.text_stories != shape.text_stories {
            if shape.source_id == 0 {
                bail!(
                    "cannot save PPTX safely: edited text in shape {} cannot be isolated in the source slide",
                    shape.name
                );
            }
            let changes = base
                .text_stories
                .iter()
                .zip(&shape.text_stories)
                .filter(|(baseline, current)| baseline != current)
                .map(|(baseline, current)| StoryChange {
                    baseline: baseline.clone(),
                    current: current.clone(),
                })
                .collect::<Vec<_>>();
            edited.insert(shape.source_id, changes);
        }
        compare_shapes(&base.children, &shape.children, edited)?;
    }
    Ok(())
}

fn package_parts(bytes: &[u8]) -> Result<BTreeMap<String, Vec<u8>>> {
    Ok(ooxml_opc::unzip_parts(bytes)
        .map_err(anyhow::Error::msg)?
        .into_iter()
        .collect())
}

fn reject_signatures(parts: &BTreeMap<String, Vec<u8>>) -> Result<()> {
    if let Some(path) = parts.keys().find(|path| {
        let normalized = path.to_ascii_lowercase();
        normalized.starts_with("_xmlsignatures/") || normalized.contains("/_xmlsignatures/")
    }) {
        bail!(
            "cannot save PPTX safely: {path} contains a package digital signature that an edit would invalidate"
        );
    }
    Ok(())
}

fn ensure_package_fidelity(source: &[u8], saved: &[u8], edited: &EditedSlides) -> Result<()> {
    let source_parts = package_parts(source)?;
    let saved_parts = package_parts(saved)?;
    if source_parts.keys().collect::<Vec<_>>() != saved_parts.keys().collect::<Vec<_>>() {
        bail!("cannot save PPTX safely: the package part set changed during the text edit");
    }
    for (path, source_bytes) in &source_parts {
        let saved_bytes = &saved_parts[path];
        let Some(shapes) = edited.get(path) else {
            if source_bytes != saved_bytes {
                bail!(
                    "cannot save PPTX safely: unrelated package part {path} changed during the engine round trip"
                );
            }
            continue;
        };
        let shape_ids = shapes.keys().copied().collect::<BTreeSet<_>>();
        let source_xml = masked_slide(source_bytes, &shape_ids, path)?;
        let saved_xml = masked_slide(saved_bytes, &shape_ids, path)?;
        let source_tokens = canonical_xml(&source_xml)?;
        let saved_tokens = canonical_xml(&saved_xml)?;
        if source_tokens != saved_tokens {
            let mismatch = source_tokens
                .iter()
                .zip(&saved_tokens)
                .position(|(left, right)| left != right)
                .unwrap_or(source_tokens.len().min(saved_tokens.len()));
            let source_token = source_tokens
                .get(mismatch)
                .map(String::as_str)
                .unwrap_or("<end>");
            let saved_token = saved_tokens
                .get(mismatch)
                .map(String::as_str)
                .unwrap_or("<end>");
            bail!(
                "cannot save PPTX safely: edited slide part {path} changed outside the edited shape text at XML token {mismatch}: {source_token:?} became {saved_token:?}"
            );
        }
    }
    Ok(())
}

fn ensure_model_fidelity(current: &DeckSnapshot, reopened: &Presentation) -> Result<()> {
    let reopened = reopened
        .snapshot()
        .context("cannot save PPTX safely: snapshot the engine output")?;
    if current.slides.len() != reopened.slides.len()
        || (current.width_emu, current.height_emu) != (reopened.width_emu, reopened.height_emu)
    {
        bail!("cannot save PPTX safely: deck geometry did not survive the engine round trip");
    }
    for (expected, actual) in current.slides.iter().zip(&reopened.slides) {
        if slide_shell(expected) != slide_shell(actual) {
            bail!("cannot save PPTX safely: slide metadata did not survive the engine round trip");
        }
        compare_round_trip_shapes(&expected.shapes, &actual.shapes)?;
    }
    Ok(())
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum StoryToken {
    Text(u16),
    ParagraphEnd,
}

#[derive(Clone, Debug, PartialEq)]
enum TokenMetadata {
    Run {
        properties: RunProperties,
        field_id: Option<String>,
        field_type: Option<String>,
        line_break: bool,
        xml: String,
    },
    Paragraph {
        properties: Box<ParagraphProperties>,
        end_properties: Option<RunProperties>,
        xml: String,
    },
}

#[derive(Clone, Debug)]
struct AnnotatedToken {
    value: StoryToken,
    metadata: TokenMetadata,
}

fn ensure_text_metadata_fidelity(
    source: &PptxPackage,
    saved: &PptxPackage,
    edited: &EditedSlides,
) -> Result<()> {
    for (path, shapes) in edited {
        let source_slide = source
            .slides
            .iter()
            .find(|slide| slide.part_path == *path)
            .with_context(|| {
                format!("cannot save PPTX safely: source slide part {path} is missing")
            })?;
        let saved_slide = saved
            .slides
            .iter()
            .find(|slide| slide.part_path == *path)
            .with_context(|| {
                format!("cannot save PPTX safely: saved slide part {path} is missing")
            })?;
        let source_xml =
            std::str::from_utf8(source.part_bytes(path).with_context(|| {
                format!("cannot save PPTX safely: source part {path} is missing")
            })?)
            .with_context(|| {
                format!("cannot save PPTX safely: source part {path} is not UTF-8 XML")
            })?;
        let saved_xml =
            std::str::from_utf8(saved.part_bytes(path).with_context(|| {
                format!("cannot save PPTX safely: saved part {path} is missing")
            })?)
            .with_context(|| {
                format!("cannot save PPTX safely: saved part {path} is not UTF-8 XML")
            })?;
        for (source_id, stories) in shapes {
            if stories.len() != 1 {
                bail!(
                    "cannot save PPTX safely: shape {source_id} in {path} has multiple edited text stories that cannot be audited independently"
                );
            }
            let source_node =
                find_source_shape(&source_slide.shapes, *source_id).with_context(|| {
                    format!(
                        "cannot save PPTX safely: source shape {source_id} is missing from {path}"
                    )
                })?;
            let saved_node =
                find_source_shape(&saved_slide.shapes, *source_id).with_context(|| {
                    format!(
                        "cannot save PPTX safely: saved shape {source_id} is missing from {path}"
                    )
                })?;
            let source_body = shape_text_body(source_node).with_context(|| {
                format!(
                    "cannot save PPTX safely: source shape {source_id} has no auditable text body"
                )
            })?;
            let saved_body = shape_text_body(saved_node).with_context(|| {
                format!(
                    "cannot save PPTX safely: saved shape {source_id} has no auditable text body"
                )
            })?;
            if text_body_shell(source_body) != text_body_shell(saved_body) {
                bail!(
                    "cannot save PPTX safely: shape {source_id} text-body properties changed outside the text edit"
                );
            }
            let source_block = source_shape_block(source_xml, *source_id, path)?;
            let saved_block = source_shape_block(saved_xml, *source_id, path)?;
            let source_body_xml = text_body_block(source_block, path, *source_id)?;
            let saved_body_xml = text_body_block(saved_block, path, *source_id)?;
            let source_tokens = annotated_story(source_body, source_body_xml, path, *source_id)?;
            let saved_tokens = annotated_story(saved_body, saved_body_xml, path, *source_id)?;
            let change = &stories[0];
            let baseline_values = story_tokens(&change.baseline);
            let current_values = story_tokens(&change.current);
            if token_values(&source_tokens) != baseline_values {
                bail!(
                    "cannot save PPTX safely: source shape {source_id} text does not match the editable story model"
                );
            }
            if token_values(&saved_tokens) != current_values {
                bail!(
                    "cannot save PPTX safely: saved shape {source_id} text does not match the edited story model"
                );
            }
            for (baseline_index, current_index) in
                unchanged_token_pairs(&baseline_values, &current_values)?
            {
                let source_metadata = &source_tokens[baseline_index].metadata;
                let saved_metadata = &saved_tokens[current_index].metadata;
                if source_metadata != saved_metadata {
                    let loss = metadata_difference(source_metadata, saved_metadata);
                    bail!(
                        "cannot save PPTX safely: shape {source_id} would lose {loss} from unchanged text at story position {baseline_index}"
                    );
                }
            }
        }
    }
    Ok(())
}

fn find_source_shape(shapes: &[ShapeNode], source_id: u32) -> Option<&ShapeNode> {
    for shape in shapes {
        if shape.id() == source_id {
            return Some(shape);
        }
        if let ShapeNode::Group(group) = shape
            && let Some(found) = find_source_shape(&group.children, source_id)
        {
            return Some(found);
        }
    }
    None
}

fn shape_text_body(shape: &ShapeNode) -> Option<&TextBody> {
    match shape {
        ShapeNode::Shape(shape) => shape.text.as_ref(),
        _ => None,
    }
}

fn text_body_shell(body: &TextBody) -> TextBody {
    let mut shell = body.clone();
    shell.paragraphs.clear();
    shell
}

fn annotated_story(
    body: &TextBody,
    xml: &str,
    path: &str,
    source_id: u32,
) -> Result<Vec<AnnotatedToken>> {
    let paragraph_blocks = element_blocks(xml, &["a:p"])?;
    if paragraph_blocks.len() != body.paragraphs.len() {
        bail!(
            "cannot save PPTX safely: shape {source_id} in {path} has text paragraphs that cannot be matched to the parsed model"
        );
    }
    let mut tokens = Vec::new();
    for (paragraph, paragraph_xml) in body.paragraphs.iter().zip(paragraph_blocks) {
        let run_blocks = element_blocks(paragraph_xml, &["a:r", "a:br", "a:fld"])?;
        if run_blocks.len() != paragraph.runs.len() {
            bail!(
                "cannot save PPTX safely: shape {source_id} in {path} has text runs that cannot be matched to the parsed model"
            );
        }
        for (run, run_xml) in paragraph.runs.iter().zip(run_blocks) {
            if run.text.is_empty() {
                bail!(
                    "cannot save PPTX safely: shape {source_id} in {path} contains an empty text run whose markup cannot be tied to an edited character"
                );
            }
            let metadata = run_metadata(run, run_fingerprint(run_xml)?);
            for unit in run.text.encode_utf16() {
                tokens.push(AnnotatedToken {
                    value: StoryToken::Text(unit),
                    metadata: metadata.clone(),
                });
            }
        }
        tokens.push(AnnotatedToken {
            value: StoryToken::ParagraphEnd,
            metadata: TokenMetadata::Paragraph {
                properties: Box::new(paragraph.properties.clone()),
                end_properties: paragraph.end_properties.clone(),
                xml: paragraph_fingerprint(paragraph_xml)?,
            },
        });
    }
    Ok(tokens)
}

fn run_metadata(run: &ModelTextRun, xml: String) -> TokenMetadata {
    TokenMetadata::Run {
        properties: run.properties.clone(),
        field_id: run.field_id.clone(),
        field_type: run.field_type.clone(),
        line_break: run.line_break,
        xml,
    }
}

fn story_tokens(story: &StorySnapshot) -> Vec<StoryToken> {
    let mut tokens = Vec::with_capacity(story.length as usize);
    for paragraph in &story.paragraphs {
        for run in &paragraph.runs {
            tokens.extend(run.text.encode_utf16().map(StoryToken::Text));
        }
        tokens.push(StoryToken::ParagraphEnd);
    }
    tokens
}

fn token_values(tokens: &[AnnotatedToken]) -> Vec<StoryToken> {
    tokens.iter().map(|token| token.value).collect()
}

fn unchanged_token_pairs(
    baseline: &[StoryToken],
    current: &[StoryToken],
) -> Result<Vec<(usize, usize)>> {
    let mut prefix = 0;
    while prefix < baseline.len() && prefix < current.len() && baseline[prefix] == current[prefix] {
        prefix += 1;
    }
    let mut suffix = 0;
    while prefix + suffix < baseline.len()
        && prefix + suffix < current.len()
        && baseline[baseline.len() - 1 - suffix] == current[current.len() - 1 - suffix]
    {
        suffix += 1;
    }
    let baseline_middle = &baseline[prefix..baseline.len() - suffix];
    let current_middle = &current[prefix..current.len() - suffix];
    let cells = baseline_middle
        .len()
        .checked_add(1)
        .and_then(|rows| {
            current_middle
                .len()
                .checked_add(1)
                .and_then(|columns| rows.checked_mul(columns))
        })
        .context("cannot save PPTX safely: text fidelity diff dimensions overflow")?;
    if cells > MAX_FIDELITY_DIFF_CELLS {
        bail!(
            "cannot save PPTX safely: the edited story is too large for an exact metadata-preservation audit"
        );
    }
    let columns = current_middle.len() + 1;
    let mut lengths = vec![0_u32; cells];
    for baseline_index in (0..baseline_middle.len()).rev() {
        for current_index in (0..current_middle.len()).rev() {
            let index = baseline_index * columns + current_index;
            lengths[index] = if baseline_middle[baseline_index] == current_middle[current_index] {
                lengths[(baseline_index + 1) * columns + current_index + 1] + 1
            } else {
                lengths[(baseline_index + 1) * columns + current_index]
                    .max(lengths[baseline_index * columns + current_index + 1])
            };
        }
    }
    let mut pairs = (0..prefix).map(|index| (index, index)).collect::<Vec<_>>();
    let mut baseline_index = 0;
    let mut current_index = 0;
    while baseline_index < baseline_middle.len() && current_index < current_middle.len() {
        if baseline_middle[baseline_index] == current_middle[current_index]
            && lengths[baseline_index * columns + current_index]
                == lengths[(baseline_index + 1) * columns + current_index + 1] + 1
        {
            pairs.push((prefix + baseline_index, prefix + current_index));
            baseline_index += 1;
            current_index += 1;
        } else if lengths[(baseline_index + 1) * columns + current_index]
            >= lengths[baseline_index * columns + current_index + 1]
        {
            baseline_index += 1;
        } else {
            current_index += 1;
        }
    }
    pairs.extend((0..suffix).map(|index| {
        (
            baseline.len() - suffix + index,
            current.len() - suffix + index,
        )
    }));
    Ok(pairs)
}

fn metadata_difference(source: &TokenMetadata, saved: &TokenMetadata) -> &'static str {
    match (source, saved) {
        (
            TokenMetadata::Run {
                properties: source_properties,
                field_id: source_field_id,
                field_type: source_field_type,
                line_break: source_line_break,
                xml: source_xml,
            },
            TokenMetadata::Run {
                properties: saved_properties,
                field_id: saved_field_id,
                field_type: saved_field_type,
                line_break: saved_line_break,
                xml: saved_xml,
            },
        ) => {
            if source_field_id != saved_field_id || source_field_type != saved_field_type {
                "field identity"
            } else if source_properties.hyperlink_relationship_id
                != saved_properties.hyperlink_relationship_id
            {
                "hyperlink identity"
            } else if source_properties.language != saved_properties.language {
                "run language"
            } else if source_properties.color != saved_properties.color {
                "run color metadata"
            } else if source_line_break != saved_line_break {
                "line-break identity"
            } else if source_properties != saved_properties {
                "run formatting"
            } else if source_xml != saved_xml {
                "unmodeled run XML"
            } else {
                "text metadata"
            }
        }
        (
            TokenMetadata::Paragraph {
                properties: source_properties,
                end_properties: source_end,
                xml: source_xml,
            },
            TokenMetadata::Paragraph {
                properties: saved_properties,
                end_properties: saved_end,
                xml: saved_xml,
            },
        ) => {
            if source_properties != saved_properties {
                "paragraph properties"
            } else if source_end != saved_end {
                "paragraph-end formatting"
            } else if source_xml != saved_xml {
                "unmodeled paragraph XML"
            } else {
                "paragraph metadata"
            }
        }
        _ => "text structure metadata",
    }
}

fn compare_round_trip_shapes(expected: &[ShapeSnapshot], actual: &[ShapeSnapshot]) -> Result<()> {
    if expected.len() != actual.len() {
        bail!("cannot save PPTX safely: shape geometry did not survive the engine round trip");
    }
    for (expected, actual) in expected.iter().zip(actual) {
        if shape_shell(expected) != shape_shell(actual) {
            bail!(
                "cannot save PPTX safely: shape {} geometry did not survive the engine round trip",
                expected.name
            );
        }
        if expected.text_stories.len() != actual.text_stories.len()
            || expected
                .text_stories
                .iter()
                .zip(&actual.text_stories)
                .any(|(left, right)| !story_content_eq(left, right))
        {
            bail!(
                "cannot save PPTX safely: edited text in shape {} did not survive the engine round trip",
                expected.name
            );
        }
        compare_round_trip_shapes(&expected.children, &actual.children)?;
    }
    Ok(())
}

fn story_content_eq(left: &StorySnapshot, right: &StorySnapshot) -> bool {
    left.length == right.length
        && left.paragraphs.len() == right.paragraphs.len()
        && left
            .paragraphs
            .iter()
            .zip(&right.paragraphs)
            .all(|(left, right)| paragraph_content(left) == paragraph_content(right))
}

fn paragraph_content(paragraph: &ParagraphSnapshot) -> ParagraphSnapshot {
    let mut content = paragraph.clone();
    content.id.clear();
    content
}

fn source_shape_block<'a>(xml: &'a str, source_id: u32, path: &str) -> Result<&'a str> {
    let mut cursor = 0;
    while let Some((start, end)) = element_range(xml, "p:sp", cursor)? {
        let block = &xml[start..end];
        if shape_source_id(block) == Some(source_id) {
            return Ok(block);
        }
        cursor = end;
    }
    bail!("cannot save PPTX safely: shape {source_id} is not exposed in {path}")
}

fn text_body_block<'a>(block: &'a str, path: &str, source_id: u32) -> Result<&'a str> {
    let (start, end) = element_range(block, "p:txBody", 0)?.with_context(|| {
        format!("cannot save PPTX safely: shape {source_id} in {path} has no complete text body")
    })?;
    Ok(&block[start..end])
}

fn element_blocks<'a>(xml: &'a str, names: &[&str]) -> Result<Vec<&'a str>> {
    let mut blocks = Vec::new();
    let mut cursor = 0;
    while let Some((_, start)) = names
        .iter()
        .filter_map(|name| find_tag(xml, name, cursor).map(|start| (*name, start)))
        .min_by_key(|(_, start)| *start)
    {
        let name = names
            .iter()
            .copied()
            .find(|name| find_tag(xml, name, cursor) == Some(start))
            .context("cannot save PPTX safely: XML element lookup became inconsistent")?;
        let (_, end) = element_range(xml, name, start)?
            .context("cannot save PPTX safely: XML element disappeared during audit")?;
        blocks.push(&xml[start..end]);
        cursor = end;
    }
    Ok(blocks)
}

fn element_range(xml: &str, name: &str, from: usize) -> Result<Option<(usize, usize)>> {
    let Some(start) = find_tag(xml, name, from) else {
        return Ok(None);
    };
    let tag_end = xml_tag_end(xml, start).with_context(|| {
        format!("cannot save PPTX safely: {name} has an unterminated start tag")
    })?;
    let self_closing = xml.as_bytes()[start..tag_end]
        .iter()
        .rev()
        .copied()
        .find(|byte| !byte.is_ascii_whitespace())
        == Some(b'/');
    if self_closing {
        return Ok(Some((start, tag_end + 1)));
    }
    let closing = format!("</{name}>");
    let end = xml[tag_end + 1..]
        .find(&closing)
        .map(|offset| tag_end + 1 + offset + closing.len())
        .with_context(|| format!("cannot save PPTX safely: {name} has no closing tag"))?;
    Ok(Some((start, end)))
}

fn xml_tag_end(xml: &str, start: usize) -> Option<usize> {
    let mut quote = None;
    for (offset, byte) in xml.as_bytes().iter().copied().enumerate().skip(start + 1) {
        match (quote, byte) {
            (None, b'\'' | b'"') => quote = Some(byte),
            (Some(open), close) if open == close => quote = None,
            (None, b'>') => return Some(offset),
            _ => {}
        }
    }
    None
}

fn strip_elements(xml: &str, names: &[&str]) -> Result<String> {
    let mut stripped = String::with_capacity(xml.len());
    let mut cursor = 0;
    loop {
        let next = names
            .iter()
            .filter_map(|name| find_tag(xml, name, cursor).map(|start| (*name, start)))
            .min_by_key(|(_, start)| *start);
        let Some((name, start)) = next else {
            stripped.push_str(&xml[cursor..]);
            break;
        };
        let (_, end) = element_range(xml, name, start)?
            .context("cannot save PPTX safely: XML element disappeared during masking")?;
        stripped.push_str(&xml[cursor..start]);
        cursor = end;
    }
    Ok(stripped)
}

fn run_fingerprint(xml: &str) -> Result<String> {
    xml_fingerprint(&strip_elements(xml, &["a:t"])?)
}

fn paragraph_fingerprint(xml: &str) -> Result<String> {
    xml_fingerprint(&strip_elements(xml, &["a:r", "a:br", "a:fld"])?)
}

fn xml_fingerprint(xml: &str) -> Result<String> {
    Ok(canonical_xml(xml)?.join("\u{1f}"))
}

fn masked_slide(bytes: &[u8], shape_ids: &BTreeSet<u32>, path: &str) -> Result<String> {
    let xml = std::str::from_utf8(bytes)
        .with_context(|| format!("cannot save PPTX safely: {path} is not UTF-8 XML"))?;
    let mut output = String::with_capacity(xml.len());
    let mut cursor = 0;
    let mut found = BTreeSet::new();
    while let Some(start) = find_tag(xml, "p:sp", cursor) {
        let Some(close) = xml[start..]
            .find("</p:sp>")
            .map(|offset| start + offset + 7)
        else {
            bail!("cannot save PPTX safely: {path} has an unterminated shape element");
        };
        output.push_str(&xml[cursor..start]);
        let block = &xml[start..close];
        let source_id = shape_source_id(block);
        if source_id.is_some_and(|id| shape_ids.contains(&id)) {
            let id = source_id.expect("checked above");
            output.push_str(&mask_text_body(block, path, id)?);
            found.insert(id);
        } else {
            output.push_str(block);
        }
        cursor = close;
    }
    output.push_str(&xml[cursor..]);
    if &found != shape_ids {
        let missing = shape_ids.difference(&found).copied().collect::<Vec<_>>();
        bail!("cannot save PPTX safely: {path} does not expose edited shape IDs {missing:?}");
    }
    Ok(output)
}

fn find_tag(xml: &str, name: &str, from: usize) -> Option<usize> {
    let needle = format!("<{name}");
    let mut cursor = from;
    while let Some(offset) = xml[cursor..].find(&needle) {
        let start = cursor + offset;
        let next = xml.as_bytes().get(start + needle.len()).copied();
        if next.is_some_and(|byte| byte == b'>' || byte == b'/' || byte.is_ascii_whitespace()) {
            return Some(start);
        }
        cursor = start + needle.len();
    }
    None
}

fn shape_source_id(block: &str) -> Option<u32> {
    let start = find_tag(block, "p:cNvPr", 0)?;
    let end = block[start..].find('>')? + start;
    attribute(&block[start..=end], "id")?.parse().ok()
}

fn attribute<'a>(tag: &'a str, name: &str) -> Option<&'a str> {
    let mut cursor = 0;
    while let Some(offset) = tag[cursor..].find(name) {
        let start = cursor + offset;
        let before = start
            .checked_sub(1)
            .and_then(|index| tag.as_bytes().get(index));
        let after = tag.as_bytes().get(start + name.len());
        if before.is_some_and(|byte| byte.is_ascii_whitespace()) && after == Some(&b'=') {
            let quote = *tag.as_bytes().get(start + name.len() + 1)?;
            if !matches!(quote, b'\'' | b'"') {
                return None;
            }
            let value_start = start + name.len() + 2;
            let value_end = tag.as_bytes()[value_start..]
                .iter()
                .position(|byte| *byte == quote)?
                + value_start;
            return Some(&tag[value_start..value_end]);
        }
        cursor = start + name.len();
    }
    None
}

fn mask_text_body(block: &str, path: &str, source_id: u32) -> Result<String> {
    let body_start = find_tag(block, "p:txBody", 0).with_context(|| {
        format!("cannot save PPTX safely: shape {source_id} in {path} has no text body")
    })?;
    let body_end = block[body_start..]
        .find("</p:txBody>")
        .map(|offset| body_start + offset + 11)
        .with_context(|| {
            format!(
                "cannot save PPTX safely: shape {source_id} in {path} has no complete text body"
            )
        })?;
    let body = &block[body_start..body_end];
    let paragraph_start = find_tag(body, "a:p", 0).with_context(|| {
        format!("cannot save PPTX safely: shape {source_id} in {path} has no text paragraph")
    })?;
    let paragraph_end = body
        .rfind("</a:p>")
        .map(|offset| offset + 6)
        .with_context(|| {
            format!("cannot save PPTX safely: shape {source_id} in {path} has no complete text paragraph")
        })?;
    let mut masked = String::with_capacity(block.len());
    masked.push_str(&block[..body_start]);
    masked.push_str(&body[..paragraph_start]);
    masked.push_str("<edited-text/>");
    masked.push_str(&body[paragraph_end..]);
    masked.push_str(&block[body_end..]);
    Ok(masked)
}

fn canonical_xml(xml: &str) -> Result<Vec<String>> {
    let mut tokens = Vec::new();
    let mut cursor = 0;
    while cursor < xml.len() {
        let Some(offset) = xml[cursor..].find('<') else {
            let text = &xml[cursor..];
            if !text.trim().is_empty() {
                tokens.push(format!("text:{text}"));
            }
            break;
        };
        let start = cursor + offset;
        let text = &xml[cursor..start];
        if !text.trim().is_empty() {
            tokens.push(format!("text:{text}"));
        }
        let end = xml[start..]
            .find('>')
            .map(|offset| start + offset)
            .context("cannot save PPTX safely: slide XML contains an unterminated tag")?;
        let raw = xml[start + 1..end].trim();
        if !raw.starts_with('?') {
            tokens.extend(canonical_tag(raw)?);
        }
        cursor = end + 1;
    }
    Ok(tokens)
}

fn canonical_tag(raw: &str) -> Result<Vec<String>> {
    if raw.starts_with('/') || raw.starts_with('!') {
        return Ok(vec![raw.split_whitespace().collect::<Vec<_>>().join(" ")]);
    }
    let self_closing = raw.ends_with('/');
    let raw = raw.trim_end_matches('/').trim_end();
    let name_end = raw.find(char::is_whitespace).unwrap_or(raw.len());
    let name = &raw[..name_end];
    let mut attributes = Vec::new();
    let mut cursor = name_end;
    while cursor < raw.len() {
        while raw.as_bytes()[cursor].is_ascii_whitespace() {
            cursor += 1;
            if cursor == raw.len() {
                break;
            }
        }
        if cursor == raw.len() {
            break;
        }
        let attr_start = cursor;
        while cursor < raw.len()
            && !raw.as_bytes()[cursor].is_ascii_whitespace()
            && raw.as_bytes()[cursor] != b'='
        {
            cursor += 1;
        }
        let attr_name = &raw[attr_start..cursor];
        while cursor < raw.len() && raw.as_bytes()[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if raw.as_bytes().get(cursor) != Some(&b'=') {
            bail!("cannot save PPTX safely: malformed XML attribute {attr_name}");
        }
        cursor += 1;
        while cursor < raw.len() && raw.as_bytes()[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *raw
            .as_bytes()
            .get(cursor)
            .context("cannot save PPTX safely: unterminated XML attribute")?;
        if !matches!(quote, b'\'' | b'"') {
            bail!("cannot save PPTX safely: unquoted XML attribute {attr_name}");
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < raw.len() && raw.as_bytes()[cursor] != quote {
            cursor += 1;
        }
        if cursor == raw.len() {
            bail!("cannot save PPTX safely: unterminated XML attribute {attr_name}");
        }
        attributes.push((attr_name.to_owned(), raw[value_start..cursor].to_owned()));
        cursor += 1;
    }
    attributes.sort();
    let mut tokens = vec![format!(
        "{}|{}",
        name,
        attributes
            .into_iter()
            .map(|(name, value)| format!("{name}={value}"))
            .collect::<Vec<_>>()
            .join("|")
    )];
    if self_closing {
        tokens.push(format!("/{name}"));
    }
    Ok(tokens)
}
