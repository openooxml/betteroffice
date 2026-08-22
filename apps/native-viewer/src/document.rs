#[cfg(feature = "docx")]
use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result, bail};
#[cfg(feature = "xlsx")]
use betteroffice_xlsx::CellRef;
#[cfg(feature = "docx")]
use docx_edit::{EditingDoc, EngineSession, SimpleFormat, seed_from_docx};
#[cfg(feature = "docx")]
use docx_layout::display_list::DisplayList;
#[cfg(feature = "docx")]
use docx_parse::{S9PackageWire, S9ParseOptions, parse_docx_s9_wire};
#[cfg(feature = "docx")]
use serde_json::{Value, json};
use vello::kurbo::Affine;

#[cfg(feature = "docx")]
use crate::chrome::Alignment;
use crate::chrome::EditingState;
#[cfg(feature = "docx")]
use crate::collaboration::BROWSER_SEED_CLIENT_ID;
#[cfg(feature = "docx")]
use crate::docx_scene::translate_document;
#[cfg(feature = "docx")]
use crate::editing::{DocxEditor, SceneChange, SelectionRect, TextLoc};
#[cfg(feature = "docx")]
use crate::fonts::FontRegistry;
#[cfg(feature = "docx")]
use crate::images::ImageRegistry;
#[cfg(feature = "pptx")]
use crate::pptx_editing::{PptxEditChange, PptxEditor, PptxHit, PptxRemoteChange, PptxTextHit};
use crate::scene_shared::PageScene;
#[cfg(feature = "xlsx")]
use crate::xlsx_editing::{CellMove, XlsxEditor};
#[cfg(feature = "xlsx")]
use crate::xlsx_scene::translate_sheet;

pub struct DocumentView {
    pub source: PathBuf,
    pub reference: ReferenceDocument,
    pub pages: Vec<PageScene>,
    pub max_texture_dimension_2d: u32,
}

pub enum ReferenceDocument {
    #[cfg(feature = "docx")]
    Docx(Box<DocxReference>),
    #[cfg(feature = "xlsx")]
    Xlsx(Box<XlsxReference>),
    #[cfg(feature = "pptx")]
    Pptx(Box<PptxReference>),
}

#[cfg(feature = "docx")]
pub struct DocxReference {
    pub display_list: DisplayList,
    pub fonts: FontRegistry,
    pub images: ImageRegistry,
    pub editor: DocxEditor,
}

#[cfg(feature = "xlsx")]
pub struct XlsxReference {
    pub editor: XlsxEditor,
    pub sheet_index: usize,
    pub sheet_count: usize,
    pub sheet_name: String,
    pub chart_count: usize,
    pub chart_placeholders: usize,
}

#[cfg(feature = "xlsx")]
pub struct XlsxOverlay {
    pub rect: xlsx_render::Rect,
    pub draft: Option<String>,
}

#[cfg(feature = "pptx")]
pub struct PptxReference {
    pub editor: PptxEditor,
}

#[derive(Clone, Copy, Debug)]
pub struct ViewerCaretGeometry {
    pub page_index: usize,
    pub x: f64,
    pub y: f64,
    pub height: f64,
    pub transform: Affine,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[allow(dead_code)]
pub enum DocumentFormat {
    Docx,
    Xlsx,
    Pptx,
}

#[cfg(any(feature = "docx", feature = "pptx"))]
#[derive(Clone, Copy, Debug)]
pub enum MoveDirection {
    Left,
    Right,
    Up,
    Down,
    Home,
    End,
}

#[cfg(any(feature = "docx", feature = "pptx"))]
#[derive(Clone, Copy, Debug)]
pub enum DeleteDirection {
    Backward,
    Forward,
}

impl DocumentFormat {
    pub fn from_path(path: &Path) -> Result<Self> {
        match path
            .extension()
            .and_then(|extension| extension.to_str())
            .map(str::to_ascii_lowercase)
            .as_deref()
        {
            Some("docx") => {
                #[cfg(feature = "docx")]
                {
                    Ok(Self::Docx)
                }
                #[cfg(not(feature = "docx"))]
                {
                    bail!(
                        ".docx files open in BetterOffice Docs; this build does not include the DOCX engine"
                    )
                }
            }
            Some("xlsx") => {
                #[cfg(feature = "xlsx")]
                {
                    Ok(Self::Xlsx)
                }
                #[cfg(not(feature = "xlsx"))]
                {
                    bail!(
                        ".xlsx files open in BetterOffice Sheets; this build does not include the XLSX engine"
                    )
                }
            }
            Some("pptx") => {
                #[cfg(feature = "pptx")]
                {
                    Ok(Self::Pptx)
                }
                #[cfg(not(feature = "pptx"))]
                {
                    bail!(
                        ".pptx files open in BetterOffice Slides; this build does not include the PPTX engine"
                    )
                }
            }
            _ => bail!("--document must be a .docx, .xlsx, or .pptx file"),
        }
    }
}

#[allow(irrefutable_let_patterns)]
impl DocumentView {
    pub fn scene_label(&self, _index: usize) -> String {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(_) => format!("page {}", _index + 1),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => {
                format!("sheet {}", reference.sheet_index + 1)
            }
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(_) => format!("slide {}", _index + 1),
        }
    }

    pub fn title_summary(&self) -> String {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(_) => format!("{} pages", self.pages.len()),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => format!(
                "sheet {} of {} ({})",
                reference.sheet_index + 1,
                reference.sheet_count,
                reference.sheet_name
            ),
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => {
                format!("{} slides", reference.editor.slide_count())
            }
        }
    }

    pub fn status_position(&self, _page_index: usize) -> String {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(_) => {
                format!("Page {} of {}", _page_index + 1, self.pages.len())
            }
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => {
                let cell = reference
                    .editor
                    .status()
                    .unwrap_or_else(|_| reference.editor.selection().to_a1());
                format!(
                    "Sheet {} of {} · {cell}",
                    reference.sheet_index + 1,
                    reference.sheet_count
                )
            }
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => {
                format!(
                    "Slide {} of {}",
                    _page_index + 1,
                    reference.editor.slide_count()
                )
            }
        }
    }

    pub fn editing_state(&self) -> Result<EditingState> {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(reference) => reference.editor.editing_state(),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => Ok(EditingState::editable_without_selection(
                reference.editor.can_undo(),
                reference.editor.can_redo(),
            )),
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => Ok(EditingState::editable_without_selection(
                reference.editor.can_undo(),
                reference.editor.can_redo(),
            )),
        }
    }

    pub fn display_item_name(&self) -> &'static str {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(_) => "primitives",
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(_) => "commands",
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(_) => "primitives",
        }
    }

    #[cfg(feature = "docx")]
    pub fn docx_hit_test(&self, page_index: usize, x: f64, y: f64) -> Result<Option<TextLoc>> {
        let ReferenceDocument::Docx(reference) = &self.reference else {
            return Ok(None);
        };
        reference.editor.hit_test(page_index, x, y)
    }

    #[cfg(feature = "docx")]
    pub fn docx_select_point(&mut self, loc: TextLoc, extend: bool, word: bool) -> Result<bool> {
        let ReferenceDocument::Docx(reference) = &mut self.reference else {
            return Ok(false);
        };
        reference.editor.select_point(loc, extend, word)?;
        Ok(true)
    }

    #[cfg(feature = "docx")]
    pub fn docx_extend_to(&mut self, loc: TextLoc) -> Result<bool> {
        let ReferenceDocument::Docx(reference) = &mut self.reference else {
            return Ok(false);
        };
        reference.editor.extend_to(loc)?;
        Ok(true)
    }

    #[cfg(feature = "docx")]
    pub fn docx_move_selection(&mut self, direction: MoveDirection, extend: bool) -> Result<bool> {
        let ReferenceDocument::Docx(reference) = &mut self.reference else {
            return Ok(false);
        };
        reference.editor.move_selection(direction, extend)
    }

    #[cfg(feature = "docx")]
    pub fn docx_insert_text(&mut self, text: &str) -> Result<bool> {
        self.apply_docx_edit(|editor| editor.insert_text(text))
    }

    #[cfg(feature = "docx")]
    pub fn docx_delete(&mut self, direction: DeleteDirection) -> Result<bool> {
        self.apply_docx_edit(|editor| editor.delete(direction))
    }

    #[cfg(feature = "docx")]
    pub fn docx_enter(&mut self) -> Result<bool> {
        self.apply_docx_edit(DocxEditor::enter)
    }

    #[cfg(feature = "docx")]
    pub fn docx_toggle_format(&mut self, format: SimpleFormat) -> Result<bool> {
        self.apply_docx_edit(|editor| editor.toggle_format(format))
    }

    #[cfg(feature = "docx")]
    pub fn docx_set_alignment(&mut self, alignment: Alignment) -> Result<bool> {
        self.apply_docx_edit(|editor| editor.set_alignment(alignment))
    }

    #[cfg(feature = "docx")]
    pub fn docx_undo(&mut self) -> Result<bool> {
        self.apply_docx_edit(DocxEditor::undo)
    }

    #[cfg(feature = "docx")]
    pub fn docx_redo(&mut self) -> Result<bool> {
        self.apply_docx_edit(DocxEditor::redo)
    }

    #[cfg(all(test, feature = "docx"))]
    pub fn docx_state_vector(&self) -> Result<Vec<u8>> {
        let ReferenceDocument::Docx(reference) = &self.reference else {
            bail!("collaboration is available only for DOCX");
        };
        Ok(reference.editor.state_vector())
    }

    #[cfg(feature = "docx")]
    pub fn docx_apply_remote_update(&mut self, update: &[u8]) -> Result<bool> {
        self.apply_docx_edit(|editor| editor.apply_remote_update(update).map(Some))
    }

    #[cfg(all(test, feature = "docx"))]
    pub fn docx_drain_local_updates(&self) -> Vec<Vec<u8>> {
        let ReferenceDocument::Docx(reference) = &self.reference else {
            return Vec::new();
        };
        reference.editor.drain_local_updates()
    }

    #[cfg(all(test, feature = "docx"))]
    pub fn docx_fingerprint(&self) -> Result<[u8; 32]> {
        let ReferenceDocument::Docx(reference) = &self.reference else {
            bail!("collaboration is available only for DOCX");
        };
        Ok(reference.editor.source_fingerprint())
    }

    pub fn collaboration_state_vector(&self) -> Result<Vec<u8>> {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(reference) => Ok(reference.editor.state_vector()),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => reference.editor.state_vector(),
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => reference.editor.state_vector(),
        }
    }

    pub fn collaboration_encode_diff(&self, state_vector: &[u8]) -> Result<Vec<u8>> {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(reference) => reference.editor.encode_diff(state_vector),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => reference.editor.encode_diff(state_vector),
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => reference.editor.encode_diff(state_vector),
        }
    }

    pub fn collaboration_apply_remote_update(&mut self, update: &[u8]) -> Result<bool> {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(_) => self.docx_apply_remote_update(update),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(_) => {
                let changed = {
                    let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
                        unreachable!();
                    };
                    reference.editor.apply_remote_update(update)?
                };
                if changed {
                    self.refresh_xlsx_scene()?;
                }
                Ok(changed)
            }
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(_) => {
                let change = {
                    let ReferenceDocument::Pptx(reference) = &mut self.reference else {
                        unreachable!();
                    };
                    reference.editor.apply_remote_update(update)?
                };
                let Some(change) = change else {
                    return Ok(false);
                };
                match change {
                    PptxRemoteChange::Slides(changes) => {
                        for change in changes {
                            self.pages[change.page_index] = change.page;
                        }
                    }
                    PptxRemoteChange::All(pages) => self.pages = pages,
                }
                Ok(true)
            }
        }
    }

    pub fn collaboration_drain_local_updates(&self) -> Vec<Vec<u8>> {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(reference) => reference.editor.drain_local_updates(),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => reference.editor.drain_local_updates(),
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => reference.editor.drain_local_updates(),
        }
    }

    pub fn collaboration_fingerprint(&self) -> Result<[u8; 32]> {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(reference) => Ok(reference.editor.source_fingerprint()),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => reference.editor.source_fingerprint(),
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => reference.editor.source_fingerprint(),
        }
    }

    #[cfg(all(test, feature = "docx"))]
    pub fn docx_canonical_checksum(&self) -> Result<[u8; 32]> {
        let ReferenceDocument::Docx(reference) = &self.reference else {
            bail!("collaboration is available only for DOCX");
        };
        reference.editor.canonical_checksum()
    }

    #[cfg(all(test, feature = "docx"))]
    pub fn docx_relayout_count(&self) -> usize {
        let ReferenceDocument::Docx(reference) = &self.reference else {
            return 0;
        };
        reference.editor.relayout_count()
    }

    #[cfg(test)]
    pub fn collaboration_canonical_checksum(&self) -> Result<[u8; 32]> {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(reference) => reference.editor.canonical_checksum(),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => reference.editor.canonical_checksum(),
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => reference.editor.canonical_checksum(),
        }
    }

    #[cfg(test)]
    pub fn collaboration_relayout_count(&self) -> usize {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(reference) => reference.editor.relayout_count(),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => reference.editor.relayout_count(),
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => reference.editor.relayout_count(),
        }
    }

    #[cfg(feature = "docx")]
    pub fn docx_selection_rects(&self) -> &[SelectionRect] {
        let ReferenceDocument::Docx(reference) = &self.reference else {
            return &[];
        };
        reference.editor.selection_rects()
    }

    pub fn caret_geometry(&self) -> Result<Option<ViewerCaretGeometry>> {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(reference) => {
                Ok(reference
                    .editor
                    .caret_geometry()?
                    .map(|caret| ViewerCaretGeometry {
                        page_index: caret.page_index,
                        x: caret.x,
                        y: caret.y,
                        height: caret.height,
                        transform: Affine::IDENTITY,
                    }))
            }
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => {
                Ok(reference
                    .editor
                    .caret_geometry()
                    .map(|caret| ViewerCaretGeometry {
                        page_index: caret.page_index,
                        x: caret.x,
                        y: caret.y,
                        height: caret.height,
                        transform: caret.transform,
                    }))
            }
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(_) => Ok(None),
        }
    }

    pub fn has_text_caret(&self) -> bool {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(reference) => reference.editor.has_collapsed_selection(),
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => reference.editor.has_caret(),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(_) => false,
        }
    }

    pub fn is_dirty(&self) -> bool {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(reference) => reference.editor.is_dirty(),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => {
                reference.editor.is_dirty() || reference.editor.is_editing()
            }
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => reference.editor.is_dirty(),
        }
    }

    pub fn is_remote_only_dirty(&self) -> bool {
        match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(reference) => reference.editor.is_remote_only_dirty(),
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => reference.editor.is_remote_only_dirty(),
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => reference.editor.is_remote_only_dirty(),
        }
    }

    #[cfg(feature = "docx")]
    pub fn save_docx_to(&mut self, path: &Path) -> Result<()> {
        let ReferenceDocument::Docx(reference) = &mut self.reference else {
            bail!("only DOCX documents are editable");
        };
        reference.editor.save_to(path)
    }

    #[cfg(feature = "pptx")]
    pub fn pptx_hit_test(&self, slide_index: usize, x: f64, y: f64) -> Option<PptxHit> {
        let ReferenceDocument::Pptx(reference) = &self.reference else {
            return None;
        };
        reference.editor.hit_test(slide_index, x as f32, y as f32)
    }

    #[cfg(feature = "pptx")]
    pub fn pptx_select_hit(&mut self, hit: PptxTextHit) -> bool {
        let ReferenceDocument::Pptx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.select_hit(hit)
    }

    #[cfg(feature = "pptx")]
    pub(crate) fn pptx_select_story_position(
        &mut self,
        story_id: &str,
        position: u32,
    ) -> Result<bool> {
        let ReferenceDocument::Pptx(reference) = &mut self.reference else {
            return Ok(false);
        };
        reference.editor.select_story_position(story_id, position)
    }

    #[cfg(feature = "pptx")]
    pub fn pptx_clear_caret(&mut self) -> bool {
        let ReferenceDocument::Pptx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.clear_caret()
    }

    #[cfg(feature = "pptx")]
    pub fn pptx_move_caret(&mut self, direction: MoveDirection) -> bool {
        let ReferenceDocument::Pptx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.move_caret(direction)
    }

    #[cfg(feature = "pptx")]
    pub fn pptx_insert_text(&mut self, text: &str) -> Result<bool> {
        self.apply_pptx_edit(|editor| editor.insert_text(text))
    }

    #[cfg(feature = "pptx")]
    pub fn pptx_delete(&mut self, direction: DeleteDirection) -> Result<bool> {
        self.apply_pptx_edit(|editor| editor.delete(direction))
    }

    #[cfg(feature = "pptx")]
    pub fn pptx_enter(&mut self) -> Result<bool> {
        self.apply_pptx_edit(PptxEditor::enter)
    }

    #[cfg(feature = "pptx")]
    pub fn pptx_undo(&mut self) -> Result<bool> {
        let pages = {
            let ReferenceDocument::Pptx(reference) = &mut self.reference else {
                return Ok(false);
            };
            reference.editor.undo()?
        };
        let Some(pages) = pages else {
            return Ok(false);
        };
        self.pages = pages;
        Ok(true)
    }

    #[cfg(feature = "pptx")]
    pub fn pptx_redo(&mut self) -> Result<bool> {
        let pages = {
            let ReferenceDocument::Pptx(reference) = &mut self.reference else {
                return Ok(false);
            };
            reference.editor.redo()?
        };
        let Some(pages) = pages else {
            return Ok(false);
        };
        self.pages = pages;
        Ok(true)
    }

    #[cfg(feature = "pptx")]
    pub fn pptx_is_dirty(&self) -> bool {
        matches!(
            &self.reference,
            ReferenceDocument::Pptx(reference) if reference.editor.is_dirty()
        )
    }

    #[cfg(feature = "pptx")]
    pub fn save_pptx_to(&mut self, path: &Path) -> Result<()> {
        let ReferenceDocument::Pptx(reference) = &mut self.reference else {
            bail!("only PPTX decks use the PPTX save path");
        };
        reference.editor.save_to(path)
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_hit_test(&self, page_index: usize, x: f64, y: f64) -> Option<CellRef> {
        let ReferenceDocument::Xlsx(reference) = &self.reference else {
            return None;
        };
        (page_index == 0)
            .then(|| reference.editor.cell_at_point(x as f32, y as f32))
            .flatten()
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_select_cell(&mut self, cell: CellRef) -> bool {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.select(cell)
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_move_selection(&mut self, movement: CellMove) -> bool {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.move_selection(movement)
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_begin_edit(&mut self, seed: Option<&str>) -> Result<bool> {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return Ok(false);
        };
        reference.editor.begin_edit(seed)
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_insert_text(&mut self, text: &str) -> bool {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.insert_text(text)
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_backspace(&mut self) -> bool {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.backspace()
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_cancel_edit(&mut self) -> bool {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.cancel()
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_commit(&mut self, movement: Option<CellMove>) -> Result<bool> {
        let changed = {
            let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
                return Ok(false);
            };
            reference.editor.commit(movement)?
        };
        if changed {
            self.refresh_xlsx_scene()?;
        }
        Ok(changed)
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_undo(&mut self) -> Result<bool> {
        let changed = {
            let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
                return Ok(false);
            };
            reference.editor.undo()?
        };
        if changed {
            self.refresh_xlsx_scene()?;
        }
        Ok(changed)
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_redo(&mut self) -> Result<bool> {
        let changed = {
            let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
                return Ok(false);
            };
            reference.editor.redo()?
        };
        if changed {
            self.refresh_xlsx_scene()?;
        }
        Ok(changed)
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_is_editing(&self) -> bool {
        matches!(
            &self.reference,
            ReferenceDocument::Xlsx(reference) if reference.editor.is_editing()
        )
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_is_dirty(&self) -> bool {
        matches!(
            &self.reference,
            ReferenceDocument::Xlsx(reference) if reference.editor.is_dirty()
        )
    }

    #[cfg(feature = "xlsx")]
    pub fn xlsx_overlay(&self) -> Option<XlsxOverlay> {
        let ReferenceDocument::Xlsx(reference) = &self.reference else {
            return None;
        };
        Some(XlsxOverlay {
            rect: reference.editor.selection_rect()?,
            draft: reference.editor.draft_value().map(str::to_owned),
        })
    }

    #[cfg(feature = "xlsx")]
    pub fn save_xlsx_to(&mut self, path: &Path) -> Result<()> {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            bail!("only XLSX workbooks use the XLSX save path");
        };
        reference.editor.save_to(path)
    }

    pub fn edited_path(&self) -> Result<PathBuf> {
        let extension = match &self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(_) => "docx",
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(_) => "xlsx",
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(_) => "pptx",
        };
        let stem = self
            .source
            .file_stem()
            .and_then(|stem| stem.to_str())
            .context("input has no UTF-8 file stem")?;
        let parent = self.source.parent().unwrap_or_else(|| Path::new("."));
        Ok(parent.join(format!("{stem}-edited.{extension}")))
    }

    pub fn recover_after_edit_error(&mut self) -> Result<()> {
        match &mut self.reference {
            #[cfg(feature = "docx")]
            ReferenceDocument::Docx(reference) => {
                reference.editor.recover_layout()?;
                let display_list = reference
                    .editor
                    .engine()
                    .with_display_list(Clone::clone)
                    .context("engine did not retain a recovered display list")?;
                self.pages =
                    translate_document(&display_list, &reference.fonts, &reference.images)?;
                reference.display_list = display_list;
            }
            #[cfg(feature = "xlsx")]
            ReferenceDocument::Xlsx(reference) => {
                reference.editor.recover_layout()?;
                let display_list = reference.editor.display_list();
                self.pages = vec![translate_sheet(display_list)?];
                reference.chart_count = display_list.charts.len();
                reference.chart_placeholders = display_list
                    .charts
                    .iter()
                    .filter(|chart| chart.placeholder)
                    .count();
            }
            #[cfg(feature = "pptx")]
            ReferenceDocument::Pptx(reference) => {
                self.pages = reference.editor.recover_layout()?;
            }
        }
        Ok(())
    }

    #[cfg(feature = "docx")]
    fn apply_docx_edit(
        &mut self,
        edit: impl FnOnce(&mut DocxEditor) -> Result<Option<SceneChange>>,
    ) -> Result<bool> {
        let ReferenceDocument::Docx(reference) = &mut self.reference else {
            return Ok(false);
        };
        let Some(change) = edit(&mut reference.editor)? else {
            return Ok(false);
        };
        let display_list = reference
            .editor
            .engine()
            .with_display_list(Clone::clone)
            .context("engine did not retain an edited display list")?;
        if change.rebuild_all || display_list.pages.len() != self.pages.len() {
            self.pages = translate_document(&display_list, &reference.fonts, &reference.images)?;
        } else {
            for page_index in change.changed_pages {
                let page = display_list
                    .pages
                    .get(page_index)
                    .with_context(|| format!("edited page {} is unavailable", page_index + 1))?;
                self.pages[page_index] =
                    crate::docx_scene::translate_page(page, &reference.fonts, &reference.images)?;
            }
        }
        reference.display_list = display_list;
        Ok(true)
    }

    #[cfg(feature = "pptx")]
    fn apply_pptx_edit(
        &mut self,
        edit: impl FnOnce(&mut PptxEditor) -> Result<Option<PptxEditChange>>,
    ) -> Result<bool> {
        let change = {
            let ReferenceDocument::Pptx(reference) = &mut self.reference else {
                return Ok(false);
            };
            edit(&mut reference.editor)?
        };
        let Some(change) = change else {
            return Ok(false);
        };
        self.pages[change.page_index] = change.page;
        Ok(true)
    }

    #[cfg(feature = "xlsx")]
    fn refresh_xlsx_scene(&mut self) -> Result<()> {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return Ok(());
        };
        let display_list = reference.editor.display_list();
        let page = translate_sheet(display_list)?;
        reference.chart_count = display_list.charts.len();
        reference.chart_placeholders = display_list
            .charts
            .iter()
            .filter(|chart| chart.placeholder)
            .count();
        self.pages[0] = page;
        Ok(())
    }
}

pub fn load_document(
    path: &Path,
    sheet_index: usize,
    max_texture_dimension_2d: u32,
) -> Result<DocumentView> {
    load_document_with_xlsx_mode(path, sheet_index, max_texture_dimension_2d, true)
}

pub fn load_document_for_export(
    path: &Path,
    sheet_index: usize,
    max_texture_dimension_2d: u32,
) -> Result<DocumentView> {
    load_document_with_xlsx_mode(path, sheet_index, max_texture_dimension_2d, false)
}

#[cfg(all(test, feature = "docx"))]
pub fn load_collaborative_docx(
    path: &Path,
    max_texture_dimension_2d: u32,
    client_id: u64,
) -> Result<DocumentView> {
    if DocumentFormat::from_path(path)? != DocumentFormat::Docx {
        bail!("--room is available only for DOCX");
    }
    load_docx(path, max_texture_dimension_2d, Some(client_id))
}

#[cfg(all(test, feature = "xlsx"))]
pub fn load_collaborative_xlsx(
    path: &Path,
    sheet_index: usize,
    max_texture_dimension_2d: u32,
    client_id: u64,
) -> Result<DocumentView> {
    if DocumentFormat::from_path(path)? != DocumentFormat::Xlsx {
        bail!("XLSX collaboration requires an .xlsx workbook");
    }
    load_xlsx(
        path,
        sheet_index,
        max_texture_dimension_2d,
        true,
        Some(client_id),
    )
}

#[cfg(all(test, feature = "pptx"))]
pub fn load_collaborative_pptx(
    path: &Path,
    max_texture_dimension_2d: u32,
    client_id: u64,
) -> Result<DocumentView> {
    if DocumentFormat::from_path(path)? != DocumentFormat::Pptx {
        bail!("PPTX collaboration requires a .pptx deck");
    }
    load_pptx(path, max_texture_dimension_2d, Some(client_id))
}

pub fn load_collaborative_document(
    path: &Path,
    _sheet_index: usize,
    max_texture_dimension_2d: u32,
    client_id: u64,
) -> Result<DocumentView> {
    match DocumentFormat::from_path(path)? {
        #[cfg(feature = "docx")]
        DocumentFormat::Docx => load_docx(path, max_texture_dimension_2d, Some(client_id)),
        #[cfg(feature = "xlsx")]
        DocumentFormat::Xlsx => load_xlsx(
            path,
            _sheet_index,
            max_texture_dimension_2d,
            true,
            Some(client_id),
        ),
        #[cfg(feature = "pptx")]
        DocumentFormat::Pptx => load_pptx(path, max_texture_dimension_2d, Some(client_id)),
        #[cfg(not(feature = "docx"))]
        DocumentFormat::Docx => unreachable!("DocumentFormat::from_path returned DOCX"),
        #[cfg(not(feature = "xlsx"))]
        DocumentFormat::Xlsx => unreachable!("DocumentFormat::from_path returned XLSX"),
        #[cfg(not(feature = "pptx"))]
        DocumentFormat::Pptx => unreachable!("DocumentFormat::from_path returned PPTX"),
    }
}

fn load_document_with_xlsx_mode(
    path: &Path,
    _sheet_index: usize,
    max_texture_dimension_2d: u32,
    _recalculate_xlsx: bool,
) -> Result<DocumentView> {
    match DocumentFormat::from_path(path)? {
        #[cfg(feature = "docx")]
        DocumentFormat::Docx => load_docx(path, max_texture_dimension_2d, None),
        #[cfg(feature = "xlsx")]
        DocumentFormat::Xlsx => load_xlsx(
            path,
            _sheet_index,
            max_texture_dimension_2d,
            _recalculate_xlsx,
            None,
        ),
        #[cfg(feature = "pptx")]
        DocumentFormat::Pptx => load_pptx(path, max_texture_dimension_2d, None),
        #[cfg(not(feature = "docx"))]
        DocumentFormat::Docx => unreachable!("DocumentFormat::from_path returned DOCX"),
        #[cfg(not(feature = "xlsx"))]
        DocumentFormat::Xlsx => unreachable!("DocumentFormat::from_path returned XLSX"),
        #[cfg(not(feature = "pptx"))]
        DocumentFormat::Pptx => unreachable!("DocumentFormat::from_path returned PPTX"),
    }
}

#[cfg(feature = "pptx")]
fn load_pptx(
    path: &Path,
    max_texture_dimension_2d: u32,
    collaboration_client_id: Option<u64>,
) -> Result<DocumentView> {
    let bytes = fs::read(path).with_context(|| format!("read PPTX {}", path.display()))?;
    let (editor, pages) = match collaboration_client_id {
        Some(client_id) => {
            PptxEditor::open_collaborative(bytes, max_texture_dimension_2d, client_id)
        }
        None => PptxEditor::open(bytes, max_texture_dimension_2d),
    }
    .with_context(|| format!("open PPTX {}", path.display()))?;
    Ok(DocumentView {
        source: path.to_owned(),
        reference: ReferenceDocument::Pptx(Box::new(PptxReference { editor })),
        pages,
        max_texture_dimension_2d,
    })
}

#[cfg(feature = "docx")]
fn load_docx(
    path: &Path,
    max_texture_dimension_2d: u32,
    collaboration_client_id: Option<u64>,
) -> Result<DocumentView> {
    let bytes = fs::read(path).with_context(|| format!("read DOCX {}", path.display()))?;
    let parsed = parse_docx_s9_wire(&bytes, S9ParseOptions::default())?;
    let package = parsed.document.package;
    let engine = EngineSession::new(collaboration_client_id.unwrap_or(0x0056_454c_4c4f));
    if collaboration_client_id.is_some() {
        let seed = EditingDoc::new(BROWSER_SEED_CLIENT_ID);
        seed_from_docx(&seed, &bytes).map_err(anyhow::Error::msg)?;
        engine
            .doc()
            .apply_update_v1(&seed.encode_state_as_update_v1())
            .map_err(anyhow::Error::msg)?;
    } else {
        seed_from_docx(engine.doc(), &bytes).map_err(anyhow::Error::msg)?;
    }

    let probe = region_request(&package, None)?;
    let requirements_json = engine
        .layout_font_requirements_json(&probe.to_string())
        .map_err(anyhow::Error::msg)?;
    let requirements: Value = serde_json::from_str(&requirements_json)?;
    let fonts = FontRegistry::load(&requirements)?;
    let request = region_request(&package, Some(&fonts.chain_ids))?;
    let layout_request = request.to_string();
    engine
        .layout_document_with_regions_json(&layout_request)
        .map_err(anyhow::Error::msg)?;
    let extras = json!({ "fontChains": fonts.chain_ids }).to_string();
    engine
        .build_display_list_frame(&extras, 0)
        .map_err(anyhow::Error::msg)?;
    let initial_frame = engine
        .apply_and_layout("body", 0)
        .map_err(anyhow::Error::msg)?;
    let display_list = engine
        .with_display_list(Clone::clone)
        .context("engine did not retain a display list")?;
    let images = ImageRegistry::load(&bytes)?;
    let pages = translate_document(&display_list, &fonts, &images)?;
    let editor = DocxEditor::new(
        engine,
        bytes,
        package,
        &initial_frame,
        layout_request,
        extras,
        collaboration_client_id.is_some(),
    )?;
    Ok(DocumentView {
        source: path.to_owned(),
        reference: ReferenceDocument::Docx(Box::new(DocxReference {
            display_list,
            fonts,
            images,
            editor,
        })),
        pages,
        max_texture_dimension_2d,
    })
}

#[cfg(feature = "xlsx")]
fn load_xlsx(
    path: &Path,
    sheet_index: usize,
    max_texture_dimension_2d: u32,
    recalculate: bool,
    collaboration_client_id: Option<u64>,
) -> Result<DocumentView> {
    let bytes = fs::read(path).with_context(|| format!("read XLSX {}", path.display()))?;
    let editor = match collaboration_client_id {
        Some(client_id) => {
            XlsxEditor::open_collaborative(bytes, sheet_index, max_texture_dimension_2d, client_id)
        }
        None => XlsxEditor::open(bytes, sheet_index, max_texture_dimension_2d, recalculate),
    }
    .map_err(|error| anyhow::anyhow!("open XLSX {}: {error:#}", path.display()))?;
    let sheet_count = editor.sheet_count();
    let sheet_name = editor.sheet_name()?.to_owned();
    let display_list = editor.display_list();
    let page = translate_sheet(display_list)?;
    let chart_count = display_list.charts.len();
    let chart_placeholders = display_list
        .charts
        .iter()
        .filter(|chart| chart.placeholder)
        .count();
    Ok(DocumentView {
        source: path.to_owned(),
        reference: ReferenceDocument::Xlsx(Box::new(XlsxReference {
            editor,
            sheet_index,
            sheet_count,
            sheet_name,
            chart_count,
            chart_placeholders,
        })),
        pages: vec![page],
        max_texture_dimension_2d,
    })
}

#[cfg(feature = "docx")]
fn region_request(
    package: &S9PackageWire,
    font_chains: Option<&BTreeMap<String, Vec<u32>>>,
) -> Result<Value> {
    let body = &package.document;
    let mut sections = body
        .sections
        .as_deref()
        .unwrap_or_default()
        .iter()
        .map(|section| {
            let properties = serde_json::to_value(&section.properties)?;
            let section_id = section
                .id
                .clone()
                .or_else(|| properties["sectionId"].as_str().map(str::to_owned));
            Ok(json!({ "sectionId": section_id, "properties": properties }))
        })
        .collect::<Result<Vec<_>>>()?;
    let final_properties = body
        .final_section_properties
        .as_ref()
        .map(serde_json::to_value)
        .transpose()?
        .unwrap_or_else(|| json!({}));
    sections.push(json!({
        "sectionId": final_properties["sectionId"].as_str(),
        "properties": final_properties
    }));
    let watermark = sections
        .last()
        .and_then(|section| section["properties"].get("watermark"))
        .cloned()
        .unwrap_or(Value::Null);
    let mut notes = Vec::new();
    push_notes(&mut notes, package.footnotes.as_deref(), "footnote");
    push_notes(&mut notes, package.endnotes.as_deref(), "endnote");
    let compatibility = &package.settings.compatibility_flags;
    let measurement = font_chains.map(|chains| {
        json!({
            "fontChains": chains,
            "defaults": { "fontSize": 11, "fontFamily": "Calibri" },
            "compat": {
                "noLeading": compatibility.no_leading,
                "doNotExpandShiftReturn": compatibility.do_not_expand_shift_return
            },
            "authoritativeShaping": true
        })
    });
    let mut request = json!({
        "bodyStory": "body",
        "options": { "pageGap": 24 },
        "regions": {
            "sections": sections,
            "settings": package.settings,
            "watermark": watermark
        },
        "notes": { "contents": notes },
        "renderEnv": {}
    });
    if let Some(measurement) = measurement {
        request["measurement"] = measurement;
    }
    Ok(request)
}

#[cfg(feature = "docx")]
fn push_notes(target: &mut Vec<Value>, notes: Option<&[docx_parse::Note]>, note_kind: &str) {
    for note in notes.unwrap_or_default() {
        if !note.is_separator() {
            target.push(json!({
                "id": note.id as i64,
                "noteKind": note_kind,
                "height": 0
            }));
        }
    }
}

#[cfg(all(test, feature = "docx", feature = "xlsx", feature = "pptx"))]
mod tests {
    use std::sync::atomic::{AtomicU64, Ordering};

    use betteroffice_pptx::{Primitive as PptxPrimitive, Transform as PptxTransform};
    use betteroffice_xlsx::{CellValue, DrawCmd, SheetId};
    use docx_edit::{
        EditCtx, EditError, EditingDoc, FormatPolicy, Position, SegmentContent, StoryRange,
    };
    use vello::kurbo::Point;
    use yrs::Update;
    use yrs::updates::decoder::Decode;

    use super::*;
    use crate::chrome::{Alignment, ToggleState};
    use crate::collaboration_protocol::{
        ProtocolMessage, decode_messages, encode_sync_step_1, encode_sync_step_2, encode_update,
    };
    use crate::editing::TextLoc;
    use crate::test_fixtures;

    static TEST_FILE_ID: AtomicU64 = AtomicU64::new(0);

    struct PptxProbe {
        hit: PptxTextHit,
        x: f32,
        y: f32,
        caret_x: f32,
        caret_y: f32,
        caret_height: f32,
        transform: Affine,
    }

    fn pptx_probe(document: &DocumentView, slide_index: usize) -> PptxProbe {
        pptx_probe_with_transform(document, slide_index, false)
    }

    fn pptx_probe_with_transform(
        document: &DocumentView,
        slide_index: usize,
        transformed: bool,
    ) -> PptxProbe {
        let ReferenceDocument::Pptx(reference) = &document.reference else {
            panic!("expected PPTX reference");
        };
        let rendered = reference.editor.rendered_slide(slide_index).unwrap();
        for primitive in rendered.display_list.primitives.iter().rev() {
            let PptxPrimitive::TextBox {
                shape_id: Some(shape_id),
                story_id: Some(story_id),
                x,
                y,
                w,
                h,
                lines,
                transform,
                ..
            } = primitive
            else {
                continue;
            };
            if (*transform != PptxTransform::default()) != transformed
                || reference.editor.story(story_id).is_err()
            {
                continue;
            }
            let affine = pptx_test_transform(*x, *y, *w, *h, *transform);
            for line in lines {
                for stop in &line.caret_stops {
                    let point = affine
                        * Point::new(f64::from(stop.x), f64::from(line.y + line.height / 2.0));
                    let Some(PptxHit::Text(hit)) =
                        reference
                            .editor
                            .hit_test(slide_index, point.x as f32, point.y as f32)
                    else {
                        continue;
                    };
                    if hit.shape_id == *shape_id
                        && hit.story_id == *story_id
                        && hit.position == stop.position
                    {
                        return PptxProbe {
                            hit,
                            x: point.x as f32,
                            y: point.y as f32,
                            caret_x: stop.x,
                            caret_y: line.y,
                            caret_height: line.height,
                            transform: affine,
                        };
                    }
                }
            }
        }
        panic!("demo PPTX has no hittable editable text stop");
    }

    fn pptx_test_transform(x: f32, y: f32, w: f32, h: f32, transform: PptxTransform) -> Affine {
        let center = Point::new(f64::from(x + w / 2.0), f64::from(y + h / 2.0));
        let flip = Affine::translate(center.to_vec2())
            * Affine::scale_non_uniform(
                if transform.flip_h { -1.0 } else { 1.0 },
                if transform.flip_v { -1.0 } else { 1.0 },
            )
            * Affine::translate(-center.to_vec2());
        Affine::rotate_about(f64::from(transform.rotation_deg).to_radians(), center) * flip
    }

    fn docx_paragraph_text(document: &DocumentView, para_id: &str) -> String {
        let ReferenceDocument::Docx(reference) = &document.reference else {
            panic!("expected DOCX reference");
        };
        reference
            .editor
            .engine()
            .doc()
            .paragraphs("body")
            .unwrap()
            .into_iter()
            .find(|paragraph| paragraph.para_id == para_id)
            .unwrap()
            .text
    }

    fn docx_paragraph_embed_count(document: &DocumentView, para_id: &str) -> usize {
        let ReferenceDocument::Docx(reference) = &document.reference else {
            panic!("expected DOCX reference");
        };
        let mut count = 0;
        for segment in reference
            .editor
            .engine()
            .doc()
            .story_segments("body")
            .unwrap()
        {
            match segment.content {
                SegmentContent::OtherEmbed { .. } => count += 1,
                SegmentContent::Pilcrow(properties) if properties.para_id == para_id => {
                    return count;
                }
                SegmentContent::Pilcrow(_) => count = 0,
                SegmentContent::Text(_) => {}
            }
        }
        panic!("paragraph {para_id} is unavailable");
    }

    fn docx_selection_range(document: &DocumentView) -> StoryRange {
        let ReferenceDocument::Docx(reference) = &document.reference else {
            panic!("expected DOCX reference");
        };
        reference.editor.selection_range().unwrap().unwrap()
    }

    fn complete_sync(sender: &DocumentView, receiver: &mut DocumentView) {
        let step_1 = encode_sync_step_1(&receiver.collaboration_state_vector().unwrap()).unwrap();
        let decoded_step_1 = decode_messages(&step_1).unwrap();
        let [ProtocolMessage::SyncStep1(state_vector)] = decoded_step_1.as_slice() else {
            unreachable!();
        };
        let step_2 =
            encode_sync_step_2(&sender.collaboration_encode_diff(state_vector).unwrap()).unwrap();
        let decoded_step_2 = decode_messages(&step_2).unwrap();
        let [ProtocolMessage::SyncStep2(update)] = decoded_step_2.as_slice() else {
            unreachable!();
        };
        receiver.collaboration_apply_remote_update(update).unwrap();
    }

    fn framed_local_updates(document: &DocumentView) -> Vec<Vec<u8>> {
        document
            .collaboration_drain_local_updates()
            .into_iter()
            .map(|update| encode_update(&update).unwrap())
            .collect()
    }

    fn deliver_room_frames(receiver: &mut DocumentView, frames: &[Vec<u8>]) {
        for frame in frames {
            let messages = decode_messages(frame).unwrap();
            let [ProtocolMessage::Update(update)] = messages.as_slice() else {
                unreachable!();
            };
            receiver.collaboration_apply_remote_update(update).unwrap();
        }
    }

    #[test]
    fn recognizes_supported_extensions() {
        assert_eq!(
            DocumentFormat::from_path(Path::new("document.DOCX")).unwrap(),
            DocumentFormat::Docx
        );
        assert_eq!(
            DocumentFormat::from_path(Path::new("workbook.XLSX")).unwrap(),
            DocumentFormat::Xlsx
        );
        assert_eq!(
            DocumentFormat::from_path(Path::new("slides.PPTX")).unwrap(),
            DocumentFormat::Pptx
        );
    }

    #[test]
    fn two_native_sessions_in_one_room_converge_to_the_same_canonical_checksum() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Shared document</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("native-room-convergence", &bytes);
        let mut left = load_collaborative_docx(&source, 8_192, 101).unwrap();
        let mut right = load_collaborative_docx(&source, 8_192, 202).unwrap();
        complete_sync(&left, &mut right);
        complete_sync(&right, &mut left);
        let baseline = left.docx_canonical_checksum().unwrap();
        assert_eq!(baseline, right.docx_canonical_checksum().unwrap());

        let caret = TextLoc {
            para_id: "11111111".to_owned(),
            offset: 6,
        };
        left.docx_select_point(caret.clone(), false, false).unwrap();
        right.docx_select_point(caret, false, false).unwrap();
        assert!(left.docx_insert_text(" left").unwrap());
        let left_first = framed_local_updates(&left);
        assert!(right.docx_insert_text(" right").unwrap());
        let right_first = framed_local_updates(&right);
        deliver_room_frames(&mut left, &right_first);
        assert!(left.docx_insert_text(" second").unwrap());
        let left_second = framed_local_updates(&left);
        assert!(right.docx_insert_text(" second").unwrap());
        let right_second = framed_local_updates(&right);
        deliver_room_frames(&mut right, &left_first);
        deliver_room_frames(&mut left, &right_second);
        deliver_room_frames(&mut right, &left_second);
        complete_sync(&left, &mut right);
        complete_sync(&right, &mut left);

        let left_checksum = left.docx_canonical_checksum().unwrap();
        let right_checksum = right.docx_canonical_checksum().unwrap();
        assert_ne!(left_checksum, baseline);
        assert_eq!(left_checksum, right_checksum);
        std::fs::remove_file(source).unwrap();
    }

    #[test]
    fn two_native_xlsx_sessions_converge_after_interleaved_cell_edits() {
        let source = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let mut left = load_collaborative_xlsx(&source, 0, 8_192, 101).unwrap();
        let mut right = load_collaborative_xlsx(&source, 0, 8_192, 202).unwrap();
        assert_eq!(
            left.collaboration_fingerprint().unwrap(),
            right.collaboration_fingerprint().unwrap()
        );
        complete_sync(&left, &mut right);
        complete_sync(&right, &mut left);
        let baseline = left.collaboration_canonical_checksum().unwrap();
        assert_eq!(baseline, right.collaboration_canonical_checksum().unwrap());

        let edit = |document: &mut DocumentView, cell: &str, value: &str| {
            document.xlsx_select_cell(CellRef::parse_a1(cell).unwrap());
            document.xlsx_begin_edit(Some(value)).unwrap();
            assert!(document.xlsx_commit(None).unwrap());
        };
        edit(&mut left, "B5", "left first");
        let left_first = framed_local_updates(&left);
        edit(&mut right, "C5", "right first");
        let right_first = framed_local_updates(&right);
        deliver_room_frames(&mut left, &right_first);
        edit(&mut left, "D5", "left second");
        let left_second = framed_local_updates(&left);
        deliver_room_frames(&mut right, &left_first);
        edit(&mut right, "E5", "right second");
        let right_second = framed_local_updates(&right);
        deliver_room_frames(&mut left, &right_second);
        deliver_room_frames(&mut right, &left_second);
        complete_sync(&left, &mut right);
        complete_sync(&right, &mut left);

        let (ReferenceDocument::Xlsx(left_reference), ReferenceDocument::Xlsx(right_reference)) =
            (&left.reference, &right.reference)
        else {
            unreachable!();
        };
        assert_eq!(
            left_reference.editor.workbook().model(),
            right_reference.editor.workbook().model()
        );
        assert_eq!(
            left_reference.editor.workbook().save().unwrap(),
            right_reference.editor.workbook().save().unwrap()
        );
        let left_checksum = left.collaboration_canonical_checksum().unwrap();
        let right_checksum = right.collaboration_canonical_checksum().unwrap();
        assert_ne!(left_checksum, baseline);
        assert_eq!(left_checksum, right_checksum);
        assert_eq!(
            left.collaboration_state_vector().unwrap(),
            right.collaboration_state_vector().unwrap()
        );
    }

    #[test]
    fn two_native_pptx_sessions_converge_after_interleaved_text_edits() {
        let source =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let mut left = load_collaborative_pptx(&source, 8_192, 101).unwrap();
        let mut right = load_collaborative_pptx(&source, 8_192, 202).unwrap();
        assert_eq!(
            left.collaboration_fingerprint().unwrap(),
            right.collaboration_fingerprint().unwrap()
        );
        complete_sync(&left, &mut right);
        complete_sync(&right, &mut left);
        let baseline = left.collaboration_canonical_checksum().unwrap();
        assert_eq!(baseline, right.collaboration_canonical_checksum().unwrap());
        let ReferenceDocument::Pptx(reference) = &left.reference else {
            unreachable!();
        };
        let story = reference
            .editor
            .snapshot()
            .unwrap()
            .slides
            .into_iter()
            .flat_map(|slide| slide.shapes)
            .flat_map(|shape| shape.text_stories)
            .find(|story| story.length > 1)
            .unwrap();
        assert!(
            left.pptx_select_story_position(&story.id, story.length - 1)
                .unwrap()
        );
        assert!(
            right
                .pptx_select_story_position(&story.id, story.length - 1)
                .unwrap()
        );

        assert!(left.pptx_insert_text(" LEFT").unwrap());
        let left_first = framed_local_updates(&left);
        assert!(right.pptx_insert_text(" RIGHT").unwrap());
        let right_first = framed_local_updates(&right);
        deliver_room_frames(&mut left, &right_first);
        assert!(left.pptx_insert_text(" SECOND").unwrap());
        let left_second = framed_local_updates(&left);
        deliver_room_frames(&mut right, &left_first);
        assert!(right.pptx_insert_text(" SECOND").unwrap());
        let right_second = framed_local_updates(&right);
        deliver_room_frames(&mut left, &right_second);
        deliver_room_frames(&mut right, &left_second);
        complete_sync(&left, &mut right);
        complete_sync(&right, &mut left);

        let (ReferenceDocument::Pptx(left_reference), ReferenceDocument::Pptx(right_reference)) =
            (&left.reference, &right.reference)
        else {
            unreachable!();
        };
        assert_eq!(
            left_reference.editor.snapshot().unwrap(),
            right_reference.editor.snapshot().unwrap()
        );
        let left_checksum = left.collaboration_canonical_checksum().unwrap();
        let right_checksum = right.collaboration_canonical_checksum().unwrap();
        assert_ne!(left_checksum, baseline);
        assert_eq!(left_checksum, right_checksum);
        assert_eq!(
            left.collaboration_state_vector().unwrap(),
            right.collaboration_state_vector().unwrap()
        );
    }

    #[test]
    fn remote_pptx_insertion_before_the_caret_preserves_its_logical_position() {
        let source =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let mut local = load_collaborative_pptx(&source, 8_192, 101).unwrap();
        let mut peer = load_collaborative_pptx(&source, 8_192, 202).unwrap();
        let ReferenceDocument::Pptx(reference) = &local.reference else {
            unreachable!();
        };
        let story = reference
            .editor
            .snapshot()
            .unwrap()
            .slides
            .into_iter()
            .flat_map(|slide| slide.shapes)
            .flat_map(|shape| shape.text_stories)
            .find(|story| story.length > 1)
            .unwrap();
        let caret_before = story.length - 1;
        assert!(
            local
                .pptx_select_story_position(&story.id, caret_before)
                .unwrap()
        );
        assert!(peer.pptx_select_story_position(&story.id, 0).unwrap());
        let inserted = "Remote ";
        assert!(peer.pptx_insert_text(inserted).unwrap());
        let before_relayout = local.collaboration_relayout_count();

        for update in peer.collaboration_drain_local_updates() {
            assert!(local.collaboration_apply_remote_update(&update).unwrap());
        }

        let ReferenceDocument::Pptx(reference) = &local.reference else {
            unreachable!();
        };
        assert_eq!(reference.editor.caret_story_id(), Some(story.id.as_str()));
        assert_eq!(
            reference.editor.caret_position(),
            Some(caret_before + inserted.encode_utf16().count() as u32)
        );
        assert_eq!(local.collaboration_relayout_count(), before_relayout + 1);
        assert!(local.caret_geometry().unwrap().is_some());
    }

    #[test]
    fn remote_xlsx_edit_preserves_the_row_column_selection() {
        let source = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let mut local = load_collaborative_xlsx(&source, 0, 8_192, 101).unwrap();
        let mut peer = load_collaborative_xlsx(&source, 0, 8_192, 202).unwrap();
        let selected = CellRef::parse_a1("D8").unwrap();
        local.xlsx_select_cell(selected);
        local.xlsx_begin_edit(None).unwrap();
        let draft_before = local.xlsx_overlay().unwrap().draft;
        peer.xlsx_select_cell(CellRef::parse_a1("A1").unwrap());
        peer.xlsx_begin_edit(Some("remote")).unwrap();
        peer.xlsx_commit(None).unwrap();

        for update in peer.collaboration_drain_local_updates() {
            local.collaboration_apply_remote_update(&update).unwrap();
        }

        let ReferenceDocument::Xlsx(reference) = &local.reference else {
            unreachable!();
        };
        assert_eq!(reference.editor.selection(), selected);
        assert_eq!(reference.editor.draft_value(), draft_before.as_deref());
    }

    #[test]
    fn hostile_utf8_update_returns_typed_errors_without_crashing() {
        let hostile = [
            0x01, 0x01, 0xca, 0x01, 0x00, 0xc4, 0x01, 0x0a, 0x01, 0x0b, 0x06, 0x20, 0x74, 0x68,
            0x65, 0xe5, 0x65, 0x00,
        ];
        let bare = Update::decode_v1(&hostile);
        assert!(matches!(bare, Err(yrs::encoding::read::Error::Custom(_))));

        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Safe</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("hostile-update", &bytes);
        let mut document = load_collaborative_docx(&source, 8_192, 101).unwrap();
        let ReferenceDocument::Docx(reference) = &document.reference else {
            unreachable!();
        };
        assert!(matches!(
            reference.editor.engine().doc().apply_update_v1(&hostile),
            Err(EditError::InvalidUpdate(_))
        ));
        assert!(document.docx_apply_remote_update(&hostile).is_err());
        assert_eq!(docx_paragraph_text(&document, "11111111"), "Safe");
        std::fs::remove_file(source).unwrap();
    }

    #[test]
    fn mutated_real_updates_never_panic() {
        let writer = EditingDoc::new(71);
        writer
            .create_story("body", "Mutation seed", "Normal", "left")
            .unwrap();
        writer
            .insert_text(
                &EditCtx::local("", ""),
                Position::new("body", 4),
                " payload",
                FormatPolicy::Plain,
            )
            .unwrap();
        let real_update = writer.encode_state_as_update_v1();
        let mut state = 0x9e37_79b9_7f4a_7c15_u64;
        for case in 0..4_096 {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            let mut update = real_update.clone();
            match case % 3 {
                0 => update.truncate((state as usize) % (update.len() + 1)),
                1 => {
                    let index = (state as usize) % update.len();
                    update[index] ^= (state >> 32) as u8 | 1;
                }
                _ => update.extend_from_slice(&state.to_le_bytes()),
            }
            let result = std::panic::catch_unwind(move || {
                let reader = EditingDoc::new(72);
                let _ = reader.apply_update_v1(&update);
            });
            assert!(result.is_ok(), "mutation {case} panicked");
        }
    }

    #[test]
    fn remote_insert_transforms_the_local_selection_before_typing() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>ABCDEFGH</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("remote-selection-transform", &bytes);
        let mut local = load_collaborative_docx(&source, 8_192, 101).unwrap();
        let mut peer = load_collaborative_docx(&source, 8_192, 202).unwrap();
        local
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 4,
                },
                false,
                false,
            )
            .unwrap();
        local
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 8,
                },
                true,
                false,
            )
            .unwrap();
        peer.docx_select_point(
            TextLoc {
                para_id: "11111111".to_owned(),
                offset: 0,
            },
            false,
            false,
        )
        .unwrap();
        assert!(peer.docx_insert_text("ZZZZZ").unwrap());
        for update in peer.docx_drain_local_updates() {
            local.docx_apply_remote_update(&update).unwrap();
        }
        assert_eq!(docx_selection_range(&local), StoryRange::new("body", 9, 13));
        assert!(local.docx_insert_text("X").unwrap());
        assert_eq!(docx_paragraph_text(&local, "11111111"), "ZZZZZABCDX");
        std::fs::remove_file(source).unwrap();
    }

    #[test]
    fn opens_document_with_footnote_and_endnote_regions() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../crates/docx-edit/tests/fixtures/footnote-anchor.docx");
        let document = load_document(&path, 0, 8_192).unwrap();
        assert!(!document.pages.is_empty());
        assert!(
            document
                .pages
                .iter()
                .any(|page| !page.scene.encoding().is_empty())
        );
    }

    #[test]
    fn alignment_only_save_refuses_to_flatten_a_simple_field() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>"#,
        );
        let source = test_fixtures::write_docx("alignment-field", &bytes);
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 0,
                },
                false,
                false,
            )
            .unwrap();
        assert!(document.docx_set_alignment(Alignment::Center).unwrap());

        let output = test_fixtures::write_docx("alignment-field-output", b"sentinel");
        let error = document.save_docx_to(&output).unwrap_err().to_string();
        assert!(error.contains("paragraph 1 (11111111)"), "{error}");
        assert!(error.contains("fields"), "{error}");
        assert_eq!(fs::read(&output).unwrap(), b"sentinel");
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn edited_paragraph_preserves_unmodelled_properties_without_a_source_id() {
        let properties = r#"<w:pPr><w:kinsoku w:val="0"/><w:wordWrap w:val="0"/><w:snapToGrid w:val="0"/><w:textAlignment w:val="center"/></w:pPr>"#;
        let bytes = test_fixtures::editing_docx(&format!(
            "<w:p>{properties}<w:r><w:t>Plain</w:t></w:r></w:p>"
        ));
        let source = test_fixtures::write_docx("unmodelled-ppr", &bytes);
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let para_id = match &document.reference {
            ReferenceDocument::Docx(reference) => {
                reference.editor.engine().doc().paragraphs("body").unwrap()[0]
                    .para_id
                    .clone()
            }
            _ => unreachable!(),
        };
        document
            .docx_select_point(
                TextLoc {
                    para_id: para_id.clone(),
                    offset: 5,
                },
                false,
                false,
            )
            .unwrap();
        assert!(document.docx_insert_text(" edited").unwrap());

        let output = test_fixtures::write_docx("unmodelled-ppr-output", b"sentinel");
        document.save_docx_to(&output).unwrap();
        let saved = fs::read(&output).unwrap();
        let xml = String::from_utf8(test_fixtures::part(&saved, "word/document.xml")).unwrap();
        assert!(xml.contains(properties), "{xml}");
        let reopened = load_document(&output, 0, 8_192).unwrap();
        assert_eq!(docx_paragraph_text(&reopened, &para_id), "Plain edited");
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn structural_edit_disables_save_with_a_reason_until_undo() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Plain</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("structural-save-state", &bytes);
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 2,
                },
                false,
                false,
            )
            .unwrap();
        assert!(document.docx_enter().unwrap());
        let blocked = document.editing_state().unwrap();
        assert!(!blocked.can_save);
        assert_eq!(
            blocked.save_disabled_reason.as_deref(),
            Some(crate::editing::STRUCTURAL_SAVE_REASON)
        );
        assert!(document.docx_undo().unwrap());
        let restored = document.editing_state().unwrap();
        assert!(restored.can_save);
        assert!(restored.save_disabled_reason.is_none());
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn double_click_selects_the_second_adjacent_emoji() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>😀😀</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("adjacent-emoji", &bytes);
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let first = TextLoc {
            para_id: "11111111".to_owned(),
            offset: 0,
        };
        document
            .docx_select_point(first.clone(), false, false)
            .unwrap();
        let paragraph_start = docx_selection_range(&document).start;
        document
            .docx_select_point(TextLoc { offset: 2, ..first }, false, true)
            .unwrap();
        assert_eq!(
            docx_selection_range(&document),
            StoryRange::new("body", paragraph_start + 2, paragraph_start + 4)
        );
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn cross_paragraph_alignment_selection_is_mixed() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>Left</w:t></w:r></w:p><w:p w14:paraId="22222222"><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Right</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("mixed-alignment", &bytes);
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 1,
                },
                false,
                false,
            )
            .unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "22222222".to_owned(),
                    offset: 1,
                },
                true,
                false,
            )
            .unwrap();
        let state = document.editing_state().unwrap();
        assert_eq!(state.alignment, None);
        assert_eq!(state.alignment_state, ToggleState::Mixed);
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn backspace_after_text_deletes_only_the_trailing_inline_embed() {
        let source = test_fixtures::write_docx(
            "inline-embed-backspace",
            &test_fixtures::inline_embed_docx(),
        );
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 10,
                },
                false,
                false,
            )
            .unwrap();
        assert!(document.docx_delete(DeleteDirection::Backward).unwrap());
        assert_eq!(docx_paragraph_text(&document, "11111111"), "Bold link");
        assert_eq!(docx_paragraph_embed_count(&document, "11111111"), 0);
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn delete_before_a_trailing_inline_embed_is_atomic() {
        let source =
            test_fixtures::write_docx("inline-embed-delete", &test_fixtures::inline_embed_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 9,
                },
                false,
                false,
            )
            .unwrap();
        assert!(document.docx_delete(DeleteDirection::Forward).unwrap());
        assert_eq!(docx_paragraph_text(&document, "11111111"), "Bold link");
        assert_eq!(docx_paragraph_embed_count(&document, "11111111"), 0);
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn backspace_maps_story_offsets_past_a_leading_inline_embed() {
        let source = test_fixtures::write_docx(
            "inline-embed-astral-backspace",
            &test_fixtures::inline_embed_docx(),
        );
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "22222222".to_owned(),
                    offset: 4,
                },
                false,
                false,
            )
            .unwrap();
        assert!(document.docx_delete(DeleteDirection::Backward).unwrap());
        assert_eq!(docx_paragraph_text(&document, "22222222"), "a");
        assert_eq!(docx_paragraph_embed_count(&document, "22222222"), 1);
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn arrows_and_words_map_across_an_inline_embed() {
        let source =
            test_fixtures::write_docx("inline-embed-arrows", &test_fixtures::inline_embed_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "33333333".to_owned(),
                    offset: 0,
                },
                false,
                false,
            )
            .unwrap();
        let paragraph_start = docx_selection_range(&document).start;
        for offset in 1..=10 {
            assert!(
                document
                    .docx_move_selection(MoveDirection::Right, false)
                    .unwrap()
            );
            assert_eq!(
                docx_selection_range(&document).start,
                paragraph_start + offset
            );
        }
        assert!(
            document
                .docx_move_selection(MoveDirection::Right, false)
                .unwrap()
        );
        assert_eq!(docx_selection_range(&document).start, paragraph_start + 11);
        document
            .docx_select_point(
                TextLoc {
                    para_id: "33333333".to_owned(),
                    offset: 6,
                },
                false,
                true,
            )
            .unwrap();
        assert_eq!(
            docx_selection_range(&document),
            StoryRange::new("body", paragraph_start + 6, paragraph_start + 10)
        );
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn emoji_round_trips_after_deleting_the_ascii_prefix() {
        let source = test_fixtures::write_docx("emoji-save", &test_fixtures::inline_embed_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "44444444".to_owned(),
                    offset: 0,
                },
                false,
                false,
            )
            .unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "44444444".to_owned(),
                    offset: 6,
                },
                true,
                false,
            )
            .unwrap();
        assert!(document.docx_delete(DeleteDirection::Backward).unwrap());

        let output = test_fixtures::write_docx("emoji-save-output", b"sentinel");
        document.save_docx_to(&output).unwrap();
        let saved = fs::read(&output).unwrap();
        let document_xml =
            String::from_utf8(test_fixtures::part(&saved, "word/document.xml")).unwrap();
        assert!(document_xml.contains("<w:t>😀</w:t>"), "{document_xml}");
        assert!(!document_xml.contains('\u{fffd}'), "{document_xml}");
        let reopened = load_document(&output, 0, 8_192).unwrap();
        assert_eq!(docx_paragraph_text(&reopened, "44444444"), "😀");

        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn full_layout_recovery_preserves_unsaved_docx_edits() {
        let source =
            test_fixtures::write_docx("edit-recovery", &test_fixtures::inline_embed_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "44444444".to_owned(),
                    offset: 8,
                },
                false,
                false,
            )
            .unwrap();
        assert!(document.docx_insert_text("!").unwrap());
        document.pages.clear();

        document.recover_after_edit_error().unwrap();
        assert!(!document.pages.is_empty());
        assert_eq!(docx_paragraph_text(&document, "44444444"), "Hello 😀!");
        let ReferenceDocument::Docx(reference) = &document.reference else {
            panic!("expected DOCX reference");
        };
        assert!(reference.editor.is_dirty());

        let output = test_fixtures::write_docx("edit-recovery-output", b"sentinel");
        document.save_docx_to(&output).unwrap();
        let reopened = load_document(&output, 0, 8_192).unwrap();
        assert_eq!(docx_paragraph_text(&reopened, "44444444"), "Hello 😀!");
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn preserves_untouched_complex_paragraph_and_refuses_editing_it() {
        let source_bytes = test_fixtures::complex_docx();
        let complex_before = test_fixtures::paragraph(&source_bytes, "11111111");
        let source = test_fixtures::write_docx("complex-save", &source_bytes);
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "22222222".to_owned(),
                    offset: 15,
                },
                false,
                false,
            )
            .unwrap();
        assert!(document.docx_insert_text(" edited").unwrap());
        let output = test_fixtures::write_docx("complex-output", b"replace me");
        document.save_docx_to(&output).unwrap();
        let saved = fs::read(&output).unwrap();
        assert_eq!(test_fixtures::paragraph(&saved, "11111111"), complex_before);
        let reopened = load_document(&output, 0, 8_192).unwrap();
        let ReferenceDocument::Docx(reference) = &reopened.reference else {
            panic!("expected DOCX reference");
        };
        let paragraphs = reference.editor.engine().doc().paragraphs("body").unwrap();
        assert_eq!(paragraphs[1].text, "Plain paragraph edited");

        let mut complex_edit = load_document(&source, 0, 8_192).unwrap();
        complex_edit
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 1,
                },
                false,
                false,
            )
            .unwrap();
        assert!(complex_edit.docx_insert_text("x").unwrap());
        let refused = test_fixtures::write_docx("complex-refused", b"sentinel");
        let error = complex_edit.save_docx_to(&refused).unwrap_err().to_string();
        assert!(error.contains("paragraph 1 (11111111)"));
        assert!(error.contains("hyperlinks"));
        assert!(error.contains("note references"));
        assert_eq!(fs::read(&refused).unwrap(), b"sentinel");

        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
        fs::remove_file(refused).unwrap();
    }

    #[test]
    fn derives_toolbar_marks_from_the_engine_selection_context() {
        let source = test_fixtures::write_docx("toolbar-state", &test_fixtures::complex_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 2,
                },
                false,
                false,
            )
            .unwrap();
        let caret = document.editing_state().unwrap();
        assert_eq!(caret.bold, ToggleState::On);
        assert!(!caret.inline_enabled);

        document
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 5,
                },
                true,
                false,
            )
            .unwrap();
        let mixed = document.editing_state().unwrap();
        assert_eq!(mixed.bold, ToggleState::Mixed);
        assert!(mixed.inline_enabled);
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn format_preserves_selection_and_undo_updates_toolbar_state() {
        let source = test_fixtures::write_docx("toolbar-format", &test_fixtures::complex_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let start = TextLoc {
            para_id: "22222222".to_owned(),
            offset: 0,
        };
        let end = TextLoc {
            para_id: "22222222".to_owned(),
            offset: 5,
        };
        document.docx_select_point(start, false, false).unwrap();
        document.docx_select_point(end, true, false).unwrap();
        let before = match &document.reference {
            ReferenceDocument::Docx(reference) => reference.editor.selection_range().unwrap(),
            _ => unreachable!(),
        };
        let initial = document.editing_state().unwrap();
        assert_eq!(initial.bold, ToggleState::Off);
        assert!(!initial.can_undo);

        assert!(document.docx_toggle_format(SimpleFormat::Bold).unwrap());
        let formatted = document.editing_state().unwrap();
        assert_eq!(formatted.bold, ToggleState::On);
        assert!(formatted.can_undo);
        let after = match &document.reference {
            ReferenceDocument::Docx(reference) => reference.editor.selection_range().unwrap(),
            _ => unreachable!(),
        };
        assert_eq!(after, before);

        assert!(document.docx_toggle_format(SimpleFormat::Italic).unwrap());
        let combined = document.editing_state().unwrap();
        assert_eq!(combined.bold, ToggleState::On);
        assert_eq!(combined.italic, ToggleState::On);
        let after_second = match &document.reference {
            ReferenceDocument::Docx(reference) => reference.editor.selection_range().unwrap(),
            _ => unreachable!(),
        };
        assert_eq!(after_second, before);

        assert!(document.docx_undo().unwrap());
        let undone = document.editing_state().unwrap();
        assert_eq!(undone.bold, ToggleState::On);
        assert_eq!(undone.italic, ToggleState::Off);
        assert!(undone.can_undo);
        assert!(undone.can_redo);
        assert!(document.docx_undo().unwrap());
        let restored = document.editing_state().unwrap();
        assert_eq!(restored.bold, ToggleState::Off);
        assert!(!restored.can_undo);
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn alignment_round_trips_through_save_and_reopen() {
        let source = test_fixtures::write_docx("toolbar-alignment", &test_fixtures::complex_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let caret = TextLoc {
            para_id: "22222222".to_owned(),
            offset: 4,
        };
        document
            .docx_select_point(caret.clone(), false, false)
            .unwrap();
        assert!(document.docx_set_alignment(Alignment::Center).unwrap());
        assert_eq!(
            document.editing_state().unwrap().alignment,
            Some(Alignment::Center)
        );

        let output = test_fixtures::write_docx("toolbar-alignment-output", b"sentinel");
        document.save_docx_to(&output).unwrap();
        let mut reopened = load_document(&output, 0, 8_192).unwrap();
        reopened.docx_select_point(caret, false, false).unwrap();
        assert_eq!(
            reopened.editing_state().unwrap().alignment,
            Some(Alignment::Center)
        );
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn inline_format_round_trips_through_save_and_reopen() {
        let source =
            test_fixtures::write_docx("toolbar-inline-save", &test_fixtures::complex_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let start = TextLoc {
            para_id: "22222222".to_owned(),
            offset: 0,
        };
        let end = TextLoc {
            para_id: "22222222".to_owned(),
            offset: 5,
        };
        document
            .docx_select_point(start.clone(), false, false)
            .unwrap();
        document
            .docx_select_point(end.clone(), true, false)
            .unwrap();
        assert!(document.docx_toggle_format(SimpleFormat::Bold).unwrap());

        let output = test_fixtures::write_docx("toolbar-inline-output", b"sentinel");
        document.save_docx_to(&output).unwrap();
        let mut reopened = load_document(&output, 0, 8_192).unwrap();
        reopened.docx_select_point(start, false, false).unwrap();
        reopened.docx_select_point(end, true, false).unwrap();
        assert_eq!(reopened.editing_state().unwrap().bold, ToggleState::On);
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn translates_embedded_and_rotated_image_fixtures() {
        for rotation in [None, Some(45.0)] {
            let source = test_fixtures::write_docx("image", &test_fixtures::image_docx(rotation));
            let document = load_document(&source, 0, 8_192).unwrap();
            let ReferenceDocument::Docx(reference) = &document.reference else {
                panic!("expected DOCX reference");
            };
            let image = reference
                .display_list
                .pages
                .iter()
                .flat_map(|page| &page.primitives)
                .find_map(|primitive| match primitive {
                    docx_layout::display_list::Primitive::Image(image) => Some(image),
                    _ => None,
                })
                .unwrap();
            assert!(image.rel_id.starts_with("data:image/png;base64,"));
            assert_eq!(image.attrs.content_frame.is_some(), rotation.is_some());
            assert_eq!(
                document
                    .pages
                    .iter()
                    .map(|page| page.skipped.total())
                    .sum::<usize>(),
                0
            );
            fs::remove_file(source).unwrap();
        }
    }

    #[test]
    fn loads_showcase_with_native_chart_commands() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let document = load_document(&path, 0, 8_192).unwrap();
        let ReferenceDocument::Xlsx(reference) = &document.reference else {
            panic!("expected XLSX reference");
        };
        assert_eq!(reference.sheet_name, "Dashboard");
        assert_eq!(reference.chart_count, 1);
        assert_eq!(reference.chart_placeholders, 0);
        assert_eq!(document.pages[0].skipped.total(), 0);
    }

    #[test]
    fn edits_a_selected_showcase_cell_and_rebuilds_the_sheet() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let cell = CellRef::parse_a1("B5").unwrap();
        assert!(document.xlsx_select_cell(cell));
        assert!(document.xlsx_move_selection(CellMove::Right));
        assert!(document.xlsx_move_selection(CellMove::Left));
        let selection = document.xlsx_overlay().unwrap().rect;
        assert!(selection.w > 0.0 && selection.h > 0.0);
        assert!(document.xlsx_begin_edit(Some("2048")).unwrap());
        assert_eq!(
            document.xlsx_overlay().unwrap().draft.as_deref(),
            Some("2048")
        );
        assert!(document.xlsx_commit(None).unwrap());

        let ReferenceDocument::Xlsx(reference) = &document.reference else {
            panic!("expected XLSX reference");
        };
        assert_eq!(
            reference
                .editor
                .workbook()
                .cell(reference.editor.sheet(), cell)
                .unwrap()
                .input,
            "2048"
        );
        assert!(
            reference
                .editor
                .display_list()
                .commands
                .iter()
                .any(|command| matches!(command, DrawCmd::Text { text, .. } if text == "2048"))
        );
        let state = document.editing_state().unwrap();
        assert!(state.can_save);
        assert!(state.can_undo);
        assert!(!state.inline_enabled);
        assert!(!state.alignment_enabled);
        assert!(document.status_position(0).contains("B5 = 2048"));
        assert!(document.xlsx_undo().unwrap());
        assert!(!document.xlsx_is_dirty());
        assert!(document.editing_state().unwrap().can_redo);
        let undone = test_fixtures::write_xlsx("undone-output", b"sentinel");
        document.save_xlsx_to(&undone).unwrap();
        assert_eq!(fs::read(&undone).unwrap(), fs::read(&path).unwrap());
        fs::remove_file(undone).unwrap();
        assert!(document.xlsx_redo().unwrap());
        assert!(document.xlsx_is_dirty());
        let ReferenceDocument::Xlsx(reference) = &document.reference else {
            panic!("expected XLSX reference");
        };
        assert_eq!(
            reference
                .editor
                .workbook()
                .cell(reference.editor.sheet(), cell)
                .unwrap()
                .input,
            "2048"
        );
    }

    #[test]
    fn formula_commit_uses_the_engine_recalculation() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let cell = CellRef::parse_a1("B5").unwrap();
        document.xlsx_select_cell(cell);
        document.xlsx_begin_edit(Some("=40+2")).unwrap();
        document.xlsx_commit(None).unwrap();

        let ReferenceDocument::Xlsx(reference) = &document.reference else {
            panic!("expected XLSX reference");
        };
        let calculated = reference
            .editor
            .workbook()
            .sheet(reference.editor.sheet())
            .unwrap()
            .cell(cell)
            .unwrap();
        assert_eq!(calculated.formula.as_deref(), Some("40+2"));
        assert_eq!(calculated.value, CellValue::Number { value: 42.0 });
        assert!(
            reference
                .editor
                .display_list()
                .commands
                .iter()
                .any(|command| matches!(command, DrawCmd::Text { text, .. } if text == "42"))
        );
        assert!(
            !reference
                .editor
                .display_list()
                .commands
                .iter()
                .any(|command| matches!(command, DrawCmd::Text { text, .. } if text == "=40+2"))
        );
    }

    #[test]
    fn escape_discards_the_cell_draft() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let cell = CellRef::parse_a1("B5").unwrap();
        document.xlsx_select_cell(cell);
        let before = match &document.reference {
            ReferenceDocument::Xlsx(reference) => {
                reference
                    .editor
                    .workbook()
                    .cell(reference.editor.sheet(), cell)
                    .unwrap()
                    .input
            }
            _ => panic!("expected XLSX reference"),
        };
        document.xlsx_begin_edit(Some("discarded")).unwrap();
        assert!(document.xlsx_insert_text(" draft"));
        assert!(document.xlsx_cancel_edit());
        let after = match &document.reference {
            ReferenceDocument::Xlsx(reference) => {
                reference
                    .editor
                    .workbook()
                    .cell(reference.editor.sheet(), cell)
                    .unwrap()
                    .input
            }
            _ => panic!("expected XLSX reference"),
        };
        assert_eq!(after, before);
        assert!(document.xlsx_overlay().unwrap().draft.is_none());
    }

    #[test]
    fn saves_reopens_and_preserves_an_untouched_sheet() {
        let source = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let source_bytes = fs::read(&source).unwrap();
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let untouched = match &document.reference {
            ReferenceDocument::Xlsx(reference) => reference
                .editor
                .workbook()
                .sheet(SheetId(2))
                .unwrap()
                .clone(),
            _ => panic!("expected XLSX reference"),
        };
        let cell = CellRef::parse_a1("B5").unwrap();
        document.xlsx_select_cell(cell);
        document.xlsx_begin_edit(Some("1776")).unwrap();
        document.xlsx_commit(None).unwrap();

        let output = test_fixtures::write_xlsx("edited-output", b"sentinel");
        document.save_xlsx_to(&output).unwrap();
        assert!(!document.xlsx_is_dirty());
        assert!(document.editing_state().unwrap().can_undo);
        assert!(document.xlsx_undo().unwrap());
        assert!(document.xlsx_is_dirty());
        assert!(document.xlsx_redo().unwrap());
        assert!(!document.xlsx_is_dirty());
        let reopened = load_document(&output, 0, 8_192).unwrap();
        let ReferenceDocument::Xlsx(reference) = &reopened.reference else {
            panic!("expected XLSX reference");
        };
        assert_eq!(
            reference
                .editor
                .workbook()
                .cell(reference.editor.sheet(), cell)
                .unwrap()
                .input,
            "1776"
        );
        assert_eq!(
            reference.editor.workbook().sheet(SheetId(2)).unwrap(),
            &untouched
        );
        let saved = fs::read(&output).unwrap();
        for part in [
            "xl/worksheets/sheet2.xml",
            "xl/worksheets/sheet3.xml",
            "xl/drawings/drawing1.xml",
            "xl/charts/chart1.xml",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
        ] {
            assert_eq!(
                test_fixtures::part(&source_bytes, part),
                test_fixtures::part(&saved, part),
                "{part} changed"
            );
        }
        assert_eq!(fs::read(&source).unwrap(), source_bytes);
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn refuses_to_flatten_shared_formulas_on_save() {
        let source =
            test_fixtures::write_xlsx("shared-formula", &test_fixtures::shared_formula_xlsx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let cell = CellRef::parse_a1("C1").unwrap();
        document.xlsx_select_cell(cell);
        document.xlsx_begin_edit(Some("edited")).unwrap();
        document.xlsx_commit(None).unwrap();

        let output = test_fixtures::write_xlsx("shared-formula-refused", b"sentinel");
        let error = document.save_xlsx_to(&output).unwrap_err().to_string();
        assert!(error.contains("shared formulas"), "{error}");
        assert!(error.contains("xl/worksheets/sheet1.xml"), "{error}");
        assert_eq!(fs::read(&output).unwrap(), b"sentinel");
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn rejects_sheet_size_over_device_limit() {
        let path = test_fixtures::write_xlsx("large", &test_fixtures::large_xlsx());
        let error = load_document(&path, 0, 8_192).err().unwrap().to_string();
        println!("{error}");
        assert!(error.contains("requested sheet size"));
        assert!(error.contains("ceiling 8192"));
        fs::remove_file(path).unwrap();
    }

    #[test]
    fn loads_demo_presentation_and_audits_positioned_glyphs() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let document = load_document(&path, 0, 8_192).unwrap();
        assert_eq!(
            document.edited_path().unwrap(),
            path.parent().unwrap().join("betteroffice-demo-edited.pptx")
        );
        let ReferenceDocument::Pptx(reference) = &document.reference else {
            panic!("expected PPTX reference");
        };
        assert_eq!(reference.editor.slide_count(), 3);
        assert_eq!(document.pages.len(), 3);
        for summary in reference.editor.summaries() {
            assert!(summary.glyph_audit.glyph_runs > 0);
            assert_eq!(summary.glyph_audit.drifted_caret_stops, 0);
            assert_eq!(summary.glyph_audit.missing_caret_stops, 0);
            assert_eq!(summary.glyph_audit.drifted_widths, 0);
        }
        assert_eq!(document.pages[0].skipped.total(), 0);
        assert_eq!(
            reference.editor.summaries()[0].structured(&document.pages[0].skipped)["primitives"]["image"]
                ["translated"],
            1
        );
        assert_eq!(document.pages[1].skipped.counts["placeholder"], 1);
    }

    #[test]
    fn translates_chart_fixture_sub_primitives() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../crates/pptx-parse/tests/fixtures/chart-deck.pptx");
        let document = load_document(&path, 0, 8_192).unwrap();
        let ReferenceDocument::Pptx(reference) = &document.reference else {
            panic!("expected PPTX reference");
        };
        assert_eq!(reference.editor.slide_count(), 2);
        let summary = reference.editor.summaries()[0].structured(&document.pages[0].skipped);
        assert_eq!(summary["primitives"]["chart"]["translated"], 1);
        assert_eq!(summary["primitives"]["placeholder"]["skipped"], 1);
        assert!(
            reference.editor.summaries()[0]
                .glyph_audit
                .missing_caret_stops
                > 0
        );
    }

    #[test]
    fn pptx_hit_test_places_the_exact_engine_caret_and_types_into_its_story() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        let Some(PptxHit::Text(hit)) = document.pptx_hit_test(
            probe.hit.slide_index,
            f64::from(probe.x),
            f64::from(probe.y),
        ) else {
            panic!("expected PPTX text hit");
        };
        assert_eq!(hit, probe.hit);
        assert!(document.pptx_select_hit(hit.clone()));
        let caret = document.caret_geometry().unwrap().unwrap();
        assert_eq!(caret.page_index, hit.slide_index);
        assert_eq!(caret.x.to_bits(), f64::from(probe.caret_x).to_bits());
        assert_eq!(caret.y.to_bits(), f64::from(probe.caret_y).to_bits());
        assert_eq!(
            caret.height.to_bits(),
            f64::from(probe.caret_height).to_bits()
        );
        assert_eq!(caret.transform, Affine::IDENTITY);

        assert!(document.pptx_insert_text("Native ").unwrap());
        let ReferenceDocument::Pptx(reference) = &document.reference else {
            panic!("expected PPTX reference");
        };
        let story = reference.editor.story(&hit.story_id).unwrap();
        assert!(story.plain_text().contains("Native "));
        assert_eq!(
            reference.editor.caret_position(),
            Some(hit.position + "Native ".encode_utf16().count() as u32)
        );
        let state = document.editing_state().unwrap();
        assert!(state.can_save);
        assert!(state.can_undo);
        assert!(!state.inline_enabled);
        assert!(!state.alignment_enabled);
    }

    #[test]
    fn pptx_transformed_caret_uses_the_exact_engine_stop() {
        let source =
            test_fixtures::write_pptx("pptx-transformed-caret", &test_fixtures::transformed_pptx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let probe = pptx_probe_with_transform(&document, 0, true);
        assert!(document.pptx_select_hit(probe.hit));
        let caret = document.caret_geometry().unwrap().unwrap();
        assert_eq!(caret.x.to_bits(), f64::from(probe.caret_x).to_bits());
        assert_eq!(caret.y.to_bits(), f64::from(probe.caret_y).to_bits());
        assert_eq!(
            caret.height.to_bits(),
            f64::from(probe.caret_height).to_bits()
        );
        assert_eq!(caret.transform, probe.transform);
        let actual_top = caret.transform * Point::new(caret.x, caret.y);
        let engine_top =
            probe.transform * Point::new(f64::from(probe.caret_x), f64::from(probe.caret_y));
        assert_eq!(actual_top.x.to_bits(), engine_top.x.to_bits());
        assert_eq!(actual_top.y.to_bits(), engine_top.y.to_bits());
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn pptx_arrow_navigation_stops_at_text_box_edges() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        for direction in [
            MoveDirection::Left,
            MoveDirection::Right,
            MoveDirection::Up,
            MoveDirection::Down,
        ] {
            assert!(document.pptx_select_hit(probe.hit.clone()));
            let mut stopped = false;
            for _ in 0..10_000 {
                if !document.pptx_move_caret(direction) {
                    stopped = true;
                    break;
                }
            }
            assert!(stopped);
            assert!(!document.pptx_move_caret(direction));
            let ReferenceDocument::Pptx(reference) = &document.reference else {
                panic!("expected PPTX reference");
            };
            assert_eq!(
                reference.editor.caret_story_id(),
                Some(probe.hit.story_id.as_str())
            );
        }
    }

    #[test]
    fn pptx_backspace_and_delete_remove_engine_caret_intervals() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        let before = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert!(document.pptx_select_hit(probe.hit.clone()));
        assert!(document.pptx_delete(DeleteDirection::Forward).unwrap());
        let after_delete = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert!(after_delete.length < before.length);

        assert!(document.pptx_undo().unwrap());
        assert!(document.pptx_select_hit(probe.hit.clone()));
        assert!(document.pptx_move_caret(MoveDirection::Right));
        assert!(document.pptx_delete(DeleteDirection::Backward).unwrap());
        let after_backspace = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert!(after_backspace.length < before.length);
    }

    #[test]
    fn pptx_enter_splits_and_backspace_at_the_start_merges_the_paragraph() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        let before = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert!(document.pptx_select_hit(probe.hit.clone()));
        assert!(document.pptx_enter().unwrap());
        let split = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                assert_eq!(
                    reference.editor.caret_position(),
                    Some(probe.hit.position + 1)
                );
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert_eq!(split.paragraphs.len(), before.paragraphs.len() + 1);
        assert_eq!(
            split.plain_text().replace('\n', ""),
            before.plain_text().replace('\n', "")
        );

        assert!(document.pptx_delete(DeleteDirection::Backward).unwrap());
        let merged = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert_eq!(merged.paragraphs.len(), before.paragraphs.len());
        assert_eq!(merged.plain_text(), before.plain_text());
    }

    #[test]
    fn pptx_save_reopens_text_and_keeps_untouched_slides_byte_identical() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let source = fs::read(&path).unwrap();
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        assert!(document.pptx_select_hit(probe.hit.clone()));
        assert!(document.pptx_insert_text("Persisted ").unwrap());

        let output = test_fixtures::write_pptx("pptx-edited-output", b"sentinel");
        document.save_pptx_to(&output).unwrap();
        assert!(!document.pptx_is_dirty());
        assert!(document.editing_state().unwrap().can_undo);
        assert!(document.pptx_undo().unwrap());
        assert!(document.pptx_is_dirty());
        assert!(document.pptx_redo().unwrap());
        assert!(!document.pptx_is_dirty());
        let saved = fs::read(&output).unwrap();
        for slide in ["ppt/slides/slide2.xml", "ppt/slides/slide3.xml"] {
            assert_eq!(
                test_fixtures::part(&source, slide),
                test_fixtures::part(&saved, slide),
                "{slide}"
            );
        }
        assert_eq!(fs::read(&path).unwrap(), source);
        let reopened = load_document(&output, 0, 8_192).unwrap();
        let ReferenceDocument::Pptx(reference) = &reopened.reference else {
            panic!("expected PPTX reference");
        };
        assert!(
            reference
                .editor
                .story(&probe.hit.story_id)
                .unwrap()
                .plain_text()
                .contains("Persisted ")
        );
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn pptx_save_refuses_to_flatten_unchanged_run_color_metadata() {
        let source =
            test_fixtures::write_pptx("pptx-theme-linked", &test_fixtures::theme_linked_pptx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        assert!(document.pptx_select_hit(probe.hit));
        assert!(document.pptx_insert_text("Theme ").unwrap());
        let output = test_fixtures::write_pptx("pptx-theme-refused", b"sentinel");

        let error = document.save_pptx_to(&output).unwrap_err().to_string();
        assert!(error.contains("run color metadata"), "{error}");
        assert!(error.contains("unchanged text"), "{error}");
        assert_eq!(fs::read(&output).unwrap(), b"sentinel");
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn pptx_save_round_trips_a_safe_paragraph_split() {
        let source = test_fixtures::write_pptx(
            "pptx-language-neutral",
            &test_fixtures::language_neutral_pptx(),
        );
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        let before = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert!(document.pptx_select_hit(probe.hit.clone()));
        assert!(document.pptx_move_caret(MoveDirection::Right));
        assert!(document.pptx_enter().unwrap());
        let output = test_fixtures::write_pptx("pptx-safe-split", b"sentinel");
        document.save_pptx_to(&output).unwrap();

        let reopened = load_document(&output, 0, 8_192).unwrap();
        let ReferenceDocument::Pptx(reference) = &reopened.reference else {
            panic!("expected PPTX reference");
        };
        let after = reference.editor.story(&probe.hit.story_id).unwrap();
        assert_eq!(after.paragraphs.len(), before.paragraphs.len() + 1);
        assert_eq!(
            after.plain_text().replace('\n', ""),
            before.plain_text().replace('\n', "")
        );
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn pptx_save_refuses_a_paragraph_split_that_would_drop_run_language() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        assert!(document.pptx_select_hit(probe.hit));
        assert!(document.pptx_move_caret(MoveDirection::Right));
        assert!(document.pptx_enter().unwrap());
        let output = test_fixtures::write_pptx("pptx-language-refused", b"sentinel");

        let error = document.save_pptx_to(&output).unwrap_err().to_string();
        assert!(error.contains("run language"), "{error}");
        assert!(error.contains("unchanged text"), "{error}");
        assert_eq!(fs::read(&output).unwrap(), b"sentinel");
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn pptx_save_refuses_a_digital_signature_before_touching_the_output() {
        let source = test_fixtures::write_pptx("pptx-signed", &test_fixtures::signed_pptx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        assert!(document.pptx_select_hit(probe.hit));
        assert!(document.pptx_insert_text("Signed ").unwrap());
        let output = test_fixtures::write_pptx("pptx-signed-output", b"sentinel");

        let error = document.save_pptx_to(&output).unwrap_err().to_string();
        assert!(error.contains("digital signature"), "{error}");
        assert!(error.contains("_xmlsignatures/sig1.xml"), "{error}");
        assert_eq!(fs::read(&output).unwrap(), b"sentinel");
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn edits_saves_and_reopens_docx_through_the_viewer_path() {
        let source =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.docx");
        let source_before = fs::read(&source).unwrap();
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let (para_id, before, paragraph_count) = match &document.reference {
            ReferenceDocument::Docx(reference) => {
                let paragraphs = reference.editor.engine().doc().paragraphs("body").unwrap();
                let paragraph = paragraphs.first().unwrap();
                (
                    paragraph.para_id.clone(),
                    paragraph.text.clone(),
                    paragraphs.len(),
                )
            }
            _ => panic!("expected DOCX reference"),
        };
        assert_eq!(before, "Welcome to BetterOffice");

        let caret = match &document.reference {
            ReferenceDocument::Docx(reference) => reference
                .editor
                .engine()
                .resident_caret_snapshot(Some((&para_id, 7)))
                .unwrap()
                .caret_rect
                .unwrap(),
            _ => panic!("expected DOCX reference"),
        };
        let loc = document
            .docx_hit_test(caret.page_index, caret.x, caret.y + caret.height * 0.5)
            .unwrap()
            .unwrap();
        assert_eq!(loc.para_id, para_id);
        assert_eq!(loc.offset, 7);
        document.docx_select_point(loc, false, false).unwrap();
        assert!(document.docx_insert_text(" native").unwrap());

        let output = std::env::temp_dir().join(format!(
            "betteroffice-native-viewer-{}-{}.docx",
            std::process::id(),
            TEST_FILE_ID.fetch_add(1, Ordering::Relaxed)
        ));
        document.save_docx_to(&output).unwrap();
        let reopened = load_document(&output, 0, 8_192).unwrap();
        let (after, reopened_count) = match &reopened.reference {
            ReferenceDocument::Docx(reference) => {
                let paragraphs = reference.editor.engine().doc().paragraphs("body").unwrap();
                (paragraphs.first().unwrap().text.clone(), paragraphs.len())
            }
            _ => panic!("expected DOCX reference"),
        };
        assert_eq!(after, "Welcome native to BetterOffice");
        assert_eq!(reopened_count, paragraph_count);
        assert_eq!(fs::read(&source).unwrap(), source_before);
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn applies_delete_and_enter_through_the_viewer_path() {
        let source =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.docx");
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let (para_id, paragraph_count) = match &document.reference {
            ReferenceDocument::Docx(reference) => {
                let paragraphs = reference.editor.engine().doc().paragraphs("body").unwrap();
                (paragraphs[0].para_id.clone(), paragraphs.len())
            }
            _ => panic!("expected DOCX reference"),
        };
        let caret = TextLoc { para_id, offset: 7 };

        document
            .docx_select_point(caret.clone(), false, false)
            .unwrap();
        let first_line = document.caret_geometry().unwrap().unwrap();
        assert!(
            document
                .docx_move_selection(MoveDirection::Down, false)
                .unwrap()
        );
        let next_line = document.caret_geometry().unwrap().unwrap();
        assert!(
            next_line.page_index > first_line.page_index
                || (next_line.page_index == first_line.page_index && next_line.y > first_line.y)
        );
        assert!(
            document
                .docx_move_selection(MoveDirection::Up, false)
                .unwrap()
        );
        document
            .docx_select_point(caret.clone(), false, false)
            .unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: caret.para_id.clone(),
                    offset: 10,
                },
                true,
                false,
            )
            .unwrap();
        assert!(!document.docx_selection_rects().is_empty());
        document
            .docx_select_point(
                TextLoc {
                    para_id: caret.para_id.clone(),
                    offset: 8,
                },
                false,
                true,
            )
            .unwrap();
        assert!(!document.docx_selection_rects().is_empty());
        document
            .docx_select_point(caret.clone(), false, false)
            .unwrap();
        assert!(
            document
                .docx_move_selection(MoveDirection::Home, false)
                .unwrap()
        );
        let home = document.caret_geometry().unwrap().unwrap();
        assert!(
            document
                .docx_move_selection(MoveDirection::End, false)
                .unwrap()
        );
        let end = document.caret_geometry().unwrap().unwrap();
        assert_eq!(home.page_index, end.page_index);
        assert!(end.x > home.x);
        document
            .docx_select_point(caret.clone(), false, false)
            .unwrap();
        assert!(document.docx_delete(DeleteDirection::Forward).unwrap());
        assert!(document.docx_insert_text(" ").unwrap());
        assert!(document.docx_insert_text("😀").unwrap());
        assert!(document.docx_delete(DeleteDirection::Backward).unwrap());
        document.docx_select_point(caret, false, false).unwrap();
        assert!(document.docx_enter().unwrap());
        assert!(document.caret_geometry().unwrap().is_some());

        let paragraphs = match &document.reference {
            ReferenceDocument::Docx(reference) => {
                reference.editor.engine().doc().paragraphs("body").unwrap()
            }
            _ => panic!("expected DOCX reference"),
        };
        assert_eq!(paragraphs.len(), paragraph_count + 1);
        assert_eq!(paragraphs[0].text, "Welcome");
        assert_eq!(paragraphs[1].text, " to BetterOffice");

        let output = std::env::temp_dir().join(format!(
            "betteroffice-native-viewer-{}-{}.docx",
            std::process::id(),
            TEST_FILE_ID.fetch_add(1, Ordering::Relaxed)
        ));
        fs::write(&output, b"sentinel").unwrap();
        let error = document.save_docx_to(&output).unwrap_err().to_string();
        assert!(error.contains("cannot save DOCX structural edits"));
        assert_eq!(fs::read(&output).unwrap(), b"sentinel");

        assert!(document.docx_delete(DeleteDirection::Backward).unwrap());
        let paragraphs = match &document.reference {
            ReferenceDocument::Docx(reference) => {
                reference.editor.engine().doc().paragraphs("body").unwrap()
            }
            _ => panic!("expected DOCX reference"),
        };
        assert_eq!(paragraphs.len(), paragraph_count);
        assert_eq!(paragraphs[0].text, "Welcome to BetterOffice");

        document.save_docx_to(&output).unwrap();
        let reopened = load_document(&output, 0, 8_192).unwrap();
        let reopened_paragraphs = match &reopened.reference {
            ReferenceDocument::Docx(reference) => {
                reference.editor.engine().doc().paragraphs("body").unwrap()
            }
            _ => panic!("expected DOCX reference"),
        };
        assert_eq!(reopened_paragraphs.len(), paragraph_count);
        assert_eq!(reopened_paragraphs[0].text, "Welcome to BetterOffice");
        fs::remove_file(output).unwrap();
    }
}
