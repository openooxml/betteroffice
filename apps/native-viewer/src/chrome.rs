use std::sync::LazyLock;

use rustybuzz::{Face, UnicodeBuffer};
use vello::kurbo::{Affine, Line, Point, Rect, Stroke};
use vello::peniko::{Blob, Color, Fill, FontData};
use vello::{Glyph, Scene};

pub const TOOLBAR_HEIGHT: f64 = 52.0;
pub const STATUS_HEIGHT: f64 = 28.0;
pub const ZOOM_MIN: f64 = 0.25;
pub const ZOOM_MAX: f64 = 5.0;

const FONT_BYTES: &[u8] =
    include_bytes!("../../../packages/fonts/assets/LiberationSans-Regular.ttf");
const BOLD_FONT_BYTES: &[u8] =
    include_bytes!("../../../packages/fonts/assets/LiberationSans-Bold.ttf");
const FONT_SIZE: f32 = 13.0;
const ICON_SIZE: f32 = 16.0;

static FACE: LazyLock<Face<'static>> =
    LazyLock::new(|| Face::from_slice(FONT_BYTES, 0).expect("embedded font is valid"));
static BOLD_FACE: LazyLock<Face<'static>> =
    LazyLock::new(|| Face::from_slice(BOLD_FONT_BYTES, 0).expect("embedded font is valid"));

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ToggleState {
    Off,
    On,
    Mixed,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
}

impl Alignment {
    #[cfg(feature = "docx")]
    pub fn from_engine(value: Option<&str>) -> Option<Self> {
        match value {
            Some("left" | "start") => Some(Self::Left),
            Some("center") => Some(Self::Center),
            Some("right" | "end") => Some(Self::Right),
            Some("justify" | "both" | "distribute") => Some(Self::Justify),
            _ => None,
        }
    }

    #[cfg(feature = "docx")]
    pub fn engine_value(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Justify => "justify",
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct EditingState {
    pub bold: ToggleState,
    pub italic: ToggleState,
    pub underline: ToggleState,
    pub alignment: Option<Alignment>,
    pub alignment_state: ToggleState,
    pub inline_enabled: bool,
    pub alignment_enabled: bool,
    pub can_undo: bool,
    pub can_redo: bool,
    pub can_save: bool,
    pub save_disabled_reason: Option<String>,
}

impl EditingState {
    pub fn read_only() -> Self {
        Self {
            bold: ToggleState::Off,
            italic: ToggleState::Off,
            underline: ToggleState::Off,
            alignment: None,
            alignment_state: ToggleState::Off,
            inline_enabled: false,
            alignment_enabled: false,
            can_undo: false,
            can_redo: false,
            can_save: false,
            save_disabled_reason: None,
        }
    }

    pub fn editable_without_selection(can_undo: bool, can_redo: bool) -> Self {
        Self {
            can_undo,
            can_redo,
            can_save: true,
            ..Self::read_only()
        }
    }
}

#[derive(Clone, Debug)]
pub struct ChromeState {
    pub editing: EditingState,
    pub zoom: f64,
    pub file_name: String,
    pub position: String,
    pub message: Option<String>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ToolbarCommand {
    Undo,
    Redo,
    Bold,
    Italic,
    Underline,
    Align(Alignment),
    ZoomOut,
    ResetZoom,
    ZoomIn,
    Save,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ChromeHit {
    Miss,
    Consumed(Option<ToolbarCommand>),
}

#[derive(Clone, Copy)]
struct Button {
    command: ToolbarCommand,
    bounds: Rect,
}

struct ChromeLayout {
    toolbar: Rect,
    status: Rect,
    buttons: Vec<Button>,
    separators: Vec<f64>,
}

pub struct Chrome {
    font: FontData,
    bold_font: FontData,
}

impl Chrome {
    pub fn new() -> Self {
        Self {
            font: FontData::new(Blob::from(FONT_BYTES.to_vec()), 0),
            bold_font: FontData::new(Blob::from(BOLD_FONT_BYTES.to_vec()), 0),
        }
    }

    pub fn hit_test(
        &self,
        width: f64,
        height: f64,
        point: Point,
        state: &ChromeState,
    ) -> ChromeHit {
        let layout = ChromeLayout::new(width, height);
        if layout.toolbar.contains(point) {
            let command = layout
                .buttons
                .iter()
                .find(|button| button.bounds.contains(point))
                .map(|button| button.command)
                .filter(|command| command_enabled(*command, state));
            return ChromeHit::Consumed(command);
        }
        if layout.status.contains(point) {
            return ChromeHit::Consumed(None);
        }
        ChromeHit::Miss
    }

    pub fn paint(
        &self,
        scene: &mut Scene,
        width: f64,
        height: f64,
        device_scale: f64,
        state: &ChromeState,
    ) {
        let layout = ChromeLayout::new(width, height);
        let transform = Affine::scale(device_scale);
        scene.fill(
            Fill::NonZero,
            transform,
            Color::from_rgba8(248, 249, 250, 255),
            None,
            &layout.toolbar,
        );
        scene.stroke(
            &Stroke::new(1.0),
            transform,
            Color::from_rgba8(218, 220, 224, 255),
            None,
            &Line::new((0.0, TOOLBAR_HEIGHT - 0.5), (width, TOOLBAR_HEIGHT - 0.5)),
        );
        for separator in &layout.separators {
            scene.stroke(
                &Stroke::new(1.0),
                transform,
                Color::from_rgba8(218, 220, 224, 255),
                None,
                &Line::new((*separator, 13.0), (*separator, TOOLBAR_HEIGHT - 13.0)),
            );
        }
        for button in &layout.buttons {
            self.paint_button(scene, transform, button, state);
        }
        self.paint_status(scene, transform, &layout, state);
    }

    fn paint_button(
        &self,
        scene: &mut Scene,
        transform: Affine,
        button: &Button,
        state: &ChromeState,
    ) {
        let enabled = command_enabled(button.command, state);
        let selected = command_selection(button.command, state);
        let background = match selected {
            ToggleState::On => Color::from_rgba8(210, 227, 252, 255),
            ToggleState::Mixed => Color::from_rgba8(226, 232, 240, 255),
            ToggleState::Off => Color::from_rgba8(255, 255, 255, 255),
        };
        let border = if selected == ToggleState::On {
            Color::from_rgba8(138, 180, 248, 255)
        } else {
            Color::from_rgba8(218, 220, 224, 255)
        };
        scene.fill(Fill::NonZero, transform, background, None, &button.bounds);
        scene.stroke(&Stroke::new(1.0), transform, border, None, &button.bounds);
        if selected == ToggleState::Mixed {
            let bar = Rect::new(
                button.bounds.x0 + 7.0,
                button.bounds.y1 - 5.0,
                button.bounds.x1 - 7.0,
                button.bounds.y1 - 3.0,
            );
            scene.fill(
                Fill::NonZero,
                transform,
                Color::from_rgba8(95, 99, 104, 255),
                None,
                &bar,
            );
        }
        let color = if enabled {
            Color::from_rgba8(32, 33, 36, 255)
        } else {
            Color::from_rgba8(154, 160, 166, 255)
        };
        match button.command {
            ToolbarCommand::Undo => self.paint_centered_text(
                scene,
                transform,
                button.bounds,
                "↶",
                ICON_SIZE,
                false,
                false,
                color,
            ),
            ToolbarCommand::Redo => self.paint_centered_text(
                scene,
                transform,
                button.bounds,
                "↷",
                ICON_SIZE,
                false,
                false,
                color,
            ),
            ToolbarCommand::Bold => self.paint_centered_text(
                scene,
                transform,
                button.bounds,
                "B",
                ICON_SIZE,
                true,
                false,
                color,
            ),
            ToolbarCommand::Italic => self.paint_centered_text(
                scene,
                transform,
                button.bounds,
                "I",
                ICON_SIZE,
                false,
                true,
                color,
            ),
            ToolbarCommand::Underline => {
                self.paint_centered_text(
                    scene,
                    transform,
                    button.bounds,
                    "U",
                    ICON_SIZE,
                    false,
                    false,
                    color,
                );
                scene.stroke(
                    &Stroke::new(1.2),
                    transform,
                    color,
                    None,
                    &Line::new(
                        (
                            button.bounds.center().x - 5.0,
                            button.bounds.center().y + 7.0,
                        ),
                        (
                            button.bounds.center().x + 5.0,
                            button.bounds.center().y + 7.0,
                        ),
                    ),
                );
            }
            ToolbarCommand::Align(alignment) => {
                paint_alignment(scene, transform, button.bounds, alignment, color)
            }
            ToolbarCommand::ZoomOut => self.paint_centered_text(
                scene,
                transform,
                button.bounds,
                "−",
                18.0,
                false,
                false,
                color,
            ),
            ToolbarCommand::ResetZoom => self.paint_centered_text(
                scene,
                transform,
                button.bounds,
                &format!("{:.0}%", state.zoom * 100.0),
                FONT_SIZE,
                false,
                false,
                color,
            ),
            ToolbarCommand::ZoomIn => self.paint_centered_text(
                scene,
                transform,
                button.bounds,
                "+",
                18.0,
                false,
                false,
                color,
            ),
            ToolbarCommand::Save => self.paint_centered_text(
                scene,
                transform,
                button.bounds,
                "Save",
                FONT_SIZE,
                false,
                false,
                color,
            ),
        }
    }

    fn paint_status(
        &self,
        scene: &mut Scene,
        transform: Affine,
        layout: &ChromeLayout,
        state: &ChromeState,
    ) {
        scene.fill(
            Fill::NonZero,
            transform,
            Color::from_rgba8(248, 249, 250, 255),
            None,
            &layout.status,
        );
        scene.stroke(
            &Stroke::new(1.0),
            transform,
            Color::from_rgba8(218, 220, 224, 255),
            None,
            &Line::new(
                (0.0, layout.status.y0 + 0.5),
                (layout.status.x1, layout.status.y0 + 0.5),
            ),
        );
        let baseline = layout.status.y0 + 18.0;
        let width = layout.status.width();
        let file_width = (width * 0.34).clamp(120.0, 300.0);
        let position_width = (width * 0.18).clamp(100.0, 180.0);
        let message_x = 24.0 + file_width + position_width;
        self.paint_fitted_text(
            scene,
            transform,
            Point::new(12.0, baseline),
            file_width,
            &state.file_name,
            Color::from_rgba8(60, 64, 67, 255),
        );
        self.paint_fitted_text(
            scene,
            transform,
            Point::new(18.0 + file_width, baseline),
            position_width,
            &state.position,
            Color::from_rgba8(95, 99, 104, 255),
        );
        if let Some(message) = &state.message {
            self.paint_fitted_text(
                scene,
                transform,
                Point::new(message_x, baseline),
                (width - message_x - 12.0).max(0.0),
                message,
                Color::from_rgba8(26, 115, 232, 255),
            );
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn paint_centered_text(
        &self,
        scene: &mut Scene,
        transform: Affine,
        bounds: Rect,
        text: &str,
        size: f32,
        bold: bool,
        italic: bool,
        color: Color,
    ) {
        let prepared = PreparedText::new(text, size, bold);
        let x = bounds.center().x - f64::from(prepared.width) / 2.0;
        let baseline = bounds.center().y + f64::from(size) * 0.36;
        prepared.paint(
            scene,
            if bold { &self.bold_font } else { &self.font },
            transform * Affine::translate((x, baseline)),
            italic,
            color,
        );
    }

    fn paint_fitted_text(
        &self,
        scene: &mut Scene,
        transform: Affine,
        origin: Point,
        max_width: f64,
        text: &str,
        color: Color,
    ) {
        if max_width <= 0.0 {
            return;
        }
        let fitted = fitted_text(text, max_width as f32, FONT_SIZE, false);
        PreparedText::new(&fitted, FONT_SIZE, false).paint(
            scene,
            &self.font,
            transform * Affine::translate((origin.x, origin.y)),
            false,
            color,
        );
    }
}

impl Default for Chrome {
    fn default() -> Self {
        Self::new()
    }
}

impl ChromeLayout {
    fn new(width: f64, height: f64) -> Self {
        let compact = width < 600.0;
        let button_width = if compact { 28.0 } else { 32.0 };
        let zoom_width = if compact { 44.0 } else { 52.0 };
        let save_width = if compact { 48.0 } else { 52.0 };
        let gap = 4.0;
        let mut buttons = Vec::new();
        let mut separators = Vec::new();
        let mut x = 12.0;
        let mut push = |command, width: f64, x: &mut f64| {
            buttons.push(Button {
                command,
                bounds: Rect::new(*x, 10.0, *x + width, 42.0),
            });
            *x += width + gap;
        };
        push(ToolbarCommand::Undo, button_width, &mut x);
        push(ToolbarCommand::Redo, button_width, &mut x);
        separators.push(x + 2.0);
        x += 12.0;
        push(ToolbarCommand::Bold, button_width, &mut x);
        push(ToolbarCommand::Italic, button_width, &mut x);
        push(ToolbarCommand::Underline, button_width, &mut x);
        separators.push(x + 2.0);
        x += 12.0;
        for alignment in [
            Alignment::Left,
            Alignment::Center,
            Alignment::Right,
            Alignment::Justify,
        ] {
            push(ToolbarCommand::Align(alignment), button_width, &mut x);
        }
        let right_width = button_width * 2.0 + zoom_width + save_width + gap * 3.0 + 12.0;
        x = (width - 12.0 - right_width).max(x + 12.0);
        push(ToolbarCommand::ZoomOut, button_width, &mut x);
        push(ToolbarCommand::ResetZoom, zoom_width, &mut x);
        push(ToolbarCommand::ZoomIn, button_width, &mut x);
        separators.push(x + 2.0);
        x += 12.0;
        push(ToolbarCommand::Save, save_width, &mut x);
        Self {
            toolbar: Rect::new(0.0, 0.0, width.max(0.0), TOOLBAR_HEIGHT),
            status: Rect::new(
                0.0,
                (height - STATUS_HEIGHT).max(TOOLBAR_HEIGHT),
                width.max(0.0),
                height.max(TOOLBAR_HEIGHT),
            ),
            buttons,
            separators,
        }
    }
}

struct PreparedText {
    glyphs: Vec<Glyph>,
    width: f32,
    size: f32,
}

impl PreparedText {
    fn new(text: &str, size: f32, bold: bool) -> Self {
        let face = if bold { &*BOLD_FACE } else { &*FACE };
        let scale = size / face.units_per_em() as f32;
        let mut buffer = UnicodeBuffer::new();
        buffer.push_str(text);
        let shaped = rustybuzz::shape(face, &[], buffer);
        let mut pen = 0.0;
        let glyphs = shaped
            .glyph_infos()
            .iter()
            .zip(shaped.glyph_positions())
            .map(|(info, position)| {
                let glyph = Glyph {
                    id: info.glyph_id,
                    x: pen + position.x_offset as f32 * scale,
                    y: -(position.y_offset as f32) * scale,
                };
                pen += position.x_advance as f32 * scale;
                glyph
            })
            .collect();
        Self {
            glyphs,
            width: pen,
            size,
        }
    }

    fn paint(
        &self,
        scene: &mut Scene,
        font: &FontData,
        transform: Affine,
        italic: bool,
        color: Color,
    ) {
        scene
            .draw_glyphs(font)
            .font_size(self.size)
            .brush(color)
            .glyph_transform(italic.then_some(Affine::skew(0.18, 0.0)))
            .transform(transform)
            .draw(Fill::NonZero, self.glyphs.iter().copied());
    }
}

fn fitted_text(text: &str, max_width: f32, size: f32, bold: bool) -> String {
    if PreparedText::new(text, size, bold).width <= max_width {
        return text.to_owned();
    }
    let ellipsis = "…";
    let ellipsis_width = PreparedText::new(ellipsis, size, bold).width;
    let mut output = String::new();
    for character in text.chars() {
        let mut candidate = output.clone();
        candidate.push(character);
        if PreparedText::new(&candidate, size, bold).width + ellipsis_width > max_width {
            break;
        }
        output = candidate;
    }
    output.push_str(ellipsis);
    output
}

fn command_enabled(command: ToolbarCommand, state: &ChromeState) -> bool {
    match command {
        ToolbarCommand::Undo => state.editing.can_undo,
        ToolbarCommand::Redo => state.editing.can_redo,
        ToolbarCommand::Bold | ToolbarCommand::Italic | ToolbarCommand::Underline => {
            state.editing.inline_enabled
        }
        ToolbarCommand::Align(_) => state.editing.alignment_enabled,
        ToolbarCommand::ZoomOut => state.zoom > ZOOM_MIN + f64::EPSILON,
        ToolbarCommand::ResetZoom => true,
        ToolbarCommand::ZoomIn => state.zoom < ZOOM_MAX - f64::EPSILON,
        ToolbarCommand::Save => state.editing.can_save,
    }
}

fn command_selection(command: ToolbarCommand, state: &ChromeState) -> ToggleState {
    match command {
        ToolbarCommand::Bold => state.editing.bold,
        ToolbarCommand::Italic => state.editing.italic,
        ToolbarCommand::Underline => state.editing.underline,
        ToolbarCommand::Align(_) if state.editing.alignment_state == ToggleState::Mixed => {
            ToggleState::Mixed
        }
        ToolbarCommand::Align(alignment) if state.editing.alignment == Some(alignment) => {
            state.editing.alignment_state
        }
        _ => ToggleState::Off,
    }
}

fn paint_alignment(
    scene: &mut Scene,
    transform: Affine,
    bounds: Rect,
    alignment: Alignment,
    color: Color,
) {
    let center = bounds.center();
    for (index, width) in [14.0, 10.0, 14.0, 8.0].into_iter().enumerate() {
        let x = match alignment {
            Alignment::Left => center.x - 7.0,
            Alignment::Center => center.x - width / 2.0,
            Alignment::Right => center.x + 7.0 - width,
            Alignment::Justify => center.x - 7.0,
        };
        let width = if alignment == Alignment::Justify {
            14.0
        } else {
            width
        };
        let y = center.y - 6.0 + index as f64 * 4.0;
        scene.stroke(
            &Stroke::new(1.4),
            transform,
            color,
            None,
            &Line::new((x, y), (x + width, y)),
        );
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn state(editing: EditingState) -> ChromeState {
        ChromeState {
            editing,
            zoom: 1.0,
            file_name: "document.docx".to_owned(),
            position: "Page 1 of 2".to_owned(),
            message: None,
        }
    }

    #[test]
    fn disabled_controls_are_consumed_without_commands() {
        let chrome = Chrome::new();
        let state = state(EditingState::read_only());
        let layout = ChromeLayout::new(800.0, 600.0);
        let bold = layout
            .buttons
            .iter()
            .find(|button| button.command == ToolbarCommand::Bold)
            .unwrap();
        assert_eq!(
            chrome.hit_test(800.0, 600.0, bold.bounds.center(), &state),
            ChromeHit::Consumed(None)
        );
    }

    #[test]
    fn chrome_scales_layout_in_logical_coordinates() {
        let state = state(EditingState::editable_without_selection(false, false));
        let chrome = Chrome::new();
        assert_eq!(
            chrome.hit_test(520.0, 420.0, Point::new(4.0, 4.0), &state),
            ChromeHit::Consumed(None)
        );
        assert_eq!(
            chrome.hit_test(900.0, 700.0, Point::new(4.0, 100.0), &state),
            ChromeHit::Miss
        );
        assert_eq!(
            chrome.hit_test(900.0, 700.0, Point::new(4.0, 699.0), &state),
            ChromeHit::Consumed(None)
        );
    }

    #[test]
    fn enabled_controls_emit_their_toolbar_commands() {
        let editing = EditingState {
            bold: ToggleState::Mixed,
            italic: ToggleState::Off,
            underline: ToggleState::On,
            alignment: Some(Alignment::Center),
            alignment_state: ToggleState::On,
            inline_enabled: true,
            alignment_enabled: true,
            can_undo: true,
            can_redo: true,
            can_save: true,
            save_disabled_reason: None,
        };
        let state = state(editing);
        let chrome = Chrome::new();
        let layout = ChromeLayout::new(900.0, 700.0);
        for expected in [
            ToolbarCommand::Undo,
            ToolbarCommand::Redo,
            ToolbarCommand::Bold,
            ToolbarCommand::Italic,
            ToolbarCommand::Underline,
            ToolbarCommand::Align(Alignment::Left),
            ToolbarCommand::Align(Alignment::Center),
            ToolbarCommand::Align(Alignment::Right),
            ToolbarCommand::Align(Alignment::Justify),
            ToolbarCommand::ZoomOut,
            ToolbarCommand::ResetZoom,
            ToolbarCommand::ZoomIn,
            ToolbarCommand::Save,
        ] {
            let button = layout
                .buttons
                .iter()
                .find(|button| button.command == expected)
                .unwrap();
            assert_eq!(
                chrome.hit_test(900.0, 700.0, button.bounds.center(), &state),
                ChromeHit::Consumed(Some(expected))
            );
        }
    }

    #[test]
    fn mixed_paragraph_alignment_marks_every_alignment_control_mixed() {
        let mut editing = EditingState::editable_without_selection(false, false);
        editing.alignment_enabled = true;
        editing.alignment_state = ToggleState::Mixed;
        let state = state(editing);
        for alignment in [
            Alignment::Left,
            Alignment::Center,
            Alignment::Right,
            Alignment::Justify,
        ] {
            assert_eq!(
                command_selection(ToolbarCommand::Align(alignment), &state),
                ToggleState::Mixed
            );
        }
    }
}
