use crate::display_list::PaintTransform;

/// CSS's conventional 96 pixels per inch, applied only at display-list paint time.
pub const PIXELS_PER_INCH: f32 = 96.0;

pub fn final_paint_transform(page_height_inches: f32) -> PaintTransform {
    PaintTransform {
        a: PIXELS_PER_INCH,
        b: 0.0,
        c: 0.0,
        d: -PIXELS_PER_INCH,
        e: 0.0,
        f: page_height_inches * PIXELS_PER_INCH,
    }
}
pub fn to_canvas(transform: PaintTransform, x: f32, y: f32) -> (f32, f32) {
    (
        transform.a * x + transform.c * y + transform.e,
        transform.b * x + transform.d * y + transform.f,
    )
}

pub fn to_canvas_length(transform: PaintTransform, length: f32) -> f32 {
    length * transform.a.hypot(transform.b)
}
