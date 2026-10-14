//! Sandboxed SVG rasterisation into the straight-alpha RGBA buffer the other
//! image formats produce. The sandbox bounds what conversion and rendering cost
//! by construction: the bytes are scanned before `roxmltree` parses them, the
//! document is audited before `usvg` builds anything, and the tree `usvg` builds
//! is priced before `resvg` paints it. Every bound below is a share of one
//! envelope, and a document past any of them is refused.

mod audit;
mod cost;
mod geometry;
mod markup;
mod reference;
mod sanitize;
mod style;

use std::panic::{AssertUnwindSafe, catch_unwind};

use resvg::usvg;
use tiny_skia::{IntSize, Pixmap, PremultipliedColorU8, Transform};

use crate::MAX_IMAGE_PIXELS;

/// Peak memory one SVG decode may use beyond its output raster and group
/// layers, which the caller charges to the slide's image budget.
pub const SVG_MEMORY_ENVELOPE: u64 = 100 << 20;
/// Single-core time one SVG decode may take, in nanoseconds of a release
/// build; the work units below are each about a nanosecond.
pub const SVG_TIME_ENVELOPE: u64 = 1_000_000_000;

/// Bytes `roxmltree` keeps per node or attribute, its final shrink included.
const XML_ITEM_BYTES: u64 = 80;
/// Bytes of `usvg`'s element and tree node per expanded element; a `use` of a
/// `symbol` with its viewport clip, the costliest, measures about 970.
const NODE_BYTES: u64 = 1_024;
/// Bytes per expanded markup byte: `usvg` stores each attribute instance in
/// 32 bytes and each byte of path data as up to 4.5 bytes of points.
const BYTE_BYTES: u64 = 8;
/// Bytes of the declaration lists `simplecss` copies per selector and byte of
/// a block: 40 per declaration of at least four bytes.
const STYLE_COPY_BYTES: u64 = 10;
/// Bytes of outline per byte of path data, which `usvg` strokes whole to
/// measure a shape: about 190 per segment with round joins, and at most one
/// segment per byte.
const STROKE_BYTES: u64 = 256;
/// Bytes of path data each cubic an arc becomes is charged as: its 25 bytes
/// of points, and the copies `svgtypes` makes draining up to 64 of them, stay
/// within what 8 bytes buy at [`BYTE_BYTES`] and [`BYTE_NS`].
const ARC_CUBIC_BYTES: u64 = 8;
/// Time the byte scan, both `roxmltree` parses, the sanitizer and the audit
/// take per source byte: about 27 ns measured over 4 MiB of path data.
const SOURCE_NS: u64 = 32;
/// Time `usvg` takes to build one expanded element.
const NODE_NS: u64 = 1_024;
/// Time `usvg` takes per expanded markup byte: path data parses and strokes
/// its outline for bounds at about 30 ns a byte.
const BYTE_NS: u64 = 32;
/// Time `resvg` takes per unit of [`MAX_SVG_RENDER_WORK`]: an anti-aliased,
/// alpha-blended pixel.
const RENDER_NS: u64 = 4;
/// Time the stroker takes per piece it may emit: 35 to 40 ns a piece measured
/// on long outlines, which the bound on pieces overstates at least twofold.
const VERB_NS: u64 = 32;

/// One document's bytes, twice what [`MAX_SVG_EXPANDED_BYTES`] lets it draw
/// so an editor's private markup, which the sanitizer drops, still fits: the
/// byte scan, both parses, the sanitizer and the audit take about 134 ms over
/// them, and `roxmltree` copies at most as many bytes again.
pub const MAX_SVG_BYTES: usize = 1 << 22;
/// The document `usvg` reads, re-serialised from the source, which is held
/// alongside it while `usvg` converts it: 3 MiB of the memory envelope, room
/// for everything [`MAX_SVG_EXPANDED_BYTES`] lets a document draw once.
pub const MAX_SVG_SANITIZED_BYTES: usize = 3 << 20;
/// Bytes of any one attribute value but path data, points and `style`, as
/// written: every inherited value is parsed anew for each element under it.
pub const MAX_SVG_VALUE_BYTES: usize = 1_024;
/// Elements one document may nest, in its markup or once its references are
/// expanded. The parsers, `usvg`'s converter and `resvg` recurse over it at
/// about 3 KiB of stack a level: 64 levels fit a 512 KiB thread stack twice.
pub const MAX_SVG_DEPTH: usize = 64;
/// Nodes `roxmltree` will build, and the `<` it reserves one for before
/// parsing: 10 MiB of the memory envelope.
pub const MAX_SVG_NODES: u32 = 1 << 17;
/// Attributes across the markup, and the `=` `roxmltree` reserves one for:
/// 10 MiB of the memory envelope.
pub const MAX_SVG_ATTRIBUTES: usize = 1 << 17;
/// Attributes on one element. `roxmltree` compares each with every earlier
/// one, so the markup costs at most 64 comparisons an attribute.
pub const MAX_SVG_ELEMENT_ATTRIBUTES: usize = 64;
/// Namespace declarations in scope at one element, and distinct ones in the
/// document: `roxmltree` looks each element and prefixed attribute up across
/// those in scope, and inserts each distinct one into a sorted list.
pub const MAX_SVG_NAMESPACES: usize = 64;
/// Elements a document expands into once every `use` and clip path
/// instantiates its target: 32 MiB of the memory envelope and 34 ms.
pub const MAX_SVG_EXPANDED_NODES: u64 = 1 << 15;
/// Markup bytes expansion hands `usvg`, which re-reads a target's attributes
/// and path data once per instance, and walks its markup once per `use`:
/// 16 MiB of the memory envelope and 67 ms.
pub const MAX_SVG_EXPANDED_BYTES: u64 = 1 << 21;
/// Path data one shape may carry, each cubic its arcs become weighed as
/// [`ARC_CUBIC_BYTES`], which `usvg` strokes whole to measure it: 16 MiB of
/// outline, the envelope's share for any one path's transients.
pub const MAX_SVG_PATH_BYTES: usize = 1 << 16;
/// That share: what the edges, outline and dashes of any one path may take
/// while `usvg` measures it or `resvg` paints it.
pub(super) const SVG_TRANSIENT_BYTES: u64 = MAX_SVG_PATH_BYTES as u64 * STROKE_BYTES;
/// Stops one gradient may carry. `usvg` drops equal offsets by shifting the
/// list, quadratic in its length, and `tiny-skia` tests every stop per pixel.
pub const MAX_SVG_GRADIENT_STOPS: usize = 256;
/// Bytes of the gradient copies `usvg` makes per shape: one for every shape
/// or `use` that paints with a gradient in its own box's units, each ~256
/// bytes plus 12 a stop. 6 MiB of the envelope.
pub const MAX_SVG_PAINT_BYTES: u64 = 6 << 20;
/// What `usvg` spends collecting distinct gradients and clip paths, which it
/// compares each reference against every one collected so far at about a
/// quarter nanosecond each: 34 ms of the time envelope.
pub const MAX_SVG_COLLECT_WORK: u64 = 1 << 25;
/// `<style>` elements plus the rules they declare. `simplecss` re-sorts every
/// rule after each sheet, so the sorts stay under ten million comparisons.
pub const MAX_SVG_STYLE_RULES: usize = 1_024;
/// Simple selectors (a type, `.class` or `#id`) one rule may compound, each
/// an attribute lookup when `simplecss` tests the rule.
pub const MAX_SVG_SELECTOR_PARTS: usize = 16;
/// Bytes of one selector, which bounds the names each lookup compares.
pub const MAX_SVG_SELECTOR_BYTES: usize = 256;
/// Selectors times the bytes of the block they share: `simplecss` copies a
/// block's declarations once per selector of its list. 2.5 MiB of the
/// envelope.
pub const MAX_SVG_STYLE_COPIES: u64 = 1 << 18;
/// What `simplecss` and `usvg` spend on CSS, in units of about a nanosecond:
/// rescans of every stylesheet and `style` attribute, selector tests, and
/// declarations applied, across the expanded document. 101 ms of the time
/// envelope.
pub const MAX_SVG_STYLE_WORK: u64 = 3 << 25;
/// Pieces the stroker may emit while `usvg` strokes every shape instance whole
/// to measure it, bounded before conversion; and, separately, the pieces it
/// emits while the render is priced by stroking what `resvg` will stroke,
/// each stroke started only while its bound still fits. 50 ms each.
pub const MAX_SVG_STROKE_VERBS: u64 = 3 << 19;
/// What `usvg` spends looking up inherited properties, in nanoseconds: each
/// instance looks through every ancestor's attributes up to the root. 34 ms
/// of the time envelope.
pub const MAX_SVG_INHERIT_WORK: u64 = 1 << 25;
/// How far from the origin, in multiples of its tolerance, the stroker may
/// meet a curve or reach with its width. Past `2^17` tolerances an `f32` step
/// is more than a sixty-fourth of one, and the stroker, splitting until its
/// arithmetic agrees with itself, emits millions of pieces for one curve.
pub const MAX_SVG_STROKE_SPAN: f64 = 131_072.0;
/// Group layers (opacity, clip, blend, isolation) one render may stack. Each
/// is a raster charged to the slide's image budget with the output.
pub const MAX_SVG_LAYER_DEPTH: usize = 8;
/// Painted pixels, gradient stops weighted in, as a multiple of the output
/// raster: a cheap first refusal before [`MAX_SVG_RENDER_WORK`] is summed.
pub const MAX_SVG_OVERDRAW: u64 = 64;
/// Painting work in painted-pixel units, each about 4 ns: 470 ms, the rest of
/// the time envelope. Past it the raster supersamples less, and a render still
/// past it at its intrinsic size is refused.
pub const MAX_SVG_RENDER_WORK: u64 = 7 * MAX_IMAGE_PIXELS / 2;
/// One rasterised SVG's longest side, and any layer's. The raster is outside
/// the envelope: the caller charges it, with its layers, to the slide's image
/// budget. Past 8191 `tiny-skia` draws a pixmap in tiles, every path per tile.
pub const MAX_SVG_RASTER_DIM: u32 = 8_191;
/// How far above its intrinsic size an SVG rasterises, so a picture frame
/// larger than the document still has pixels to stretch.
const SVG_SUPERSAMPLE: u32 = 4;

const _: () = assert!(
    MAX_SVG_BYTES as u64
        + MAX_SVG_SANITIZED_BYTES as u64
        + (MAX_SVG_NODES as u64 + MAX_SVG_ATTRIBUTES as u64) * XML_ITEM_BYTES
        + MAX_SVG_EXPANDED_NODES * NODE_BYTES
        + MAX_SVG_EXPANDED_BYTES * BYTE_BYTES
        + MAX_SVG_PATH_BYTES as u64 * STROKE_BYTES
        + MAX_SVG_PAINT_BYTES
        + MAX_SVG_STYLE_COPIES * STYLE_COPY_BYTES
        <= SVG_MEMORY_ENVELOPE
);
const _: () = assert!(
    MAX_SVG_BYTES as u64 * SOURCE_NS
        + MAX_SVG_EXPANDED_NODES * NODE_NS
        + MAX_SVG_EXPANDED_BYTES * BYTE_NS
        + MAX_SVG_STYLE_WORK
        + MAX_SVG_COLLECT_WORK
        + 2 * MAX_SVG_STROKE_VERBS * VERB_NS
        + MAX_SVG_INHERIT_WORK
        + MAX_SVG_RENDER_WORK * RENDER_NS
        <= SVG_TIME_ENVELOPE
);

/// Why the sandbox declined a document. Structural only: no markup, no
/// attribute value and no reference target is carried out of the decoder.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SvgRefusal {
    /// Not UTF-8, or the root element is not `svg`.
    NotSvg,
    /// Past [`MAX_SVG_BYTES`], or markup past [`MAX_SVG_NODES`],
    /// [`MAX_SVG_ELEMENT_ATTRIBUTES`], [`MAX_SVG_ATTRIBUTES`] or
    /// [`MAX_SVG_NAMESPACES`].
    DocumentTooLarge,
    /// Carries a `<!DOCTYPE>`, so it may declare entities.
    DoctypeDeclared,
    /// Past [`MAX_SVG_DEPTH`], in the markup or once references are expanded.
    TooDeeplyNested,
    /// References something outside itself: a network or filesystem href, or a
    /// `url()` that is not a same-document fragment.
    ExternalReference,
    /// Malformed, past [`MAX_SVG_NODES`], or without a usable intrinsic size.
    Unparsable,
    /// Rasterises past [`MAX_SVG_RASTER_DIM`].
    RasterTooLarge,
    /// An element outside the drawing allowlist, an attribute in the SVG, XLink
    /// or XML namespace other than `xlink:href`, `xlink:title`, `xml:space` and
    /// `xml:lang`, a reference into content the audit skips, or a mask, filter,
    /// pattern, image or text node in the tree.
    UnsupportedElement,
    /// A stylesheet past [`MAX_SVG_STYLE_RULES`] or beyond plain type, class and
    /// id rules within [`MAX_SVG_SELECTOR_PARTS`] and [`MAX_SVG_SELECTOR_BYTES`],
    /// a `filter` in any form, `inherit` in any CSS, or a clip path inherited
    /// from whatever element a copy lands under.
    UnsupportedStyle,
    /// A reference that leads back to itself.
    ReferenceCycle,
    /// Past [`MAX_SVG_EXPANDED_NODES`], [`MAX_SVG_EXPANDED_BYTES`],
    /// [`MAX_SVG_STYLE_WORK`], [`MAX_SVG_STYLE_COPIES`], [`MAX_SVG_PAINT_BYTES`],
    /// [`MAX_SVG_COLLECT_WORK`], [`MAX_SVG_STROKE_VERBS`] or
    /// [`MAX_SVG_INHERIT_WORK`] once references are expanded; a gradient past
    /// [`MAX_SVG_GRADIENT_STOPS`], a shape past [`MAX_SVG_PATH_BYTES`] or an
    /// arc of more than 64 cubics; a curved shape, stroke width or dash list in
    /// relative units; or, where anything is stroked, a rotation or skew, or a
    /// curve or width past [`MAX_SVG_STROKE_SPAN`].
    ExpansionTooLarge,
    /// Group layers past [`MAX_SVG_LAYER_DEPTH`], or a clip path that is
    /// itself clipped.
    TooManyLayers,
    /// Paints past [`MAX_SVG_OVERDRAW`], or even at its intrinsic size past
    /// [`MAX_SVG_RENDER_WORK`] or with a path whose edges, outline and dashes
    /// would outgrow the envelope's share for one path; or a stroke too wide
    /// for the rasteriser's fixed point.
    RenderTooCostly,
    /// `usvg` or `resvg` panicked.
    Panicked,
}

/// A parsed SVG and the raster it will render into.
pub struct SvgImage {
    tree: usvg::Tree,
    size: IntSize,
    pixels: u64,
}

/// Whether `bytes` are worth handing to [`parse`]. `image`'s sniffer never
/// claims SVG, so this is the only thing that routes a picture here.
pub fn looks_like_svg(bytes: &[u8]) -> bool {
    let head = &bytes[..bytes.len().min(1024)];
    let head = head.strip_prefix(&[0xef, 0xbb, 0xbf]).unwrap_or(head);
    let start = head.iter().position(|byte| !byte.is_ascii_whitespace());
    start.is_some_and(|start| head[start] == b'<')
        && head.windows(4).any(|window| window == b"<svg")
}

/// Parses under the sandbox: no DTD, no external reference, an allowlisted and
/// bounded document, and a render priced before it runs.
#[cfg(any(test, feature = "fuzzing"))]
pub fn parse(bytes: &[u8]) -> Result<SvgImage, SvgRefusal> {
    parse_within(bytes, MAX_IMAGE_PIXELS)
}

/// [`parse`], supersampling only as far as `pixels` of raster and layers.
pub(crate) fn parse_within(bytes: &[u8], pixels: u64) -> Result<SvgImage, SvgRefusal> {
    if bytes.len() > MAX_SVG_BYTES {
        return Err(SvgRefusal::DocumentTooLarge);
    }
    let text = std::str::from_utf8(bytes).map_err(|_| SvgRefusal::NotSvg)?;
    markup::scan(bytes)?;
    let document = usvg::roxmltree::Document::parse_with_options(
        text,
        usvg::roxmltree::ParsingOptions {
            allow_dtd: false,
            nodes_limit: MAX_SVG_NODES,
            ..Default::default()
        },
    )
    .map_err(|error| match error {
        usvg::roxmltree::Error::DtdDetected => SvgRefusal::DoctypeDeclared,
        _ => SvgRefusal::Unparsable,
    })?;
    if document.root_element().tag_name().name() != "svg" {
        return Err(SvgRefusal::NotSvg);
    }
    let sanitized = sanitize::sanitize(&document)?;
    drop(document);
    markup::scan(sanitized.as_bytes())?;
    let document = usvg::roxmltree::Document::parse_with_options(
        &sanitized,
        usvg::roxmltree::ParsingOptions {
            allow_dtd: false,
            nodes_limit: MAX_SVG_NODES,
            ..Default::default()
        },
    )
    .map_err(|_| SvgRefusal::Unparsable)?;
    audit::audit(&document)?;
    let tree = guarded(|| usvg::Tree::from_xmltree(&document, &sandbox()))?
        .map_err(|_| SvgRefusal::Unparsable)?;
    let (size, pixels) = guarded(|| raster(&tree, pixels))??;
    Ok(SvgImage { tree, size, pixels })
}

impl SvgImage {
    /// Pixels the render will allocate, for the caller's decode budget: the
    /// output raster plus the deepest stack of group layers and clip masks.
    pub fn pixels(&self) -> u64 {
        self.pixels
    }

    /// The document's own size in CSS px, which a tiled fill repeats at.
    pub fn intrinsic(&self) -> (f32, f32) {
        (self.tree.size().width(), self.tree.size().height())
    }

    /// Straight-alpha RGBA, matching what the raster formats hand back.
    pub fn render(&self) -> Result<(Vec<u8>, IntSize), SvgRefusal> {
        let mut pixmap =
            Pixmap::new(self.size.width(), self.size.height()).ok_or(SvgRefusal::RasterTooLarge)?;
        let scale = Transform::from_scale(
            self.size.width() as f32 / self.tree.size().width(),
            self.size.height() as f32 / self.tree.size().height(),
        );
        guarded(|| resvg::render(&self.tree, scale, &mut pixmap.as_mut()))?;
        let mut data = pixmap.take();
        let (pixels, _) = data.as_chunks_mut::<4>();
        for pixel in pixels {
            let straight = PremultipliedColorU8::from_rgba(pixel[0], pixel[1], pixel[2], pixel[3])
                .map(|color| color.demultiply());
            if let Some(color) = straight {
                *pixel = [color.red(), color.green(), color.blue(), color.alpha()];
            }
        }
        Ok((data, self.size))
    }
}

/// Neither resolver may reach a file or a network, whatever the audit missed.
fn sandbox() -> usvg::Options<'static> {
    usvg::Options {
        resources_dir: None,
        image_href_resolver: usvg::ImageHrefResolver {
            resolve_data: Box::new(|_, _, _| None),
            resolve_string: Box::new(|_, _| None),
        },
        ..usvg::Options::default()
    }
}

/// Runs a `usvg` or `resvg` step with a panic mapped to a refusal. The panic
/// hook is left alone, so the panic is still reported the usual way.
fn guarded<T>(step: impl FnOnce() -> T) -> Result<T, SvgRefusal> {
    catch_unwind(AssertUnwindSafe(step)).map_err(|_| SvgRefusal::Panicked)
}

/// The raster an intrinsic size renders into and the pixels it allocates: the
/// largest supersample whose output and layers fit `budget` and whose
/// painting fits [`MAX_SVG_RENDER_WORK`]. Sizes are priced first without
/// stroking anything; the largest that fits and then the intrinsic one are
/// priced again with their strokes drawn, out of [`MAX_SVG_STROKE_VERBS`]
/// pieces split between the two. At the intrinsic size only the painting
/// refuses; a raster too large for the budget is the caller's to skip.
fn raster(tree: &usvg::Tree, budget: u64) -> Result<(IntSize, u64), SvgRefusal> {
    let width = tree.size().width();
    let height = tree.size().height();
    if !width.is_finite() || !height.is_finite() || width <= 0.0 || height <= 0.0 {
        return Err(SvgRefusal::Unparsable);
    }
    let limit = MAX_SVG_RASTER_DIM as f32;
    let (width, height) = (width.ceil().max(1.0), height.ceil().max(1.0));
    if width > limit || height > limit {
        return Err(SvgRefusal::RasterTooLarge);
    }
    let mut sizes = Vec::new();
    for factor in (1..=SVG_SUPERSAMPLE).rev() {
        let (scaled_width, scaled_height) = (width * factor as f32, height * factor as f32);
        if factor == 1 || (scaled_width <= limit && scaled_height <= limit) {
            let size = IntSize::from_wh(scaled_width as u32, scaled_height as u32);
            sizes.push(size.ok_or(SvgRefusal::Unparsable)?);
        }
    }
    let intrinsic = sizes[sizes.len() - 1];
    let fits = |cost: &cost::Cost, size: IntSize| {
        cost.work <= MAX_SVG_RENDER_WORK && (cost.pixels <= budget || size == intrinsic)
    };
    let mut first = intrinsic;
    for &size in &sizes {
        if fits(&cost::measure(tree, size, None)?, size) {
            first = size;
            break;
        }
    }
    let share = (MAX_SVG_STROKE_VERBS / 2) as f64;
    let mut strokes = share;
    for size in [first, intrinsic] {
        let cost = cost::measure(tree, size, Some(&mut strokes))?;
        if fits(&cost, size) {
            return Ok((size, cost.pixels));
        }
        if size == intrinsic {
            break;
        }
        strokes += share;
    }
    Err(SvgRefusal::RenderTooCostly)
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;

    fn refusal(bytes: &[u8]) -> Option<SvgRefusal> {
        parse(bytes).err()
    }

    pub(crate) fn document(body: &str) -> String {
        format!(r##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96">{body}</svg>"##)
    }

    #[test]
    fn a_bounded_document_rasterizes_above_its_intrinsic_size() {
        let source = document(r##"<rect width="96" height="96" fill="#00ff00"/>"##);
        let image = parse(source.as_bytes()).expect("parse");
        let (data, size) = image.render().expect("render");
        assert_eq!((size.width(), size.height()), (384, 384));
        assert_eq!(image.pixels(), 384 * 384);
        assert_eq!(data.len(), 384 * 384 * 4);
        assert_eq!(&data[..4], &[0, 255, 0, 255]);
    }

    #[test]
    fn an_intrinsic_size_at_the_bound_rasterizes_without_supersampling() {
        let source = format!(
            r##"<svg xmlns="http://www.w3.org/2000/svg" width="{dim}" height="{dim}"><rect width="{dim}" height="{dim}" fill="#000"/></svg>"##,
            dim = MAX_SVG_RASTER_DIM
        );
        let size = parse(source.as_bytes()).expect("parse").size;
        assert_eq!((size.width(), size.height()), (8191, 8191));
        let past = source.replace("8191", "8192");
        assert_eq!(refusal(past.as_bytes()), Some(SvgRefusal::RasterTooLarge));
    }

    #[test]
    fn a_layer_wider_than_a_tile_steps_the_supersample_down() {
        let layer = |width: u32| {
            format!(
                r##"<svg xmlns="http://www.w3.org/2000/svg" width="4000" height="10"><g opacity="0.5"><rect x="-{width}" width="{}" height="10" fill="#000"/></g></svg>"##,
                2 * width + 4000
            )
        };
        let wide = parse(layer(2_000).as_bytes()).expect("parse").size;
        assert_eq!(wide.width(), 4_000, "an 8000-pixel layer fits one tile");
        assert_eq!(
            refusal(layer(4_100).as_bytes()),
            Some(SvgRefusal::RenderTooCostly),
            "a 12200-pixel layer would be drawn in tiles"
        );
    }

    #[test]
    fn a_budget_steps_the_supersample_down_before_the_caller_refuses() {
        let source = document(r##"<rect width="96" height="96" fill="#000"/>"##);
        let image = parse_within(source.as_bytes(), 96 * 96 * 4).expect("parse");
        assert_eq!((image.size.width(), image.pixels()), (192, 192 * 192));
        let image = parse_within(source.as_bytes(), 1).expect("parse");
        assert_eq!(
            image.size.width(),
            96,
            "the intrinsic size is the caller's to skip"
        );
    }

    #[test]
    fn a_document_past_the_byte_bound_is_refused() {
        let padding = " ".repeat(MAX_SVG_BYTES);
        let source = document(&format!("<desc>{padding}</desc>"));
        assert_eq!(
            refusal(source.as_bytes()),
            Some(SvgRefusal::DocumentTooLarge)
        );
    }

    #[test]
    fn nesting_past_the_depth_bound_is_refused() {
        let nest =
            |depth: usize| document(&format!("{}{}", "<g>".repeat(depth), "</g>".repeat(depth)));
        assert_eq!(
            refusal(nest(MAX_SVG_DEPTH + 1).as_bytes()),
            Some(SvgRefusal::TooDeeplyNested)
        );
        assert!(parse(nest(MAX_SVG_DEPTH - 1).as_bytes()).is_ok());
    }

    #[test]
    fn siblings_comments_and_quoted_angle_brackets_do_not_count_as_nesting() {
        let body = concat!(
            r##"<!-- <g><g><g> --><![CDATA[<g><g>]]>"##,
            r##"<desc title="a &gt; b"/><path d="M 0 0"/><path d="M 1 1"/>"##
        );
        let source = document(&body.repeat(64));
        assert!(parse(source.as_bytes()).is_ok());
    }

    #[test]
    fn a_prologue_that_mentions_a_tag_is_not_the_root_element() {
        let prologue = concat!(
            r##"<!-- exported by <svg width="1"><g><g><g> -->"##,
            r##"<?xml-stylesheet href="a.css" type="text/css"?>"##
        );
        let body = r##"<rect width="96" height="96" fill="#00ff00"/>"##;
        assert!(parse(format!("{prologue}{}", document(body)).as_bytes()).is_ok());
        assert_eq!(
            refusal(format!("{prologue}<html><svg/></html>").as_bytes()),
            Some(SvgRefusal::NotSvg)
        );
    }

    #[test]
    fn a_declared_entity_is_refused_before_it_can_expand() {
        let source = concat!(
            r##"<!DOCTYPE svg [<!ENTITY secret SYSTEM "file:///etc/passwd">]>"##,
            r##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96"><desc>&secret;</desc></svg>"##
        );
        assert_eq!(
            refusal(source.as_bytes()),
            Some(SvgRefusal::DoctypeDeclared)
        );
    }

    #[test]
    fn a_reference_outside_the_document_is_refused() {
        for body in [
            r##"<use href="https://example.invalid/sprite.svg#icon"/>"##,
            r##"<linearGradient xmlns:xlink="http://www.w3.org/1999/xlink" id="g" xlink:href="../../secret.svg#g"/>"##,
            r##"<rect width="96" height="96" fill="url(https://example.invalid/paint.svg#g)"/>"##,
            r##"<rect width="96" height="96" style="fill:url('/etc/paint.svg#g')"/>"##,
            r##"<rect width="96" height="96" fill="URL(http://example.invalid/paint.svg#g)"/>"##,
            r##"<rect width="96" height="96" style="clip-path:Url( data:image/svg+xml,x)"/>"##,
            r##"<style>rect{fill:uRl(https://example.invalid/paint.svg#g)}</style><rect width="96" height="96"/>"##,
        ] {
            assert_eq!(
                refusal(document(body).as_bytes()),
                Some(SvgRefusal::ExternalReference),
                "{body}"
            );
        }
    }

    #[test]
    fn a_same_document_reference_is_kept() {
        let body = concat!(
            r##"<defs><linearGradient id="g"><stop offset="0" stop-color="#f00"/></linearGradient></defs>"##,
            r##"<rect width="96" height="96" fill="url(#g)"/><use href="#g"/>"##
        );
        assert!(parse(document(body).as_bytes()).is_ok());
    }

    #[test]
    fn an_output_raster_past_the_dimension_bound_is_refused() {
        let source = r##"<svg xmlns="http://www.w3.org/2000/svg" width="100000" height="100000"><rect width="100000" height="100000" fill="#000"/></svg>"##;
        assert_eq!(refusal(source.as_bytes()), Some(SvgRefusal::RasterTooLarge));
    }

    #[test]
    fn malformed_and_foreign_documents_are_refused_without_panicking() {
        assert_eq!(refusal(b"<svg><g></svg>"), Some(SvgRefusal::Unparsable));
        assert_eq!(refusal(b"<html><svg/></html>"), Some(SvgRefusal::NotSvg));
        assert_eq!(refusal(&[0xff, 0xfe, 0x3c, 0x73]), Some(SvgRefusal::NotSvg));
        assert_eq!(
            refusal(br#"<svg xmlns="urn:not-svg" viewBox="0 0 4 4"/>"#),
            Some(SvgRefusal::NotSvg)
        );
    }

    #[test]
    fn only_a_document_that_opens_as_markup_reaches_the_parser() {
        assert!(looks_like_svg(
            b"  <?xml version=\"1.0\"?><svg xmlns=\"x\"/>"
        ));
        assert!(looks_like_svg(b"\xef\xbb\xbf<svg/>"));
        assert!(!looks_like_svg(b"\x89PNG\r\n\x1a\n"));
        assert!(!looks_like_svg(b"not markup, mentions <svg> in prose"));
        assert!(!looks_like_svg(b""));
    }

    pub(crate) fn marker_chain(vertices: usize, levels: usize) -> String {
        let mut d = String::from("M0 0");
        for index in 1..vertices {
            d.push_str(&format!(" L{} {}", index % 10, index / 10));
        }
        let mut defs = String::from(
            r##"<marker id="m0" markerWidth="1" markerHeight="1" overflow="visible"><rect width="1" height="1" fill="#f00"/></marker>"##,
        );
        for level in 1..=levels {
            defs.push_str(&format!(
                r##"<marker id="m{level}" markerWidth="1" markerHeight="1" overflow="visible" markerUnits="userSpaceOnUse"><path d="{d}" fill="none" stroke="#000" marker-mid="url(#m{})"/></marker>"##,
                level - 1
            ));
        }
        document(&format!(
            r##"<defs>{defs}</defs><path d="{d}" fill="none" stroke="#000" marker-mid="url(#m{levels})"/>"##
        ))
    }

    #[test]
    fn a_marker_chain_is_refused_before_usvg_multiplies_it() {
        assert_eq!(
            refusal(marker_chain(6, 12).as_bytes()),
            Some(SvgRefusal::UnsupportedElement)
        );
    }

    #[test]
    fn a_marker_sized_to_overflow_is_refused_instead_of_panicking() {
        let source = concat!(
            r##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">"##,
            r##"<marker id="m" markerWidth="1e30" markerHeight="1e30" viewBox="0 0 1 1"><rect width="1" height="1"/></marker>"##,
            r##"<path d="M0 0L5 5" stroke="#000" stroke-width="1e30" marker-end="url(#m)"/></svg>"##
        );
        assert_eq!(
            refusal(source.as_bytes()),
            Some(SvgRefusal::UnsupportedElement)
        );
    }

    #[test]
    fn elements_outside_the_allowlist_are_refused() {
        for body in [
            r##"<filter id="f"><feGaussianBlur stdDeviation="9"/></filter>"##,
            r##"<mask id="m"><rect width="9" height="9"/></mask>"##,
            r##"<pattern id="p" width="1" height="1"><rect width="1" height="1"/></pattern>"##,
            r##"<foreignObject width="9" height="9"/>"##,
            r##"<script>alert(1)</script>"##,
            r##"<switch><rect width="9" height="9"/></switch>"##,
            r##"<text><textPath href="#p">x</textPath></text>"##,
            r##"<rect width="9" height="9"><animate attributeName="x" to="9"/></rect>"##,
            r##"<desc><marker id="m"/></desc>"##,
        ] {
            assert_eq!(
                refusal(document(body).as_bytes()),
                Some(SvgRefusal::UnsupportedElement),
                "{body}"
            );
        }
    }

    #[test]
    fn metadata_and_foreign_markup_are_skipped_but_never_instantiated() {
        let skipped = concat!(
            r##"<metadata><rdf:RDF xmlns:rdf="urn:rdf"><filter id="f"/></rdf:RDF><marker id="m"/></metadata>"##,
            r##"<sodipodi:namedview xmlns:sodipodi="urn:sodipodi" id="base"><marker/></sodipodi:namedview>"##,
            r##"<rect width="96" height="96" fill="#00f"/>"##
        );
        assert!(parse(document(skipped).as_bytes()).is_ok());
        for body in [
            r##"<metadata><g id="x"><rect width="9" height="9"/></g></metadata><use href="#x"/>"##,
            r##"<x:y xmlns:x="urn:x"><svg:g xmlns:svg="http://www.w3.org/2000/svg" id="x"/></x:y><use href="#x"/>"##,
        ] {
            assert_eq!(
                refusal(document(body).as_bytes()),
                Some(SvgRefusal::UnsupportedElement),
                "{body}"
            );
        }
    }

    pub(crate) fn use_fan_out(levels: usize, uses: usize) -> String {
        let mut defs = String::from(r##"<rect id="l0" width="1" height="1" fill="#f00"/>"##);
        for level in 1..=levels {
            let copies = format!(r##"<use href="#l{}"/>"##, level - 1).repeat(uses);
            defs.push_str(&format!(r##"<g id="l{level}">{copies}</g>"##));
        }
        document(&format!(r##"<defs>{defs}</defs><use href="#l{levels}"/>"##))
    }

    #[test]
    fn an_exponential_use_fan_out_is_refused_before_it_expands() {
        assert_eq!(
            refusal(use_fan_out(10, 10).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        assert!(parse(use_fan_out(2, 10).as_bytes()).is_ok());
    }

    pub(crate) fn use_chain(hops: usize) -> String {
        let mut defs = String::from(r##"<g id="g0"><rect width="9" height="9" fill="#f00"/></g>"##);
        for hop in 1..=hops {
            defs.push_str(&format!(
                r##"<g id="g{hop}"><use href="#g{}"/></g>"##,
                hop - 1
            ));
        }
        document(&format!(r##"<defs>{defs}</defs><use href="#g{hops}"/>"##))
    }

    #[test]
    fn a_use_chain_past_the_depth_bound_is_refused() {
        assert_eq!(
            refusal(use_chain(MAX_SVG_DEPTH).as_bytes()),
            Some(SvgRefusal::TooDeeplyNested)
        );
        assert_eq!(
            refusal(use_chain(256).as_bytes()),
            Some(SvgRefusal::TooDeeplyNested)
        );
        assert!(parse(use_chain(8).as_bytes()).is_ok());
    }

    #[test]
    fn a_reference_cycle_is_refused() {
        for body in [
            r##"<g id="a"><use href="#b"/></g><g id="b"><use href="#a"/></g>"##,
            r##"<g id="a"><rect width="9" height="9"/><use href="#a"/></g>"##,
            r##"<linearGradient id="a" href="#b"/><linearGradient id="b" href="#c"/><linearGradient id="c" href="#b"/><rect width="9" height="9" fill="url(#a)"/>"##,
            r##"<clipPath id="c"><rect width="9" height="9" clip-path="url(#c)"/></clipPath>"##,
        ] {
            assert_eq!(
                refusal(document(body).as_bytes()),
                Some(SvgRefusal::ReferenceCycle),
                "{body}"
            );
        }
    }

    #[test]
    fn a_reference_cycle_is_found_through_targets_exactly_as_usvg_reads_them() {
        let clips = |ids: [&str; 3], references: [&str; 3]| {
            let mut body = String::new();
            for (id, reference) in ids.iter().zip(references) {
                body.push_str(&format!(
                    r##"<clipPath id="{id}" {reference}><rect width="9" height="9"/></clipPath>"##
                ));
            }
            body.push_str(&format!(
                r##"<rect width="9" height="9" clip-path="url(#{})"/>"##,
                ids[0]
            ));
            document(&body)
        };
        let gradients = |ids: [&str; 3]| {
            document(&format!(
                concat!(
                    r##"<linearGradient id="a" href="#{b}"/><linearGradient id="{b}" href="#{c}"/>"##,
                    r##"<linearGradient id="{c}" xlink:href="#{b}" xmlns:xlink="http://www.w3.org/1999/xlink"/>"##,
                    r##"<rect width="9" height="9" fill="url(#a)"/>"##
                ),
                b = ids[1],
                c = ids[2]
            ))
        };
        for source in [
            clips(
                ["a1", "b1", "c1"],
                [
                    r##"clip-path="url(#b1)""##,
                    r##"clip-path="url(#c1)""##,
                    r##"clip-path="url(#a1)""##,
                ],
            ),
            clips(
                ["a2", "b2", "c2"],
                [
                    r##"clip-path="url('#b2')""##,
                    r##"clip-path=" url( &quot;#c2 &quot; ) ""##,
                    r##"style="clip-path:url(#a2)""##,
                ],
            ),
            gradients(["a", "b3", "c3"]),
        ] {
            assert_eq!(
                refusal(source.as_bytes()),
                Some(SvgRefusal::ReferenceCycle),
                "{source}"
            );
        }
        for source in [
            clips(
                ["a&#9;", "b&#9;", "c&#9;"],
                [
                    r##"clip-path="url(#b&#9;)""##,
                    r##"clip-path="url(#c&#9;)""##,
                    r##"clip-path="url(#a&#9;)""##,
                ],
            ),
            gradients(["a", "b&#9;", "c&#9;"]),
        ] {
            assert_eq!(
                refusal(source.as_bytes()),
                Some(SvgRefusal::UnsupportedElement),
                "an id that cannot be written back unchanged: {source}"
            );
        }
    }

    #[test]
    fn an_attribute_usvg_would_read_from_a_prefix_is_refused() {
        let radius = "6".repeat(57) + "1";
        let points = "1 1 ".repeat(250_000);
        for body in [
            format!(r#"<path xml:d="M0 0a{radius}1 1 0 0 0 1 1"/>"#),
            r#"<circle xml:r="3e38"/>"#.to_owned(),
            format!(r##"<polygon xml:points="{points}" stroke="#000"/>"##),
            r##"<linearGradient id="g"><stop offset="0"/></linearGradient><g fill="url(#g)"><rect width="1" height="1" fill="inherit" xml:fill="red"/></g>"##.to_owned(),
            r##"<rect xmlns:s="http://www.w3.org/2000/svg" width="1" height="1" s:fill="red"/>"##.to_owned(),
            r##"<rect width="1" height="1" xlink:style="fill:red" xmlns:xlink="http://www.w3.org/1999/xlink"/>"##.to_owned(),
        ] {
            let started = std::time::Instant::now();
            assert_eq!(
                refusal(document(&body).as_bytes()),
                Some(SvgRefusal::UnsupportedElement),
                "{}",
                &body[..body.len().min(120)]
            );
            assert!(started.elapsed() < std::time::Duration::from_secs(30));
        }
        let allowed = concat!(
            r##"<g xml:space="preserve" xml:lang="en"><rect id="r" width="1" height="1"/></g>"##,
            r##"<use xmlns:xlink="http://www.w3.org/1999/xlink" xlink:href="#r" xlink:title="copy"/>"##,
            r##"<g xmlns:x="urn:x" x:fill="url(https://example.invalid/g)"/>"##
        );
        assert!(parse(document(allowed).as_bytes()).is_ok());
    }

    #[test]
    fn a_clip_path_inherited_into_a_copy_is_refused() {
        for body in [
            r##"<g clip-path="url(#t)"><clipPath id="x" clip-path="inherit"><rect width="9" height="9"/></clipPath></g><clipPath id="t"><rect width="9" height="9" clip-path="url(#x)"/></clipPath>"##,
            r##"<rect width="9" height="9" style="clip-path: inherit"/>"##,
            r##"<style>.a{clip-path:inherit}</style><rect class="a" width="9" height="9"/>"##,
        ] {
            assert_eq!(
                refusal(document(body).as_bytes()),
                Some(SvgRefusal::UnsupportedStyle),
                "{body}"
            );
        }
    }

    #[test]
    fn a_reference_reaches_only_the_kind_of_element_usvg_converts_for_it() {
        let body = concat!(
            r##"<linearGradient id="g"><stop offset="0" stop-color="#0f0"/></linearGradient>"##,
            r##"<clipPath id="c"><rect width="96" height="96" fill="url(#g)"/></clipPath>"##,
            r##"<rect width="96" height="96" fill="url(#g)" clip-path="url(#c)"/>"##
        );
        assert!(parse(document(body).as_bytes()).is_ok());
        let wrong = format!(
            r##"<clipPath id="c">{}</clipPath><g id="x">{}</g><rect clip-path="url(#c)"/>"##,
            r##"<rect width="1" height="1" clip-path="url(#x)"/>"##.repeat(2_000),
            "<g/>".repeat(10_000)
        );
        assert_eq!(
            refusal(document(&wrong).as_bytes()),
            Some(SvgRefusal::UnsupportedElement),
            "usvg walks a group named as a clip path before it checks its kind"
        );
    }

    #[test]
    fn the_audit_reads_a_long_reference_list_in_one_pass() {
        let list = "url(#".repeat(500_000);
        for (attribute, outcome) in [
            ("fill", Some(SvgRefusal::ExternalReference)),
            ("clip-path", Some(SvgRefusal::ExternalReference)),
            ("style", Some(SvgRefusal::UnsupportedStyle)),
            ("data-x", None),
        ] {
            let value = if attribute == "style" {
                format!("fill:{list}")
            } else {
                list.clone()
            };
            let source = document(&format!(
                r##"<rect width="9" height="9" {attribute}="{value}"/>"##
            ));
            assert!(source.len() < MAX_SVG_BYTES);
            let started = std::time::Instant::now();
            assert_eq!(refusal(source.as_bytes()), outcome, "{attribute}");
            let elapsed = started.elapsed();
            assert!(
                elapsed < std::time::Duration::from_secs(30),
                "{attribute}: {elapsed:?}"
            );
        }
    }

    #[test]
    fn attributes_are_bounded_before_roxmltree_compares_them_pairwise() {
        let element = |count: usize| {
            let attributes: String = (0..count).map(|index| format!(" a{index}=\"\"")).collect();
            document(&format!(r##"<rect width="9" height="9"{attributes}/>"##))
        };
        assert!(parse(element(MAX_SVG_ELEMENT_ATTRIBUTES - 2).as_bytes()).is_ok());
        let started = std::time::Instant::now();
        assert_eq!(
            refusal(element(150_000).as_bytes()),
            Some(SvgRefusal::DocumentTooLarge)
        );
        assert!(started.elapsed() < std::time::Duration::from_secs(30));
    }

    #[test]
    fn a_use_is_charged_for_walking_its_whole_target() {
        for (filler, outcome) in [
            ("<g/>", Some(SvgRefusal::ExpansionTooLarge)),
            ("<x:a/>", None),
            ("<!---->", None),
            ("x", None),
        ] {
            let source = document(&format!(
                r##"<defs><g id="t" xmlns:x="urn:x">{}<rect width="1" height="1"/></g></defs>{}"##,
                filler.repeat(20_000),
                r##"<use href="#t"/>"##.repeat(10_000)
            ));
            assert_eq!(
                refusal(source.as_bytes()),
                outcome,
                "{filler}: foreign markup, comments and text never reach usvg"
            );
        }
    }

    #[test]
    fn a_shape_is_held_to_the_path_data_usvg_strokes_whole() {
        let path = |segments: usize| {
            document(&format!(
                r##"<path d="M0 0{}" stroke="#000" stroke-width="0.1"/><polygon points="{}"/>"##,
                "h1".repeat(segments),
                "1 1 ".repeat(segments / 2)
            ))
        };
        assert!(parse(path(MAX_SVG_PATH_BYTES / 2 - 4).as_bytes()).is_ok());
        assert_eq!(
            refusal(path(MAX_SVG_PATH_BYTES / 2).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
    }

    #[test]
    fn an_arc_is_refused_before_kurbo_subdivides_it_by_its_radius() {
        let fuzzed = format!(
            r##"<path d="M0 0{}a{}1 1 0 0 0 1 1{}" fill="#f00"/>"##,
            "a1 1 0 0 0 1 1".repeat(1_000),
            "6".repeat(57) + "1",
            "a1 1 0 0 0 1 1".repeat(1_000)
        );
        for body in [
            fuzzed.as_str(),
            r#"<path d="M0 0A1 1 0 0 0 1e12 0"/>"#,
            r#"<circle r="1e30"/>"#,
            r#"<ellipse rx="1" ry="1e20"/>"#,
            r#"<rect width="1e30" height="1e30" rx="1e30"/>"#,
            r#"<circle r="50%"/>"#,
        ] {
            let started = std::time::Instant::now();
            assert_eq!(
                refusal(document(body).as_bytes()),
                Some(SvgRefusal::ExpansionTooLarge),
                "{body}"
            );
            assert!(started.elapsed() < std::time::Duration::from_secs(30));
        }
        let arcs = format!(r#"<path d="M1 1{}"/>"#, "a4 4 0 0 1 8 0".repeat(1_000));
        assert!(parse(document(&arcs).as_bytes()).is_ok());
        let heavy = format!(r#"<path d="M1 1{}"/>"#, "a4 4 0 1 1 8 0".repeat(1_700));
        assert_eq!(
            refusal(document(&heavy).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge),
            "each large arc weighs four cubics"
        );
    }

    #[test]
    fn a_path_is_priced_for_the_edges_and_outline_it_draws() {
        let zigzag = |segments: usize, stroke: &str| {
            document(&format!(
                r##"<path d="M40 40{}" fill="none" stroke="#000" {stroke}/>"##,
                "l9 9l-9-9".repeat(segments / 2)
            ))
        };
        assert!(parse(zigzag(200, r#"stroke-width="4""#).as_bytes()).is_ok());
        for stroke in [
            r#"stroke-width="4""#,
            r#"stroke-width="4" stroke-linejoin="round""#,
            r#"stroke-width="4" stroke-dasharray="0.1 0.1""#,
        ] {
            let started = std::time::Instant::now();
            assert_eq!(
                refusal(zigzag(12_000, stroke).as_bytes()),
                Some(SvgRefusal::RenderTooCostly),
                "{stroke}"
            );
            assert!(started.elapsed() < std::time::Duration::from_secs(30));
        }
        assert!(
            parse(zigzag(12_000, r#"stroke-width="0.1""#).as_bytes()).is_ok(),
            "a hairline is drawn without an outline"
        );
    }

    #[test]
    fn an_id_many_elements_share_is_walked_within_a_bound() {
        let source = document(&format!(
            "{}{}",
            r##"<clipPath id="x"/>"##.repeat(12_000),
            r##"<rect width="1" height="1" clip-path="url(#x)"/>"##.repeat(12_000)
        ));
        let started = std::time::Instant::now();
        assert_eq!(
            refusal(source.as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        assert!(started.elapsed() < std::time::Duration::from_secs(30));
    }

    #[test]
    fn a_stroke_too_wide_for_the_rasteriser_is_refused_before_it_panics() {
        let fuzzed = concat!(
            r##"<svg><g d="e_2"><g id="e1-2"><path d="M .23.11.08  0L5 5" stroke="#000" "##,
            r##"stroke-width="1e30" arke=""/></g></g></svg>"##
        );
        assert_eq!(
            refusal(fuzzed.as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge),
            "refused before usvg strokes it to measure it"
        );
    }

    #[test]
    fn a_clip_path_chain_counts_every_instance() {
        let mut defs = String::from(
            r##"<clipPath id="c0" clipPathUnits="objectBoundingBox"><rect width="1" height="1"/></clipPath>"##,
        );
        for level in 1..=8 {
            let children = format!(
                r##"<rect width="1" height="1" clip-path="url(#c{})"/>"##,
                level - 1
            )
            .repeat(10);
            defs.push_str(&format!(
                r##"<clipPath id="c{level}" clipPathUnits="objectBoundingBox">{children}</clipPath>"##
            ));
        }
        let source = document(&format!(
            r##"<defs>{defs}</defs><rect width="96" height="96" clip-path="url(#c8)"/>"##
        ));
        assert_eq!(
            refusal(source.as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
    }

    #[test]
    fn path_data_reread_per_use_counts_against_the_expansion() {
        let d = "M0 0L1 1".repeat(20_000);
        let uses = r##"<use href="#p"/>"##.repeat(64);
        let source = document(&format!(
            r##"<defs><path id="p" d="{d}" stroke="#000"/></defs>{uses}"##
        ));
        assert!(source.len() < MAX_SVG_BYTES);
        assert_eq!(
            refusal(source.as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
    }

    #[test]
    fn a_reference_to_a_repeated_id_counts_every_element_that_carries_it() {
        let targets = r##"<linearGradient id="x"/>"##.repeat(1_000);
        let references = r##"<rect width="9" height="9" fill="url(#x)"/>"##.repeat(200);
        let source = document(&format!(r##"<defs>{targets}</defs>{references}"##));
        assert_eq!(
            refusal(source.as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
    }

    #[test]
    fn a_stylesheet_is_held_to_plain_rules_and_a_matching_budget() {
        let deep = format!(
            "<style>.x {} {{fill:red}}</style>{}<rect width=\"9\" height=\"9\"/>{}",
            "g ".repeat(32),
            "<g>".repeat(60),
            "</g>".repeat(60)
        );
        assert_eq!(
            refusal(document(&deep).as_bytes()),
            Some(SvgRefusal::UnsupportedStyle)
        );
        let rules: String = (0..1_000)
            .map(|index| format!(".c{index}{{fill:red}}"))
            .collect();
        let rects = r##"<rect width="9" height="9"/>"##.repeat(20_000);
        let wide = document(&format!("<style>{rules}</style>{rects}"));
        assert_eq!(
            refusal(wide.as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        let many: String = (0..MAX_SVG_STYLE_RULES)
            .map(|index| format!(".c{index}{{fill:red}}"))
            .collect();
        assert_eq!(
            refusal(document(&format!("<style>{many}</style>")).as_bytes()),
            Some(SvgRefusal::UnsupportedStyle)
        );
    }

    #[test]
    fn a_selector_bomb_is_refused_before_anything_matches_it() {
        let bomb = |parts: usize, elements: usize| {
            document(&format!(
                "<style>{}{{fill:red}}</style>{}",
                ".a".repeat(parts),
                r##"<g class="a"/>"##.repeat(elements)
            ))
        };
        let started = std::time::Instant::now();
        assert_ne!(refusal(bomb(1_000_000, 90_000).as_bytes()), None);
        assert!(started.elapsed() < std::time::Duration::from_secs(30));
        for (parts, outcome) in [
            (800_000, SvgRefusal::ExpansionTooLarge),
            (5_000, SvgRefusal::UnsupportedStyle),
        ] {
            let source = bomb(parts, 20_000);
            assert!(source.len() < MAX_SVG_BYTES);
            let started = std::time::Instant::now();
            assert_eq!(refusal(source.as_bytes()), Some(outcome), "{parts}");
            assert!(started.elapsed() < std::time::Duration::from_secs(30));
        }
    }

    #[test]
    fn rules_times_elements_times_parts_is_charged_before_usvg_matches_them() {
        let styled = |elements: usize| {
            let rules: String = (0..300)
                .map(|index| format!(".x{index}{}{{fill:red}}", ".a".repeat(15)))
                .collect();
            document(&format!(
                "<style>{rules}</style>{}",
                r##"<g class="a"/>"##.repeat(elements)
            ))
        };
        assert!(parse(styled(100).as_bytes()).is_ok());
        assert_eq!(
            refusal(styled(10_000).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
    }

    #[test]
    fn css_text_is_charged_the_rescans_simplecss_makes_of_it() {
        let style = |declarations: usize| {
            document(&format!(
                r##"<rect width="9" height="9" style="{}"/>"##,
                "fill:red;".repeat(declarations)
            ))
        };
        assert!(parse(style(100).as_bytes()).is_ok());
        assert_eq!(
            refusal(style(8_000).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        let sheet = document(&format!(
            "<style>.a{{{}}}</style>",
            "fill:red;".repeat(8_000)
        ));
        assert_eq!(
            refusal(sheet.as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        let copied = document(&format!(
            "<defs><rect id=\"r\" width=\"9\" height=\"9\" style=\"{}\"/></defs>{}",
            "fill:red;".repeat(400),
            r##"<use href="#r"/>"##.repeat(2_000)
        ));
        assert_eq!(
            refusal(copied.as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge),
            "each copy re-reads its style attribute"
        );
    }

    #[test]
    fn a_filter_in_any_form_is_refused() {
        for body in [
            r##"<rect width="9" height="9" filter="blur(4)"/>"##,
            r##"<rect width="9" height="9" style="fill:red;filter:drop-shadow(1 1 1 red)"/>"##,
            r##"<style>rect{filter:blur(4)}</style><rect width="9" height="9"/>"##,
        ] {
            assert_eq!(
                refusal(document(body).as_bytes()),
                Some(SvgRefusal::UnsupportedStyle),
                "{body}"
            );
        }
        assert!(
            parse(document(r##"<rect width="9" height="9" filter="none"/>"##).as_bytes()).is_ok()
        );
    }

    fn gradient_fills(stops: usize, fills: usize) -> String {
        let stops: String = (0..stops)
            .map(|index| {
                let offset = index as f32 / stops as f32;
                format!(r##"<stop offset="{offset}" stop-color="#f00"/>"##)
            })
            .collect();
        let fills = r##"<rect width="96" height="96" fill="url(#g)"/>"##.repeat(fills);
        document(&format!(
            r##"<linearGradient id="g">{stops}</linearGradient>{fills}"##
        ))
    }

    #[test]
    fn a_gradient_past_the_stop_bound_is_refused() {
        assert_eq!(
            refusal(gradient_fills(MAX_SVG_GRADIENT_STOPS + 1, 1).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        assert!(parse(gradient_fills(MAX_SVG_GRADIENT_STOPS, 1).as_bytes()).is_ok());
    }

    fn inherited(gradient: &str, stops: usize, paint: &str, shapes: usize) -> String {
        let stops: String = (0..stops)
            .map(|index| {
                let offset = index as f32 / stops as f32;
                format!(r##"<stop offset="{offset}" stop-color="#f00"/>"##)
            })
            .collect();
        document(&format!(
            r##"<linearGradient id="g" {gradient}>{stops}</linearGradient><g {paint}="url(#g)">{}</g>"##,
            r##"<rect width="1" height="1"/>"##.repeat(shapes)
        ))
    }

    #[test]
    fn an_inherited_gradient_is_charged_per_shape_that_paints_with_it() {
        let started = std::time::Instant::now();
        assert_ne!(
            refusal(inherited("", MAX_SVG_GRADIENT_STOPS, "fill", 90_000).as_bytes()),
            None
        );
        assert!(started.elapsed() < std::time::Duration::from_secs(30));
        for paint in ["fill", "stroke"] {
            let source = inherited("", MAX_SVG_GRADIENT_STOPS, paint, 3_000);
            assert_eq!(
                refusal(source.as_bytes()),
                Some(SvgRefusal::ExpansionTooLarge),
                "{paint}"
            );
            assert!(parse(inherited("", MAX_SVG_GRADIENT_STOPS, paint, 16).as_bytes()).is_ok());
        }
        let shared = inherited(r##"gradientUnits="userSpaceOnUse""##, 2, "fill", 20_000);
        assert!(
            parse(shared.as_bytes()).is_ok(),
            "a gradient in user units is shared, not copied"
        );
        assert_eq!(
            refusal(inherited("", 2, "fill", 20_000).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge),
            "every copy is compared against every other when usvg collects them"
        );
        assert!(parse(inherited("", 2, "fill", 2_000).as_bytes()).is_ok());
    }

    #[test]
    fn css_cannot_hand_a_shape_the_paint_its_attribute_set_aside() {
        let stops: String = (0..MAX_SVG_GRADIENT_STOPS)
            .map(|index| format!(r##"<stop offset="{index}" stop-color="#f00"/>"##))
            .collect();
        let shapes = r##"<rect width="1" height="1" fill="red"/>"##.repeat(3_000);
        for css in [
            "<style>rect { fill:inherit }</style>",
            "<style>rect { stroke: INHERIT }</style>",
        ] {
            let body = format!(
                r##"{css}<linearGradient id="g">{stops}</linearGradient><g fill="url(#g)" stroke="url(#g)">{shapes}</g>"##
            );
            assert_eq!(
                refusal(document(&body).as_bytes()),
                Some(SvgRefusal::UnsupportedStyle),
                "{css}"
            );
        }
        let attribute =
            r##"<g fill="url(#g)"><rect width="1" height="1" style="fill: inherit"/></g>"##;
        assert_eq!(
            refusal(document(attribute).as_bytes()),
            Some(SvgRefusal::UnsupportedStyle)
        );
        assert!(parse(inherited("", 16, "fill", 16).as_bytes()).is_ok());
    }

    #[test]
    fn a_context_paint_is_charged_per_use_that_hands_it_down() {
        let copies = |uses: usize| {
            let stops: String = (0..MAX_SVG_GRADIENT_STOPS)
                .map(|index| format!(r##"<stop offset="{}" stop-color="#f00"/>"##, index))
                .collect();
            document(&format!(
                concat!(
                    r##"<linearGradient id="g" gradientUnits="userSpaceOnUse">{}</linearGradient>"##,
                    r##"<defs><g id="r"><rect width="1" height="1" fill="context-fill"/></g></defs>{}"##
                ),
                stops,
                r##"<use href="#r" fill="url(#g)"/>"##.repeat(uses)
            ))
        };
        for uses in [4, 3_000] {
            assert_eq!(
                refusal(copies(uses).as_bytes()),
                Some(SvgRefusal::UnsupportedStyle),
                "context paint is refused outright"
            );
        }
    }

    #[test]
    fn full_canvas_fills_past_the_overdraw_bound_are_refused() {
        let fills = r##"<rect width="96" height="96" fill="#f00"/>"##.repeat(10_000);
        assert_eq!(
            refusal(document(&fills).as_bytes()),
            Some(SvgRefusal::RenderTooCostly)
        );
        assert_eq!(
            refusal(gradient_fills(MAX_SVG_GRADIENT_STOPS, 2).as_bytes()),
            Some(SvgRefusal::RenderTooCostly),
            "each stop is a test per painted pixel"
        );
        let few = r##"<rect width="96" height="96" fill="#f00"/>"##.repeat(32);
        assert!(parse(document(&few).as_bytes()).is_ok());
    }

    #[test]
    fn dashes_count_against_the_render_work() {
        let dashed =
            r##"<path d="M0 48H100000" stroke="#000" stroke-dasharray="0.5 0.5"/>"##.repeat(20);
        assert_eq!(
            refusal(document(&dashed).as_bytes()),
            Some(SvgRefusal::RenderTooCostly)
        );
        let few = r##"<path d="M0 48H96" stroke="#000" stroke-dasharray="4 4"/>"##;
        assert!(parse(document(few).as_bytes()).is_ok());
    }

    #[test]
    fn a_dash_list_is_charged_to_every_shape_that_may_stroke_with_it() {
        let group = |pattern: &str, shapes: usize| {
            document(&format!(
                r##"<g stroke="#000" stroke-dasharray="{pattern}">{}</g>"##,
                r##"<rect width="1" height="1"/>"##.repeat(shapes)
            ))
        };
        let long = "1 1 ".repeat(5_000);
        let started = std::time::Instant::now();
        assert_eq!(
            refusal(group(&long, 5_000).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        assert!(started.elapsed() < std::time::Duration::from_secs(30));
        assert!(parse(group("4 2", 5_000).as_bytes()).is_ok());
        let css = format!(
            "<style>g {{ stroke-dasharray: {long} }}</style>{}",
            group("1", 5_000)
        );
        assert_eq!(
            refusal(document(&css).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        assert_eq!(
            refusal(group("1em 1em", 1).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
    }

    #[test]
    fn css_declarations_are_read_as_simplecss_reads_them() {
        let wide = |style: &str| {
            document(&format!(
                r##"<path d="M10 10C60 90 20 -50 80 40" fill="none" stroke="#000" style="{style}"/>"##
            ))
        };
        for style in [
            "stroke-width:1e5 !important",
            "*stroke-width:1e5",
            "stroke-width/* x */:/* y */1e5",
        ] {
            assert_eq!(
                refusal(wide(style).as_bytes()),
                Some(SvgRefusal::ExpansionTooLarge),
                "{style}"
            );
        }
        let dashes = format!(
            r##"<g stroke="#000" style="*stroke-dasharray:{}">{}</g>"##,
            "1 1 ".repeat(2_000),
            r##"<rect width="1" height="1"/>"##.repeat(5_000)
        );
        assert_eq!(
            refusal(document(&dashes).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        assert!(parse(wide("stroke-width:2 !important").as_bytes()).is_ok());
        let long = format!("stroke-width:1;{}", "fill:red;".repeat(8_000));
        let started = std::time::Instant::now();
        assert_eq!(
            refusal(wide(&long).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge),
            "charged before the tokenizer rescans it"
        );
        assert!(started.elapsed() < std::time::Duration::from_secs(30));
    }

    #[test]
    fn a_dense_pattern_head_is_charged_on_every_contour_it_restarts() {
        let pattern = format!("{}1e9", "0.001 0.001 ".repeat(80));
        let body = format!(
            r##"<path d="{}" stroke="#000" stroke-width="2" stroke-dasharray="{pattern}"/>"##,
            "M1 1h1".repeat(8_000)
        );
        let started = std::time::Instant::now();
        assert_eq!(
            refusal(document(&body).as_bytes()),
            Some(SvgRefusal::RenderTooCostly)
        );
        assert!(started.elapsed() < std::time::Duration::from_secs(30));
    }

    #[test]
    fn a_clipped_curve_is_charged_the_edges_the_clipper_cuts_it_into() {
        let curves = |x: i32| {
            let (near, far) = (x - 2, x + 2);
            document(&format!(
                r#"<path d="M{near} 40{}"/>"#,
                format!("C{far} 40 {far} 44 {near} 44C{far} 44 {far} 40 {near} 40").repeat(500)
            ))
        };
        assert_eq!(
            refusal(curves(0).as_bytes()),
            Some(SvgRefusal::RenderTooCostly),
            "every curve straddles the left side"
        );
        assert!(parse(curves(48).as_bytes()).is_ok());
    }

    #[test]
    fn a_stroke_is_bounded_before_the_stroker_subdivides_it() {
        let cusp = "M687.08 -722.33C660.37 -517.96 673.22 -539.96 -10.98 434.07";
        let cusps = format!(
            r##"<g transform="scale(0.05) translate(800 800)"><path d="{}" fill="none" stroke="#000"/></g>"##,
            cusp.repeat(2_000)
        );
        for body in [
            r##"<path d="M10 10C60 90 20 -50 80 40" stroke="#000" stroke-width="1e7"/>"##,
            r##"<path d="M1000000 1000000c50 80 10 -60 70 30" stroke="#000"/>"##,
            r##"<g transform="scale(0.003)"><path d="M0 0C30000 60000 -20000 -50000 32000 30000" stroke="#000" stroke-width="10"/></g>"##,
            r##"<g transform="rotate(30 48 48)"><path d="M10 10C60 90 20 -50 80 40" stroke="#000"/></g>"##,
            r##"<style>path { transform: rotate(30deg) }</style><path d="M10 10L80 40" stroke="#000"/>"##,
            r##"<path d="M10 10L80 40" stroke="#000" stroke-width="10%"/>"##,
            cusps.as_str(),
        ] {
            let started = std::time::Instant::now();
            assert_eq!(
                refusal(document(body).as_bytes()),
                Some(SvgRefusal::ExpansionTooLarge),
                "{body}"
            );
            assert!(started.elapsed() < std::time::Duration::from_secs(30));
        }
        let icon = format!(
            r##"<circle cx="48" cy="48" r="40" fill="none" stroke="#000" stroke-width="2"/><rect x="10" y="10" width="76" height="76" rx="8" fill="none" stroke="#000" stroke-dasharray="4 2"/><g transform="scale(0.05) translate(800 800)"><path d="{cusp}" fill="none" stroke="#000" stroke-width="20"/></g>"##
        );
        assert!(parse(document(&icon).as_bytes()).is_ok());
        let markers =
            r##"<circle cx="48" cy="48" r="3" fill="none" stroke="#000"/>"##.repeat(1_000);
        assert!(parse(document(&markers).as_bytes()).is_ok());
    }

    #[test]
    fn a_stroke_is_priced_where_resvg_strokes_it_on_the_surface_or_off() {
        let curves = "M0 0C20 60 40 -60 50 0".repeat(200);
        for place in ["translate(-50000 0) scale(100)", "scale(100)"] {
            let body = format!(
                r##"<g transform="{place}"><path d="{curves}" fill="none" stroke="#000"/></g>"##
            );
            let started = std::time::Instant::now();
            assert_eq!(
                refusal(document(&body).as_bytes()),
                Some(SvgRefusal::RenderTooCostly),
                "{place}"
            );
            assert!(started.elapsed() < std::time::Duration::from_secs(30));
        }
        let dashed = concat!(
            r##"<g transform="translate(-3e7 0) scale(1000)"><path d="M30000 0C30010 20 30030 -10 30030 30" "##,
            r##"stroke="#000" stroke-width="0.0005" stroke-dasharray="0.001 0.001"/></g>"##
        );
        assert_eq!(
            refusal(document(dashed).as_bytes()),
            Some(SvgRefusal::RenderTooCostly),
            "a dashed hairline far out is measured past the float precision of its tolerance"
        );
    }

    #[test]
    fn inherited_lookups_through_deep_attribute_chains_are_charged() {
        let names = [
            "fill-opacity",
            "stroke-opacity",
            "clip-rule",
            "fill-rule",
            "color-rendering",
            "direction",
            "dominant-baseline",
            "flood-color",
            "flood-opacity",
            "font-family",
            "font-size",
            "font-stretch",
            "font-style",
            "font-variant",
            "font-weight",
            "image-rendering",
            "letter-spacing",
            "lighting-color",
            "shape-rendering",
            "stop-color",
            "stop-opacity",
            "stroke-dashoffset",
            "stroke-linecap",
            "stroke-linejoin",
            "stroke-miterlimit",
            "text-anchor",
            "text-rendering",
            "word-spacing",
            "writing-mode",
            "baseline-shift",
            "x",
            "y",
            "width",
            "height",
            "cx",
            "cy",
            "r",
            "rx",
            "ry",
            "x1",
            "y1",
            "x2",
            "y2",
            "dx",
            "dy",
            "rotate",
            "offset",
            "fx",
            "fy",
            "k1",
            "k2",
            "k3",
            "k4",
            "z",
            "order",
            "radius",
            "scale",
            "seed",
            "slope",
            "specularExponent",
            "stdDeviation",
        ];
        let chain = |depth: usize, attributes: usize, shapes: usize| {
            let attributes: String = names[..attributes]
                .iter()
                .map(|name| format!(r#" {name}="1""#))
                .collect();
            document(&format!(
                "{}{}{}",
                format!("<g{attributes}>").repeat(depth),
                r#"<rect width="1" height="1"/>"#.repeat(shapes),
                "</g>".repeat(depth)
            ))
        };
        let started = std::time::Instant::now();
        assert_eq!(
            refusal(chain(60, 60, 30_000).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        assert!(started.elapsed() < std::time::Duration::from_secs(30));
        let attributes: String = names.iter().map(|name| format!(r#" {name}="1""#)).collect();
        let fanned = |uses: usize| {
            document(&format!(
                r##"<defs><g id="c"{attributes}>{}{}{}</g></defs>{}"##,
                format!("<g{attributes}>").repeat(58),
                r#"<rect width="1" height="1"/>"#.repeat(100),
                "</g>".repeat(58),
                r##"<use href="#c"/>"##.repeat(uses)
            ))
        };
        let started = std::time::Instant::now();
        assert_eq!(
            refusal(fanned(190).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        assert!(started.elapsed() < std::time::Duration::from_secs(30));
        assert!(parse(fanned(2).as_bytes()).is_ok());
        assert!(parse(chain(8, 10, 5_000).as_bytes()).is_ok());
    }

    #[test]
    fn a_foreign_style_attribute_never_reaches_usvg() {
        let body = concat!(
            r##"<g xmlns:f="urn:f" f:style="stroke-width:100000;stroke:#000;fill:inherit">"##,
            r##"<path d="M10 10C60 90 20 -50 80 40"/></g>"##
        );
        assert!(parse(document(body).as_bytes()).is_ok());
    }

    #[test]
    fn a_dash_list_is_charged_to_every_use_that_resolves_a_stroke() {
        let uses = |count: usize| {
            document(&format!(
                r##"<defs><g id="e"/></defs><g stroke="#000" stroke-dasharray="{}">{}</g>"##,
                "1 ".repeat(500),
                r##"<use href="#e"/>"##.repeat(count)
            ))
        };
        assert!(parse(uses(10).as_bytes()).is_ok());
        assert_eq!(
            refusal(uses(4_000).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
    }

    #[test]
    fn inherited_values_are_capped_and_charged_per_inheriting_shape() {
        let padded = format!(r##"<g fill="{}red"/>"##, " ".repeat(1 << 20));
        let started = std::time::Instant::now();
        assert_eq!(
            refusal(document(&padded).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge),
            "a value past MAX_SVG_VALUE_BYTES"
        );
        assert!(started.elapsed() < std::time::Duration::from_secs(30));
        let long = |shapes: usize| {
            let target = "i".repeat(250);
            document(&format!(
                r##"<g fill="url(#{target}) #f00" stroke="url(#{target}) #00f" clip-path="url(#{target})" mask="url(#{target})">{}</g>"##,
                r##"<rect width="1" height="1"/>"##.repeat(shapes)
            ))
        };
        assert!(parse(long(3_000).as_bytes()).is_ok());
        assert_eq!(
            refusal(long(20_000).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge)
        );
    }

    #[test]
    fn a_gradient_chain_is_capped_and_charged_per_shape() {
        let chain = |links: usize, shapes: usize| {
            let gradients: String = (0..links)
                .map(|index| {
                    format!(
                        r##"<linearGradient id="g{index}" href="#g{}"/>"##,
                        index + 1
                    )
                })
                .collect();
            document(&format!(
                "{gradients}{}",
                r##"<rect width="1" height="1" fill="url(#g0)"/>"##.repeat(shapes)
            ))
        };
        assert!(parse(chain(4, 10_000).as_bytes()).is_ok());
        assert_eq!(
            refusal(chain(5, 1).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge),
            "past the chain bound"
        );
    }

    #[test]
    fn the_line_closing_each_open_contour_is_priced() {
        let price = |close: &str| {
            let source = document(&format!(
                r##"<path d="{}" fill="#000"/>"##,
                format!("M0 0L96 48L0 96{close}").repeat(200)
            ));
            let image = parse(source.as_bytes()).expect("parse");
            cost::measure(&image.tree, image.size, None)
                .expect("price")
                .work
        };
        assert_eq!(
            price(""),
            price("Z"),
            "tiny-skia closes an open contour with the line Z draws"
        );
    }

    #[test]
    fn a_gradient_is_charged_its_stops_for_every_shape_and_use_resolving_it() {
        let stops: String = (0..MAX_SVG_GRADIENT_STOPS)
            .map(|index| format!(r##"<stop offset="{index}" stop-color="#f00"/>"##))
            .collect();
        let radial = |users: &str, count: usize| {
            document(&format!(
                r##"<radialGradient id="g" gradientUnits="userSpaceOnUse" r="0">{stops}</radialGradient><defs><g id="e"/></defs>{}"##,
                users.repeat(count)
            ))
        };
        let shape = r##"<rect width="1" height="1" fill="url(#g)"/>"##;
        let user = r##"<use href="#e" fill="url(#g)"/>"##;
        assert!(parse(radial(shape, 100).as_bytes()).is_ok());
        for users in [shape, user] {
            assert_eq!(
                refusal(radial(users, 5_000).as_bytes()),
                Some(SvgRefusal::ExpansionTooLarge),
                "{users}"
            );
        }
    }

    #[test]
    fn a_sanitised_document_is_held_to_the_byte_bounds_again() {
        let class = "&#61;".repeat(200);
        let body = format!(r#"<g class="{class}"/>"#).repeat(700);
        let source = document(&body);
        assert!(source.bytes().filter(|byte| *byte == b'=').count() < MAX_SVG_ATTRIBUTES);
        assert_eq!(
            refusal(source.as_bytes()),
            Some(SvgRefusal::DocumentTooLarge),
            "each reference becomes an `=` roxmltree reserves an attribute for"
        );
    }

    #[test]
    fn a_reference_mentioning_context_does_not_hide_inherited_paint() {
        let stops: String = (0..MAX_SVG_GRADIENT_STOPS)
            .map(|index| format!(r##"<stop offset="{index}" stop-color="#f00"/>"##))
            .collect();
        let body = format!(
            r##"<linearGradient id="g">{stops}</linearGradient><g stroke="url(#g)">{}</g>"##,
            r##"<rect width="1" height="1" style="fill:url(#context-missing)"/>"##.repeat(3_000)
        );
        assert_eq!(
            refusal(document(&body).as_bytes()),
            Some(SvgRefusal::ExpansionTooLarge),
            "every shape still copies the stroke gradient it inherits"
        );
    }

    #[test]
    fn an_id_carried_by_many_elements_is_checked_at_the_cost_of_one() {
        let body = format!(
            "{}{}",
            r##"<linearGradient id="g"/>"##.repeat(40_000),
            r##"<rect style="fill:url(#g);stroke:url(#g);fill:url(#g);stroke:url(#g)"/>"##
                .repeat(40_000)
        );
        let source = document(&body);
        assert!(source.len() <= MAX_SVG_BYTES);
        let started = std::time::Instant::now();
        assert_eq!(
            refusal(source.as_bytes()),
            Some(SvgRefusal::DocumentTooLarge),
            "the sanitised text outgrows its bound after every reference is checked"
        );
        assert!(started.elapsed() < std::time::Duration::from_secs(30));
    }

    pub(crate) fn opacity_nest(depth: usize) -> String {
        document(&format!(
            r##"{}<rect width="96" height="96" fill="#0000ff"/>{}"##,
            r#"<g opacity="0.99">"#.repeat(depth),
            "</g>".repeat(depth)
        ))
    }

    #[test]
    fn nested_layers_past_the_bound_are_refused_and_within_it_charge_their_rasters() {
        assert_eq!(
            refusal(opacity_nest(MAX_SVG_LAYER_DEPTH + 1).as_bytes()),
            Some(SvgRefusal::TooManyLayers)
        );
        let image = parse(opacity_nest(MAX_SVG_LAYER_DEPTH).as_bytes()).expect("parse");
        let layers = MAX_SVG_LAYER_DEPTH as u64;
        assert!(
            (384 * 384 + layers * 388 * 388..=384 * 384 + layers * 390 * 390)
                .contains(&image.pixels()),
            "each layer is the canvas grown two pixels a side: {}",
            image.pixels()
        );
        let (data, _) = image.render().expect("render");
        assert_eq!(data[2], 255, "the nested rect still draws");
    }

    #[test]
    fn a_clip_too_large_to_stack_steps_the_supersample_down() {
        let source = concat!(
            r##"<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="960" viewBox="0 0 1920 960">"##,
            r##"<defs><clipPath id="c"><path d="M0 0H1920V960H0Z"/></clipPath></defs>"##,
            r##"<g clip-path="url(#c)"><rect width="1920" height="960" fill="#00adef"/></g></svg>"##
        );
        let image = parse(source.as_bytes()).expect("parse");
        assert_eq!((image.size.width(), image.size.height()), (3840, 1920));
        assert!(image.pixels() <= MAX_IMAGE_PIXELS);
    }

    #[test]
    fn a_panicking_step_is_a_refusal() {
        assert_eq!(guarded(|| 7), Ok(7));
        assert_eq!(
            guarded(|| -> u8 { panic!("usvg fell over") }),
            Err(SvgRefusal::Panicked)
        );
    }
}
