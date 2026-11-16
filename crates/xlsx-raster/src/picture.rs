use std::collections::HashMap;
use std::io::Cursor;

use tiny_skia::{
    FilterQuality, IntRect, IntSize, Paint, Pattern, Pixmap, Rect, SpreadMode, Transform,
};
use xlsx_render::DrawCmd;

use crate::{ClipSurface, reach};

/// Encoded image bytes keyed by the `asset_id` an image command carries.
pub type ImageMap<'a> = HashMap<&'a str, &'a [u8]>;

/// One decoded image.
pub const MAX_IMAGE_PIXELS: u64 = 33_554_432;
/// Every image one frame decodes.
pub const MAX_FRAME_IMAGE_PIXELS: u64 = 67_108_864;
/// One decode's source buffer plus its pixmap, at [`MAX_IMAGE_PIXELS`] of RGBA8.
const MAX_IMAGE_BYTES: u64 = 268_435_456;
/// The same, summed across one frame.
const MAX_FRAME_IMAGE_BYTES: u64 = 536_870_912;

/// The images one frame paints, each decoded once within the frame's budget.
pub(crate) struct Pictures<'a> {
    images: &'a ImageMap<'a>,
    decoded: HashMap<&'a str, Option<Pixmap>>,
    pixels: u64,
    bytes: u64,
}

impl<'a> Pictures<'a> {
    pub(crate) fn new(images: &'a ImageMap<'a>) -> Self {
        Self {
            images,
            decoded: HashMap::new(),
            pixels: 0,
            bytes: 0,
        }
    }

    /// Paints an image command; one without bytes, or that will not decode, paints nothing.
    pub(crate) fn paint(
        &mut self,
        target: &mut Pixmap,
        clip_surface: &mut ClipSurface,
        command: &DrawCmd,
    ) -> Result<(), String> {
        let DrawCmd::Image {
            x,
            y,
            w,
            h,
            asset_id,
            crop,
            clip,
        } = command
        else {
            return Ok(());
        };
        let Some((&key, &bytes)) = self.images.get_key_value(&**asset_id) else {
            return Ok(());
        };
        if !self.decoded.contains_key(key) {
            let decoded = self.decode(bytes);
            self.decoded.insert(key, decoded);
        }
        let Some(Some(image)) = self.decoded.get(key) else {
            return Ok(());
        };
        let (width, height) = (image.width() as f32, image.height() as f32);
        let kept = IntRect::from_ltrb(
            (crop.left * width).round() as i32,
            (crop.top * height).round() as i32,
            ((1.0 - crop.right) * width).round() as i32,
            ((1.0 - crop.bottom) * height).round() as i32,
        );
        let (Some(rect), Some(kept)) = (Rect::from_xywh(*x, *y, *w, *h), kept) else {
            return Ok(());
        };
        let cropped;
        let source = if kept.width() == image.width() && kept.height() == image.height() {
            image.as_ref()
        } else {
            let Some(pixels) = image.clone_rect(kept) else {
                return Ok(());
            };
            cropped = pixels;
            cropped.as_ref()
        };
        let paint = Paint {
            shader: Pattern::new(
                source,
                SpreadMode::Pad,
                FilterQuality::Bilinear,
                1.0,
                Transform::from_row(
                    w / source.width() as f32,
                    0.0,
                    0.0,
                    h / source.height() as f32,
                    *x,
                    *y,
                ),
            ),
            anti_alias: true,
            ..Paint::default()
        };
        match clip {
            Some(clip) => {
                let bounds = reach(rect.left(), rect.top(), rect.right(), rect.bottom(), 1.0);
                clip_surface.paint(target, *clip, bounds, |surface, transform, mask| {
                    surface.fill_rect(rect, &paint, transform, mask);
                    Ok(())
                })
            }
            None => {
                target.fill_rect(rect, &paint, Transform::identity(), None);
                Ok(())
            }
        }
    }

    /// Premultiplied pixels, or `None` past the image or frame budget.
    fn decode(&mut self, bytes: &[u8]) -> Option<Pixmap> {
        use image::ImageDecoder as _;

        let bytes = ooxml_raster::png_with_end(bytes);
        let mut decoder = image::ImageReader::new(Cursor::new(bytes.as_ref()))
            .with_guessed_format()
            .ok()?
            .into_decoder()
            .ok()?;
        let (width, height) = decoder.dimensions();
        let pixels = u64::from(width) * u64::from(height);
        let cost = decoder
            .total_bytes()
            .saturating_add(pixels.saturating_mul(4));
        if pixels > MAX_IMAGE_PIXELS
            || cost > MAX_IMAGE_BYTES
            || self.pixels + pixels > MAX_FRAME_IMAGE_PIXELS
            || self.bytes + cost > MAX_FRAME_IMAGE_BYTES
        {
            return None;
        }
        self.pixels += pixels;
        self.bytes += cost;
        let mut limits = image::Limits::default();
        limits.max_image_width = Some(width);
        limits.max_image_height = Some(height);
        limits.max_alloc = Some((MAX_IMAGE_BYTES - cost).min(MAX_FRAME_IMAGE_BYTES - self.bytes));
        decoder.set_limits(limits).ok()?;
        let orientation = decoder.orientation().ok()?;
        let mut decoded = image::DynamicImage::from_decoder(decoder).ok()?;
        decoded.apply_orientation(orientation);
        let size = IntSize::from_wh(decoded.width(), decoded.height())?;
        let mut data = decoded.into_rgba8().into_raw();
        ooxml_raster::premultiply_rgba(&mut data);
        Pixmap::from_vec(data, size)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_render::{DisplayList, ImageCrop};

    fn red_then_blue_png() -> Vec<u8> {
        let image = image::RgbaImage::from_fn(2, 1, |x, _| {
            image::Rgba(if x == 0 {
                [255, 0, 0, 255]
            } else {
                [0, 0, 255, 255]
            })
        });
        let mut bytes = Vec::new();
        image
            .write_to(&mut Cursor::new(&mut bytes), image::ImageFormat::Png)
            .unwrap();
        bytes
    }

    /// A red GIF whose one frame starts at the second column of its screen.
    fn red_gif_inset_by_one_column(width: u16, height: u16) -> Vec<u8> {
        let mut bytes = b"GIF89a".to_vec();
        bytes.extend(width.to_le_bytes());
        bytes.extend(height.to_le_bytes());
        bytes.extend([0x80, 0, 0, 255, 0, 0, 255, 0, 0, 0x2C]);
        bytes.extend(1u16.to_le_bytes());
        bytes.extend(0u16.to_le_bytes());
        bytes.extend((width - 1).to_le_bytes());
        bytes.extend(height.to_le_bytes());
        bytes.extend([0, 2]);
        let pairs = (u32::from(width) - 1) * u32::from(height) / 2;
        let (mut data, mut acc, mut bits) = (Vec::new(), 0u32, 0u32);
        for code in (0..pairs).flat_map(|_| [4u32, 0, 0]).chain([5]) {
            acc |= code << bits;
            bits += 3;
            while bits >= 8 {
                data.push(acc as u8);
                acc >>= 8;
                bits -= 8;
            }
        }
        if bits > 0 {
            data.push(acc as u8);
        }
        for block in data.chunks(255) {
            bytes.push(block.len() as u8);
            bytes.extend(block);
        }
        bytes.extend([0, 0x3B]);
        bytes
    }

    fn paint(crop: ImageCrop, clip: Option<xlsx_render::Rect>, images: &ImageMap) -> Pixmap {
        let list = DisplayList {
            truncated: false,
            width: 40.0,
            height: 20.0,
            commands: vec![DrawCmd::Image {
                x: 0.0,
                y: 0.0,
                w: 40.0,
                h: 20.0,
                asset_id: "xl/media/image1.png".into(),
                crop,
                clip,
            }],
            grid: xlsx_render::GridMeta::default(),
            hyperlinks: Vec::new(),
            charts: Vec::new(),
        };
        Pixmap::decode_png(&crate::render_png_with_images(&list, images).unwrap()).unwrap()
    }

    fn rgb(pixmap: &Pixmap, x: u32) -> (u8, u8, u8) {
        let pixel = pixmap.pixel(x, 10).unwrap();
        (pixel.red(), pixel.green(), pixel.blue())
    }

    #[test]
    fn an_image_stretches_its_kept_crop_over_the_rect_inside_its_clip() {
        let png = red_then_blue_png();
        let images = ImageMap::from([("xl/media/image1.png", png.as_slice())]);
        let whole = paint(ImageCrop::default(), None, &images);
        assert_eq!(
            (rgb(&whole, 2), rgb(&whole, 37)),
            ((255, 0, 0), (0, 0, 255))
        );
        let cropped = paint(
            ImageCrop {
                left: 0.5,
                ..ImageCrop::default()
            },
            None,
            &images,
        );
        assert_eq!(
            (rgb(&cropped, 2), rgb(&cropped, 37)),
            ((0, 0, 255), (0, 0, 255))
        );
        let clipped = paint(
            ImageCrop::default(),
            Some(xlsx_render::Rect {
                x: 20.0,
                y: 0.0,
                w: 20.0,
                h: 20.0,
            }),
            &images,
        );
        assert_eq!(
            (rgb(&clipped, 2), rgb(&clipped, 37)),
            ((255, 255, 255), (0, 0, 255))
        );
    }

    #[test]
    fn an_image_within_the_pixel_cap_decodes_inside_its_own_budget() {
        let gif = red_gif_inset_by_one_column(4800, 4800);
        let images = ImageMap::from([("xl/media/image1.png", gif.as_slice())]);
        let pixmap = paint(ImageCrop::default(), None, &images);
        assert_eq!(rgb(&pixmap, 20), (255, 0, 0));
    }

    #[test]
    fn an_image_without_decodable_bytes_paints_nothing() {
        for images in [
            ImageMap::new(),
            ImageMap::from([("xl/media/image1.png", b"not an image".as_slice())]),
        ] {
            let pixmap = paint(ImageCrop::default(), None, &images);
            assert_eq!(rgb(&pixmap, 2), (255, 255, 255));
        }
    }
}
