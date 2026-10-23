use std::io::Cursor;

use image::codecs::gif::GifEncoder;
use image::codecs::jpeg::JpegEncoder;
use image::codecs::png::{CompressionType, FilterType, PngEncoder};
use image::{
    ColorType, DynamicImage, ImageBuffer, ImageDecoder as _, ImageEncoder as _, ImageFormat,
    ImageReader, Rgb, Rgba,
};

use crate::{
    Format, RedactError, RedactionOptions, RedactionReport, media, redact_with_report,
    redact_with_report_and_options,
};

fn inspect(bytes: &[u8]) -> (ImageFormat, (u32, u32), ColorType) {
    let format = image::guess_format(bytes).expect("valid image format");
    let decoder = ImageReader::with_format(Cursor::new(bytes), format)
        .into_decoder()
        .expect("valid image header");
    (format, decoder.dimensions(), decoder.color_type())
}

fn assert_size_within_25_percent(actual: usize, original: usize, label: &str) {
    let diff = actual.abs_diff(original) as f64;
    let max_diff = (original as f64) * 0.25;
    assert!(
        diff <= max_diff,
        "{label}: actual size {actual} is not within +/-25% of original {original} (diff={diff:.0}, max={max_diff:.0})"
    );
}

fn synthetic_noisy_png(width: u32, height: u32) -> Vec<u8> {
    let stride = (width as usize) * 3;
    let mut raw = vec![0u8; stride * (height as usize)];
    let mut state = 0x1234_5678_9ABC_DEF0u64;
    for y in 0..height as usize {
        for x in 0..width as usize {
            let idx = y * stride + x * 3;
            if x < 100 && y < 100 {
                raw[idx] = 251;
                raw[idx + 1] = 13;
                raw[idx + 2] = 249;
            } else if ((x / 16 + y / 8) % 5) == 0 {
                state ^= state >> 12;
                state ^= state << 25;
                state ^= state >> 27;
                let w = state.wrapping_mul(0x2545_F491_4F6C_DD1D);
                raw[idx] = (w & 0xFF) as u8;
                raw[idx + 1] = ((w >> 8) & 0xFF) as u8;
                raw[idx + 2] = ((w >> 16) & 0xFF) as u8;
            } else {
                let g = ((x / 16) * 7 + (y / 8) * 11) as u8;
                raw[idx] = g;
                raw[idx + 1] = g.wrapping_add(40);
                raw[idx + 2] = g.wrapping_add(80);
            }
        }
    }
    let mut out = Vec::new();
    PngEncoder::new_with_quality(&mut out, CompressionType::Fast, FilterType::Sub)
        .write_image(&raw, width, height, ColorType::Rgb8.into())
        .unwrap();
    let iend_pos = out.len() - 12;
    let mut with_meta = Vec::with_capacity(out.len() + 64);
    with_meta.extend_from_slice(&out[..iend_pos]);
    let text_data = b"Comment\0SECRET_PNG_TEXT_METADATA";
    with_meta.extend_from_slice(&(text_data.len() as u32).to_be_bytes());
    with_meta.extend_from_slice(b"tEXt");
    with_meta.extend_from_slice(text_data);
    let mut crc = 0xFFFF_FFFFu32;
    for b in b"tEXt".iter().chain(text_data.iter()) {
        crc ^= u32::from(*b);
        for _ in 0..8 {
            crc = if crc & 1 == 0 {
                crc >> 1
            } else {
                (crc >> 1) ^ 0xEDB8_8320
            };
        }
    }
    with_meta.extend_from_slice(&(!crc).to_be_bytes());
    with_meta.extend_from_slice(&out[iend_pos..]);
    with_meta
}

fn synthetic_noisy_jpeg(width: u32, height: u32) -> Vec<u8> {
    let stride = (width as usize) * 3;
    let mut raw = vec![0u8; stride * (height as usize)];
    let mut state = 0xCAFE_BABE_DEAD_BEEFu64;
    for y in 0..height as usize {
        for x in 0..width as usize {
            let idx = y * stride + x * 3;
            if x < 128 && y < 128 {
                raw[idx] = 255;
                raw[idx + 1] = 0;
                raw[idx + 2] = 255;
            } else if ((x / 16 + y / 8) % 5) == 0 {
                state ^= state >> 12;
                state ^= state << 25;
                state ^= state >> 27;
                let w = state.wrapping_mul(0x2545_F491_4F6C_DD1D);
                raw[idx] = (w & 0xFF) as u8;
                raw[idx + 1] = ((w >> 8) & 0xFF) as u8;
                raw[idx + 2] = ((w >> 16) & 0xFF) as u8;
            } else {
                let gx = (x / 16) as u8;
                let gy = (y / 8) as u8;
                raw[idx] = gx.wrapping_mul(5).wrapping_add(gy.wrapping_mul(3));
                raw[idx + 1] = gx.wrapping_mul(3).wrapping_add(gy.wrapping_mul(7));
                raw[idx + 2] = gx.wrapping_mul(7).wrapping_add(gy.wrapping_mul(5));
            }
        }
    }
    let mut jpeg = Vec::new();
    JpegEncoder::new_with_quality(&mut jpeg, 80)
        .write_image(&raw, width, height, ColorType::Rgb8.into())
        .unwrap();
    let comment = b"SECRET_JPEG_EXIF_METADATA";
    let seg_len = (comment.len() + 2) as u16;
    let mut out = Vec::with_capacity(jpeg.len() + comment.len() + 4);
    out.extend_from_slice(&jpeg[0..2]);
    out.extend_from_slice(&[0xFF, 0xFE]);
    out.extend_from_slice(&seg_len.to_be_bytes());
    out.extend_from_slice(comment);
    out.extend_from_slice(&jpeg[2..]);
    out
}

fn synthetic_noisy_gif(width: u32, height: u32) -> Vec<u8> {
    let stride = (width as usize) * 4;
    let mut raw = vec![0u8; stride * (height as usize)];
    let mut state = 0x7777_8888_9999_AAAAu64;
    for y in 0..height as usize {
        for x in 0..width as usize {
            let idx = y * stride + x * 4;
            if x < 64 && y < 64 {
                raw[idx] = 240;
                raw[idx + 1] = 20;
                raw[idx + 2] = 240;
                raw[idx + 3] = 255;
            } else if ((x / 16 + y / 8) % 2) == 0 {
                state ^= state >> 12;
                state ^= state << 25;
                state ^= state >> 27;
                let w = state.wrapping_mul(0x2545_F491_4F6C_DD1D);
                raw[idx] = (w & 0xFF) as u8;
                raw[idx + 1] = ((w >> 8) & 0xFF) as u8;
                raw[idx + 2] = ((w >> 16) & 0xFF) as u8;
                raw[idx + 3] = 255;
            } else {
                let g = ((x / 16) * 9 + (y / 8) * 5) as u8;
                raw[idx] = g;
                raw[idx + 1] = g.wrapping_add(60);
                raw[idx + 2] = g.wrapping_add(120);
                raw[idx + 3] = 255;
            }
        }
    }
    let mut out = Vec::new();
    GifEncoder::new_with_speed(&mut out, 30)
        .encode(&raw, width, height, ColorType::Rgba8.into())
        .unwrap();
    out
}

fn synthetic_bmp(width: u32, height: u32) -> Vec<u8> {
    let image = DynamicImage::ImageRgb8(ImageBuffer::from_fn(width, height, |x, y| {
        if x < 64 && y < 64 {
            Rgb([252, 11, 248])
        } else {
            Rgb([(x & 0xFF) as u8, (y & 0xFF) as u8, ((x ^ y) & 0xFF) as u8])
        }
    }));
    let mut out = Cursor::new(Vec::new());
    image.write_to(&mut out, ImageFormat::Bmp).unwrap();
    out.into_inner()
}

fn wrap_docx_with_media(media_parts: &[(&str, Vec<u8>)]) -> Vec<u8> {
    let mut parts: Vec<(String, Vec<u8>)> = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="bmp" ContentType="image/bmp"/><Default Extension="tiff" ContentType="image/tiff"/><Default Extension="svg" ContentType="image/svg+xml"/><Default Extension="emf" ContentType="image/x-emf"/><Default Extension="wmf" ContentType="image/x-wmf"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>"#.to_vec(),
        ),
    ];
    for (name, bytes) in media_parts {
        parts.push(((*name).to_owned(), bytes.clone()));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn preserve_media_size_keeps_dimensions_format_color_type_and_approximate_bytes() {
    let png_orig = synthetic_noisy_png(1200, 800);
    let jpeg_orig = synthetic_noisy_jpeg(2000, 1500);
    let gif_orig = synthetic_noisy_gif(800, 600);
    let bmp_orig = synthetic_bmp(400, 300);

    assert!(
        (100_000..5_000_000).contains(&png_orig.len()),
        "unexpected synthetic PNG size: {}",
        png_orig.len()
    );
    assert!(
        (100_000..5_000_000).contains(&jpeg_orig.len()),
        "unexpected synthetic JPEG size: {}",
        jpeg_orig.len()
    );
    assert!(
        (100_000..5_000_000).contains(&gif_orig.len()),
        "unexpected synthetic GIF size: {}",
        gif_orig.len()
    );
    assert!(
        (300_000..400_000).contains(&bmp_orig.len()),
        "unexpected synthetic BMP size: {}",
        bmp_orig.len()
    );

    let package = wrap_docx_with_media(&[
        ("word/media/photo.png", png_orig.clone()),
        ("word/media/camera.jpg", jpeg_orig.clone()),
        ("word/media/anim.gif", gif_orig.clone()),
        ("word/media/bitmap.bmp", bmp_orig.clone()),
    ]);

    let (default_out, default_report) = redact_with_report(&package, Format::Docx).unwrap();
    assert_eq!(default_report.media_parts, 4);
    let default_parts = ooxml_opc::unzip_parts(&default_out).unwrap();
    for name in [
        "word/media/photo.png",
        "word/media/camera.jpg",
        "word/media/anim.gif",
        "word/media/bitmap.bmp",
    ] {
        let bytes = &default_parts.iter().find(|(p, _)| p == name).unwrap().1;
        let (_, dims, _) = inspect(bytes);
        assert_eq!(dims, (media::PLACEHOLDER_SIZE, media::PLACEHOLDER_SIZE));
    }

    let options = RedactionOptions {
        random_characters: false,
        preserve_media_size: true,
    };
    let (preserved_out, preserved_report) =
        redact_with_report_and_options(&package, Format::Docx, &options).unwrap();
    assert_eq!(preserved_report.media_parts, 4);
    let (preserved_out_2, _) =
        redact_with_report_and_options(&package, Format::Docx, &options).unwrap();

    let parts = ooxml_opc::unzip_parts(&preserved_out).unwrap();
    let parts_2 = ooxml_opc::unzip_parts(&preserved_out_2).unwrap();
    let get_part = |list: &[(String, Vec<u8>)], name: &str| -> Vec<u8> {
        list.iter().find(|(p, _)| p == name).unwrap().1.clone()
    };

    let png_out = get_part(&parts, "word/media/photo.png");
    let png_out_2 = get_part(&parts_2, "word/media/photo.png");
    assert_eq!(
        inspect(&png_out),
        (ImageFormat::Png, (1200, 800), ColorType::Rgb8)
    );
    assert_size_within_25_percent(png_out.len(), png_orig.len(), "1200x800 PNG");
    assert_ne!(png_out, png_orig);
    assert_ne!(png_out, png_out_2);
    assert!(
        !png_out
            .windows(b"SECRET_PNG_TEXT_METADATA".len())
            .any(|w| w == b"SECRET_PNG_TEXT_METADATA")
    );
    let png_decoded = image::load_from_memory(&png_out).unwrap().to_rgb8();
    let sentinel_matches = (0..100)
        .flat_map(|y| (0..100).map(move |x| (x, y)))
        .filter(|&(x, y)| png_decoded.get_pixel(x, y).0 == [251, 13, 249])
        .count();
    assert!(
        sentinel_matches < 5,
        "source sentinel pattern survived in PNG ({sentinel_matches} matches)"
    );

    let jpeg_out = get_part(&parts, "word/media/camera.jpg");
    let jpeg_out_2 = get_part(&parts_2, "word/media/camera.jpg");
    assert_eq!(
        inspect(&jpeg_out),
        (ImageFormat::Jpeg, (2000, 1500), ColorType::Rgb8)
    );
    assert_size_within_25_percent(jpeg_out.len(), jpeg_orig.len(), "2000x1500 JPEG");
    assert_ne!(jpeg_out, jpeg_orig);
    assert_ne!(jpeg_out, jpeg_out_2);
    assert!(
        !jpeg_out
            .windows(b"SECRET_JPEG_EXIF_METADATA".len())
            .any(|w| w == b"SECRET_JPEG_EXIF_METADATA")
    );
    let jpeg_decoded = image::load_from_memory(&jpeg_out).unwrap().to_rgb8();
    let magenta_matches = (0..128)
        .flat_map(|y| (0..128).map(move |x| (x, y)))
        .filter(|&(x, y)| {
            let [r, g, b] = jpeg_decoded.get_pixel(x, y).0;
            r > 230 && g < 25 && b > 230
        })
        .count();
    assert!(
        magenta_matches < 200,
        "source magenta patch survived in JPEG ({magenta_matches} / 16384 pixels)"
    );

    let gif_out = get_part(&parts, "word/media/anim.gif");
    assert_eq!(
        inspect(&gif_out),
        (ImageFormat::Gif, (800, 600), ColorType::Rgba8)
    );
    assert_size_within_25_percent(gif_out.len(), gif_orig.len(), "800x600 GIF");
    assert_ne!(gif_out, gif_orig);

    let bmp_out = get_part(&parts, "word/media/bitmap.bmp");
    assert_eq!(
        inspect(&bmp_out),
        (ImageFormat::Bmp, (400, 300), ColorType::Rgb8)
    );
    assert_size_within_25_percent(bmp_out.len(), bmp_orig.len(), "400x300 BMP");
    assert_ne!(bmp_out, bmp_orig);
    let bmp_decoded = image::load_from_memory(&bmp_out).unwrap().to_rgb8();
    let bmp_sentinel = (0..64)
        .flat_map(|y| (0..64).map(move |x| (x, y)))
        .filter(|&(x, y)| bmp_decoded.get_pixel(x, y).0 == [252, 11, 248])
        .count();
    assert!(bmp_sentinel < 5);

    docx_parse::parse_docx_s9_wire(&preserved_out, docx_parse::S9ParseOptions::default()).unwrap();
}

#[test]
fn preserve_media_size_preserves_color_types_alpha_and_bit_depth() {
    let options = RedactionOptions {
        random_characters: false,
        preserve_media_size: true,
    };

    let mut rgba_png = Vec::new();
    PngEncoder::new(&mut rgba_png)
        .write_image(&[120u8; 160 * 120 * 4], 160, 120, ColorType::Rgba8.into())
        .unwrap();
    let mut report = RedactionReport::default();
    let out_rgba =
        media::replace_media_with_options("word/media/alpha.png", &rgba_png, &options, &mut report)
            .unwrap();
    assert_eq!(
        inspect(&out_rgba),
        (ImageFormat::Png, (160, 120), ColorType::Rgba8)
    );

    let mut l8_png = Vec::new();
    PngEncoder::new(&mut l8_png)
        .write_image(&[90u8; 200 * 100], 200, 100, ColorType::L8.into())
        .unwrap();
    let out_l8 =
        media::replace_media_with_options("word/media/gray.png", &l8_png, &options, &mut report)
            .unwrap();
    assert_eq!(
        inspect(&out_l8),
        (ImageFormat::Png, (200, 100), ColorType::L8)
    );

    let mut rgb16_png = Vec::new();
    PngEncoder::new(&mut rgb16_png)
        .write_image(&[40u8; 120 * 90 * 6], 120, 90, ColorType::Rgb16.into())
        .unwrap();
    let out_rgb16 =
        media::replace_media_with_options("word/media/deep.png", &rgb16_png, &options, &mut report)
            .unwrap();
    assert_eq!(
        inspect(&out_rgb16),
        (ImageFormat::Png, (120, 90), ColorType::Rgb16)
    );

    let mut l8_jpeg = Vec::new();
    JpegEncoder::new_with_quality(&mut l8_jpeg, 80)
        .write_image(&[128u8; 320 * 240], 320, 240, ColorType::L8.into())
        .unwrap();
    let out_l8_jpeg =
        media::replace_media_with_options("word/media/gray.jpg", &l8_jpeg, &options, &mut report)
            .unwrap();
    assert_eq!(
        inspect(&out_l8_jpeg),
        (ImageFormat::Jpeg, (320, 240), ColorType::L8)
    );

    let mut rgba_tiff = Cursor::new(Vec::new());
    DynamicImage::ImageRgba8(ImageBuffer::from_pixel(180, 140, Rgba([10, 20, 30, 200])))
        .write_to(&mut rgba_tiff, ImageFormat::Tiff)
        .unwrap();
    let rgba_tiff = rgba_tiff.into_inner();
    let out_tiff = media::replace_media_with_options(
        "word/media/scan.tiff",
        &rgba_tiff,
        &options,
        &mut report,
    )
    .unwrap();
    assert_eq!(
        inspect(&out_tiff),
        (ImageFormat::Tiff, (180, 140), ColorType::Rgba8)
    );
}

#[test]
fn preserve_media_size_reads_header_when_body_is_corrupt_and_bounds_invalid_or_oversized() {
    let options = RedactionOptions {
        random_characters: false,
        preserve_media_size: true,
    };
    let mut report = RedactionReport::default();

    let mut header_only_png = Vec::new();
    header_only_png.extend_from_slice(b"\x89PNG\r\n\x1a\n");
    header_only_png.extend_from_slice(&13u32.to_be_bytes());
    header_only_png.extend_from_slice(b"IHDR");
    header_only_png.extend_from_slice(&640u32.to_be_bytes());
    header_only_png.extend_from_slice(&400u32.to_be_bytes());
    header_only_png.push(8); // bit depth
    header_only_png.push(6); // RGBA
    header_only_png.extend_from_slice(&[0, 0, 0]);
    header_only_png.extend_from_slice(b"CORRUPT_PAYLOAD_SECRET");
    let out_corrupt_png = media::replace_media_with_options(
        "word/media/corrupt_body.png",
        &header_only_png,
        &options,
        &mut report,
    )
    .unwrap();
    assert_eq!(
        inspect(&out_corrupt_png),
        (ImageFormat::Png, (640, 400), ColorType::Rgba8)
    );
    assert!(
        !out_corrupt_png
            .windows(b"CORRUPT_PAYLOAD_SECRET".len())
            .any(|w| w == b"CORRUPT_PAYLOAD_SECRET")
    );

    let garbage = b"NOT_AN_IMAGE_AT_ALL_SECRET_BYTES";
    let out_fallback =
        media::replace_media_with_options("word/media/broken.png", garbage, &options, &mut report)
            .unwrap();
    assert_eq!(
        inspect(&out_fallback),
        (
            ImageFormat::Png,
            (media::PLACEHOLDER_SIZE, media::PLACEHOLDER_SIZE),
            ColorType::Rgb8
        )
    );
    assert!(
        !out_fallback
            .windows(b"SECRET_BYTES".len())
            .any(|w| w == b"SECRET_BYTES")
    );

    let mut oversized_png = header_only_png;
    oversized_png[16..20].copy_from_slice(&8000u32.to_be_bytes());
    oversized_png[20..24].copy_from_slice(&8000u32.to_be_bytes());
    let out_oversized = media::replace_media_with_options(
        "word/media/huge.png",
        &oversized_png,
        &options,
        &mut report,
    )
    .unwrap();
    assert_eq!(
        inspect(&out_oversized),
        (
            ImageFormat::Png,
            (media::PLACEHOLDER_SIZE, media::PLACEHOLDER_SIZE),
            ColorType::Rgb8
        )
    );
    assert!(out_oversized.len() < 1024);
    assert_eq!(report.media_parts, 3);

    let err = media::replace_media_with_options(
        "word/media/unknown_blob",
        b"UNKNOWN_BINARY_SECRET",
        &options,
        &mut report,
    )
    .unwrap_err();
    assert!(matches!(err, RedactError::Image { .. }));
}

#[test]
fn preserve_media_size_preserves_emf_wmf_and_svg_bounds_and_byte_size() {
    let options = RedactionOptions {
        random_characters: false,
        preserve_media_size: true,
    };
    let mut report = RedactionReport::default();

    let mut emf = Vec::with_capacity(48_000);
    emf.extend_from_slice(&1u32.to_le_bytes()); // EMR_HEADER
    emf.extend_from_slice(&88u32.to_le_bytes());
    emf.extend_from_slice(&0i32.to_le_bytes());
    emf.extend_from_slice(&0i32.to_le_bytes());
    emf.extend_from_slice(&1280i32.to_le_bytes());
    emf.extend_from_slice(&960i32.to_le_bytes());
    emf.extend_from_slice(&0i32.to_le_bytes());
    emf.extend_from_slice(&0i32.to_le_bytes());
    emf.extend_from_slice(&33_867i32.to_le_bytes());
    emf.extend_from_slice(&25_400i32.to_le_bytes());
    emf.extend_from_slice(&0x464D_4520u32.to_le_bytes());
    emf.extend_from_slice(&0x0001_0000u32.to_le_bytes());
    emf.extend_from_slice(&48_000u32.to_le_bytes());
    emf.extend_from_slice(&2u32.to_le_bytes());
    emf.extend_from_slice(&1u16.to_le_bytes());
    emf.extend_from_slice(&0u16.to_le_bytes());
    emf.extend_from_slice(&0u32.to_le_bytes());
    emf.extend_from_slice(&0u32.to_le_bytes());
    emf.extend_from_slice(&0u32.to_le_bytes());
    emf.extend_from_slice(&1280i32.to_le_bytes());
    emf.extend_from_slice(&960i32.to_le_bytes());
    emf.extend_from_slice(&338i32.to_le_bytes());
    emf.extend_from_slice(&254i32.to_le_bytes());
    emf.extend_from_slice(b"SECRET_EMF_DRAWING_PAYLOAD");
    emf.resize(48_000, 0xAA);

    let out_emf =
        media::replace_media_with_options("ppt/media/chart.emf", &emf, &options, &mut report)
            .unwrap();
    assert_size_within_25_percent(out_emf.len(), emf.len(), "EMF");
    assert_eq!(&out_emf[8..24], &emf[8..24], "EMF rclBounds preserved");
    assert_eq!(&out_emf[24..40], &emf[24..40], "EMF rclFrame preserved");
    assert!(
        !out_emf
            .windows(b"SECRET_EMF_DRAWING_PAYLOAD".len())
            .any(|w| w == b"SECRET_EMF_DRAWING_PAYLOAD")
    );

    let mut wmf = Vec::with_capacity(36_000);
    wmf.extend_from_slice(&0x9AC6_CDD7u32.to_le_bytes());
    wmf.extend_from_slice(&0u16.to_le_bytes());
    wmf.extend_from_slice(&10i16.to_le_bytes());
    wmf.extend_from_slice(&20i16.to_le_bytes());
    wmf.extend_from_slice(&3010i16.to_le_bytes());
    wmf.extend_from_slice(&1820i16.to_le_bytes());
    wmf.extend_from_slice(&1440u16.to_le_bytes());
    wmf.extend_from_slice(&[0; 6]);
    wmf.extend_from_slice(b"SECRET_WMF_RECORDS");
    wmf.resize(36_000, 0x55);

    let out_wmf =
        media::replace_media_with_options("word/media/clip.wmf", &wmf, &options, &mut report)
            .unwrap();
    assert_size_within_25_percent(out_wmf.len(), wmf.len(), "WMF");
    assert_eq!(
        &out_wmf[6..16],
        &wmf[6..16],
        "WMF placeable bbox+inch preserved"
    );
    assert!(
        !out_wmf
            .windows(b"SECRET_WMF_RECORDS".len())
            .any(|w| w == b"SECRET_WMF_RECORDS")
    );

    let mut svg = String::from(
        r#"<svg xmlns="http://www.w3.org/2000/svg" width="640px" height="360px" viewBox="0 0 640 360"><text>SECRET_SVG_LABEL</text><path d=""#,
    );
    while svg.len() < 24_000 {
        svg.push_str("M0 0L10 10");
    }
    svg.push_str(r#""/></svg>"#);
    let out_svg = media::replace_media_with_options(
        "ppt/media/icon.svg",
        svg.as_bytes(),
        &options,
        &mut report,
    )
    .unwrap();
    assert_eq!(out_svg.len(), svg.len());
    let svg_text = std::str::from_utf8(&out_svg).unwrap();
    assert!(svg_text.contains(r#"width="640px""#));
    assert!(svg_text.contains(r#"height="360px""#));
    assert!(svg_text.contains(r#"viewBox="0 0 640 360""#));
    assert!(!svg_text.contains("SECRET_SVG_LABEL"));

    let pkg = wrap_docx_with_media(&[
        ("word/media/chart.emf", emf),
        ("word/media/clip.wmf", wmf),
        ("word/media/icon.svg", svg.into_bytes()),
    ]);
    let (redacted_pkg, _) = redact_with_report_and_options(&pkg, Format::Docx, &options).unwrap();
    docx_parse::parse_docx_s9_wire(&redacted_pkg, docx_parse::S9ParseOptions::default()).unwrap();
}

fn bmp_header_only(width: u32, height: u32) -> Vec<u8> {
    let raw = width * height * 3;
    let mut out = Vec::with_capacity(54);
    out.extend_from_slice(b"BM");
    out.extend_from_slice(&(54 + raw).to_le_bytes());
    out.extend_from_slice(&0u32.to_le_bytes());
    out.extend_from_slice(&54u32.to_le_bytes());
    out.extend_from_slice(&40u32.to_le_bytes());
    out.extend_from_slice(&(width as i32).to_le_bytes());
    out.extend_from_slice(&(height as i32).to_le_bytes());
    out.extend_from_slice(&1u16.to_le_bytes());
    out.extend_from_slice(&24u16.to_le_bytes());
    out.extend_from_slice(&[0u8; 24]);
    out
}

#[test]
fn preserve_media_size_refuses_header_only_parts_past_the_package_budget() {
    // Each 54-byte header declares 48 MiB of pixels; twelve of them exceed the 512 MiB package
    // limit. The budget must refuse before retaining them, not fail later when rezipping.
    let header = bmp_header_only(4096, 4096);
    let names: Vec<String> = (0..12)
        .map(|i| format!("word/media/image{i}.bmp"))
        .collect();
    let media: Vec<(&str, Vec<u8>)> = names
        .iter()
        .map(|name| (name.as_str(), header.clone()))
        .collect();
    let package = wrap_docx_with_media(&media);
    let options = RedactionOptions {
        preserve_media_size: true,
        ..RedactionOptions::default()
    };
    let error = redact_with_report_and_options(&package, Format::Docx, &options).unwrap_err();
    assert!(
        matches!(&error, RedactError::Image { message, .. } if message.contains("package size limit")),
        "expected the media budget to refuse, got {error:?}"
    );
}

fn preserve_options() -> RedactionOptions {
    RedactionOptions {
        preserve_media_size: true,
        ..RedactionOptions::default()
    }
}

#[test]
fn preserve_media_size_charges_encoded_bytes_against_the_budget() {
    // 23-byte JPEG headers declaring 512x512 RGB: each input is tiny, every placeholder is not.
    let header: Vec<u8> = [
        0xFF, 0xD8, 0xFF, 0xC0, 0x00, 0x11, 0x08, 0x02, 0x00, 0x02, 0x00, 0x03, 0x01, 0x11, 0x00,
        0x02, 0x11, 0x00, 0x03, 0x11, 0x00, 0xFF, 0xD9,
    ]
    .to_vec();
    let limit = 256 * 1024;
    let mut budget = media::MediaBudget { remaining: limit };
    let mut retained = 0u64;
    let mut refused = false;
    for i in 0..60 {
        match media::replace_media_budgeted(
            &format!("word/media/image{i}.jpg"),
            &header,
            &preserve_options(),
            &mut budget,
            &mut RedactionReport::default(),
        ) {
            Ok(out) => retained += out.len() as u64,
            Err(RedactError::Image { .. }) => {
                refused = true;
                break;
            }
            Err(other) => panic!("unexpected error {other:?}"),
        }
    }
    assert!(refused, "budget never refused; retained {retained} bytes");
    assert!(
        retained <= limit,
        "retained {retained} bytes over a {limit}-byte budget"
    );
}

#[test]
fn preserve_media_size_keeps_compressed_tiff_near_its_byte_size() {
    let mut original = Cursor::new(Vec::new());
    let pixels = vec![200u8; 1024 * 1024 * 3];
    {
        let mut encoder = tiff::encoder::TiffEncoder::new(&mut original)
            .unwrap()
            .with_compression(tiff::encoder::Compression::Deflate(
                tiff::encoder::DeflateLevel::Best,
            ));
        let mut image = encoder
            .new_image::<tiff::encoder::colortype::RGB8>(1024, 1024)
            .unwrap();
        image.rows_per_strip(1024).unwrap();
        image.write_data(&pixels).unwrap();
    }
    let original = original.into_inner();
    assert!(
        original.len() < 64 * 1024,
        "fixture is {} bytes",
        original.len()
    );
    let out = media::replace_media_with_options(
        "word/media/flat.tiff",
        &original,
        &preserve_options(),
        &mut RedactionReport::default(),
    )
    .unwrap();
    let (format, dims, color) = inspect(&out);
    assert_eq!(
        (format, dims, color),
        (ImageFormat::Tiff, (1024, 1024), ColorType::Rgb8)
    );
    assert!(
        out.len() <= original.len() * 4,
        "compressed TIFF of {} bytes became {} bytes",
        original.len(),
        out.len()
    );
}

#[test]
fn preserve_media_size_unescapes_svg_dimensions() {
    let svg = br#"<svg xmlns="http://www.w3.org/2000/svg" width="&#54;40" height="&#51;60"/>"#;
    let out = media::replace_media_with_options(
        "word/media/image1.svg",
        svg,
        &preserve_options(),
        &mut RedactionReport::default(),
    )
    .unwrap();
    let text = String::from_utf8(out).unwrap();
    assert!(
        text.contains(r#"width="640" height="360""#),
        "dimensions lost: {}",
        &text[..text.len().min(200)]
    );
}

#[test]
fn greptile_svg_view_box_keeps_missing_dimensions_absent() {
    for dimensions in ["", r#" width="800""#, r#" height="600""#] {
        let svg = format!(
            r#"<svg xmlns="http://www.w3.org/2000/svg"{dimensions} viewBox="0 0 800 600"/>"#
        );
        let out = media::replace_media_with_options(
            "word/media/viewbox.svg",
            svg.as_bytes(),
            &preserve_options(),
            &mut RedactionReport::default(),
        )
        .unwrap();
        let text = String::from_utf8(out).unwrap();
        let root = text.split('>').next().unwrap();
        assert_eq!(
            root.contains(" width="),
            dimensions.contains(" width="),
            "{root}"
        );
        assert_eq!(
            root.contains(" height="),
            dimensions.contains(" height="),
            "{root}"
        );
        assert!(root.contains(r#"viewBox="0 0 800 600""#));
    }
}

#[test]
fn greptile_svg_exponent_lengths_preserve_numbers_and_units() {
    for (width, height, expected_width, expected_height) in [
        ("8e2", "6e2", "800", "600"),
        ("8E+2px", "6e+2pt", "800px", "600pt"),
        ("8e-1em", "6e-1ex", "0.8em", "0.6ex"),
    ] {
        let svg = format!(r#"<svg width="{width}" height="{height}"/>"#);
        let out = media::replace_media_with_options(
            "word/media/exponent.svg",
            svg.as_bytes(),
            &preserve_options(),
            &mut RedactionReport::default(),
        )
        .unwrap();
        let text = String::from_utf8(out).unwrap();
        assert!(
            text.contains(&format!(
                r#"width="{expected_width}" height="{expected_height}""#
            )),
            "{text}"
        );
    }
}

#[test]
fn greptile_standard_wmf_keeps_window_bounds_and_padding() {
    for (origin, extent) in [([20i16, 10], [1800i16, 3000]), ([0, 0], [-1800, 3000])] {
        let mut wmf = Vec::new();
        wmf.extend_from_slice(&1u16.to_le_bytes());
        wmf.extend_from_slice(&9u16.to_le_bytes());
        wmf.extend_from_slice(&0x0300u16.to_le_bytes());
        wmf.extend_from_slice(&18000u32.to_le_bytes());
        wmf.extend_from_slice(&0u16.to_le_bytes());
        wmf.extend_from_slice(&17978u32.to_le_bytes());
        wmf.extend_from_slice(&0u16.to_le_bytes());
        for (function, values) in [(0x020Bu16, origin), (0x020C, extent)] {
            wmf.extend_from_slice(&5u32.to_le_bytes());
            wmf.extend_from_slice(&function.to_le_bytes());
            for value in values {
                wmf.extend_from_slice(&value.to_le_bytes());
            }
        }
        wmf.extend_from_slice(&17978u32.to_le_bytes());
        wmf.extend_from_slice(&0x0626u16.to_le_bytes());
        wmf.extend_from_slice(&0u16.to_le_bytes());
        wmf.extend_from_slice(&35946u16.to_le_bytes());
        wmf.resize(35994, 0x55);
        wmf.extend_from_slice(&3u32.to_le_bytes());
        wmf.extend_from_slice(&0u16.to_le_bytes());
        let out = media::replace_media_with_options(
            "word/media/standard.wmf",
            &wmf,
            &preserve_options(),
            &mut RedactionReport::default(),
        )
        .unwrap();
        assert_size_within_25_percent(out.len(), wmf.len(), "standard WMF");
        assert_eq!(&out[..6], &wmf[..6]);
        assert_eq!(&out[18..38], &wmf[18..38]);
        assert_eq!(
            u32::from_le_bytes(out[6..10].try_into().unwrap()) as usize * 2,
            out.len()
        );
        assert!(!out.windows(64).any(|bytes| bytes == [0x55; 64]));
        let mut offset = 18;
        let mut max_record = 0;
        loop {
            let words = u32::from_le_bytes(out[offset..offset + 4].try_into().unwrap()) as usize;
            assert!(words >= 3);
            max_record = max_record.max(words);
            let function = u16::from_le_bytes(out[offset + 4..offset + 6].try_into().unwrap());
            offset += words * 2;
            assert!(offset <= out.len());
            if function == 0 {
                break;
            }
        }
        assert_eq!(offset, out.len());
        assert_eq!(
            u32::from_le_bytes(out[12..16].try_into().unwrap()) as usize,
            max_record
        );
    }
}

#[test]
fn greptile_package_budget_includes_retained_xml_after_media() {
    let header = bmp_header_only(4096, 4096);
    let names: Vec<String> = (0..10)
        .map(|i| format!("word/media/image{i}.bmp"))
        .collect();
    let mut media: Vec<(&str, Vec<u8>)> = names
        .iter()
        .map(|name| (name.as_str(), header.clone()))
        .collect();
    let mut xml = b"<retained>".to_vec();
    xml.resize(40 * 1024 * 1024, b' ');
    xml.extend_from_slice(b"</retained>");
    media.push(("word/retained.xml", xml));
    let package = wrap_docx_with_media(&media);
    let error =
        redact_with_report_and_options(&package, Format::Docx, &preserve_options()).unwrap_err();
    assert!(
        matches!(&error, RedactError::Image { message, .. } if message.contains("package size limit")),
        "expected an early media budget refusal, got {error:?}"
    );
}

#[test]
fn preserve_media_size_charges_raster_fallbacks_against_the_budget() {
    for (path, bytes) in [
        ("word/media/broken.png", b"invalid image".to_vec()),
        ("word/media/oversized.bmp", bmp_header_only(4097, 1)),
    ] {
        let mut budget = media::MediaBudget { remaining: 0 };
        let mut report = RedactionReport::default();
        let error = media::replace_media_budgeted(
            path,
            &bytes,
            &preserve_options(),
            &mut budget,
            &mut report,
        )
        .unwrap_err();
        assert!(
            matches!(error, RedactError::Image { message, .. } if message.contains("package size limit"))
        );
        assert_eq!(report.media_parts, 0);
    }
}
