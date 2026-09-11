use crate::ImageEffect;

/// Applies bitmap effects to straight RGBA pixels.
pub fn apply_image_effects(data: &mut [u8], effects: &[ImageEffect]) {
    for effect in effects {
        let (pixels, _) = data.as_chunks_mut::<4>();
        match effect {
            ImageEffect::BiLevel { threshold } => {
                let threshold = f64::from(threshold.clamp(0.0, 1.0)) * 255.0;
                for pixel in pixels {
                    let value = if luma(pixel) < threshold { 0 } else { 255 };
                    pixel[..3].fill(value);
                }
            }
            ImageEffect::Grayscale => {
                for pixel in pixels {
                    let value = luma(pixel).round() as u8;
                    pixel[..3].fill(value);
                }
            }
            ImageEffect::Luminance {
                brightness,
                contrast,
            } => {
                let map = luminance_map(f64::from(*brightness), f64::from(*contrast));
                for pixel in pixels {
                    for channel in &mut pixel[..3] {
                        *channel = map[usize::from(*channel)];
                    }
                }
            }
            ImageEffect::Duotone { shadow, highlight } => {
                let (Some(shadow), Some(highlight)) = (rgba(shadow), rgba(highlight)) else {
                    continue;
                };
                for pixel in pixels {
                    let ratio = luma(pixel) / 255.0;
                    for channel in 0..3 {
                        pixel[channel] = (f64::from(shadow[channel]) * (1.0 - ratio)
                            + f64::from(highlight[channel]) * ratio)
                            .round() as u8;
                    }
                    // An endpoint may be translucent. It modulates the source's own alpha
                    // rather than replacing it, or a transparent pixel would turn opaque.
                    let endpoint =
                        f64::from(shadow[3]) * (1.0 - ratio) + f64::from(highlight[3]) * ratio;
                    pixel[3] = (f64::from(pixel[3]) * endpoint / 255.0).round() as u8;
                }
            }
            ImageEffect::ColorChange {
                from,
                to,
                use_alpha,
            } => {
                let (Some(from), Some(to)) = (rgba(from), rgba(to)) else {
                    continue;
                };
                for pixel in pixels {
                    if pixel[..3] == from[..3] && (!use_alpha || pixel[3] == from[3]) {
                        pixel[..3].copy_from_slice(&to[..3]);
                        if *use_alpha {
                            pixel[3] = to[3];
                        }
                    }
                }
            }
        }
    }
}

/// `a:lum` as a 256-entry ramp, fitted to what LibreOffice draws for the same
/// brightness and contrast.
fn luminance_map(brightness: f64, contrast: f64) -> [u8; 256] {
    let contrast = contrast.clamp(-1.0, 1.0);
    let slope = if contrast >= 0.0 {
        128.0 / (128.0 - 127.0 * contrast)
    } else {
        (128.0 + 127.0 * contrast) / 128.0
    };
    let offset = 128.0 - 128.0 * slope + 255.0 * brightness.clamp(-1.0, 1.0) * (1.0 + slope) / 2.0;
    std::array::from_fn(|value| (slope * value as f64 + offset).round().clamp(0.0, 255.0) as u8)
}

fn luma(pixel: &[u8; 4]) -> f64 {
    0.299 * f64::from(pixel[0]) + 0.587 * f64::from(pixel[1]) + 0.114 * f64::from(pixel[2])
}

fn rgba(value: &str) -> Option<[u8; 4]> {
    let hex = value.strip_prefix('#').unwrap_or(value);
    let expanded;
    let hex = match hex.len() {
        3 | 4 => {
            expanded = hex.chars().flat_map(|c| [c, c]).collect::<String>();
            expanded.as_str()
        }
        6 | 8 => hex,
        _ => return None,
    };
    let byte = |i| u8::from_str_radix(hex.get(i..i + 2)?, 16).ok();
    Some([
        byte(0)?,
        byte(2)?,
        byte(4)?,
        if hex.len() == 8 { byte(6)? } else { 255 },
    ])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_duotone_endpoint_modulates_alpha_instead_of_replacing_it() {
        let mut data = [255, 255, 255, 200, 0, 0, 0, 0];
        apply_image_effects(
            &mut data,
            &[ImageEffect::Duotone {
                shadow: "#000000ff".to_owned(),
                highlight: "#ffffff80".to_owned(),
            }],
        );
        assert_eq!(data, [255, 255, 255, 100, 0, 0, 0, 0]);

        let mut opaque = [255, 255, 255, 200];
        apply_image_effects(
            &mut opaque,
            &[ImageEffect::Duotone {
                shadow: "#000000".to_owned(),
                highlight: "#ffffff".to_owned(),
            }],
        );
        assert_eq!(opaque, [255, 255, 255, 200]);
    }

    #[test]
    fn lum_maps_each_channel_through_the_reference_ramp() {
        let lum = |brightness, contrast| ImageEffect::Luminance {
            brightness,
            contrast,
        };
        let cases = [
            ([0, 3, 64, 128], lum(0.7, -0.7), [205, 206, 225, 128]),
            ([167, 223, 255, 64], lum(0.7, -0.7), [255, 255, 255, 64]),
            ([0, 3, 167, 255], lum(0.0, -0.5), [64, 65, 148, 255]),
            ([0, 3, 167, 255], lum(0.03, 0.77), [0, 0, 255, 255]),
            ([0, 127, 128, 255], lum(0.0, 1.0), [0, 0, 128, 255]),
            ([0, 127, 255, 255], lum(0.0, -1.0), [127, 128, 129, 255]),
            ([0, 127, 255, 255], lum(1.0, 0.0), [255, 255, 255, 255]),
            ([0, 127, 255, 255], lum(-1.0, -1.0), [0, 0, 0, 255]),
            ([3, 167, 223, 128], lum(0.0, 0.0), [3, 167, 223, 128]),
        ];
        for (mut actual, effect, expected) in cases {
            apply_image_effects(&mut actual, std::slice::from_ref(&effect));
            assert_eq!(actual, expected, "{effect:?}");
        }
    }

    #[test]
    fn effects_preserve_order_luma_and_alpha_semantics() {
        let duotone = ImageEffect::Duotone {
            shadow: "#737373ff".into(),
            highlight: "#ffffffff".into(),
        };
        let change = |use_alpha| ImageEffect::ColorChange {
            from: "#ffffffff".into(),
            to: "#ff000000".into(),
            use_alpha,
        };
        let cases = [
            (
                [3, 167, 223, 128],
                vec![ImageEffect::BiLevel { threshold: 0.5 }],
                [0, 0, 0, 128],
            ),
            (
                [3, 167, 223, 128],
                vec![ImageEffect::BiLevel { threshold: 0.25 }],
                [255, 255, 255, 128],
            ),
            (
                [3, 167, 223, 128],
                vec![ImageEffect::Grayscale],
                [124, 124, 124, 128],
            ),
            (
                [3, 167, 223, 128],
                vec![duotone.clone()],
                [183, 183, 183, 128],
            ),
            (
                [255, 255, 255, 255],
                vec![change(true), duotone.clone()],
                [157, 157, 157, 0],
            ),
            (
                [255, 255, 255, 128],
                vec![change(true)],
                [255, 255, 255, 128],
            ),
            ([255, 255, 255, 0], vec![change(true)], [255, 255, 255, 0]),
            ([255, 255, 255, 128], vec![change(false)], [255, 0, 0, 128]),
            ([255, 255, 255, 0], vec![change(false)], [255, 0, 0, 0]),
            (
                [255, 254, 255, 255],
                vec![change(false)],
                [255, 254, 255, 255],
            ),
            ([3, 167, 223, 128], vec![], [3, 167, 223, 128]),
        ];
        for (mut actual, effects, expected) in cases {
            apply_image_effects(&mut actual, &effects);
            assert_eq!(actual, expected, "{effects:?}");
        }
        let ramp = ImageEffect::Duotone {
            shadow: "#000000ff".into(),
            highlight: "#ff0000ff".into(),
        };
        let knockout = ImageEffect::ColorChange {
            from: "#ffffffff".into(),
            to: "#ffffff00".into(),
            use_alpha: true,
        };
        let mut forward = [255; 4];
        apply_image_effects(&mut forward, &[knockout.clone(), ramp.clone()]);
        assert_eq!(forward, [255, 0, 0, 0]);
        let mut backward = [255; 4];
        apply_image_effects(&mut backward, &[ramp, knockout]);
        assert_eq!(backward, [255, 0, 0, 255]);
    }
}
