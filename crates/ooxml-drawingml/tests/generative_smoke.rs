use ooxml_drawingml::{ColorValue, resolve_color_value_to_hex};
use proptest::prelude::*;

proptest! {
    #[test]
    fn a_plain_rgb_colour_resolves_to_itself(rgb in "[0-9A-F]{6}") {
        let resolved = resolve_color_value_to_hex(Some(&ColorValue {
            rgb: Some(rgb.clone()),
            ..ColorValue::default()
        }));
        prop_assert_eq!(resolved, Some(format!("#{rgb}")));
    }
}
