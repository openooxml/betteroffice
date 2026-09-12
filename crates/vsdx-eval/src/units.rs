//! ShapeSheet units.

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Unit {
    Number,
    Bool,
    Inches,
    Radians,
    Seconds,
}

pub fn unit(s: &str) -> Option<(Unit, f64)> {
    match s.to_ascii_lowercase().as_str() {
        "" => Some((Unit::Number, 1.)),
        "in" => Some((Unit::Inches, 1.)),
        "dl" => Some((Unit::Inches, 1.)),
        "cm" => Some((Unit::Inches, 1. / 2.54)),
        "mm" => Some((Unit::Inches, 1. / 25.4)),
        "pt" => Some((Unit::Inches, 1. / 72.)),
        "pica" => Some((Unit::Inches, 1. / 6.)),
        "ft" => Some((Unit::Inches, 12.)),
        "m" => Some((Unit::Inches, 100. / 2.54)),
        "deg" => Some((Unit::Radians, std::f64::consts::PI / 180.)),
        "rad" => Some((Unit::Radians, 1.)),
        "da" => Some((Unit::Radians, 1.)),
        "es" => Some((Unit::Seconds, 1.)),
        "em" => Some((Unit::Seconds, 60.)),
        "ed" => Some((Unit::Seconds, 24. * 60. * 60.)),
        "ew" => Some((Unit::Seconds, 7. * 24. * 60. * 60.)),
        "bool" => Some((Unit::Bool, 1.)),
        _ => None,
    }
}
