use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, XmlVersion};

use crate::rels::attribute_local;
use crate::{Format, RedactError};

pub(crate) struct VisioContentTypes {
    pub(crate) accepted_drawing: bool,
    pub(crate) accepted_template: bool,
    pub(crate) refused: bool,
}

pub(crate) fn classify_content_types(bytes: &[u8]) -> Result<VisioContentTypes, RedactError> {
    let mut reader = NsReader::from_reader(bytes);
    let error = |message: String| RedactError::Xml {
        part: "[Content_Types].xml".to_owned(),
        message,
    };
    let mut result = VisioContentTypes {
        accepted_drawing: false,
        accepted_template: false,
        refused: false,
    };
    loop {
        match reader
            .read_event()
            .map_err(|value| error(value.to_string()))?
        {
            Event::Start(start) | Event::Empty(start)
                if matches!(start.local_name().as_ref(), b"Override" | b"Default") =>
            {
                let namespace = reader.resolver().resolve_element(start.name()).0;
                if !matches!(namespace, ResolveResult::Unbound)
                    && !matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == b"http://schemas.openxmlformats.org/package/2006/content-types")
                {
                    continue;
                }
                for attribute in start.attributes() {
                    let attribute = attribute.map_err(|value| error(value.to_string()))?;
                    if attribute.key.as_ref() == b"ContentType" {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(|value| error(value.to_string()))?
                            .trim()
                            .to_ascii_lowercase();
                        if value == "application/vnd.ms-visio.drawing.main+xml" {
                            result.accepted_drawing = true;
                        } else if value == "application/vnd.ms-visio.template.main+xml" {
                            result.accepted_template = true;
                        } else if value.starts_with("application/vnd.ms-visio.")
                            && value.ends_with(".main+xml")
                        {
                            result.refused = true;
                        }
                    }
                }
            }
            Event::Eof => return Ok(result),
            _ => {}
        }
    }
}

pub(crate) fn is_visio(format: Format) -> bool {
    matches!(format, Format::Vsdx | Format::Vstx)
}

/// Whether a package path holds Visio drawing XML governed by the allowlist.
pub(crate) fn is_visio_part(path: &str) -> bool {
    path.to_ascii_lowercase().starts_with("visio/")
}

pub(crate) fn is_section(element: &str) -> bool {
    element.eq_ignore_ascii_case("Section")
}

pub(crate) fn is_cell(element: &str) -> bool {
    element.eq_ignore_ascii_case("Cell")
}

/// Deny-by-default redacts every Visio text node outside package plumbing.
pub(crate) fn redact_text(path: &str) -> bool {
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".rels") || lower.ends_with("[content_types].xml") {
        return false;
    }
    true
}

pub(crate) fn attribute_named(attributes: &[(String, String)], expected: &str) -> Option<String> {
    for (key, value) in attributes {
        if attribute_local(key).eq_ignore_ascii_case(expected) {
            return Some(value.clone());
        }
    }
    None
}

pub(crate) fn ambiguous(part: &str, message: &str) -> RedactError {
    RedactError::AmbiguousVisio {
        part: part.to_owned(),
        message: message.to_owned(),
    }
}

/// Fixed ShapeSheet section vocabulary. Unknown section names are redacted.
fn is_known_section(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "action"
            | "actiontag"
            | "alignment"
            | "bevel"
            | "changeshapebehavior"
            | "character"
            | "connection"
            | "connectionabcd"
            | "control"
            | "documentproperties"
            | "field"
            | "firstcomponent"
            | "geometry"
            | "hyperlink"
            | "layer"
            | "layout"
            | "pagelayout"
            | "pageproperties"
            | "paragraph"
            | "printproperties"
            | "property"
            | "reviewer"
            | "rulergrid"
            | "scratch"
            | "shapelayout"
            | "tabs"
            | "textblock"
            | "texttransform"
            | "themeproperties"
            | "user"
    )
}

/// Sections whose cached values and formulas are provably numeric formatting.
/// Every other section, including any unknown one, has its V and F redacted.
fn is_safe_section(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "alignment"
            | "character"
            | "connection"
            | "connectionabcd"
            | "documentproperties"
            | "geometry"
            | "layout"
            | "pagelayout"
            | "pageproperties"
            | "paragraph"
            | "printproperties"
            | "rulergrid"
            | "tabs"
    )
}

/// Fixed ShapeSheet cell vocabulary. Unknown cell names are redacted.
fn is_known_cell(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "a" | "address"
            | "alignbottom"
            | "alignleft"
            | "alignright"
            | "aligntop"
            | "angle"
            | "autogen"
            | "avenue"
            | "b"
            | "beginarrow"
            | "beginarrowsize"
            | "beginx"
            | "beginy"
            | "bottommargin"
            | "bullet"
            | "bulletstr"
            | "c"
            | "case"
            | "color"
            | "colortrans"
            | "comment"
            | "compoundtype"
            | "d"
            | "data1"
            | "data2"
            | "data3"
            | "defaulttabstop"
            | "description"
            | "dirx"
            | "diry"
            | "displaymode"
            | "drawingresizetype"
            | "drawingscale"
            | "drawingsizetype"
            | "endarrow"
            | "endarrowsize"
            | "endx"
            | "endy"
            | "extrainfo"
            | "fillbkgnd"
            | "fillbkgndtrans"
            | "fillforegnd"
            | "fillforegndtrans"
            | "fillgradientangle"
            | "fillgradientdir"
            | "fillgradientenabled"
            | "fillgradientstopcount"
            | "fillgradientstops"
            | "fillpattern"
            | "flags"
            | "flipx"
            | "flipy"
            | "font"
            | "fontscale"
            | "height"
            | "horzalign"
            | "indfirst"
            | "indleft"
            | "indright"
            | "inplace"
            | "label"
            | "langid"
            | "leftmargin"
            | "letterspace"
            | "linecap"
            | "linecolor"
            | "linecolortrans"
            | "linegradientangle"
            | "linegradientdir"
            | "linegradientenabled"
            | "linepattern"
            | "lineweight"
            | "locpinx"
            | "locpiny"
            | "menu"
            | "nofill"
            | "noline"
            | "noquickdrag"
            | "noshow"
            | "nosnap"
            | "oned"
            | "pageheight"
            | "pagewidth"
            | "pinx"
            | "piny"
            | "pos"
            | "prompt"
            | "resizemode"
            | "rightmargin"
            | "rounding"
            | "shdwbkgnd"
            | "shdwbkgndtrans"
            | "shdwforegnd"
            | "shdwforegndtrans"
            | "shdwoffsetx"
            | "shdwoffsety"
            | "shdwobliqueangle"
            | "shdwpattern"
            | "shdwscalefactor"
            | "shdwtype"
            | "shapeshdwtype"
            | "shapeshdwoffsetx"
            | "shapeshdwoffsety"
            | "size"
            | "spafter"
            | "spbefore"
            | "spline"
            | "splineknot"
            | "splinestart"
            | "style"
            | "subaddress"
            | "themeindex"
            | "tooltip"
            | "topmargin"
            | "txtangle"
            | "txtheight"
            | "txtlocpinx"
            | "txtlocpiny"
            | "txtpinx"
            | "txtpiny"
            | "txtwidth"
            | "type"
            | "value"
            | "verticalalign"
            | "width"
            | "x"
            | "y"
    )
}

fn is_known_row_type(name: &str) -> bool {
    matches!(
        name,
        "MoveTo"
            | "RelMoveTo"
            | "LineTo"
            | "RelLineTo"
            | "ArcTo"
            | "RelArcTo"
            | "EllipticalArcTo"
            | "RelEllipticalArcTo"
            | "Ellipse"
            | "RelEllipse"
            | "SplineStart"
            | "RelSplineStart"
            | "SplineKnot"
            | "RelSplineKnot"
            | "PolylineTo"
            | "RelPolylineTo"
            | "InfiniteLine"
            | "RelInfiniteLine"
            | "Close"
            | "Connection"
    )
}

fn is_known_shape_type(name: &str) -> bool {
    matches!(name, "Shape" | "Group" | "Guide" | "Foreign" | "Bitmap")
}

fn is_known_foreign_type(name: &str) -> bool {
    matches!(name, "Bitmap" | "Metafile" | "OLE" | "EMF" | "Foreign")
}

fn is_known_unit(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "mm" | "cm" | "m" | "in" | "pt" | "pc" | "deg" | "rad" | "dl" | "dp"
    )
}

fn is_decimal(value: &str) -> bool {
    let value = value.trim();
    if value.is_empty() {
        return false;
    }
    let value = value.strip_prefix(['+', '-']).unwrap_or(value);
    if value.is_empty() {
        return false;
    }
    let (head, exponent) = match value.split_once(['e', 'E']) {
        Some((head, exponent)) => {
            let exponent = exponent.strip_prefix(['+', '-']).unwrap_or(exponent);
            if head.is_empty()
                || exponent.is_empty()
                || !exponent.bytes().all(|b| b.is_ascii_digit())
            {
                return false;
            }
            (head, Some(exponent))
        }
        None => (value, None),
    };
    let _ = exponent;
    let mut dot = false;
    let mut digits = false;
    for byte in head.bytes() {
        match byte {
            b'0'..=b'9' => digits = true,
            b'.' if !dot => dot = true,
            _ => return false,
        }
    }
    digits
}

/// Cached values survive only as numbers or booleans.
fn is_safe_value(value: &str) -> bool {
    is_decimal(value)
        || value.trim().eq_ignore_ascii_case("true")
        || value.trim().eq_ignore_ascii_case("false")
}

/// Cell names whose values and formulas are always user authored.
fn is_user_cell(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "value"
            | "prompt"
            | "label"
            | "address"
            | "subaddress"
            | "description"
            | "extrainfo"
            | "menu"
            | "tooltip"
            | "comment"
            | "bulletstr"
    )
}

/// Formulas survive only without string literals and without references
/// outside the structural vocabulary. Anything else is redacted.
fn is_safe_formula(value: &str) -> bool {
    if value.contains(['"', '\'']) {
        return false;
    }
    let mut body = value.trim();
    body = body.strip_prefix('=').unwrap_or(body);
    if body.trim().is_empty() {
        return true;
    }
    let mut tokens = Vec::new();
    let mut current = String::new();
    for character in body.chars() {
        if character.is_ascii_alphabetic() || character == '_' {
            current.push(character);
        } else {
            if !current.is_empty() {
                tokens.push(std::mem::take(&mut current));
            }
            if !(character.is_ascii_digit()
                || character.is_whitespace()
                || matches!(
                    character,
                    '+' | '-'
                        | '*'
                        | '/'
                        | '^'
                        | '%'
                        | '('
                        | ')'
                        | ','
                        | '.'
                        | '!'
                        | '<'
                        | '>'
                        | '&'
                        | ':'
                        | ';'
                ))
            {
                return false;
            }
        }
    }
    if !current.is_empty() {
        tokens.push(current);
    }
    tokens.iter().all(|token| {
        is_known_cell(token)
            || matches!(
                token.to_ascii_lowercase().as_str(),
                "true"
                    | "false"
                    | "inh"
                    | "noformula"
                    | "sheet"
                    | "pagesheet"
                    | "documentsheet"
                    | "connections"
                    | "themeval"
                    | "rgb"
            )
    })
}

fn is_connection_cell(value: &str) -> bool {
    if is_known_cell(value) {
        return true;
    }
    let lower = value.to_ascii_lowercase();
    if let Some(tail) = lower.strip_prefix("connections.") {
        let tail = tail.trim();
        if tail.len() < 2 {
            return false;
        }
        let (head, digits) = tail.split_at(1);
        return matches!(head, "x" | "y")
            && !digits.is_empty()
            && digits.bytes().all(|b| b.is_ascii_digit());
    }
    false
}

fn is_hex_color(value: &str) -> bool {
    let value = value.trim();
    let body = value.strip_prefix('#').unwrap_or(value);
    matches!(body.len(), 6 | 8)
        && body.bytes().all(|b| b.is_ascii_hexdigit())
        && value.starts_with('#')
}

/// Whether an alphabetic run surviving redaction is provably structural.
/// Used by the unexpected-text scan so the allowlist cannot drift from the
/// redactor: every preserved word must be recognised here.
#[cfg(test)]
pub(crate) fn is_structural_word(word: &str) -> bool {
    if word.len() < 4 {
        return true;
    }
    if word.bytes().all(|b| b == b'x' || b == b'X') {
        return true;
    }
    let lower = word.to_ascii_lowercase();
    if is_known_section(&lower) || is_known_cell(&lower) {
        return true;
    }
    matches!(
        lower.as_str(),
        "moveto"
            | "relmoveto"
            | "lineto"
            | "rellineto"
            | "arcto"
            | "relarcto"
            | "ellipticalarcto"
            | "relellipticalarcto"
            | "ellipse"
            | "relellipse"
            | "splinestart"
            | "relsplinestart"
            | "splineknot"
            | "relsplineknot"
            | "polylineto"
            | "relpolylineto"
            | "infiniteline"
            | "relinfiniteline"
            | "close"
            | "connection"
            | "shape"
            | "group"
            | "guide"
            | "foreign"
            | "bitmap"
            | "metafile"
            | "true"
            | "false"
            | "inh"
            | "noformula"
            | "sheet"
            | "pagesheet"
            | "documentsheet"
            | "connections"
            | "themeval"
            | "external"
            | "example"
            | "https"
            | "redactedproperty"
            | "redactedstyle"
            | "normal"
            | "page"
            | "pages"
            | "master"
            | "masters"
            | "theme"
            | "themes"
            | "window"
            | "windows"
            | "document"
            | "visio"
            | "comments"
            | "comment"
            | "recordsets"
            | "recordset"
            | "dataconnections"
            | "dataconnection"
            | "application"
            | "relationship"
            | "relationships"
            | "types"
            | "content"
            | "override"
            | "default"
            | "extension"
            | "partname"
            | "contenttype"
            | "target"
            | "targetmode"
            | "type"
            | "office"
            | "drawing"
            | "template"
            | "main"
            | "image"
            | "hyperlink"
            | "microsoft"
            | "schemas"
            | "openxmlformats"
            | "package"
            | "http"
            | "rels"
            | "style"
            | "styles"
            | "basedon"
            | "refby"
            | "core"
            | "custom"
            | "docprops"
            | "extended"
            | "officedocument"
            | "properties"
    )
}

/// Whether a Visio attribute value is provably structural and survives.
/// `key` is the full attribute name, `section` the enclosing Section N,
/// `cell` the enclosing Cell N.
pub(crate) fn preserve_attribute(
    element: &str,
    key: &str,
    value: &str,
    section: Option<&str>,
    cell: Option<&str>,
) -> bool {
    let local = attribute_local(key);
    if key == "r:id" {
        return true;
    }
    if local.eq_ignore_ascii_case("ID") || local.eq_ignore_ascii_case("IX") {
        let trimmed = value.trim();
        return !trimmed.is_empty() && trimmed.bytes().all(|b| b.is_ascii_digit() || b == b' ');
    }
    if local.eq_ignore_ascii_case("Del") {
        return matches!(value.trim(), "0" | "1");
    }
    if element.eq_ignore_ascii_case("Cell") {
        if local.eq_ignore_ascii_case("N") {
            return is_known_cell(value);
        }
        if local.eq_ignore_ascii_case("V") {
            let Some(cell) = cell else { return false };
            if !is_known_cell(cell) || is_user_cell(cell) {
                return false;
            }
            let safe = match section {
                None => true,
                Some(section) => is_safe_section(section),
            };
            return safe && is_safe_value(value);
        }
        if local.eq_ignore_ascii_case("F") {
            let Some(cell) = cell else { return false };
            if !is_known_cell(cell) || is_user_cell(cell) {
                return false;
            }
            let safe = match section {
                None => true,
                Some(section) => is_safe_section(section),
            };
            return safe && is_safe_formula(value);
        }
        if local.eq_ignore_ascii_case("U") {
            return is_known_unit(value);
        }
        if local.eq_ignore_ascii_case("E") {
            return false;
        }
        return false;
    }
    if element.eq_ignore_ascii_case("Section") {
        if local.eq_ignore_ascii_case("N") {
            return is_known_section(value);
        }
        return false;
    }
    if element.eq_ignore_ascii_case("Row") {
        if local.eq_ignore_ascii_case("T") {
            return is_known_row_type(value);
        }
        return false;
    }
    if element.eq_ignore_ascii_case("Shape") {
        if local.eq_ignore_ascii_case("Type") {
            return is_known_shape_type(value);
        }
        if local.eq_ignore_ascii_case("Master")
            || local.eq_ignore_ascii_case("MasterShape")
            || local.eq_ignore_ascii_case("LineStyle")
            || local.eq_ignore_ascii_case("FillStyle")
            || local.eq_ignore_ascii_case("TextStyle")
        {
            let trimmed = value.trim();
            return !trimmed.is_empty() && trimmed.bytes().all(|b| b.is_ascii_digit());
        }
        return false;
    }
    if element.eq_ignore_ascii_case("Connect") {
        if local.eq_ignore_ascii_case("FromSheet") || local.eq_ignore_ascii_case("ToSheet") {
            let trimmed = value.trim();
            return !trimmed.is_empty() && trimmed.bytes().all(|b| b.is_ascii_digit());
        }
        if local.eq_ignore_ascii_case("FromPart") || local.eq_ignore_ascii_case("ToPart") {
            let trimmed = value.trim();
            let digits = trimmed.strip_prefix('-').unwrap_or(trimmed);
            return !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit());
        }
        if local.eq_ignore_ascii_case("FromCell") || local.eq_ignore_ascii_case("ToCell") {
            return is_connection_cell(value);
        }
        return false;
    }
    if element.eq_ignore_ascii_case("FaceName") {
        return false;
    }
    if element.eq_ignore_ascii_case("ColorEntry") {
        if local.eq_ignore_ascii_case("RGB") {
            return is_hex_color(value);
        }
        return false;
    }
    if element.eq_ignore_ascii_case("StyleSheet") {
        if local.eq_ignore_ascii_case("BasedOn") {
            let trimmed = value.trim();
            return !trimmed.is_empty() && trimmed.bytes().all(|b| b.is_ascii_digit());
        }
        return false;
    }
    if element.eq_ignore_ascii_case("ForeignData") {
        if local.eq_ignore_ascii_case("ForeignType") {
            return is_known_foreign_type(value);
        }
        if local.eq_ignore_ascii_case("CompressionType") {
            return matches!(value, "0" | "1" | "None" | "GZip");
        }
        return false;
    }
    if element.eq_ignore_ascii_case("Trigger") {
        return false;
    }
    if element.eq_ignore_ascii_case("RefBy") {
        if local.eq_ignore_ascii_case("T") {
            return matches!(value, "Page" | "Master" | "Shape" | "Style" | "Document");
        }
        return false;
    }
    if element.eq_ignore_ascii_case("Rel") {
        return local.eq_ignore_ascii_case("id");
    }
    false
}

#[cfg(test)]
mod vocabulary_tests {
    use super::{
        is_known_cell, is_known_row_type, is_known_section, is_safe_section, is_user_cell,
    };

    /// Visio spells these cells in mixed case; vocabularies that lowercase before
    /// matching make an arm carrying a capital unreachable, so the name is
    /// redacted as if it were author-defined.
    #[test]
    fn canonical_visio_spellings_are_recognised() {
        for name in [
            "AutoGen",
            "DrawingResizeType",
            "ExtraInfo",
            "FillForegnd",
            "ResizeMode",
            "PinX",
            "PinY",
            "LocPinX",
            "LocPinY",
            "BeginX",
            "EndY",
            "LineWeight",
            "FillPattern",
            "NoShow",
        ] {
            assert!(
                is_known_cell(name),
                "cell {name} fell out of the vocabulary"
            );
        }
        assert!(is_known_section("Geometry"));
        assert!(is_safe_section("Character"));
        assert!(is_known_row_type("EllipticalArcTo"));
        assert!(is_user_cell("Prompt"));
    }

    /// A `matches!` arm containing a capital is unreachable when the scrutinee is
    /// lowercased first. Guards the whole class, not just the five that were wrong.
    #[test]
    fn lowercased_match_arms_have_no_unreachable_capitals() {
        let whole = include_str!("visio.rs");
        let source = &whole[..whole.find("#[cfg(test)]").unwrap_or(whole.len())];
        let needle = ".to_ascii_lowercase().as_str(),";
        let mut offenders = Vec::new();
        let mut checked = 0;
        for (index, _) in source.match_indices(needle) {
            let rest = &source[index + needle.len()..];
            let Some(end) = rest.find(
                "
    )",
            ) else {
                continue;
            };
            checked += 1;
            for arm in rest[..end].split('"').skip(1).step_by(2) {
                if arm.chars().any(|character| character.is_ascii_uppercase()) {
                    offenders.push(arm.to_owned());
                }
            }
        }
        assert!(
            checked >= 5,
            "guard found only {checked} lowercased matches"
        );
        assert!(
            offenders.is_empty(),
            "unreachable arms (scrutinee is lowercased): {offenders:?}"
        );
    }
}
