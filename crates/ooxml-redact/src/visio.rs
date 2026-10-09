use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, XmlVersion};

use crate::mask::TextMasker;
use crate::rels::attribute_local;
use crate::{Format, RedactError};

mod relationships;
pub(crate) use relationships::normalize_relationships;

pub(crate) const NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/visio/2012/main";

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

pub(crate) fn package_plumbing(path: &str) -> bool {
    path == "[content_types].xml" || path.ends_with(".rels")
}

pub(crate) fn preserve_package_attribute(element: &str, key: &str) -> bool {
    matches!(
        (element, key),
        ("Relationship", "Id" | "Type" | "Target" | "TargetMode")
            | ("Override", "PartName" | "ContentType")
            | ("Default", "Extension" | "ContentType")
    )
}

pub(crate) fn relationship_namespace(namespace: &[u8]) -> bool {
    matches!(
        namespace,
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            | b"http://purl.oclc.org/ooxml/officeDocument/relationships"
    )
}

pub(crate) fn preserve_other_attribute(
    element: &str,
    key: &str,
    value: &str,
    drawingml: bool,
    relationship: bool,
) -> bool {
    if relationship && matches!(attribute_local(key), "id" | "embed" | "link") {
        return is_rel_id(value);
    }
    if key == "xml:space" {
        return matches!(value, "default" | "preserve");
    }
    if element == "property" && key == "fmtid" {
        return value == "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
    }
    if element == "property" && key == "pid" {
        return value.bytes().all(|byte| byte.is_ascii_digit());
    }
    if !drawingml || !crate::rels::is_unqualified(key) {
        return false;
    }
    if matches!(
        key,
        "val"
            | "pos"
            | "ang"
            | "scaled"
            | "rotWithShape"
            | "w"
            | "lim"
            | "dpi"
            | "fov"
            | "zoom"
            | "x"
            | "y"
            | "z"
            | "lat"
            | "lon"
            | "rev"
            | "dist"
            | "dir"
            | "sx"
            | "sy"
            | "kx"
            | "ky"
            | "blurRad"
            | "endA"
            | "stA"
            | "endPos"
            | "stPos"
    ) {
        if is_decimal(value) || matches!(value, "true" | "false") {
            return true;
        }
        if matches!(element, "srgbClr" | "sysClr") && key == "val" {
            return value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit())
                || matches!(
                    value,
                    "window"
                        | "windowText"
                        | "menu"
                        | "menuText"
                        | "highlight"
                        | "highlightText"
                        | "btnFace"
                        | "btnText"
                );
        }
        if element == "schemeClr" && key == "val" {
            return matches!(
                value,
                "dk1"
                    | "lt1"
                    | "dk2"
                    | "lt2"
                    | "accent1"
                    | "accent2"
                    | "accent3"
                    | "accent4"
                    | "accent5"
                    | "accent6"
                    | "hlink"
                    | "folHlink"
                    | "phClr"
                    | "bg1"
                    | "bg2"
                    | "tx1"
                    | "tx2"
            );
        }
        if element == "prstDash" && key == "val" {
            return matches!(
                value,
                "solid"
                    | "dot"
                    | "dash"
                    | "lgDash"
                    | "dashDot"
                    | "lgDashDot"
                    | "lgDashDotDot"
                    | "sysDash"
                    | "sysDot"
                    | "sysDashDot"
                    | "sysDashDotDot"
            );
        }
    }
    if key == "lastClr" {
        return value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit());
    }
    matches!(
        (key, value),
        ("cap", "rnd" | "sq" | "flat")
            | ("cmpd", "sng" | "dbl" | "thickThin" | "thinThick" | "tri")
            | ("algn", "ctr" | "in")
            | ("flip", "none" | "x" | "y" | "xy")
            | ("path", "shape" | "circle" | "rect")
            | (
                "type",
                "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow"
            )
            | ("w" | "len", "sm" | "med" | "lg")
    )
}

pub(crate) fn attribute_named(attributes: &[(String, String)], expected: &str) -> Option<String> {
    for (key, value) in attributes {
        if key == expected {
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

/// Sections eligible for preserving numeric layout and formatting cells.
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
            | "linepatterntrans"
            | "lineweight"
            | "locpinx"
            | "locpiny"
            | "locale"
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
            | "textbkgnd"
            | "textdirection"
            | "textblockverticalalign"
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
            | "Ellipse"
            | "EllipticalArcTo"
            | "InfiniteLine"
            | "NURBSTo"
            | "PolylineTo"
            | "RelCubBezTo"
            | "RelEllipticalArcTo"
            | "RelQuadBezTo"
            | "SplineStart"
            | "SplineKnot"
            | "Close"
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
        "mm" | "cm"
            | "m"
            | "in"
            | "ft"
            | "pt"
            | "pc"
            | "pica"
            | "deg"
            | "rad"
            | "dl"
            | "dp"
            | "da"
            | "bool"
            | "str"
            | "es"
            | "em"
            | "ed"
            | "ew"
    )
}

fn is_decimal(value: &str) -> bool {
    value.trim().parse::<f64>().is_ok_and(f64::is_finite)
}

/// Cached numbers, colors, booleans, and inheritance markers survive.
fn is_safe_value(value: &str) -> bool {
    is_decimal(value)
        || is_hex_color(value)
        || matches!(value, "Themed" | "Inh")
        || value.trim().eq_ignore_ascii_case("true")
        || value.trim().eq_ignore_ascii_case("false")
}

/// Cell names whose values and formulas are always user authored.
fn is_user_cell(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "value"
            | "data1"
            | "data2"
            | "data3"
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

/// Formulas survive only without string literals or non-structural references.
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
            || is_known_unit(token)
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
                    | "guard"
                    | "themeguard"
                    | "if"
                    | "and"
                    | "or"
                    | "not"
                    | "abs"
                    | "min"
                    | "max"
                    | "sqrt"
                    | "sin"
                    | "cos"
                    | "tan"
                    | "atan2"
                    | "pi"
                    | "int"
                    | "round"
                    | "nurbs"
                    | "polyline"
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

/// Structural Visio attribute values survive; everything else is redacted.
pub(crate) fn preserve_attribute(
    element: &str,
    key: &str,
    value: &str,
    section: Option<&str>,
    cell: Option<&str>,
    is_relationship: bool,
) -> bool {
    let local = attribute_local(key);
    if is_relationship && matches!(local, "id" | "embed" | "link") && is_rel_id(value) {
        return true;
    }
    if !crate::rels::is_unqualified(key) {
        return key == "xml:space" && matches!(value, "default" | "preserve");
    }
    if local.eq_ignore_ascii_case("ID") || local.eq_ignore_ascii_case("IX") {
        let trimmed = value.trim();
        return !trimmed.is_empty() && trimmed.bytes().all(|b| b.is_ascii_digit());
    }
    if local.eq_ignore_ascii_case("Del") {
        return matches!(value.trim(), "0" | "1");
    }
    if is_style_reference(local) {
        let trimmed = value.trim();
        return !trimmed.is_empty() && trimmed.bytes().all(|b| b.is_ascii_digit());
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
        if local == "N" {
            return value
                .strip_prefix("Row")
                .is_some_and(|tail| !tail.is_empty() && tail.bytes().all(|b| b.is_ascii_digit()));
        }
        if local.eq_ignore_ascii_case("T") {
            return is_known_row_type(value);
        }
        return false;
    }
    if element.eq_ignore_ascii_case("Shape") {
        if local.eq_ignore_ascii_case("Type") {
            return is_known_shape_type(value);
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
            return matches!(
                value,
                "0" | "1" | "None" | "GZip" | "JPEG" | "PNG" | "GIF" | "TIFF" | "BMP"
            );
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
    false
}

pub(crate) fn cached_formula(
    cell: Option<&str>,
    section: Option<&str>,
    cached: Option<&str>,
) -> String {
    if cell.is_some_and(|name| is_known_cell(name) && !is_user_cell(name))
        && section.is_none_or(is_safe_section)
        && let Some(value) = cached.filter(|value| is_decimal(value))
    {
        return value.to_owned();
    }
    "0".to_owned()
}

/// Replace private attributes using typed defaults where needed.
pub(crate) fn redacted_value(
    element: &str,
    key: &str,
    value: &str,
    masker: &mut TextMasker,
) -> Result<String, RedactError> {
    let local = attribute_local(key);
    if local == "typeface" || element == "FaceName" && local == "Name" {
        return Ok("Arial".to_owned());
    }
    if local.eq_ignore_ascii_case("Date") || local.eq_ignore_ascii_case("dateUtc") {
        return Ok("1970-01-01T00:00:00Z".to_owned());
    }
    if local.eq_ignore_ascii_case("UniqueID")
        || local.eq_ignore_ascii_case("BaseID")
        || local.to_ascii_lowercase().ends_with("guid")
    {
        if value.trim_start().starts_with('{') && value.trim_end().ends_with('}') {
            return Ok("{00000000-0000-0000-0000-000000000000}".to_owned());
        }
        return Ok("00000000-0000-0000-0000-000000000000".to_owned());
    }
    if element.eq_ignore_ascii_case("ColorEntry") && local.eq_ignore_ascii_case("RGB") {
        return Ok("#000000".to_owned());
    }
    if element.eq_ignore_ascii_case("ForeignData") && local.eq_ignore_ascii_case("CompressionType")
    {
        return Ok("0".to_owned());
    }
    if element.eq_ignore_ascii_case("RefBy") && local.eq_ignore_ascii_case("T") {
        return Ok("Page".to_owned());
    }
    if is_numeric_attribute(local) {
        return Ok("0".to_owned());
    }
    masker.replace(value)
}

/// Relationship reference ids are emitter-assigned counters without author text.
fn is_rel_id(value: &str) -> bool {
    value
        .strip_prefix("rId")
        .is_some_and(|tail| !tail.is_empty() && tail.bytes().all(|b| b.is_ascii_digit()))
}

/// Style references are numeric ids on any element carrying a stylesheet.
fn is_style_reference(local: &str) -> bool {
    matches!(
        local.to_ascii_lowercase().as_str(),
        "master"
            | "mastershape"
            | "linestyle"
            | "fillstyle"
            | "textstyle"
            | "defaultlinestyle"
            | "defaultfillstyle"
            | "defaulttextstyle"
            | "defaultguidestyle"
    )
}

/// Attribute names whose values are integers, decimals, or booleans.
fn is_numeric_attribute(local: &str) -> bool {
    matches!(
        local.to_ascii_lowercase().as_str(),
        "id" | "ix"
            | "del"
            | "master"
            | "mastershape"
            | "linestyle"
            | "fillstyle"
            | "textstyle"
            | "defaultlinestyle"
            | "defaultfillstyle"
            | "defaulttextstyle"
            | "defaultguidestyle"
            | "fromsheet"
            | "tosheet"
            | "frompart"
            | "topart"
            | "basedon"
            | "originalid"
            | "parentwindow"
            | "page"
            | "toppage"
            | "iconsize"
            | "alignname"
            | "patternflags"
            | "mastertype"
            | "iconupdate"
            | "hidden"
            | "iscustomname"
            | "iscustomnameu"
            | "matchbyname"
            | "windowstate"
            | "windowleft"
            | "windowtop"
            | "windowwidth"
            | "windowheight"
            | "clientwidth"
            | "clientheight"
            | "viewscale"
            | "viewcenterx"
            | "viewcentery"
            | "unicoderanges"
            | "charsets"
            | "panose"
            | "flags"
            | "schemeenum"
            | "fontidx"
            | "fillidx"
            | "lineidx"
            | "effectidx"
    )
}
