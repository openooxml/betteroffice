//! The document audit that runs before `usvg` sees a document: elements come
//! from an allowlist, every reference names a fragment of the document, and the
//! tree those references expand into is measured before anything builds it.

use std::collections::HashMap;
use std::str::FromStr;

use resvg::usvg::roxmltree::{Document, Node};

use super::geometry::{self, MEASURE_TOLERANCE, OUTLINE_VERB_BYTES, Outline};
use super::style::{Strokes, StyleSheet};
use super::{
    ARC_CUBIC_BYTES, MAX_SVG_ATTRIBUTES, MAX_SVG_COLLECT_WORK, MAX_SVG_DEPTH,
    MAX_SVG_EXPANDED_BYTES, MAX_SVG_EXPANDED_NODES, MAX_SVG_GRADIENT_STOPS, MAX_SVG_INHERIT_WORK,
    MAX_SVG_PAINT_BYTES, MAX_SVG_PATH_BYTES, MAX_SVG_STROKE_SPAN, MAX_SVG_STROKE_VERBS,
    MAX_SVG_STYLE_WORK, SVG_TRANSIENT_BYTES, SvgRefusal, reference,
};

pub(super) const SVG_NS: &str = "http://www.w3.org/2000/svg";
const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";

/// What a document may draw with. Office icons stay within `svg g defs style
/// path linearGradient stop`; markers, filters, masks, patterns, scripts and
/// animation are not here.
pub(super) const ALLOWED: [&str; 24] = [
    "svg",
    "g",
    "defs",
    "title",
    "desc",
    "metadata",
    "style",
    "path",
    "rect",
    "circle",
    "ellipse",
    "line",
    "polyline",
    "polygon",
    "linearGradient",
    "radialGradient",
    "stop",
    "clipPath",
    "symbol",
    "use",
    "a",
    "image",
    "text",
    "tspan",
];

/// Time `usvg` spends on each ancestor it looks through for an element's
/// inherited properties, and on each attribute of that ancestor: about 17 ns
/// and 1.6 ns, measured over 60-deep chains of 62 attributes.
const ANCESTOR_NS: u64 = 20;
const ANCESTOR_ATTRIBUTE_NS: u64 = 2;
/// Time per byte of an ancestor's values, which `usvg` parses anew for each
/// instance that inherits one: about 2.5 ns measured for the nearest, charged
/// for every ancestor.
const ANCESTOR_VALUE_NS: u64 = 2;
/// Gradients one `href` chain may link: `usvg` walks the chain again for every
/// shape that paints with its head, and does not remember one that failed.
const MAX_GRADIENT_CHAIN: u64 = 4;
/// Time `usvg` spends on each gradient of a chain per shape painting with it,
/// looking its attributes up through the chain: about 50 ns measured.
const CHAIN_LINK_NS: u64 = 128;
/// Time per stop `usvg` spends converting a gradient for one shape or `use`
/// resolving it, which it repeats for a conversion that fails or yields only
/// a colour, as a radial with no radius does: about 32 ns measured.
const STOP_NS: u64 = 64;
/// Attributes `usvg` keeps on one element at most, one per name it knows.
const KEPT_ATTRIBUTES: u64 = 256;

/// Bytes `usvg` spends on one copy of a gradient, besides its stops.
const PAINT_COPY_BYTES: u64 = 256;
/// Bytes of one copied stop.
const STOP_BYTES: u64 = 12;

#[derive(Clone, Copy)]
enum Slot {
    /// Inside `metadata` or a foreign namespace: `usvg` never builds it, so a
    /// reference into it is refused rather than audited.
    Skipped,
    Element(usize),
}

/// How `usvg` follows a reference, which decides what its target must be.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Link {
    /// A `use` instantiates any element.
    Use,
    /// A `clip-path` converts a `clipPath`, and only that.
    Clip,
    /// A `fill` converts a gradient, and only that.
    Fill,
    /// A `stroke`, likewise.
    Stroke,
    /// A gradient's `href` reads the gradient it names for stops and
    /// attributes; the chain goes on only through gradients.
    Chain,
}

/// How an element takes part in painting.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Role {
    /// Converted to a path that resolves its own fill and stroke.
    Shape,
    /// Converted with its children, which inherit its paint.
    Group,
    /// A group around the element it instantiates, which also hands its own
    /// fill and stroke to any `context-fill` or `context-stroke` inside.
    Use,
    /// Never painted directly, so nothing inside inherits paint through it.
    Inert,
}

/// One of an element's `fill` and `stroke`.
#[derive(Clone, Copy, Default)]
struct Paint {
    /// Set on the element itself, so nothing inside inherits it from above.
    set: bool,
    /// Stops of the largest gradient it may name.
    stops: u64,
    /// Whether that gradient may be in `objectBoundingBox` units, which
    /// `usvg` copies for every shape that paints with it.
    per_shape: bool,
    /// Gradients the longest chain it names links, stops or none.
    chain: u64,
}

struct Element<'a> {
    node: Node<'a, 'a>,
    role: Role,
    opaque: bool,
    children: Vec<usize>,
    links: Vec<(Link, &'a str)>,
    bytes: u64,
    style: u64,
    paint: [Paint; 2],
    outline: Outline,
    parent: Option<usize>,
    /// Attributes `usvg` may keep on the element, its CSS included.
    kept: u64,
    /// Bytes of the element's values `usvg` may parse for an inheriting
    /// instance: its presentation attributes and the CSS applied to it.
    values: u64,
    /// What looking through the element's markup ancestors costs.
    above: u64,
    /// Bytes of the longest dash list the element declares.
    dash: u64,
    strokes: Strokes,
    /// Pieces the stroker may emit when `usvg` strokes this shape to measure
    /// it.
    verbs: u64,
}

/// An element's instance as `usvg` expands it. Consumers are shapes that will
/// resolve a paint above this element, in the order `[fill, stroke]`.
#[derive(Clone, Copy, Default)]
struct Expanded {
    nodes: u64,
    bytes: u64,
    style: u64,
    verbs: u64,
    /// Time `usvg` spends looking up inherited properties through ancestors,
    /// summed over every instance.
    looks: u64,
    depth: u64,
    /// Consumers inheriting the paint set above.
    open: [u64; 2],
    /// Gradient copies, the stops they carry, and every gradient paint.
    copies: u64,
    stops: u64,
    paints: u64,
    /// Clip paths `usvg` attaches, each collected against every other.
    clips: u64,
}

/// An element's edges: the first `render` targets are the children and `use`
/// instances whose paint it passes down, up to `expand` more are clip paths it
/// instantiates, and the rest are gradients it reads through a chain.
struct Edges {
    targets: Vec<usize>,
    render: usize,
    expand: usize,
}

/// Refuses a document outside the sandbox before `usvg` converts it.
pub(super) fn audit(document: &Document<'_>) -> Result<(), SvgRefusal> {
    let root = document.root_element();
    if !is_svg(root) {
        return Err(SvgRefusal::NotSvg);
    }
    let subtrees = subtrees(document);
    let mut slots: Vec<Option<Slot>> = vec![None; subtrees.len()];
    let mut ids: HashMap<&str, Vec<usize>> = HashMap::new();
    let mut elements: Vec<Element<'_>> = Vec::new();
    let mut css = 0u64;
    for node in root.descendants().filter(|node| node.is_element()) {
        for attribute in node
            .attributes()
            .filter(|attribute| attribute.name() == "id")
        {
            ids.entry(attribute.value())
                .or_default()
                .push(node.id().get_usize());
        }
        let parent = node
            .parent_element()
            .and_then(|parent| slots[parent.id().get_usize()]);
        let skipped = match parent {
            Some(Slot::Element(parent)) => elements[parent].opaque,
            Some(Slot::Skipped) => true,
            None => false,
        };
        if skipped || !is_svg(node) {
            slots[node.id().get_usize()] = Some(Slot::Skipped);
            continue;
        }
        let name = node.tag_name().name();
        if !ALLOWED.contains(&name) {
            return Err(SvgRefusal::UnsupportedElement);
        }
        if elements.len() as u64 >= MAX_SVG_EXPANDED_NODES {
            return Err(SvgRefusal::ExpansionTooLarge);
        }
        let index = elements.len();
        let mut element = Element {
            node,
            role: role(name),
            opaque: name == "metadata",
            children: Vec::new(),
            links: Vec::new(),
            bytes: 0,
            style: 0,
            paint: [Paint::default(); 2],
            outline: Outline::default(),
            parent: match parent {
                Some(Slot::Element(parent)) => Some(parent),
                _ => None,
            },
            kept: node.attributes().len() as u64,
            values: node
                .attributes()
                .filter(|attribute| {
                    !matches!(
                        attribute.name(),
                        "id" | "class" | "style" | "href" | "d" | "points"
                    )
                })
                .map(|attribute| attribute.value().len() as u64)
                .sum(),
            above: 0,
            dash: 0,
            strokes: Strokes::default(),
            verbs: 0,
        };
        audit_attributes(&mut element, &mut css)?;
        element.bytes += node.children().count() as u64;
        elements.push(element);
        if let Some(Slot::Element(parent)) = parent {
            elements[parent].children.push(index);
        }
        slots[node.id().get_usize()] = Some(Slot::Element(index));
    }

    let mut stops = vec![0u64; elements.len()];
    let mut gradients = 0u64;
    for (index, element) in elements.iter().enumerate() {
        if is_gradient(element.node) {
            gradients += 1;
            stops[index] = element
                .children
                .iter()
                .filter(|child| elements[**child].node.tag_name().name() == "stop")
                .count() as u64;
            if stops[index] > MAX_SVG_GRADIENT_STOPS as u64 {
                return Err(SvgRefusal::ExpansionTooLarge);
            }
        }
    }

    let sheet = StyleSheet::collect(document)?;
    let mut work = sheet.work();
    for element in &mut elements {
        element.style = element.style.saturating_add(sheet.tests(element.node));
        work = work.saturating_add(element.style);
    }
    if work > MAX_SVG_STYLE_WORK {
        return Err(SvgRefusal::ExpansionTooLarge);
    }
    elements[0].style = elements[0].style.saturating_add(sheet.work());
    let dash = elements
        .iter()
        .map(|element| element.dash)
        .fold(sheet.dash(), u64::max);
    for element in elements
        .iter_mut()
        .filter(|element| matches!(element.role, Role::Shape | Role::Use))
    {
        element.bytes = element.bytes.saturating_add(dash);
    }
    measure_strokes(&mut elements, sheet.strokes())?;
    let room = MAX_SVG_EXPANDED_NODES as usize;
    let mut links = 0usize;
    for element in &mut elements {
        let mut overflow = false;
        let insert = declaration_work(element.node);
        sheet.each_match(element.node, |declarations, colons, references| {
            element.style = element
                .style
                .saturating_add(declarations.saturating_mul(insert));
            element.kept = element.kept.saturating_add(colons);
            element.values = element.values.saturating_add(declarations);
            if element.links.len() + 3 * references.len() > room {
                overflow = true;
                return;
            }
            for &target in references {
                css_links(&mut element.links, target);
            }
        });
        links += element.links.len();
        if overflow || links > room {
            return Err(SvgRefusal::ExpansionTooLarge);
        }
    }
    for index in 0..elements.len() {
        if let Some(parent) = elements[index].parent {
            elements[index].above = elements[parent]
                .above
                .saturating_add(weight(&elements[parent]));
        }
    }

    let mut total = 0usize;
    let mut visits = MAX_SVG_ATTRIBUTES;
    let mut edges = Vec::with_capacity(elements.len());
    let mut paints = Vec::with_capacity(elements.len());
    let mut scans = Vec::with_capacity(elements.len());
    for element in &elements {
        let mut resolve = |link: Link, targets: &mut Vec<usize>| {
            let links = &element.links;
            resolve(links, link, &ids, &slots, &elements, &mut visits, targets)
        };
        let mut targets = element.children.clone();
        resolve(Link::Use, &mut targets)?;
        let render = targets.len();
        resolve(Link::Clip, &mut targets)?;
        let expand = targets.len();
        resolve(Link::Chain, &mut targets)?;
        let mut painted = Vec::new();
        resolve(Link::Fill, &mut painted)?;
        let fills = painted.len();
        resolve(Link::Stroke, &mut painted)?;
        total += targets.len() + painted.len();
        if total > room {
            return Err(SvgRefusal::ExpansionTooLarge);
        }
        let scanned = targets[element.children.len()..render]
            .iter()
            .map(|target| subtrees[elements[*target].node.id().get_usize()])
            .fold(0u64, u64::saturating_add);
        scans.push(scanned);
        edges.push(Edges {
            targets,
            render,
            expand,
        });
        paints.push((painted, fills));
    }
    for (element, scanned) in elements.iter_mut().zip(scans) {
        element.bytes = element.bytes.saturating_add(scanned);
    }

    let chains = chain_stops(&edges, &mut stops)?;
    let per_shape: Vec<bool> = elements
        .iter()
        .map(|element| is_gradient(element.node) && !user_units(element.node))
        .collect();
    for (element, (painted, fills)) in elements.iter_mut().zip(&paints) {
        for (at, &target) in painted.iter().enumerate() {
            let paint = &mut element.paint[usize::from(at >= *fills)];
            paint.stops = paint.stops.max(stops[target]);
            paint.chain = paint.chain.max(chains[target]);
            paint.per_shape |= per_shape[target];
        }
    }
    expand(&elements, &edges, gradients)
}

/// Pushes the element every `link` of this kind reaches, as far as `usvg`
/// follows it, refusing one that lands in content the audit skipped. Every
/// element carrying a named id is visited, followed or not, out of `visits`
/// shared across the document, so an id many elements repeat costs the audit
/// no more than [`MAX_SVG_ATTRIBUTES`] steps in all.
fn resolve(
    links: &[(Link, &str)],
    link: Link,
    ids: &HashMap<&str, Vec<usize>>,
    slots: &[Option<Slot>],
    elements: &[Element<'_>],
    visits: &mut usize,
    targets: &mut Vec<usize>,
) -> Result<(), SvgRefusal> {
    for &(_, id) in links.iter().filter(|(kind, _)| *kind == link) {
        for &node in ids.get(id).into_iter().flatten() {
            *visits = visits.checked_sub(1).ok_or(SvgRefusal::ExpansionTooLarge)?;
            let Some(Slot::Element(target)) = slots[node] else {
                return Err(SvgRefusal::UnsupportedElement);
            };
            if follows(link, elements[target].node) {
                targets.push(target);
            }
        }
    }
    Ok(())
}

/// Nodes of every kind in each node's markup subtree, itself included, by node
/// id: what `usvg` walks through a `use` target, text and foreign markup too,
/// looking for a `use` inside that leads back.
fn subtrees(document: &Document<'_>) -> Vec<u64> {
    let mut parents = Vec::new();
    for node in document.descendants() {
        parents.push(node.parent().map(|parent| parent.id().get_usize()));
    }
    let mut sizes = vec![1u64; parents.len()];
    for index in (0..parents.len()).rev() {
        if let Some(parent) = parents[index] {
            sizes[parent] += sizes[index];
        }
    }
    sizes
}

/// Work collecting distinct paints and clip paths: `usvg` compares each one it
/// meets against every one collected so far, about a quarter nanosecond each.
fn collection(size: &Expanded, gradients: u64) -> u64 {
    let paints = size
        .paints
        .saturating_mul(size.copies.saturating_add(gradients));
    paints.saturating_add(size.clips.saturating_mul(size.clips)) / 4
}

fn role(name: &str) -> Role {
    match name {
        "rect" | "circle" | "ellipse" | "line" | "polyline" | "polygon" | "path" => Role::Shape,
        "svg" | "g" | "a" | "symbol" => Role::Group,
        "use" => Role::Use,
        _ => Role::Inert,
    }
}

/// `usvg` reads an element as SVG when it has no namespace or the SVG one.
fn is_svg(node: Node<'_, '_>) -> bool {
    matches!(node.tag_name().namespace(), None | Some(SVG_NS))
}

fn is_gradient(node: Node<'_, '_>) -> bool {
    matches!(node.tag_name().name(), "linearGradient" | "radialGradient")
}

/// Whether a gradient is in `userSpaceOnUse` units on its own, which `usvg`
/// shares between the shapes it paints instead of copying it per shape.
fn user_units(node: Node<'_, '_>) -> bool {
    node.attributes()
        .find(|attribute| {
            attribute.name() == "gradientUnits"
                && matches!(
                    attribute.namespace(),
                    None | Some(SVG_NS | XLINK_NS | XML_NS)
                )
        })
        .is_some_and(|attribute| attribute.value() == "userSpaceOnUse")
}

/// Whether `usvg` goes on into `target` along a `link`: it converts only a
/// `clipPath` for a clip and only a gradient for a paint or a gradient chain.
fn follows(link: Link, target: Node<'_, '_>) -> bool {
    match link {
        Link::Use => true,
        Link::Clip => target.tag_name().name() == "clipPath",
        Link::Fill | Link::Stroke | Link::Chain => is_gradient(target),
    }
}

/// A target named in CSS text, whose declarations are not split out: it
/// counts as a clip, a fill and a stroke alike.
fn css_links<'a>(links: &mut Vec<(Link, &'a str)>, target: &'a str) {
    links.push((Link::Clip, target));
    links.push((Link::Fill, target));
    links.push((Link::Stroke, target));
}

/// Collects the element's markup bytes, the style work a `style` attribute
/// costs per instance, its paint, and the references `usvg` follows from it,
/// each read by the parser `usvg` reads it with. An `href` is followed only on
/// `use` and on gradients: on `a` and `image` it is inert, since nothing
/// follows a link and both image resolvers return `None`.
///
/// `usvg` reads an attribute by its local name from the SVG, XLink and XML
/// namespaces as well as from none, so a prefixed copy could shadow or stand
/// in for the one audited. Only `xlink:href`, whose precedence `usvg` settles,
/// and the inert `xlink:title`, `xml:space` and `xml:lang` may carry one.
fn audit_attributes(element: &mut Element<'_>, css: &mut u64) -> Result<(), SvgRefusal> {
    let node = element.node;
    let name = node.tag_name().name();
    element.bytes = name.len() as u64;
    let (mut href, mut xlink_href) = (None, None);
    for attribute in node.attributes() {
        let (local, value) = (attribute.name(), attribute.value());
        element.bytes += (local.len() + value.len()) as u64 + 4;
        let namespace = attribute.namespace();
        match (namespace, local) {
            (None, _) | (Some(XLINK_NS), "href" | "title") | (Some(XML_NS), "space" | "lang") => {}
            (Some(SVG_NS | XLINK_NS | XML_NS), _) => return Err(SvgRefusal::UnsupportedElement),
            _ => continue,
        }
        if local == "filter" && value != "none" {
            return Err(SvgRefusal::UnsupportedStyle);
        }
        if local == "style" {
            let rescans = super::style::rescans(value.len());
            *css = css
                .saturating_add(rescans)
                .saturating_add(super::style::tokenized(value));
            if *css > MAX_SVG_STYLE_WORK {
                return Err(SvgRefusal::ExpansionTooLarge);
            }
            element.kept += value.bytes().filter(|byte| *byte == b':').count() as u64;
            element.values += value.len() as u64;
            let applied = (value.len() as u64).saturating_mul(declaration_work(node));
            element.style = element
                .style
                .saturating_add(rescans)
                .saturating_add(applied);
            super::style::screen(value)?;
            element.dash = element.dash.max(element.strokes.read(value)?);
            let mut targets = Vec::new();
            reference::css(value, &mut targets)?;
            for target in targets {
                css_links(&mut element.links, target);
            }
            continue;
        }
        match local {
            "href" if namespace.is_none() => href = href.or(Some(value)),
            "href" if namespace == Some(XLINK_NS) => xlink_href = xlink_href.or(Some(value)),
            "clip-path" if value == "inherit" => return Err(SvgRefusal::UnsupportedStyle),
            "clip-path" => {
                if let Some(target) = reference::func_iri(value)? {
                    element.links.push((Link::Clip, target));
                }
            }
            "fill" | "stroke" => {
                let (link, slot) = match local {
                    "fill" => (Link::Fill, 0),
                    _ => (Link::Stroke, 1),
                };
                element.paint[slot].set |= value != "inherit";
                element.strokes.stroked |= slot == 1 && value.trim() != "none";
                if let Some(target) = reference::paint(value)? {
                    element.links.push((link, target));
                }
            }
            "stroke-dasharray" => {
                element.dash = element.dash.max(geometry::dash_list(value)?);
            }
            "stroke-width" => {
                let width = geometry::stroke_width(value)?;
                element.strokes.width = element.strokes.width.max(width);
            }
            "transform" => {
                element.strokes.turned |= svgtypes::Transform::from_str(value)
                    .is_ok_and(|transform| transform.b != 0.0 || transform.c != 0.0);
            }
            "mask" | "marker-start" | "marker-mid" | "marker-end" => {
                reference::func_iri(value)?;
            }
            _ => {}
        }
    }
    let link = match name {
        "use" => Some(Link::Use),
        "linearGradient" | "radialGradient" => Some(Link::Chain),
        _ => None,
    };
    if let (Some(link), Some(value)) = (link, href.or(xlink_href))
        && let Some(target) = reference::href(value)?
    {
        element.links.push((link, target));
    }
    for text in node.children().filter(|child| child.is_text()) {
        element.bytes += text.text().map_or(0, str::len) as u64;
    }
    element.outline = outline(node)?;
    let arcs = (element.outline.arc_cubics as u64).saturating_mul(ARC_CUBIC_BYTES);
    let data = ["d", "points"]
        .iter()
        .flat_map(|name| geometry::attributes(node, name))
        .map(str::len)
        .sum::<usize>() as u64;
    if data.saturating_add(arcs) > MAX_SVG_PATH_BYTES as u64 {
        return Err(SvgRefusal::ExpansionTooLarge);
    }
    element.bytes = element.bytes.saturating_add(arcs);
    Ok(())
}

/// Prices what `usvg` spends stroking every shape whole to measure it, when
/// anything may be stroked: each shape is charged the pieces its outline may
/// take at the widest width declared anywhere. A rotation or skew anywhere is
/// refused, since `usvg` then strokes in canvas units the audit cannot see,
/// and so is a curve too far out for the stroker's `f32` arithmetic.
fn measure_strokes(elements: &mut [Element<'_>], sheet: Strokes) -> Result<(), SvgRefusal> {
    let mut strokes = sheet;
    for element in elements.iter() {
        strokes.join(element.strokes);
    }
    if !strokes.stroked {
        return Ok(());
    }
    if strokes.turned {
        return Err(SvgRefusal::ExpansionTooLarge);
    }
    let radius = strokes.width.max(1.0) / 2.0;
    for element in elements
        .iter_mut()
        .filter(|element| element.role == Role::Shape)
    {
        let outline = &element.outline;
        let span = (outline.reach + radius) / MEASURE_TOLERANCE;
        let verbs = geometry::stroke_verbs(outline, radius, MEASURE_TOLERANCE);
        if span.is_nan()
            || span > MAX_SVG_STROKE_SPAN
            || verbs * OUTLINE_VERB_BYTES > SVG_TRANSIENT_BYTES as f64
        {
            return Err(SvgRefusal::ExpansionTooLarge);
        }
        element.verbs = verbs as u64;
    }
    Ok(())
}

/// The outline `usvg` builds for a shape, every `d` it may read summed.
fn outline(node: Node<'_, '_>) -> Result<Outline, SvgRefusal> {
    if node.tag_name().name() != "path" {
        return geometry::shape(node);
    }
    let mut sum = Outline::default();
    for data in geometry::attributes(node, "d") {
        let outline = geometry::path(data)?;
        sum.segments += outline.segments;
        sum.contours += outline.contours;
        sum.curves += outline.curves;
        sum.length += outline.length;
        sum.reach = sum.reach.max(outline.reach);
        sum.arc_cubics += outline.arc_cubics;
    }
    Ok(sum)
}

/// Time `usvg` spends on `element` as an ancestor of one instance it looks
/// up inherited properties for.
fn weight(element: &Element<'_>) -> u64 {
    ANCESTOR_NS
        + ANCESTOR_ATTRIBUTE_NS * element.kept.min(KEPT_ATTRIBUTES)
        + ANCESTOR_VALUE_NS.saturating_mul(element.values)
}

/// Style work per byte of declarations applied to one instance of `node`:
/// each declaration looks its property up among the element's attributes
/// and copies its value.
fn declaration_work(node: Node<'_, '_>) -> u64 {
    32 + node.attributes().len() as u64
}

/// Raises each gradient's stops to the most any gradient its `href` chain
/// reaches carries, since `usvg` takes the stops of the first one with any,
/// and returns how many gradients each chain links. A chain that returns to
/// itself is refused, `usvg` would walk it forever, and so is one linking more
/// than [`MAX_GRADIENT_CHAIN`].
fn chain_stops(edges: &[Edges], stops: &mut [u64]) -> Result<Vec<u64>, SvgRefusal> {
    const NEW: u8 = 0;
    const OPEN: u8 = 1;
    const DONE: u8 = 2;
    let mut state = vec![NEW; stops.len()];
    let mut chains = vec![1u64; stops.len()];
    for start in 0..stops.len() {
        if state[start] != NEW || edges[start].targets.len() == edges[start].expand {
            continue;
        }
        let mut stack = vec![(start, edges[start].expand)];
        state[start] = OPEN;
        while let Some((index, next)) = stack.last_mut() {
            let index = *index;
            if let Some(&target) = edges[index].targets.get(*next) {
                *next += 1;
                match state[target] {
                    NEW => {
                        state[target] = OPEN;
                        stack.push((target, edges[target].expand));
                    }
                    OPEN => return Err(SvgRefusal::ReferenceCycle),
                    _ => {}
                }
                continue;
            }
            let chained = edges[index].targets[edges[index].expand..]
                .iter()
                .map(|target| stops[*target])
                .max()
                .unwrap_or(0);
            stops[index] = stops[index].max(chained);
            chains[index] = 1 + edges[index].targets[edges[index].expand..]
                .iter()
                .map(|target| chains[*target])
                .max()
                .unwrap_or(0);
            if chains[index] > MAX_GRADIENT_CHAIN {
                return Err(SvgRefusal::ExpansionTooLarge);
            }
            state[index] = DONE;
            stack.pop();
        }
    }
    Ok(chains)
}

/// Sizes the tree `usvg` would build, where every reference instantiates its
/// target again: a memoised depth-first walk, iterative so a long chain cannot
/// overflow the stack, that refuses a cycle of any length and stops at the
/// first bound. `gradients` is how many gradients the document declares, each
/// converted once and collected alongside the copies.
fn expand(elements: &[Element<'_>], edges: &[Edges], gradients: u64) -> Result<(), SvgRefusal> {
    const NEW: u8 = 0;
    const OPEN: u8 = 1;
    const DONE: u8 = 2;
    let mut state = vec![NEW; elements.len()];
    let mut sizes = vec![Expanded::default(); elements.len()];
    let mut stack = vec![(0usize, 0usize)];
    state[0] = OPEN;
    while let Some((index, next)) = stack.last_mut() {
        let index = *index;
        if let Some(&target) = edges[index].targets.get(*next) {
            *next += 1;
            match state[target] {
                NEW => {
                    state[target] = OPEN;
                    stack.push((target, 0));
                }
                OPEN => return Err(SvgRefusal::ReferenceCycle),
                _ => {}
            }
            continue;
        }
        let size = finish(elements, index, &edges[index], &sizes);
        if size.nodes > MAX_SVG_EXPANDED_NODES
            || size.bytes > MAX_SVG_EXPANDED_BYTES
            || size.style > MAX_SVG_STYLE_WORK
            || size.verbs > MAX_SVG_STROKE_VERBS
            || size.looks > MAX_SVG_INHERIT_WORK
            || size.stops.saturating_mul(STOP_BYTES) + size.copies.saturating_mul(PAINT_COPY_BYTES)
                > MAX_SVG_PAINT_BYTES
            || collection(&size, gradients) > MAX_SVG_COLLECT_WORK
        {
            return Err(SvgRefusal::ExpansionTooLarge);
        }
        if size.depth > MAX_SVG_DEPTH as u64 {
            return Err(SvgRefusal::TooDeeplyNested);
        }
        sizes[index] = size;
        state[index] = DONE;
        stack.pop();
    }
    Ok(())
}

/// One element's instance, from the instances of the elements it expands.
/// Every instance under it looks through it for inherited properties, and
/// a clip path's content also through the clip path's own ancestors. A
/// shape's fill and stroke resolve at the nearest element up its instance
/// that sets them, and a gradient there is copied per shape when its units are
/// the shape's box. A `use` resolves both itself, for any context paint inside.
fn finish(elements: &[Element<'_>], index: usize, edges: &Edges, sizes: &[Expanded]) -> Expanded {
    let own = &elements[index];
    let mut size = Expanded {
        nodes: 1,
        bytes: own.bytes,
        style: own.style,
        verbs: own.verbs,
        ..Expanded::default()
    };
    for &target in &edges.targets[..edges.expand] {
        let inner = &sizes[target];
        size.nodes = size.nodes.saturating_add(inner.nodes);
        size.bytes = size.bytes.saturating_add(inner.bytes);
        size.style = size.style.saturating_add(inner.style);
        size.verbs = size.verbs.saturating_add(inner.verbs);
        size.looks = size.looks.saturating_add(inner.looks);
        size.depth = size.depth.max(inner.depth);
        size.copies = size.copies.saturating_add(inner.copies);
        size.stops = size.stops.saturating_add(inner.stops);
        size.paints = size.paints.saturating_add(inner.paints);
        size.clips = size.clips.saturating_add(inner.clips);
    }
    size.depth += 1;
    size.looks = size
        .looks
        .saturating_add(weight(own).saturating_mul(size.nodes));
    for &target in &edges.targets[edges.render..edges.expand] {
        let inner = &sizes[target];
        size.looks = size
            .looks
            .saturating_add(inner.nodes.saturating_mul(elements[target].above));
    }
    let clipped = edges.expand > edges.render;
    let viewport = own.role == Role::Use
        || (own.node.tag_name().name() == "svg" && own.node.parent_element().is_some());
    size.clips = size
        .clips
        .saturating_add(u64::from(clipped) + u64::from(viewport));
    let mut open = [0u64; 2];
    match own.role {
        Role::Shape => open = [1, 1],
        Role::Group | Role::Use => {
            for &target in &edges.targets[..edges.render] {
                for (slot, inner) in open.iter_mut().zip(sizes[target].open) {
                    *slot = slot.saturating_add(inner);
                }
            }
        }
        Role::Inert => {}
    }
    if own.role == Role::Use {
        for slot in &mut open {
            *slot = slot.saturating_add(1);
        }
    }
    for (slot, paint) in own.paint.iter().enumerate() {
        let consumers = open[slot];
        let convert = paint
            .chain
            .saturating_mul(CHAIN_LINK_NS)
            .saturating_add(paint.stops.saturating_mul(STOP_NS));
        size.looks = size.looks.saturating_add(consumers.saturating_mul(convert));
        if paint.stops > 0 {
            let copies = if paint.per_shape { open[slot] } else { 0 };
            size.copies = size.copies.saturating_add(copies);
            size.stops = size
                .stops
                .saturating_add(copies.saturating_mul(paint.stops));
            size.paints = size.paints.saturating_add(open[slot]);
        }
    }
    for (slot, paint) in own.paint.iter().enumerate() {
        if paint.set {
            open[slot] = 0;
        }
    }
    size.open = open;
    size
}
