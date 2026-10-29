use quick_xml::Reader;
use quick_xml::events::Event;

use crate::PreflightError;

#[derive(Default)]
pub(crate) struct Node {
    pub(crate) name: &'static str,
    attrs: Vec<(&'static str, String)>,
    pub(crate) children: Vec<Node>,
    pub(crate) text: String,
}

impl Node {
    pub(crate) fn attr(&self, name: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(key, _)| *key == name)
            .map(|(_, value)| value.as_str())
    }
    pub(crate) fn child(&self, name: &str) -> Option<&Self> {
        self.children.iter().find(|child| child.name == name)
    }
    pub(crate) fn val(&self, name: &str) -> Option<&str> {
        self.child(name)?.attr("val")
    }
    pub(crate) fn path(&self, names: &[&str]) -> Option<&Self> {
        names.iter().try_fold(self, |node, name| node.child(name))
    }
}

fn name(namespace: u8, local: &[u8]) -> &'static str {
    match (namespace, local) {
        (b'w', b"styles") => "w:styles",
        (b'w', b"settings") => "w:settings",
        (b'w', b"numbering") => "w:numbering",
        (b'w', b"footnotes") => "w:footnotes",
        (b'w', b"endnotes") => "w:endnotes",
        (b'a', b"theme") => "a:theme",
        (_, b"Relationships") => "Relationships",
        (b'a', b"cs") => "a:cs",
        (b'a', b"ea") => "a:ea",
        (b'a', b"font") => "a:font",
        (b'a', b"fontScheme") => "a:fontScheme",
        (b'a', b"latin") => "a:latin",
        (b'a', b"majorFont") => "a:majorFont",
        (b'a', b"minorFont") => "a:minorFont",
        (b'a', b"themeElements") => "a:themeElements",
        (b'm', b"oMath") => "m:oMath",
        (b'm', b"oMathPara") => "m:oMathPara",
        (b'm', b"t") => "m:t",
        (b'w', b"abstractNum") => "w:abstractNum",
        (b'w', b"abstractNumId") => "w:abstractNumId",
        (b'w', b"b") => "w:b",
        (b'w', b"bCs") => "w:bCs",
        (b'w', b"basedOn") => "w:basedOn",
        (b'w', b"body") => "w:body",
        (b'w', b"cs") => "w:cs",
        (b'w', b"del") => "w:del",
        (b'w', b"delText") => "w:delText",
        (b'w', b"docDefaults") => "w:docDefaults",
        (b'w', b"document") => "w:document",
        (b'w', b"drawing") => "w:drawing",
        (b'w', b"endnote") => "w:endnote",
        (b'w', b"endnoteReference") => "w:endnoteReference",
        (b'w', b"fldChar") => "w:fldChar",
        (b'w', b"fldSimple") => "w:fldSimple",
        (b'w', b"footerReference") => "w:footerReference",
        (b'w', b"footnote") => "w:footnote",
        (b'w', b"footnoteReference") => "w:footnoteReference",
        (b'w', b"ftr") => "w:ftr",
        (b'w', b"gridBefore") => "w:gridBefore",
        (b'w', b"gridCol") => "w:gridCol",
        (b'w', b"gridSpan") => "w:gridSpan",
        (b'w', b"hdr") => "w:hdr",
        (b'w', b"headerReference") => "w:headerReference",
        (b'w', b"i") => "w:i",
        (b'w', b"iCs") => "w:iCs",
        (b'w', b"ilvl") => "w:ilvl",
        (b'w', b"ins") => "w:ins",
        (b'w', b"instrText") => "w:instrText",
        (b'w', b"lvl") => "w:lvl",
        (b'w', b"lvlOverride") => "w:lvlOverride",
        (b'w', b"lvlText") => "w:lvlText",
        (b'w', b"moveFrom") => "w:moveFrom",
        (b'w', b"moveTo") => "w:moveTo",
        (b'w', b"noBreakHyphen") => "w:noBreakHyphen",
        (b'w', b"num") => "w:num",
        (b'w', b"numId") => "w:numId",
        (b'w', b"numPr") => "w:numPr",
        (b'w', b"numStyleLink") => "w:numStyleLink",
        (b'w', b"p") => "w:p",
        (b'w', b"pPr") => "w:pPr",
        (b'w', b"pStyle") => "w:pStyle",
        (b'w', b"pict") => "w:pict",
        (b'w', b"r") => "w:r",
        (b'w', b"rFonts") => "w:rFonts",
        (b'w', b"rPr") => "w:rPr",
        (b'w', b"rPrDefault") => "w:rPrDefault",
        (b'w', b"rStyle") => "w:rStyle",
        (b'w', b"rtl") => "w:rtl",
        (b'w', b"sdt") => "w:sdt",
        (b'w', b"sdtContent") => "w:sdtContent",
        (b'w', b"sectPr") => "w:sectPr",
        (b'w', b"softHyphen") => "w:softHyphen",
        (b'w', b"style") => "w:style",
        (b'w', b"styleLink") => "w:styleLink",
        (b'w', b"sym") => "w:sym",
        (b'w', b"t") => "w:t",
        (b'w', b"tab") => "w:tab",
        (b'w', b"tbl") => "w:tbl",
        (b'w', b"tblGrid") => "w:tblGrid",
        (b'w', b"tblLook") => "w:tblLook",
        (b'w', b"tblPr") => "w:tblPr",
        (b'w', b"tblStyle") => "w:tblStyle",
        (b'w', b"tblStyleColBandSize") => "w:tblStyleColBandSize",
        (b'w', b"tblStylePr") => "w:tblStylePr",
        (b'w', b"tblStyleRowBandSize") => "w:tblStyleRowBandSize",
        (b'w', b"tc") => "w:tc",
        (b'w', b"tcPr") => "w:tcPr",
        (b'w', b"themeFontLang") => "w:themeFontLang",
        (b'w', b"tr") => "w:tr",
        (b'w', b"trPr") => "w:trPr",
        (b'w', b"txbxContent") => "w:txbxContent",
        (b'w', b"vMerge") => "w:vMerge",
        (b'w', b"vanish") => "w:vanish",
        (_, b"Relationship") => "Relationship",
        (_, b"AlternateContent") => "AlternateContent",
        (_, b"Choice") => "Choice",
        (_, b"Fallback") => "Fallback",
        _ => "",
    }
}

fn attribute(local: &[u8]) -> Option<&'static str> {
    Some(match local {
        b"Id" => "Id",
        b"Type" => "Type",
        b"Target" => "Target",
        b"TargetMode" => "TargetMode",
        b"id" => "id",
        b"val" => "val",
        b"type" => "type",
        b"styleId" => "styleId",
        b"default" => "default",
        b"ascii" => "ascii",
        b"hAnsi" => "hAnsi",
        b"eastAsia" => "eastAsia",
        b"cs" => "cs",
        b"bidi" => "bidi",
        b"asciiTheme" => "asciiTheme",
        b"hAnsiTheme" => "hAnsiTheme",
        b"eastAsiaTheme" => "eastAsiaTheme",
        b"cstheme" => "cstheme",
        b"hint" => "hint",
        b"script" => "script",
        b"typeface" => "typeface",
        b"numId" => "numId",
        b"abstractNumId" => "abstractNumId",
        b"ilvl" => "ilvl",
        b"char" => "char",
        b"font" => "font",
        b"firstRow" => "firstRow",
        b"lastRow" => "lastRow",
        b"firstColumn" => "firstColumn",
        b"lastColumn" => "lastColumn",
        b"noHBand" => "noHBand",
        b"noVBand" => "noVBand",
        b"fldCharType" => "fldCharType",
        b"instr" => "instr",
        b"format" => "format",
        _ => return None,
    })
}

fn namespace(uri: &[u8]) -> u8 {
    match uri {
        b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        | b"http://purl.oclc.org/ooxml/wordprocessingml/main" => b'w',
        b"http://schemas.openxmlformats.org/drawingml/2006/main"
        | b"http://purl.oclc.org/ooxml/drawingml/main" => b'a',
        b"http://schemas.openxmlformats.org/officeDocument/2006/math"
        | b"http://purl.oclc.org/ooxml/officeDocument/math" => b'm',
        _ => 0,
    }
}

#[inline(never)]
pub(crate) fn parse(bytes: &[u8]) -> Result<Node, PreflightError> {
    std::str::from_utf8(bytes).map_err(|_| PreflightError::Xml)?;
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().expand_empty_elements = true;
    let mut stack = vec![(Node::default(), 0)];
    let mut bindings = Vec::<(Vec<u8>, u8)>::new();
    let mut count = 0usize;
    loop {
        match reader.read_event().map_err(|_| PreflightError::Xml)? {
            Event::Start(start) => {
                count += 1;
                if stack.len() > 128 || count > 250_000 {
                    return Err(PreflightError::Limit);
                }
                let scope = bindings.len();
                for attr in start.attributes() {
                    let attr = attr.map_err(|_| PreflightError::Xml)?;
                    let prefix = if attr.key.as_ref() == b"xmlns" {
                        Some(&b""[..])
                    } else {
                        attr.key.as_ref().strip_prefix(b"xmlns:")
                    };
                    if let Some(prefix) = prefix {
                        let value = attr
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Implicit1_0,
                                reader.decoder(),
                            )
                            .map_err(|_| PreflightError::Xml)?;
                        bindings.push((prefix.to_vec(), namespace(value.as_bytes())));
                    }
                }
                let qualified = start.name();
                let qualified = qualified.as_ref();
                let (prefix, local) = match qualified.iter().position(|byte| *byte == b':') {
                    Some(at) => (&qualified[..at], &qualified[at + 1..]),
                    None => (&b""[..], qualified),
                };
                let ns = bindings
                    .iter()
                    .rev()
                    .find(|(bound, _)| bound.as_slice() == prefix)
                    .map_or(0, |(_, ns)| *ns);
                let mut node = Node {
                    name: name(ns, local),
                    ..Node::default()
                };
                if !node.name.is_empty() {
                    for attr in start.attributes() {
                        let attr = attr.map_err(|_| PreflightError::Xml)?;
                        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:")
                        {
                            continue;
                        }
                        let Some(key) = attribute(attr.key.local_name().as_ref()) else {
                            continue;
                        };
                        let value = attr
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Implicit1_0,
                                reader.decoder(),
                            )
                            .map_err(|_| PreflightError::Xml)?;
                        node.attrs.push((key, value.into_owned()));
                    }
                }
                stack.push((node, scope));
            }
            Event::End(_) => {
                if stack.len() < 2 {
                    return Err(PreflightError::Xml);
                }
                let (node, scope) = stack.pop().ok_or(PreflightError::Xml)?;
                if stack.len() == 1 && node.name.is_empty() {
                    return Err(PreflightError::Xml);
                }
                bindings.truncate(scope);
                let (parent, _) = stack.last_mut().ok_or(PreflightError::Xml)?;
                if node.name.is_empty() {
                    parent.children.extend(node.children);
                } else {
                    parent.children.push(node);
                }
            }
            Event::Text(text) => {
                let (node, _) = stack.last_mut().ok_or(PreflightError::Xml)?;
                if matches!(node.name, "w:t" | "w:delText" | "m:t") {
                    let value = text.decode().map_err(|_| PreflightError::Xml)?;
                    node.text.push_str(&value);
                }
            }
            Event::CData(text) => {
                let (node, _) = stack.last_mut().ok_or(PreflightError::Xml)?;
                if matches!(node.name, "w:t" | "w:delText" | "m:t") {
                    let value = text.decode().map_err(|_| PreflightError::Xml)?;
                    node.text.push_str(&value);
                }
            }
            Event::GeneralRef(reference) => {
                let (node, _) = stack.last_mut().ok_or(PreflightError::Xml)?;
                if matches!(node.name, "w:t" | "w:delText" | "m:t") {
                    if let Some(ch) = reference
                        .resolve_char_ref()
                        .map_err(|_| PreflightError::Xml)?
                    {
                        node.text.push(ch);
                    } else {
                        node.text.push_str(
                            match reference
                                .decode()
                                .map_err(|_| PreflightError::Xml)?
                                .as_ref()
                            {
                                "amp" => "&",
                                "lt" => "<",
                                "gt" => ">",
                                "quot" => "\"",
                                "apos" => "'",
                                _ => return Err(PreflightError::Xml),
                            },
                        );
                    }
                }
            }
            Event::DocType(_) => return Err(PreflightError::Xml),
            Event::Eof => break,
            _ => {}
        }
    }
    if stack.len() != 1 || stack[0].0.children.len() != 1 {
        return Err(PreflightError::Xml);
    }
    stack
        .pop()
        .and_then(|(mut root, _)| root.children.pop())
        .ok_or(PreflightError::Xml)
}
