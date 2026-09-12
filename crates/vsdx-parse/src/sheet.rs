use crate::VsdxError;
use crate::xml::ParseBudget;
use crate::xml::{XmlElement, XmlNode};

#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Cell {
    pub name: String,
    pub formula: Option<String>,
    pub value: Option<String>,
    pub unit: Option<String>,
    pub del: bool,
    pub other_attrs: Vec<(String, String)>,
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Section {
    pub name: String,
    pub index: Option<u32>,
    pub del: bool,
    pub children: Vec<SectionChild>,
    pub other_attrs: Vec<(String, String)>,
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    pub index: Option<u32>,
    pub name: Option<String>,
    pub local_name: Option<String>,
    pub row_type: Option<String>,
    pub del: bool,
    pub children: Vec<RowChild>,
    pub other_attrs: Vec<(String, String)>,
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Shape {
    pub id: u32,
    pub name: Option<String>,
    pub name_u: Option<String>,
    pub shape_type: Option<String>,
    pub master: Option<u32>,
    pub master_shape: Option<u32>,
    pub line_style: Option<u32>,
    pub fill_style: Option<u32>,
    pub text_style: Option<u32>,
    pub children: Vec<ShapeChild>,
    pub del: bool,
    pub other_attrs: Vec<(String, String)>,
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ShapeChild {
    Cell(Cell),
    Section(Section),
    Text(Vec<TextToken>),
    ForeignData(ForeignData),
    Shapes(Vec<ShapesChild>),
    Unknown(OpaqueXml),
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ForeignData {
    pub foreign_type: Option<String>,
    pub compression_type: Option<String>,
    pub relationship_id: Option<String>,
    pub other_attrs: Vec<(String, String)>,
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SectionChild {
    Row(Row),
    Unknown(OpaqueXml),
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum RowChild {
    Cell(Cell),
    Unknown(OpaqueXml),
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ShapesChild {
    Shape(Shape),
    Unknown(OpaqueXml),
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ConnectsChild {
    Connect(Connect),
    Unknown(OpaqueXml),
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OpaqueXml {
    pub name: String,
    pub attributes: Vec<(String, String)>,
    pub children: Vec<OpaqueXmlNode>,
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum OpaqueXmlNode {
    Element(OpaqueXml),
    Text(String),
}
impl Shape {
    pub fn cells(&self) -> impl Iterator<Item = &Cell> {
        self.children.iter().filter_map(|child| {
            if let ShapeChild::Cell(value) = child {
                Some(value)
            } else {
                None
            }
        })
    }
    pub fn sections(&self) -> impl Iterator<Item = &Section> {
        self.children.iter().filter_map(|child| {
            if let ShapeChild::Section(value) = child {
                Some(value)
            } else {
                None
            }
        })
    }
    pub fn text(&self) -> Option<&[TextToken]> {
        self.children.iter().find_map(|child| {
            if let ShapeChild::Text(value) = child {
                Some(value.as_slice())
            } else {
                None
            }
        })
    }
    pub fn foreign_data(&self) -> Option<&ForeignData> {
        self.children.iter().find_map(|child| match child {
            ShapeChild::ForeignData(value) => Some(value),
            _ => None,
        })
    }
    pub fn shapes(&self) -> impl Iterator<Item = &Shape> {
        self.children
            .iter()
            .filter_map(|child| {
                if let ShapeChild::Shapes(values) = child {
                    Some(values.iter().filter_map(|value| {
                        if let ShapesChild::Shape(value) = value {
                            Some(value)
                        } else {
                            None
                        }
                    }))
                } else {
                    None
                }
            })
            .flatten()
    }
}
impl Section {
    pub fn rows(&self) -> impl Iterator<Item = &Row> {
        self.children.iter().filter_map(|child| {
            if let SectionChild::Row(value) = child {
                Some(value)
            } else {
                None
            }
        })
    }
}
impl Row {
    pub fn cells(&self) -> impl Iterator<Item = &Cell> {
        self.children.iter().filter_map(|child| {
            if let RowChild::Cell(value) = child {
                Some(value)
            } else {
                None
            }
        })
    }
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TextToken {
    Literal(String),
    CharacterRun(u32),
    ParagraphRun(u32),
    Tab(u32),
    Field(u32),
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Connect {
    pub from_sheet: u32,
    pub from_cell: Option<String>,
    pub from_part: Option<i32>,
    pub to_sheet: u32,
    pub to_cell: Option<String>,
    pub to_part: Option<i32>,
    pub other_attrs: Vec<(String, String)>,
}
#[derive(Clone, Debug, Default, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Sheet {
    pub id: Option<u32>,
    pub children: Vec<SheetChild>,
    pub other_attrs: Vec<(String, String)>,
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SheetChild {
    Cell(Cell),
    Section(Section),
    Shapes(Vec<ShapesChild>),
    Connects(Vec<ConnectsChild>),
    Unknown(OpaqueXml),
}
impl Sheet {
    pub fn cells(&self) -> impl Iterator<Item = &Cell> {
        self.children.iter().filter_map(|child| {
            if let SheetChild::Cell(value) = child {
                Some(value)
            } else {
                None
            }
        })
    }
    pub fn sections(&self) -> impl Iterator<Item = &Section> {
        self.children.iter().filter_map(|child| {
            if let SheetChild::Section(value) = child {
                Some(value)
            } else {
                None
            }
        })
    }
    pub fn shapes(&self) -> impl Iterator<Item = &Shape> {
        self.children
            .iter()
            .filter_map(|child| {
                if let SheetChild::Shapes(values) = child {
                    Some(values.iter().filter_map(|value| {
                        if let ShapesChild::Shape(value) = value {
                            Some(value)
                        } else {
                            None
                        }
                    }))
                } else {
                    None
                }
            })
            .flatten()
    }
    pub fn connects(&self) -> impl Iterator<Item = &Connect> {
        self.children
            .iter()
            .filter_map(|child| {
                if let SheetChild::Connects(values) = child {
                    Some(values.iter().filter_map(|value| {
                        if let ConnectsChild::Connect(value) = value {
                            Some(value)
                        } else {
                            None
                        }
                    }))
                } else {
                    None
                }
            })
            .flatten()
    }
}
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct XmlRecord {
    pub name: String,
    pub attributes: Vec<(String, String)>,
}
pub(crate) fn parse_records(parent: Option<&XmlElement>) -> Vec<XmlRecord> {
    parent
        .into_iter()
        .flat_map(elements)
        .map(|element| XmlRecord {
            name: element.local_name().to_owned(),
            attributes: element.attributes.clone(),
        })
        .collect()
}
pub(crate) fn parse_sheet(
    root: &XmlElement,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Sheet, VsdxError> {
    let mut sheet = Sheet {
        id: optional_u32(root, "ID", part)?,
        other_attrs: other(root, &["ID"]),
        ..Sheet::default()
    };
    for child in elements(root) {
        sheet.children.push(match child.local_name() {
            "Cell" => SheetChild::Cell(parse_cell(child, part, budget)?),
            "Section" => SheetChild::Section(parse_section(child, part, budget)?),
            "Shapes" => SheetChild::Shapes(parse_shapes(child, part, budget)?),
            "Connects" => SheetChild::Connects(parse_connects(child, part)?),
            _ => SheetChild::Unknown(opaque(child)),
        });
    }
    Ok(sheet)
}
fn parse_shape(
    element: &XmlElement,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Shape, VsdxError> {
    budget.charge_shape(part)?;
    let mut shape = Shape {
        id: required_u32(element, "ID", part)?,
        name: attr(element, "Name"),
        name_u: attr(element, "NameU"),
        shape_type: attr(element, "Type"),
        master: optional_u32(element, "Master", part)?,
        master_shape: optional_u32(element, "MasterShape", part)?,
        line_style: optional_u32(element, "LineStyle", part)?,
        fill_style: optional_u32(element, "FillStyle", part)?,
        text_style: optional_u32(element, "TextStyle", part)?,
        children: Vec::new(),
        del: deleted(element),
        other_attrs: other(
            element,
            &[
                "ID",
                "Name",
                "NameU",
                "Type",
                "Master",
                "MasterShape",
                "LineStyle",
                "FillStyle",
                "TextStyle",
                "Del",
            ],
        ),
    };
    for child in elements(element) {
        shape.children.push(match child.local_name() {
            "Cell" => ShapeChild::Cell(parse_cell(child, part, budget)?),
            "Section" => ShapeChild::Section(parse_section(child, part, budget)?),
            "Text" => ShapeChild::Text(parse_text(child, part)?),
            "ForeignData" => ShapeChild::ForeignData(parse_foreign_data(child)),
            "Shapes" => ShapeChild::Shapes(parse_shapes(child, part, budget)?),
            _ => ShapeChild::Unknown(opaque(child)),
        });
    }
    Ok(shape)
}
fn parse_foreign_data(element: &XmlElement) -> ForeignData {
    let relationship_id = elements(element)
        .find(|child| child.local_name() == "Rel")
        .and_then(|child| {
            child
                .attributes
                .iter()
                .find(|(name, _)| name == "id" || name.ends_with(":id"))
                .map(|(_, value)| value.clone())
        });
    ForeignData {
        foreign_type: attr(element, "ForeignType"),
        compression_type: attr(element, "CompressionType"),
        relationship_id,
        other_attrs: other(element, &["ForeignType", "CompressionType"]),
    }
}
fn parse_cell(
    element: &XmlElement,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Cell, VsdxError> {
    budget.charge_cell(part)?;
    Ok(Cell {
        name: required(element, "N", part)?,
        formula: attr(element, "F"),
        value: attr(element, "V"),
        unit: attr(element, "U"),
        del: deleted(element),
        other_attrs: other(element, &["N", "F", "V", "U", "Del"]),
    })
}
fn parse_section(
    element: &XmlElement,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Section, VsdxError> {
    budget.charge_section(part)?;
    Ok(Section {
        name: required(element, "N", part)?,
        index: optional_u32(element, "IX", part)?,
        del: deleted(element),
        children: elements(element)
            .map(|child| match child.local_name() {
                "Row" => Ok(SectionChild::Row(parse_row(child, part, budget)?)),
                _ => Ok(SectionChild::Unknown(opaque(child))),
            })
            .collect::<Result<_, VsdxError>>()?,
        other_attrs: other(element, &["N", "IX", "Del"]),
    })
}
fn parse_row(
    element: &XmlElement,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Row, VsdxError> {
    budget.charge_row(part)?;
    Ok(Row {
        index: optional_u32(element, "IX", part)?,
        name: attr(element, "N"),
        local_name: attr(element, "LocalName"),
        row_type: attr(element, "T"),
        del: deleted(element),
        children: elements(element)
            .map(|child| match child.local_name() {
                "Cell" => Ok(RowChild::Cell(parse_cell(child, part, budget)?)),
                _ => Ok(RowChild::Unknown(opaque(child))),
            })
            .collect::<Result<_, VsdxError>>()?,
        other_attrs: other(element, &["IX", "N", "LocalName", "T", "Del"]),
    })
}
fn parse_shapes(
    element: &XmlElement,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Vec<ShapesChild>, VsdxError> {
    elements(element)
        .map(|child| match child.local_name() {
            "Shape" => Ok(ShapesChild::Shape(parse_shape(child, part, budget)?)),
            _ => Ok(ShapesChild::Unknown(opaque(child))),
        })
        .collect()
}
fn parse_connects(element: &XmlElement, part: &str) -> Result<Vec<ConnectsChild>, VsdxError> {
    elements(element)
        .map(|child| match child.local_name() {
            "Connect" => Ok(ConnectsChild::Connect(parse_connect(child, part)?)),
            _ => Ok(ConnectsChild::Unknown(opaque(child))),
        })
        .collect()
}
fn parse_text(element: &XmlElement, part: &str) -> Result<Vec<TextToken>, VsdxError> {
    element
        .children
        .iter()
        .map(|node| match node {
            XmlNode::Text(text) => Ok(TextToken::Literal(text.clone())),
            XmlNode::Element(marker) => match marker.local_name() {
                "cp" => Ok(TextToken::CharacterRun(required_u32(marker, "IX", part)?)),
                "pp" => Ok(TextToken::ParagraphRun(required_u32(marker, "IX", part)?)),
                "tp" => Ok(TextToken::Tab(required_u32(marker, "IX", part)?)),
                "fld" => Ok(TextToken::Field(required_u32(marker, "IX", part)?)),
                name => Err(malformed(part, format!("unknown Text child {name}"))),
            },
        })
        .collect()
}
fn parse_connect(element: &XmlElement, part: &str) -> Result<Connect, VsdxError> {
    Ok(Connect {
        from_sheet: required_u32(element, "FromSheet", part)?,
        from_cell: attr(element, "FromCell"),
        from_part: optional_i32(element, "FromPart", part)?,
        to_sheet: required_u32(element, "ToSheet", part)?,
        to_cell: attr(element, "ToCell"),
        to_part: optional_i32(element, "ToPart", part)?,
        other_attrs: other(
            element,
            &[
                "FromSheet",
                "FromCell",
                "FromPart",
                "ToSheet",
                "ToCell",
                "ToPart",
            ],
        ),
    })
}
fn opaque(element: &XmlElement) -> OpaqueXml {
    OpaqueXml {
        name: element.name.clone(),
        attributes: element.attributes.clone(),
        children: element
            .children
            .iter()
            .map(|node| match node {
                XmlNode::Text(text) => OpaqueXmlNode::Text(text.clone()),
                XmlNode::Element(element) => OpaqueXmlNode::Element(opaque(element)),
            })
            .collect(),
    }
}
fn elements(element: &XmlElement) -> impl Iterator<Item = &XmlElement> {
    element.children.iter().filter_map(|node| {
        if let XmlNode::Element(element) = node {
            Some(element)
        } else {
            None
        }
    })
}
fn attr(element: &XmlElement, name: &str) -> Option<String> {
    element.attribute(name).map(str::to_owned)
}
fn required(element: &XmlElement, name: &str, part: &str) -> Result<String, VsdxError> {
    attr(element, name).ok_or_else(|| malformed(part, format!("missing {name}")))
}
fn required_u32(element: &XmlElement, name: &str, part: &str) -> Result<u32, VsdxError> {
    required(element, name, part)?
        .parse()
        .map_err(|_| malformed(part, format!("invalid {name}")))
}
fn optional_u32(element: &XmlElement, name: &str, part: &str) -> Result<Option<u32>, VsdxError> {
    attr(element, name)
        .map(|value| {
            value
                .parse()
                .map_err(|_| malformed(part, format!("invalid {name}")))
        })
        .transpose()
}
fn optional_i32(element: &XmlElement, name: &str, part: &str) -> Result<Option<i32>, VsdxError> {
    attr(element, name)
        .map(|value| {
            value
                .parse()
                .map_err(|_| malformed(part, format!("invalid {name}")))
        })
        .transpose()
}
fn deleted(element: &XmlElement) -> bool {
    element.attribute("Del") == Some("1")
}
fn other(element: &XmlElement, known: &[&str]) -> Vec<(String, String)> {
    element
        .attributes
        .iter()
        .filter(|(name, _)| !known.contains(&name.as_str()))
        .cloned()
        .collect()
}
fn malformed(part: &str, message: String) -> VsdxError {
    VsdxError::MalformedXml {
        part: part.to_owned(),
        offset: 0,
        message,
    }
}

#[cfg(test)]
pub(crate) fn serialize_sheet(root: &str, sheet: &Sheet) -> String {
    let mut output = String::new();
    let mut root_attrs = sheet.other_attrs.clone();
    if let Some(id) = sheet.id {
        root_attrs.push(("ID".to_owned(), id.to_string()));
    }
    element_open(&mut output, root, &root_attrs);
    for child in &sheet.children {
        match child {
            SheetChild::Cell(value) => serialize_cell(&mut output, value),
            SheetChild::Section(value) => serialize_section(&mut output, value),
            SheetChild::Shapes(values) => serialize_shapes(&mut output, values),
            SheetChild::Connects(values) => serialize_connects(&mut output, values),
            SheetChild::Unknown(value) => serialize_opaque(&mut output, value),
        }
    }
    output.push_str("</");
    output.push_str(root);
    output.push('>');
    output
}
#[cfg(test)]
fn element_open(output: &mut String, name: &str, attrs: &[(String, String)]) {
    output.push('<');
    output.push_str(name);
    for (name, value) in attrs {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        escape(output, value, true);
        output.push('"');
    }
    output.push('>');
}
#[cfg(test)]
fn attrs(mut attrs: Vec<(String, String)>, other: &[(String, String)]) -> Vec<(String, String)> {
    attrs.extend(other.iter().cloned());
    attrs
}
#[cfg(test)]
fn option(attrs: &mut Vec<(String, String)>, name: &str, value: &Option<String>) {
    if let Some(value) = value {
        attrs.push((name.to_owned(), value.clone()));
    }
}
#[cfg(test)]
fn number<T: ToString>(attrs: &mut Vec<(String, String)>, name: &str, value: Option<T>) {
    if let Some(value) = value {
        attrs.push((name.to_owned(), value.to_string()));
    }
}
#[cfg(test)]
fn serialize_cell(output: &mut String, cell: &Cell) {
    let mut a = vec![("N".to_owned(), cell.name.clone())];
    option(&mut a, "F", &cell.formula);
    option(&mut a, "V", &cell.value);
    option(&mut a, "U", &cell.unit);
    if cell.del {
        a.push(("Del".to_owned(), "1".to_owned()))
    };
    element_open(output, "Cell", &attrs(a, &cell.other_attrs));
    output.push_str("</Cell>");
}
#[cfg(test)]
fn serialize_section(output: &mut String, section: &Section) {
    let mut a = vec![("N".to_owned(), section.name.clone())];
    number(&mut a, "IX", section.index);
    if section.del {
        a.push(("Del".to_owned(), "1".to_owned()))
    };
    element_open(output, "Section", &attrs(a, &section.other_attrs));
    for child in &section.children {
        match child {
            SectionChild::Row(value) => serialize_row(output, value),
            SectionChild::Unknown(value) => serialize_opaque(output, value),
        }
    }
    output.push_str("</Section>");
}
#[cfg(test)]
fn serialize_row(output: &mut String, row: &Row) {
    let mut a = Vec::new();
    number(&mut a, "IX", row.index);
    option(&mut a, "N", &row.name);
    option(&mut a, "LocalName", &row.local_name);
    option(&mut a, "T", &row.row_type);
    if row.del {
        a.push(("Del".to_owned(), "1".to_owned()))
    };
    element_open(output, "Row", &attrs(a, &row.other_attrs));
    for child in &row.children {
        match child {
            RowChild::Cell(value) => serialize_cell(output, value),
            RowChild::Unknown(value) => serialize_opaque(output, value),
        }
    }
    output.push_str("</Row>");
}
#[cfg(test)]
fn serialize_shape(output: &mut String, shape: &Shape) {
    let mut a = vec![("ID".to_owned(), shape.id.to_string())];
    option(&mut a, "Name", &shape.name);
    option(&mut a, "NameU", &shape.name_u);
    option(&mut a, "Type", &shape.shape_type);
    number(&mut a, "Master", shape.master);
    number(&mut a, "MasterShape", shape.master_shape);
    number(&mut a, "LineStyle", shape.line_style);
    number(&mut a, "FillStyle", shape.fill_style);
    number(&mut a, "TextStyle", shape.text_style);
    if shape.del {
        a.push(("Del".to_owned(), "1".to_owned()))
    };
    element_open(output, "Shape", &attrs(a, &shape.other_attrs));
    for child in &shape.children {
        match child {
            ShapeChild::Cell(value) => serialize_cell(output, value),
            ShapeChild::Section(value) => serialize_section(output, value),
            ShapeChild::Text(value) => serialize_text(output, value),
            ShapeChild::ForeignData(value) => serialize_foreign_data(output, value),
            ShapeChild::Shapes(values) => serialize_shapes(output, values),
            ShapeChild::Unknown(value) => serialize_opaque(output, value),
        }
    }
    output.push_str("</Shape>");
}
#[cfg(test)]
fn serialize_foreign_data(output: &mut String, value: &ForeignData) {
    let mut attributes = value.other_attrs.clone();
    option(&mut attributes, "ForeignType", &value.foreign_type);
    option(&mut attributes, "CompressionType", &value.compression_type);
    element_open(output, "ForeignData", &attributes);
    if let Some(id) = &value.relationship_id {
        element_open(output, "Rel", &[("r:id".to_owned(), id.clone())]);
        output.push_str("</Rel>");
    }
    output.push_str("</ForeignData>");
}
#[cfg(test)]
fn serialize_text(output: &mut String, tokens: &[TextToken]) {
    output.push_str("<Text>");
    for token in tokens {
        match token {
            TextToken::Literal(value) => escape(output, value, false),
            TextToken::CharacterRun(ix) => marker(output, "cp", *ix),
            TextToken::ParagraphRun(ix) => marker(output, "pp", *ix),
            TextToken::Tab(ix) => marker(output, "tp", *ix),
            TextToken::Field(ix) => marker(output, "fld", *ix),
        }
    }
    output.push_str("</Text>");
}
#[cfg(test)]
fn marker(output: &mut String, name: &str, ix: u32) {
    output.push('<');
    output.push_str(name);
    output.push_str(" IX=\"");
    output.push_str(&ix.to_string());
    output.push_str("\"></");
    output.push_str(name);
    output.push('>');
}
#[cfg(test)]
fn serialize_connect(output: &mut String, connect: &Connect) {
    let mut a = vec![("FromSheet".to_owned(), connect.from_sheet.to_string())];
    option(&mut a, "FromCell", &connect.from_cell);
    number(&mut a, "FromPart", connect.from_part);
    a.push(("ToSheet".to_owned(), connect.to_sheet.to_string()));
    option(&mut a, "ToCell", &connect.to_cell);
    number(&mut a, "ToPart", connect.to_part);
    element_open(output, "Connect", &attrs(a, &connect.other_attrs));
    output.push_str("</Connect>");
}
#[cfg(test)]
fn serialize_shapes(output: &mut String, values: &[ShapesChild]) {
    output.push_str("<Shapes>");
    for value in values {
        match value {
            ShapesChild::Shape(value) => serialize_shape(output, value),
            ShapesChild::Unknown(value) => serialize_opaque(output, value),
        }
    }
    output.push_str("</Shapes>");
}
#[cfg(test)]
fn serialize_connects(output: &mut String, values: &[ConnectsChild]) {
    output.push_str("<Connects>");
    for value in values {
        match value {
            ConnectsChild::Connect(value) => serialize_connect(output, value),
            ConnectsChild::Unknown(value) => serialize_opaque(output, value),
        }
    }
    output.push_str("</Connects>");
}
#[cfg(test)]
fn serialize_opaque(output: &mut String, value: &OpaqueXml) {
    element_open(output, &value.name, &value.attributes);
    for child in &value.children {
        match child {
            OpaqueXmlNode::Text(value) => escape(output, value, false),
            OpaqueXmlNode::Element(value) => serialize_opaque(output, value),
        }
    }
    output.push_str("</");
    output.push_str(&value.name);
    output.push('>');
}
#[cfg(test)]
fn escape(output: &mut String, value: &str, attribute: bool) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' if attribute => output.push_str("&quot;"),
            _ => output.push(character),
        }
    }
}
