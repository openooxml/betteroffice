//! Typed native facade for inspecting VSDX diagrams.

use vsdx_parse::{ParseLimits, Shape, VsdxError, VsdxPackage};
use vsdx_resolve::{ResolveError, ResolvedShape, Resolver};

#[derive(Debug)]
pub enum Error {
    Parse(VsdxError),
    Resolve(ResolveError),
}

impl From<VsdxError> for Error {
    fn from(value: VsdxError) -> Self {
        Self::Parse(value)
    }
}
impl From<ResolveError> for Error {
    fn from(value: ResolveError) -> Self {
        Self::Resolve(value)
    }
}

pub type Result<T> = std::result::Result<T, Error>;

pub struct Diagram {
    package: VsdxPackage,
}

impl Diagram {
    pub fn open(bytes: &[u8]) -> Result<Self> {
        Ok(Self {
            package: vsdx_parse::parse_vsdx(bytes)?,
        })
    }
    pub fn open_with_limits(bytes: &[u8], limits: &ParseLimits) -> Result<Self> {
        Ok(Self {
            package: vsdx_parse::parse_vsdx_with_limits(bytes, limits)?,
        })
    }
    pub fn package(&self) -> &VsdxPackage {
        &self.package
    }
    pub fn pages(&self) -> impl Iterator<Item = Page<'_>> {
        self.package.page_contents.keys().map(|part| Page {
            diagram: self,
            part,
        })
    }
}

pub struct Page<'a> {
    diagram: &'a Diagram,
    part: &'a String,
}

impl<'a> Page<'a> {
    pub fn part_path(&self) -> &str {
        self.part
    }
    pub fn shapes(&'a self) -> impl Iterator<Item = ShapeView<'a>> {
        self.diagram.package.page_contents[self.part]
            .shapes()
            .map(move |shape| ShapeView { page: self, shape })
    }
}

pub struct ShapeView<'a> {
    page: &'a Page<'a>,
    shape: &'a Shape,
}

impl<'a> ShapeView<'a> {
    pub fn model(&self) -> &Shape {
        self.shape
    }
    pub fn resolved(&self) -> Result<ResolvedShape> {
        Ok(Resolver::new(&self.page.diagram.package)
            .resolve_shape(self.page.part, self.shape.id)?)
    }
}

pub type PageRef<'a> = Page<'a>;
pub type ShapeRef<'a> = ShapeView<'a>;
