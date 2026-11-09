use std::collections::{BTreeMap, BTreeSet};

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::PreflightError;

pub(crate) const MAX_NODES: usize = 250_000;
pub(crate) const MAX_VALUE_BYTES: usize = 256;
pub(crate) const MAX_NAMESPACE_DECLARATIONS: usize = 64;
pub(crate) const MAX_NAMESPACES: usize = 256;

#[derive(Default)]
pub(crate) struct Node {
    pub(crate) name: &'static str,
    attrs: Vec<(&'static str, String)>,
    pub(crate) children: Vec<Node>,
    pub(crate) text: String,
    pub(crate) incomplete: bool,
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
    #[cfg(any(feature = "docx", feature = "xlsx"))]
    pub(crate) fn val(&self, name: &str) -> Option<&str> {
        self.child(name)?.attr("val")
    }
    pub(crate) fn path(&self, names: &[&str]) -> Option<&Self> {
        names.iter().try_fold(self, |node, name| node.child(name))
    }
}

fn name(namespace: u8, local: &[u8]) -> &'static str {
    match (namespace, local) {
        #[cfg(feature = "pptx")]
        (b'p', b"presentation") => "p:presentation",
        #[cfg(feature = "pptx")]
        (b'p', b"sldIdLst") => "p:sldIdLst",
        #[cfg(feature = "pptx")]
        (b'p', b"sldId") => "p:sldId",
        #[cfg(feature = "pptx")]
        (b'p', b"defaultTextStyle") => "p:defaultTextStyle",
        #[cfg(feature = "pptx")]
        (b'p', b"sld") => "p:sld",
        #[cfg(feature = "pptx")]
        (b'p', b"sldLayout") => "p:sldLayout",
        #[cfg(feature = "pptx")]
        (b'p', b"sldMaster") => "p:sldMaster",
        #[cfg(feature = "pptx")]
        (b'p', b"cSld") => "p:cSld",
        #[cfg(feature = "pptx")]
        (b'p', b"spTree") => "p:spTree",
        #[cfg(feature = "pptx")]
        (b'p', b"sp") => "p:sp",
        #[cfg(feature = "pptx")]
        (b'p', b"grpSp") => "p:grpSp",
        #[cfg(feature = "pptx")]
        (b'p', b"graphicFrame") => "p:graphicFrame",
        #[cfg(feature = "pptx")]
        (b'p', b"pic") => "p:pic",
        #[cfg(feature = "pptx")]
        (b'p', b"cxnSp") => "p:cxnSp",
        #[cfg(feature = "pptx")]
        (b'p', b"nvSpPr") => "p:nvSpPr",
        #[cfg(feature = "pptx")]
        (b'p', b"nvGrpSpPr") => "p:nvGrpSpPr",
        #[cfg(feature = "pptx")]
        (b'p', b"nvGraphicFramePr") => "p:nvGraphicFramePr",
        #[cfg(feature = "pptx")]
        (b'p', b"nvPicPr") => "p:nvPicPr",
        #[cfg(feature = "pptx")]
        (b'p', b"nvCxnSpPr") => "p:nvCxnSpPr",
        #[cfg(feature = "pptx")]
        (b'p', b"cNvPr") => "p:cNvPr",
        #[cfg(feature = "pptx")]
        (b'p', b"nvPr") => "p:nvPr",
        #[cfg(feature = "pptx")]
        (b'p', b"ph") => "p:ph",
        #[cfg(feature = "pptx")]
        (b'p', b"txBody") => "p:txBody",
        #[cfg(feature = "pptx")]
        (b'p', b"txStyles") => "p:txStyles",
        #[cfg(feature = "pptx")]
        (b'p', b"titleStyle") => "p:titleStyle",
        #[cfg(feature = "pptx")]
        (b'p', b"bodyStyle") => "p:bodyStyle",
        #[cfg(feature = "pptx")]
        (b'p', b"otherStyle") => "p:otherStyle",
        #[cfg(feature = "pptx")]
        (b'p', b"style") => "p:style",
        #[cfg(feature = "pptx")]
        (b'p', b"extLst") => "p:extLst",
        #[cfg(feature = "pptx")]
        (b'a', b"p") => "a:p",
        #[cfg(feature = "pptx")]
        (b'a', b"pPr") => "a:pPr",
        #[cfg(feature = "pptx")]
        (b'a', b"r") => "a:r",
        #[cfg(feature = "pptx")]
        (b'a', b"rPr") => "a:rPr",
        #[cfg(feature = "pptx")]
        (b'a', b"defRPr") => "a:defRPr",
        #[cfg(feature = "pptx")]
        (b'a', b"endParaRPr") => "a:endParaRPr",
        #[cfg(feature = "pptx")]
        (b'a', b"t") => "a:t",
        #[cfg(feature = "pptx")]
        (b'a', b"fld") => "a:fld",
        #[cfg(feature = "pptx")]
        (b'a', b"br") => "a:br",
        #[cfg(feature = "pptx")]
        (b'a', b"lstStyle") => "a:lstStyle",
        #[cfg(feature = "pptx")]
        (b'a', b"defPPr") => "a:defPPr",
        #[cfg(feature = "pptx")]
        (b'a', b"lvl1pPr") => "a:lvl1pPr",
        #[cfg(feature = "pptx")]
        (b'a', b"lvl2pPr") => "a:lvl2pPr",
        #[cfg(feature = "pptx")]
        (b'a', b"lvl3pPr") => "a:lvl3pPr",
        #[cfg(feature = "pptx")]
        (b'a', b"lvl4pPr") => "a:lvl4pPr",
        #[cfg(feature = "pptx")]
        (b'a', b"lvl5pPr") => "a:lvl5pPr",
        #[cfg(feature = "pptx")]
        (b'a', b"lvl6pPr") => "a:lvl6pPr",
        #[cfg(feature = "pptx")]
        (b'a', b"lvl7pPr") => "a:lvl7pPr",
        #[cfg(feature = "pptx")]
        (b'a', b"lvl8pPr") => "a:lvl8pPr",
        #[cfg(feature = "pptx")]
        (b'a', b"lvl9pPr") => "a:lvl9pPr",
        #[cfg(feature = "pptx")]
        (b'a', b"buFont") => "a:buFont",
        #[cfg(feature = "pptx")]
        (b'a', b"buFontTx") => "a:buFontTx",
        #[cfg(feature = "pptx")]
        (b'a', b"buChar") => "a:buChar",
        #[cfg(feature = "pptx")]
        (b'a', b"buAutoNum") => "a:buAutoNum",
        #[cfg(feature = "pptx")]
        (b'a', b"buNone") => "a:buNone",
        #[cfg(feature = "pptx")]
        (b'a', b"graphic") => "a:graphic",
        #[cfg(feature = "pptx")]
        (b'a', b"graphicData") => "a:graphicData",
        #[cfg(feature = "pptx")]
        (b'a', b"tbl") => "a:tbl",
        #[cfg(feature = "pptx")]
        (b'a', b"tblPr") => "a:tblPr",
        #[cfg(feature = "pptx")]
        (b'a', b"tblGrid") => "a:tblGrid",
        #[cfg(feature = "pptx")]
        (b'a', b"gridCol") => "a:gridCol",
        #[cfg(feature = "pptx")]
        (b'a', b"tr") => "a:tr",
        #[cfg(feature = "pptx")]
        (b'a', b"tc") => "a:tc",
        #[cfg(feature = "pptx")]
        (b'a', b"tcPr") => "a:tcPr",
        #[cfg(feature = "pptx")]
        (b'a', b"txBody") => "a:txBody",
        #[cfg(feature = "pptx")]
        (b'a', b"tableStyleId") => "a:tableStyleId",
        #[cfg(feature = "pptx")]
        (b'a', b"tblStyleLst") => "a:tblStyleLst",
        #[cfg(feature = "pptx")]
        (b'a', b"tblStyle") => "a:tblStyle",
        #[cfg(feature = "pptx")]
        (b'a', b"wholeTbl") => "a:wholeTbl",
        #[cfg(feature = "pptx")]
        (b'a', b"band1H") => "a:band1H",
        #[cfg(feature = "pptx")]
        (b'a', b"band2H") => "a:band2H",
        #[cfg(feature = "pptx")]
        (b'a', b"band1V") => "a:band1V",
        #[cfg(feature = "pptx")]
        (b'a', b"band2V") => "a:band2V",
        #[cfg(feature = "pptx")]
        (b'a', b"firstRow") => "a:firstRow",
        #[cfg(feature = "pptx")]
        (b'a', b"lastRow") => "a:lastRow",
        #[cfg(feature = "pptx")]
        (b'a', b"firstCol") => "a:firstCol",
        #[cfg(feature = "pptx")]
        (b'a', b"lastCol") => "a:lastCol",
        #[cfg(feature = "pptx")]
        (b'a', b"nwCell") => "a:nwCell",
        #[cfg(feature = "pptx")]
        (b'a', b"neCell") => "a:neCell",
        #[cfg(feature = "pptx")]
        (b'a', b"swCell") => "a:swCell",
        #[cfg(feature = "pptx")]
        (b'a', b"seCell") => "a:seCell",
        #[cfg(feature = "pptx")]
        (b'a', b"tcTxStyle") => "a:tcTxStyle",
        #[cfg(feature = "pptx")]
        (b'a', b"fontRef") => "a:fontRef",
        #[cfg(feature = "pptx")]
        (b'a', b"extLst") => "a:extLst",
        #[cfg(feature = "xlsx")]
        (b's', b"workbook") => "s:workbook",
        #[cfg(feature = "xlsx")]
        (b's', b"bookViews") => "s:bookViews",
        #[cfg(feature = "xlsx")]
        (b's', b"workbookView") => "s:workbookView",
        #[cfg(feature = "xlsx")]
        (b's', b"sheets") => "s:sheets",
        #[cfg(feature = "xlsx")]
        (b's', b"sheet") => "s:sheet",
        #[cfg(feature = "xlsx")]
        (b's', b"worksheet") => "s:worksheet",
        #[cfg(feature = "xlsx")]
        (b's', b"sheetViews") => "s:sheetViews",
        #[cfg(feature = "xlsx")]
        (b's', b"sheetView") => "s:sheetView",
        #[cfg(feature = "xlsx")]
        (b's', b"pane") => "s:pane",
        #[cfg(feature = "xlsx")]
        (b's', b"sheetData") => "s:sheetData",
        #[cfg(feature = "xlsx")]
        (b's', b"row") => "s:row",
        #[cfg(feature = "xlsx")]
        (b's', b"c") => "s:c",
        #[cfg(feature = "xlsx")]
        (b's', b"v") => "s:v",
        #[cfg(feature = "xlsx")]
        (b's', b"f") => "s:f",
        #[cfg(feature = "xlsx")]
        (b's', b"is") => "s:is",
        #[cfg(feature = "xlsx")]
        (b's', b"t") => "s:t",
        #[cfg(feature = "xlsx")]
        (b's', b"r") => "s:r",
        #[cfg(feature = "xlsx")]
        (b's', b"rPr") => "s:rPr",
        #[cfg(feature = "xlsx")]
        (b's', b"rPh") => "s:rPh",
        #[cfg(feature = "xlsx")]
        (b's', b"sst") => "s:sst",
        #[cfg(feature = "xlsx")]
        (b's', b"si") => "s:si",
        #[cfg(feature = "xlsx")]
        (b's', b"styleSheet") => "s:styleSheet",
        #[cfg(feature = "xlsx")]
        (b's', b"fonts") => "s:fonts",
        #[cfg(feature = "xlsx")]
        (b's', b"font") => "s:font",
        #[cfg(feature = "xlsx")]
        (b's', b"name") => "s:name",
        #[cfg(feature = "xlsx")]
        (b's', b"b") => "s:b",
        #[cfg(feature = "xlsx")]
        (b's', b"i") => "s:i",
        #[cfg(feature = "xlsx")]
        (b's', b"sz") => "s:sz",
        #[cfg(feature = "xlsx")]
        (b's', b"scheme") => "s:scheme",
        #[cfg(feature = "xlsx")]
        (b's', b"cellXfs") => "s:cellXfs",
        #[cfg(feature = "xlsx")]
        (b's', b"xf") => "s:xf",
        #[cfg(feature = "xlsx")]
        (b's', b"cellStyleXfs") => "s:cellStyleXfs",
        #[cfg(feature = "xlsx")]
        (b's', b"numFmts") => "s:numFmts",
        #[cfg(feature = "xlsx")]
        (b's', b"numFmt") => "s:numFmt",
        #[cfg(feature = "xlsx")]
        (b's', b"cols") => "s:cols",
        #[cfg(feature = "xlsx")]
        (b's', b"col") => "s:col",
        #[cfg(feature = "xlsx")]
        (b's', b"mergeCells") => "s:mergeCells",
        #[cfg(feature = "xlsx")]
        (b's', b"mergeCell") => "s:mergeCell",
        #[cfg(feature = "xlsx")]
        (b's', b"sheetFormatPr") => "s:sheetFormatPr",
        #[cfg(feature = "xlsx")]
        (b's', b"hyperlinks") => "s:hyperlinks",
        #[cfg(feature = "xlsx")]
        (b's', b"hyperlink") => "s:hyperlink",
        #[cfg(feature = "xlsx")]
        (b's', b"conditionalFormatting") => "s:conditionalFormatting",
        #[cfg(feature = "xlsx")]
        (b's', b"drawing") => "s:drawing",
        #[cfg(feature = "xlsx")]
        (b's', b"extLst") => "s:extLst",
        #[cfg(feature = "pptx")]
        (b'c', b"chart") => "c:chart",
        #[cfg(feature = "pptx")]
        (b'd', b"relIds") => "d:relIds",
        #[cfg(feature = "docx")]
        (b'w', b"styles") => "w:styles",
        #[cfg(feature = "docx")]
        (b'w', b"settings") => "w:settings",
        #[cfg(feature = "docx")]
        (b'w', b"numbering") => "w:numbering",
        #[cfg(feature = "docx")]
        (b'w', b"footnotes") => "w:footnotes",
        #[cfg(feature = "docx")]
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
        #[cfg(feature = "docx")]
        (b'm', b"oMath") => "m:oMath",
        #[cfg(feature = "docx")]
        (b'm', b"oMathPara") => "m:oMathPara",
        #[cfg(feature = "docx")]
        (b'm', b"t") => "m:t",
        #[cfg(feature = "docx")]
        (b'w', b"abstractNum") => "w:abstractNum",
        #[cfg(feature = "docx")]
        (b'w', b"abstractNumId") => "w:abstractNumId",
        #[cfg(feature = "docx")]
        (b'w', b"b") => "w:b",
        #[cfg(feature = "docx")]
        (b'w', b"bCs") => "w:bCs",
        #[cfg(feature = "docx")]
        (b'w', b"basedOn") => "w:basedOn",
        #[cfg(feature = "docx")]
        (b'w', b"body") => "w:body",
        #[cfg(feature = "docx")]
        (b'w', b"cs") => "w:cs",
        #[cfg(feature = "docx")]
        (b'w', b"del") => "w:del",
        #[cfg(feature = "docx")]
        (b'w', b"delText") => "w:delText",
        #[cfg(feature = "docx")]
        (b'w', b"docDefaults") => "w:docDefaults",
        #[cfg(feature = "docx")]
        (b'w', b"document") => "w:document",
        #[cfg(feature = "docx")]
        (b'w', b"drawing") => "w:drawing",
        #[cfg(feature = "docx")]
        (b'w', b"endnote") => "w:endnote",
        #[cfg(feature = "docx")]
        (b'w', b"endnoteReference") => "w:endnoteReference",
        #[cfg(feature = "docx")]
        (b'w', b"fldChar") => "w:fldChar",
        #[cfg(feature = "docx")]
        (b'w', b"fldSimple") => "w:fldSimple",
        #[cfg(feature = "docx")]
        (b'w', b"footerReference") => "w:footerReference",
        #[cfg(feature = "docx")]
        (b'w', b"footnote") => "w:footnote",
        #[cfg(feature = "docx")]
        (b'w', b"footnoteReference") => "w:footnoteReference",
        #[cfg(feature = "docx")]
        (b'w', b"ftr") => "w:ftr",
        #[cfg(feature = "docx")]
        (b'w', b"gridBefore") => "w:gridBefore",
        #[cfg(feature = "docx")]
        (b'w', b"gridCol") => "w:gridCol",
        #[cfg(feature = "docx")]
        (b'w', b"gridSpan") => "w:gridSpan",
        #[cfg(feature = "docx")]
        (b'w', b"hdr") => "w:hdr",
        #[cfg(feature = "docx")]
        (b'w', b"headerReference") => "w:headerReference",
        #[cfg(feature = "docx")]
        (b'w', b"i") => "w:i",
        #[cfg(feature = "docx")]
        (b'w', b"iCs") => "w:iCs",
        #[cfg(feature = "docx")]
        (b'w', b"ilvl") => "w:ilvl",
        #[cfg(feature = "docx")]
        (b'w', b"ins") => "w:ins",
        #[cfg(feature = "docx")]
        (b'w', b"instrText") => "w:instrText",
        #[cfg(feature = "docx")]
        (b'w', b"lvl") => "w:lvl",
        #[cfg(feature = "docx")]
        (b'w', b"lvlOverride") => "w:lvlOverride",
        #[cfg(feature = "docx")]
        (b'w', b"lvlText") => "w:lvlText",
        #[cfg(feature = "docx")]
        (b'w', b"moveFrom") => "w:moveFrom",
        #[cfg(feature = "docx")]
        (b'w', b"moveTo") => "w:moveTo",
        #[cfg(feature = "docx")]
        (b'w', b"noBreakHyphen") => "w:noBreakHyphen",
        #[cfg(feature = "docx")]
        (b'w', b"num") => "w:num",
        #[cfg(feature = "docx")]
        (b'w', b"numId") => "w:numId",
        #[cfg(feature = "docx")]
        (b'w', b"numPr") => "w:numPr",
        #[cfg(feature = "docx")]
        (b'w', b"numStyleLink") => "w:numStyleLink",
        #[cfg(feature = "docx")]
        (b'w', b"p") => "w:p",
        #[cfg(feature = "docx")]
        (b'w', b"pPr") => "w:pPr",
        #[cfg(feature = "docx")]
        (b'w', b"pStyle") => "w:pStyle",
        #[cfg(feature = "docx")]
        (b'w', b"pict") => "w:pict",
        #[cfg(feature = "docx")]
        (b'w', b"r") => "w:r",
        #[cfg(feature = "docx")]
        (b'w', b"rFonts") => "w:rFonts",
        #[cfg(feature = "docx")]
        (b'w', b"rPr") => "w:rPr",
        #[cfg(feature = "docx")]
        (b'w', b"rPrDefault") => "w:rPrDefault",
        #[cfg(feature = "docx")]
        (b'w', b"rStyle") => "w:rStyle",
        #[cfg(feature = "docx")]
        (b'w', b"rtl") => "w:rtl",
        #[cfg(feature = "docx")]
        (b'w', b"sdt") => "w:sdt",
        #[cfg(feature = "docx")]
        (b'w', b"sdtContent") => "w:sdtContent",
        #[cfg(feature = "docx")]
        (b'w', b"sectPr") => "w:sectPr",
        #[cfg(feature = "docx")]
        (b'w', b"softHyphen") => "w:softHyphen",
        #[cfg(feature = "docx")]
        (b'w', b"style") => "w:style",
        #[cfg(feature = "docx")]
        (b'w', b"styleLink") => "w:styleLink",
        #[cfg(feature = "docx")]
        (b'w', b"sym") => "w:sym",
        #[cfg(feature = "docx")]
        (b'w', b"t") => "w:t",
        #[cfg(feature = "docx")]
        (b'w', b"tab") => "w:tab",
        #[cfg(feature = "docx")]
        (b'w', b"tbl") => "w:tbl",
        #[cfg(feature = "docx")]
        (b'w', b"tblGrid") => "w:tblGrid",
        #[cfg(feature = "docx")]
        (b'w', b"tblLook") => "w:tblLook",
        #[cfg(feature = "docx")]
        (b'w', b"tblPr") => "w:tblPr",
        #[cfg(feature = "docx")]
        (b'w', b"tblStyle") => "w:tblStyle",
        #[cfg(feature = "docx")]
        (b'w', b"tblStyleColBandSize") => "w:tblStyleColBandSize",
        #[cfg(feature = "docx")]
        (b'w', b"tblStylePr") => "w:tblStylePr",
        #[cfg(feature = "docx")]
        (b'w', b"tblStyleRowBandSize") => "w:tblStyleRowBandSize",
        #[cfg(feature = "docx")]
        (b'w', b"tc") => "w:tc",
        #[cfg(feature = "docx")]
        (b'w', b"tcPr") => "w:tcPr",
        #[cfg(feature = "docx")]
        (b'w', b"themeFontLang") => "w:themeFontLang",
        #[cfg(feature = "docx")]
        (b'w', b"lang") => "w:lang",
        #[cfg(feature = "docx")]
        (b'w', b"tr") => "w:tr",
        #[cfg(feature = "docx")]
        (b'w', b"trPr") => "w:trPr",
        #[cfg(feature = "docx")]
        (b'w', b"txbxContent") => "w:txbxContent",
        #[cfg(feature = "docx")]
        (b'w', b"vMerge") => "w:vMerge",
        #[cfg(feature = "docx")]
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
        b"activeTab" => "activeTab",
        b"state" => "state",
        b"sheetId" => "sheetId",
        b"r" => "r",
        b"s" => "s",
        b"t" => "t",
        b"fontId" => "fontId",
        b"applyFont" => "applyFont",
        b"numFmtId" => "numFmtId",
        b"formatCode" => "formatCode",
        b"min" => "min",
        b"max" => "max",
        b"style" => "style",
        b"customFormat" => "customFormat",
        b"hidden" => "hidden",
        b"ht" => "ht",
        b"width" => "width",
        b"topLeftCell" => "topLeftCell",
        b"xSplit" => "xSplit",
        b"ySplit" => "ySplit",
        b"ref" => "ref",
        b"idx" => "idx",
        b"lvl" => "lvl",
        b"b" => "b",
        b"i" => "i",
        b"lang" => "lang",
        b"showMasterSp" => "showMasterSp",
        b"gridSpan" => "gridSpan",
        b"rowSpan" => "rowSpan",
        b"hMerge" => "hMerge",
        b"vMerge" => "vMerge",
        b"firstCol" => "firstCol",
        b"lastCol" => "lastCol",
        b"bandRow" => "bandRow",
        b"bandCol" => "bandCol",
        b"def" => "def",
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
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        | b"http://purl.oclc.org/ooxml/officeDocument/relationships" => b'r',
        b"http://schemas.openxmlformats.org/presentationml/2006/main"
        | b"http://purl.oclc.org/ooxml/presentationml/main" => b'p',
        b"http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        | b"http://purl.oclc.org/ooxml/spreadsheetml/main" => b's',
        b"http://schemas.openxmlformats.org/drawingml/2006/chart"
        | b"http://purl.oclc.org/ooxml/drawingml/chart" => b'c',
        b"http://schemas.openxmlformats.org/drawingml/2006/diagram"
        | b"http://purl.oclc.org/ooxml/drawingml/diagram" => b'd',
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
pub(crate) fn parse(bytes: &[u8], nodes: &mut usize) -> Result<Node, PreflightError> {
    std::str::from_utf8(bytes).map_err(|_| PreflightError::Xml)?;
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().expand_empty_elements = true;
    let mut stack = vec![(Node::default(), 0)];
    let mut bindings = BTreeMap::<Vec<u8>, u8>::new();
    let mut history = Vec::new();
    let mut incomplete = false;
    loop {
        match reader.read_event().map_err(|_| PreflightError::Xml)? {
            Event::Start(start) => {
                *nodes = nodes.checked_sub(1).ok_or(PreflightError::Limit)?;
                if stack.len() > 128 {
                    return Err(PreflightError::Limit);
                }
                let scope = history.len();
                let mut names = BTreeSet::new();
                for (index, attr) in start.attributes().with_checks(false).enumerate() {
                    if index == 128 {
                        return Err(PreflightError::Limit);
                    }
                    let attr = attr.map_err(|_| PreflightError::Xml)?;
                    if attr.key.as_ref().len() > 128 {
                        return Err(PreflightError::Limit);
                    }
                    if !names.insert(attr.key.0) {
                        return Err(PreflightError::Xml);
                    }
                    let prefix = if attr.key.as_ref() == b"xmlns" {
                        Some(&b""[..])
                    } else {
                        attr.key.as_ref().strip_prefix(b"xmlns:")
                    };
                    if let Some(prefix) = prefix {
                        if history.len() == MAX_NAMESPACES
                            || history.len() - scope == MAX_NAMESPACE_DECLARATIONS
                            || prefix.len() > 64
                            || attr.value.len() > 512
                        {
                            return Err(PreflightError::Limit);
                        }
                        let value = attr
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Implicit1_0,
                                reader.decoder(),
                            )
                            .map_err(|_| PreflightError::Xml)?;
                        let prefix = prefix.to_vec();
                        let previous = bindings.insert(prefix.clone(), namespace(value.as_bytes()));
                        history.push((prefix, previous));
                    }
                }
                let qualified = start.name();
                let qualified = qualified.as_ref();
                let (prefix, local) = match qualified.iter().position(|byte| *byte == b':') {
                    Some(at) => (&qualified[..at], &qualified[at + 1..]),
                    None => (&b""[..], qualified),
                };
                let ns = bindings.get(prefix).copied().unwrap_or(0);
                let mut node = Node {
                    name: name(ns, local),
                    ..Node::default()
                };
                if !node.name.is_empty() {
                    for attr in start.attributes().with_checks(false) {
                        let attr = attr.map_err(|_| PreflightError::Xml)?;
                        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:")
                        {
                            continue;
                        }
                        let Some(key) = attribute(attr.key.local_name().as_ref()) else {
                            continue;
                        };
                        if attr.value.len() > MAX_VALUE_BYTES * 8 {
                            incomplete = true;
                            continue;
                        }
                        let value = attr
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Implicit1_0,
                                reader.decoder(),
                            )
                            .map_err(|_| PreflightError::Xml)?;
                        if value.len() > MAX_VALUE_BYTES {
                            incomplete = true;
                        } else {
                            if key == "id" {
                                let qualified = attr.key.as_ref();
                                if let Some(at) = qualified.iter().position(|byte| *byte == b':')
                                    && bindings.get(&qualified[..at]) == Some(&b'r')
                                {
                                    node.attrs.push(("relId", value.to_string()));
                                }
                            }
                            node.attrs.push((key, value.into_owned()));
                        }
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
                while history.len() > scope {
                    let (prefix, previous) = history.pop().ok_or(PreflightError::Xml)?;
                    if let Some(previous) = previous {
                        bindings.insert(prefix, previous);
                    } else {
                        bindings.remove(&prefix);
                    }
                }
                let (parent, _) = stack.last_mut().ok_or(PreflightError::Xml)?;
                if node.name.is_empty() {
                    parent.children.extend(node.children);
                } else {
                    parent.children.push(node);
                }
            }
            Event::Text(text) => {
                let (node, _) = stack.last_mut().ok_or(PreflightError::Xml)?;
                if matches!(
                    node.name,
                    "w:t" | "w:delText" | "m:t" | "a:t" | "a:tableStyleId" | "s:t" | "s:v" | "s:f"
                ) {
                    let value = text.decode().map_err(|_| PreflightError::Xml)?;
                    node.text.push_str(&value);
                }
            }
            Event::CData(text) => {
                let (node, _) = stack.last_mut().ok_or(PreflightError::Xml)?;
                if matches!(
                    node.name,
                    "w:t" | "w:delText" | "m:t" | "a:t" | "a:tableStyleId" | "s:t" | "s:v" | "s:f"
                ) {
                    let value = text.decode().map_err(|_| PreflightError::Xml)?;
                    node.text.push_str(&value);
                }
            }
            Event::GeneralRef(reference) => {
                let (node, _) = stack.last_mut().ok_or(PreflightError::Xml)?;
                if matches!(
                    node.name,
                    "w:t" | "w:delText" | "m:t" | "a:t" | "a:tableStyleId" | "s:t" | "s:v" | "s:f"
                ) {
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
        .map(|mut root| {
            root.incomplete = incomplete;
            root
        })
        .ok_or(PreflightError::Xml)
}
