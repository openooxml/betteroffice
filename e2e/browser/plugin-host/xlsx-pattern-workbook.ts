import type JSZip from 'jszip';

export function addPatternSheet(zip: JSZip): void {
  const styles = ['mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'];
  zip.file('xl/styles.xml', `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts><font/></fonts><fills><fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="lightUp"><fgColor rgb="FF123456"/><bgColor rgb="FFFEDCBA"/></patternFill></fill>
    <fill><patternFill patternType="gray125"><fgColor rgb="FF123456"/><bgColor rgb="FFFEDCBA"/></patternFill></fill></fills>
    <borders><border/>${styles.map((style) => `<border><top style="${style}"><color rgb="FF123456"/></top></border>`).join('')}</borders>
    <cellXfs><xf fontId="0" fillId="0" borderId="0"/><xf fillId="1"/><xf fillId="2"/>${styles.map((_, i) => `<xf borderId="${i + 1}"/>`).join('')}</cellXfs></styleSheet>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
    <sheetFormatPr defaultRowHeight="24"/><cols><col min="1" max="4" width="24"/></cols>
    <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>initial</t></is></c></row>
    <row r="2"><c r="B2" s="1"/><c r="C2" s="1"/><c r="D2" s="2"/></row>
    ${styles.map((_, i) => `<row r="${i + 4}"><c r="B${i + 4}" s="${i + 3}"/></row>`).join('')}
    </sheetData></worksheet>`);
}
