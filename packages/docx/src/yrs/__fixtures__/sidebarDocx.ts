import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../../docx/rezip/parts';

export function sidebarDocx(): Uint8Array {
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>')],
    ['word/comments.xml', toBytes('<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="7" w:author="Reader" w:date="2026-10-01T00:00:00Z"><w:p><w:r><w:t>Check this paragraph.</w:t></w:r></w:p></w:comment></w:comments>')],
    ['word/document.xml', toBytes('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>' +
      '<w:p w14:paraId="00000001"><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t xml:space="preserve"> First heading </w:t></w:r></w:p>' +
      '<w:p w14:paraId="00000002"><w:commentRangeStart w:id="7"/><w:r><w:t>Commented text</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>' +
      '<w:p w14:paraId="00000003"><w:ins w:id="11" w:author="Writer" w:date="2026-10-01T00:00:00Z"><w:r><w:t>Inserted text</w:t></w:r></w:ins><w:del w:id="12" w:author="Other writer" w:date="2026-10-02T00:00:00Z"><w:r><w:delText>Deleted text</w:delText></w:r></w:del></w:p>' +
      '<w:p w14:paraId="00000004"><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>Second heading</w:t></w:r></w:p>' +
      '<w:p w14:paraId="00000005"><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t xml:space="preserve">   </w:t></w:r></w:p>' +
      '<w:sectPr/></w:body></w:document>')],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}
