import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:bofx="urn:fidelity"`;
const paragraph = '<w:p><w:r><w:t>Text</w:t></w:r></w:p>';
const raw = '<bofx:block bofx:value="opaque"><bofx:child>hidden</bofx:child></bofx:block>';

function fixture(body: string, stories = false): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  const storyNames = stories ? ['header', 'footer', 'footnotes', 'endnotes'] : [];
  const target = (name: string) => `${name}${name === 'header' || name === 'footer' ? '1' : ''}.xml`;
  const relationships = ['styles', 'numbering', ...storyNames];
  set('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${relationships.map((name) => `<Override PartName="/word/${target(name)}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name}+xml"/>`).join('')}</Types>`);
  set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.map((name) => `<Relationship Id="${name}" Type="${R}/${name}" Target="${target(name)}"/>`).join('')}</Relationships>`);
  set('word/document.xml', `<w:document ${NS}><w:body>${body}<w:sectPr>${stories ? '<w:headerReference w:type="default" r:id="header"/><w:footerReference w:type="default" r:id="footer"/>' : ''}</w:sectPr></w:body></w:document>`);
  set('word/styles.xml', `<w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="List"><w:name w:val="List"/><w:pPr><w:ind w:left="720" w:firstLine="180"/></w:pPr></w:style></w:styles>`);
  set('word/numbering.xml', `<w:numbering ${NS}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  if (stories) {
    set('word/header1.xml', `<w:hdr ${NS}>${raw}${paragraph}</w:hdr>`);
    set('word/footer1.xml', `<w:ftr ${NS}>${paragraph}${raw}</w:ftr>`);
    set('word/footnotes.xml', `<w:footnotes ${NS}><w:footnote w:id="1">${raw}${paragraph}</w:footnote></w:footnotes>`);
    set('word/endnotes.xml', `<w:endnotes ${NS}><w:endnote w:id="1">${paragraph}${raw}</w:endnote></w:endnotes>`);
  }
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function strictRawParts(bytes: ArrayBuffer): unknown {
  const result = spawnSync('python3', ['-c', `import sys,io,zipfile,xml.etree.ElementTree as E,json
z=zipfile.ZipFile(io.BytesIO(sys.stdin.buffer.read()))
result={}
for name in ['document.xml','header1.xml','footer1.xml','footnotes.xml','endnotes.xml']:
 root=E.fromstring(z.read('word/'+name))
 def raw_paths(node,path=''):
  found=[]
  if node.tag=='{urn:fidelity}block': found.append((path,E.tostring(node,encoding='unicode')))
  children=[child for child in node if child.tag.rsplit('}',1)[-1] not in ['pPr','tcPr','trPr','tblPr','tblGrid','sdtPr']]
  for i,child in enumerate(children): found.extend(raw_paths(child,path+'/'+child.tag+':'+str(i)))
  return found
 result[name]=raw_paths(root)
 assert result[name],name
print(json.dumps(result,sort_keys=True))`], { input: Buffer.from(bytes), encoding: 'utf8' });
  expect(result.stderr).toBe('');
  expect(result.status).toBe(0);
  return JSON.parse(result.stdout);
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))));

for (const seeder of ['native', 'projected']) {
  it(`${seeder} preserves opaque blocks in body, header, footer, cells, controls and notes`, async () => {
    const table = `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${raw}${paragraph}</w:tc></w:tr></w:tbl>`;
    const sdt = `<w:sdt><w:sdtPr><w:tag w:val="test"/></w:sdtPr><w:sdtContent>${raw}${paragraph}</w:sdtContent></w:sdt>`;
    const bytes = fixture(`${paragraph}${raw}${paragraph}${table}${sdt}`, true);
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const direct = strictRawParts(await repackDocx(parsed));
    const session = await createYrsSession({ clientId: 74001 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      expect(session.storyIds().filter((id) => id.includes(':sdt'))).toEqual(['body:sdt0']);
      const saved = yrsToDocument(session, parsed);
      expect(strictRawParts(await repackDocx(saved))).toEqual(direct);
      const first = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
      expect(strictRawParts(await repackDocx(yrsToDocument(session, parsed)))).toEqual(direct);
    } finally {
      session.destroy();
    }
  });

  it(`${seeder} preserves opaque inline nodes and their surrounding runs`, async () => {
    const bytes = fixture(`<w:p><w:r><w:t>A</w:t></w:r>${raw}<w:r><w:t>B</w:t></w:r></w:p>${paragraph}`, true);
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const direct = strictRawParts(await repackDocx(parsed));
    const session = await createYrsSession({ clientId: 74005 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      expect(session.paragraphs('body')[0]!.properties._originalRunBoundaries).toHaveLength(2);
      expect(strictRawParts(await repackDocx(yrsToDocument(session, parsed)))).toEqual(direct);
      const last = session.paragraphs('body').at(-1)!;
      session.insertText({ story: 'body', paraId: last.paraId, offset: 0 }, 'Edited ');
      expect(strictRawParts(await repackDocx(yrsToDocument(session, parsed)))).toEqual(direct);
    } finally {
      session.destroy();
    }
  });

  for (const direct of ['', '<w:ind w:firstLine="0"/>', '<w:ind w:hanging="0"/>']) {
    it(`${seeder} applies numbering before style with direct indent ${direct || 'absent'}`, async () => {
      const bytes = fixture(`<w:p><w:pPr><w:pStyle w:val="List"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>${direct}</w:pPr><w:r><w:t>List</w:t></w:r></w:p>`);
      const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
      const session = await createYrsSession({ clientId: 74002 });
      try {
        if (seeder === 'native') session.seedFromDocx(bytes);
        else documentToYrs(session, parsed);
        const properties = session.paragraphs('body')[0]!.properties;
        expect(properties.indentLeft).toBe(1440);
        expect(properties.indentFirstLine).toBe(-360);
        expect(properties.hangingIndent).toBe(true);
      } finally {
        session.destroy();
      }
    });
  }
}

for (const seeder of ['native', 'projected']) {
  it(`${seeder} renders and saves hyperlinks inside complex fields`, async () => {
    const body = '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="_Toc1" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Heading</w:t></w:r></w:hyperlink><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    const bytes = fixture(body + body.replaceAll('Heading', 'Second').replaceAll('_Toc1', '_Toc2'));
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 74004 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      const text = session.storySegments('body').find((segment) => segment.kind === 'text');
      expect(text?.attributes.hyperlink).toMatchObject({ href: '#_Toc1' });
      expect(text?.attributes.textColor).toMatchObject({ rgb: '0563C1' });
      expect(session.storySegments('body').some((segment) => segment.kind === 'embed' && segment.embedKind === 'field' && segment.payload.fieldType === 'PAGE')).toBe(true);
      const saved = yrsToDocument(session, parsed);
      const paragraph = saved.package.document.content[0]!;
      expect(paragraph.type).toBe('paragraph');
      if (paragraph.type !== 'paragraph') throw new Error('missing paragraph');
      const original = parsed.package.document.content[0]!;
      if (original.type !== 'paragraph') throw new Error('missing original paragraph');
      expect(paragraph.content).toEqual(original.content);
      const first = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId: first.paraId, offset: 1 }, '!');
      const edited = yrsToDocument(session, parsed).package.document.content[0]!;
      if (edited.type !== 'paragraph') throw new Error('missing edited paragraph');
      expect(edited.content).toHaveLength(1);
      const field = edited.content[0]!;
      if (field.type !== 'complexField') throw new Error('missing saved field');
      const link = field.structuredResult?.inline?.[0];
      if (link?.type !== 'hyperlink') throw new Error('missing saved hyperlink');
      expect(link.anchor).toBe('_Toc1');
      expect(JSON.stringify(link)).toContain('H!eading');
      session.mergeParagraphs('body', first.paraId);
      const merged = yrsToDocument(session, parsed).package.document.content[0]!;
      if (merged.type !== 'paragraph') throw new Error('missing merged paragraph');
      expect(merged.content).toHaveLength(2);
      const links = merged.content.map((content) => content.type === 'complexField' ? content.structuredResult?.inline?.[0] : undefined);
      expect(links[0]).toEqual(link);
      expect(JSON.stringify(links[1])).toContain('Second');
      expect(JSON.stringify(links[1])).not.toContain('H!eading');
      await repackDocx(yrsToDocument(session, parsed));
    } finally {
      session.destroy();
    }
  });
}
