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
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:bofx="urn:fidelity"`;
const paragraph = '<w:p><w:r><w:t>Text</w:t></w:r></w:p>';
const raw = '<bofx:block bofx:value="opaque"><bofx:child>hidden</bofx:child></bofx:block>';

function fixture(body: string, stories = false, documentNamespaces = NS): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  const storyNames = stories ? ['header', 'footer', 'footnotes', 'endnotes'] : [];
  const target = (name: string) => `${name}${name === 'header' || name === 'footer' ? '1' : ''}.xml`;
  const relationships = ['styles', 'numbering', ...storyNames];
  set('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${relationships.map((name) => `<Override PartName="/word/${target(name)}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name}+xml"/>`).join('')}</Types>`);
  set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.map((name) => `<Relationship Id="${name}" Type="${R}/${name}" Target="${target(name)}"/>`).join('')}</Relationships>`);
  set('word/document.xml', `<w:document ${documentNamespaces}><w:body>${body}<w:sectPr>${stories ? '<w:headerReference w:type="default" r:id="header"/><w:footerReference w:type="default" r:id="footer"/>' : ''}</w:sectPr></w:body></w:document>`);
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
  for (const history of ['', '0', '1']) {
    it(`${seeder} full save preserves hyperlink history ${history || 'absent'}`, async () => {
      const historyAttribute = history ? ` w:history="${history}"` : '';
      const bytes = fixture(`<w:p><w:hyperlink w:anchor="target"${historyAttribute}><w:r><w:t>Link</w:t></w:r></w:hyperlink><w:r><w:t> tail</w:t></w:r></w:p>`);
      const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
      const session = await createYrsSession({ clientId: 74011 });
      const peer = await createYrsSession({ clientId: 74013 });
      try {
        if (seeder === 'native') session.seedFromDocx(bytes);
        else documentToYrs(session, parsed);
        const first = session.paragraphs('body')[0]!;
        session.insertText({ story: 'body', paraId: first.paraId, offset: 8 }, '!');
        peer.loadState(session.encodeState());
        const saved = await repackDocx(yrsToDocument(peer, parsed));
        const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
        if (history) expect(xml).toContain(`w:history="${history}"`);
        else expect(xml).not.toContain('w:history=');
        const reopened = await parseDocx(saved, { preloadFonts: false });
        const para = reopened.package.document.content[0]!;
        if (para.type !== 'paragraph') throw new Error('missing paragraph');
        const link = para.content.find(item => item.type === 'hyperlink');
        if (history) expect(link).toMatchObject({ history: history === '1' });
        else expect(link).not.toHaveProperty('history');
      } finally { peer.destroy(); session.destroy(); }
    });
  }

  it(`${seeder} full save preserves adjacent hyperlink histories`, async () => {
    const link = (history: string) => `<w:hyperlink w:anchor="target" w:history="${history}"><w:r><w:t>Link</w:t></w:r></w:hyperlink>`;
    const bytes = fixture(`<w:p>${link('0')}${link('1')}<w:r><w:t> tail</w:t></w:r></w:p>`);
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 74015 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      const first = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId: first.paraId, offset: 12 }, '!');
      const saved = await repackDocx(yrsToDocument(session, parsed));
      const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
      expect(xml.match(/w:history="[01]"/g)).toEqual(['w:history="0"', 'w:history="1"']);
    } finally { session.destroy(); }
  });

  it(`${seeder} full save preserves row start attributes`, async () => {
    const attrs = 'w:rsidR="00112233" w:rsidRPr="44556677" w:rsidTr="8899AABB" w:rsidDel="CCDDEEFF" bofx:paraId="row&amp;1"';
    const bytes = fixture(`${paragraph}<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr ${attrs}><w:tc>${paragraph}</w:tc></w:tr><w:tr><w:tc>${paragraph}</w:tc></w:tr></w:tbl>`);
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 74012 });
    const peer = await createYrsSession({ clientId: 74014 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      const first = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
      peer.loadState(session.encodeState());
      const saved = await repackDocx(yrsToDocument(peer, parsed));
      const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
      expect(xml.match(/<w:tr\b[^>]*>/g)).toEqual([`<w:tr ${attrs}>`, '<w:tr>']);
      const reopened = await parseDocx(saved, { preloadFonts: false });
      const resaved = await repackDocx(reopened);
      const reopenedXml = await (await JSZip.loadAsync(resaved)).file('word/document.xml')!.async('string');
      expect(reopenedXml.match(/<w:tr\b[^>]*>/g)).toEqual([`<w:tr ${attrs}>`, '<w:tr>']);
      session.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'below');
      const inserted = await repackDocx(yrsToDocument(session, parsed));
      const insertedXml = await (await JSZip.loadAsync(inserted)).file('word/document.xml')!.async('string');
      expect(insertedXml.match(/<w:tr\b[^>]*>/g)).toEqual([`<w:tr ${attrs}>`, '<w:tr>', '<w:tr>']);
      session.splitCell({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 2, 1);
      const split = await repackDocx(yrsToDocument(session, parsed));
      const splitXml = await (await JSZip.loadAsync(split)).file('word/document.xml')!.async('string');
      expect(splitXml.match(/<w:tr\b[^>]*>/g)).toEqual([`<w:tr ${attrs}>`, '<w:tr>', '<w:tr>', '<w:tr>']);
    } finally { peer.destroy(); session.destroy(); }
  });

  it(`${seeder} full save preserves table-scoped row attribute namespaces`, async () => {
    const bytes = fixture(`${paragraph}<w:tbl xmlns:custom="urn:row"><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr custom:keep="row&amp;1"><w:tc>${paragraph}</w:tc></w:tr></w:tbl>`);
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 74016 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      const first = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
      const saved = await repackDocx(yrsToDocument(session, parsed));
      const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
      expect(xml).toContain('<w:tr custom:keep="row&amp;1" xmlns:custom="urn:row">');
      const checked = spawnSync('python3', ['-c', 'import sys,xml.etree.ElementTree as E; E.fromstring(sys.stdin.read())'], { input: xml, encoding: 'utf8' });
      expect(checked.stderr).toBe('');
      expect(checked.status).toBe(0);
      const reopened = await parseDocx(saved, { preloadFonts: false });
      expect(reopened.package.document.content[1]).toMatchObject({ rows: [{ extraAttributes: [
        { name: 'custom:keep', value: 'row&1' },
        { name: 'xmlns:custom', value: 'urn:row' },
      ] }] });
    } finally { session.destroy(); }
  });

  it(`${seeder} full save keeps a row-rebound w14 prefix off cell paragraph ids`, async () => {
    const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
    const bytes = fixture(`${paragraph}<w:tbl xmlns:w14="urn:custom"><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr w14:keep="row"><w:tc><w:p xmlns:x14="${w14}" x14:paraId="1A2B3C4D"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`);
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 74017 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      const first = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
      const saved = await repackDocx(yrsToDocument(session, parsed));
      const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
      const checked = spawnSync('python3', ['-c', `import sys,json,xml.etree.ElementTree as E
root=E.fromstring(sys.stdin.read())
tr=root.find('.//{${W}}tr')
print(json.dumps({'row':tr.attrib,'paragraph':tr.find('.//{${W}}p').attrib}))`], { input: xml, encoding: 'utf8' });
      expect(checked.stderr).toBe('');
      expect(checked.status).toBe(0);
      const resolved = JSON.parse(checked.stdout) as { row: Record<string, string>; paragraph: Record<string, string> };
      expect(resolved.row).toEqual({ '{urn:custom}keep': 'row' });
      expect(resolved.paragraph[`{${w14}}paraId`]).toBe('1A2B3C4D');
      expect(resolved.paragraph['{urn:custom}paraId']).toBeUndefined();
    } finally { session.destroy(); }
  });

  it(`${seeder} full save picks row aliases no ancestor or nested binding uses, stable across saves`, async () => {
    const grid = '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>';
    const ancestor = `<w:tbl xmlns:w14="urn:custom">${grid}<w:tr w14:keep="row"><w:tc><w:p w14r0:tag="cell"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const nested = `<w:tbl xmlns:w14="urn:inner">${grid}<w:tr w14:keep="inner"><w:tc><w:p w14r0:from="outer"><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const outer = `<w:tbl xmlns:w14r0="urn:outer">${grid}<w:tr w14r0:keep="outer"><w:tc>${nested}<w:p/></w:tc></w:tr></w:tbl>`;
    const bytes = fixture(`${paragraph}${ancestor}${paragraph}${outer}`, false, `${NS} xmlns:w14r0="urn:root"`);
    const save = async (source: Uint8Array<ArrayBuffer>, clientId: number) => {
      const parsed = await parseDocx(source.buffer, { preloadFonts: false });
      const session = await createYrsSession({ clientId });
      try {
        if (seeder === 'native') session.seedFromDocx(source);
        else documentToYrs(session, parsed);
        const first = session.paragraphs('body')[0]!;
        session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
        const saved = new Uint8Array(await repackDocx(yrsToDocument(session, parsed)));
        return { saved, xml: await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string') };
      } finally { session.destroy(); }
    };
    const first = await save(bytes, 74018);
    const checked = spawnSync('python3', ['-c', `import sys,json,xml.etree.ElementTree as E
root=E.fromstring(sys.stdin.read())
rows={k:v for tr in root.iter('{${W}}tr') for k,v in tr.attrib.items()}
paragraphs={k:v for p in root.iter('{${W}}p') for k,v in p.attrib.items() if k.startswith('{urn:')}
print(json.dumps({'rows':rows,'paragraphs':paragraphs}))`], { input: first.xml, encoding: 'utf8' });
    expect(checked.stderr).toBe('');
    expect(checked.status).toBe(0);
    expect(JSON.parse(checked.stdout)).toEqual({
      rows: { '{urn:custom}keep': 'row', '{urn:outer}keep': 'outer', '{urn:inner}keep': 'inner' },
      paragraphs: { '{urn:root}tag': 'cell', '{urn:outer}from': 'outer' },
    });
    const rowTags = (xml: string) => xml.match(/<w:tr\b[^>]*>/g);
    const second = await save(first.saved, 74019);
    expect(rowTags(second.xml)).toEqual(rowTags(first.xml));
  });
}

for (const style of ['none', 'nil']) {
  it(`full save preserves ${style} border attributes in paragraphs, tables and cells`, async () => {
    const border = `<w:top w:val="${style}" w:sz="8" w:space="2" w:color="AABBCC"/>`;
    const bytes = fixture(`<w:p><w:pPr><w:pBdr>${border}</w:pBdr></w:pPr><w:r><w:t>Text</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblBorders>${border}</w:tblBorders></w:tblPr><w:tr><w:tc><w:tcPr><w:tcBorders>${border}</w:tcBorders></w:tcPr>${paragraph}</w:tc></w:tr></w:tbl>`);
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const first = parsed.package.document.content[0]!;
    if (first.type !== 'paragraph') throw new Error('missing paragraph');
    first.content.push({ type: 'run', content: [{ type: 'text', text: ' edited' }] });
    const saved = await repackDocx(parsed);
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
    expect(xml.split(border)).toHaveLength(4);
    const reopened = await parseDocx(saved, { preloadFonts: false });
    const resaved = await repackDocx(reopened);
    const reopenedXml = await (await JSZip.loadAsync(resaved)).file('word/document.xml')!.async('string');
    expect(reopenedXml.split(border)).toHaveLength(4);
  });
}

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
