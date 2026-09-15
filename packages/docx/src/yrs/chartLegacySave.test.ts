import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { readDocxContainer } from '../docx/zipContainer';
import type { Chart, Document, Paragraph } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const NAMESPACES = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;
const CHART_DRAWING =
  '<w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing>';
const BAR_CHART =
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>';

/** A body paragraph followed by a paragraph holding one bar chart. */
function fixture(): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  set('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="${OFFICE_DOC}.drawingml.chart+xml"/></Types>`);
  set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart1" Type="${R}/chart" Target="charts/chart1.xml"/></Relationships>`);
  set('word/document.xml', `<w:document ${NAMESPACES}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:p><w:r>${CHART_DRAWING}</w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  set('word/charts/chart1.xml', BAR_CHART);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function legacyChart(withoutDrawing: Omit<Chart, 'drawingXml'>): Chart {
  return withoutDrawing as Chart;
}

function documentWithLegacyChart(): Document {
  const chart = legacyChart({
    type: 'chart',
    chartType: 'column',
    rId: 'rIdChart1',
    path: 'word/charts/chart1.xml',
    series: [],
  });
  return {
    originalBuffer: fixture().buffer as ArrayBuffer,
    package: {
      document: {
        content: [
          { type: 'paragraph', paraId: '00000001', content: [{ type: 'run', content: [{ type: 'text', text: 'Text' }] }] },
          { type: 'paragraph', paraId: '00000002', content: [{ type: 'run', content: [{ type: 'chart', chart }] }] },
        ],
      },
    },
  };
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))));

for (const seeder of ['native', 'projected']) {
  it(`${seeder} saves a session whose chart runs predate drawing replay`, async () => {
    const bytes = fixture();
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 74030 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      const chart = session.storySegments('body').find((segment) => segment.kind === 'embed' && segment.embedKind === 'chart');
      if (chart?.kind !== 'embed') throw new Error('expected a seeded chart embed');
      const stored = JSON.parse(chart.payload.chartJson as string) as Chart;
      expect(stored.drawingXml).toContain('<w:drawing>');
      delete stored.drawingXml;
      session.applyRawOps('body', [
        { op: 'setEmbedAttr', index: 5, key: 'chartJson', value: JSON.stringify(stored) },
      ]);
      const first = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
      const projected = yrsToDocument(session, parsed);
      const chartParagraph = projected.package.document.content[1] as Paragraph;
      const chartRun = chartParagraph.content.find((child) => child.type === 'run');
      if (chartRun?.type !== 'run') throw new Error('expected a saved chart run');
      const chartContent = chartRun.content.find((content) => content.type === 'chart');
      if (chartContent?.type !== 'chart') throw new Error('expected saved chart content');
      expect(chartContent.chart.drawingXml).toBe(CHART_DRAWING);
      expect(projected.warnings ?? []).toEqual([]);
      const saved = readDocxContainer(await repackDocx(projected));
      const documentXml = saved.text('word/document.xml') ?? '';
      expect(documentXml).toContain('Edited Text');
      expect(documentXml).toContain(CHART_DRAWING);
      expect(saved.text('word/charts/chart1.xml')).toBe(BAR_CHART);
    } finally {
      session.destroy();
    }
  });
}

it('drops an unrecoverable legacy chart with a warning instead of throwing', async () => {
  const document = documentWithLegacyChart();
  const session = await createYrsSession({ clientId: 74031 });
  try {
    documentToYrs(session, document);
    const saved = yrsToDocument(session, document);
    expect(saved.warnings?.join('\n')).toContain('chart');
    const chartParagraph = saved.package.document.content[1] as Paragraph;
    expect(JSON.stringify(chartParagraph)).not.toContain('"type":"chart"');
    await repackDocx(saved);
  } finally {
    session.destroy();
  }
});

function chartParagraphWith(chart: Chart): Paragraph {
  return {
    type: 'paragraph',
    paraId: '00000001',
    content: [{ type: 'run', content: [{ type: 'chart', chart }] }],
  };
}

function chartRunXml(saved: Paragraph): string | undefined {
  const run = saved.content.find((child) => child.type === 'run');
  if (run?.type !== 'run') return undefined;
  const content = run.content.find((entry) => entry.type === 'chart');
  if (content?.type !== 'chart') return undefined;
  return content.chart.drawingXml;
}

it('recovers duplicate relationship ids with the first placement', async () => {
  const first = CHART_DRAWING.replace('cx="5486400"', 'cx="111"');
  const last = CHART_DRAWING.replace('cx="5486400"', 'cx="999"');
  const chartWith = (drawingXml?: string): Chart => ({
    type: 'chart',
    chartType: 'column',
    rId: 'rIdDup',
    path: 'word/charts/chartDup.xml',
    series: [],
    ...(drawingXml === undefined ? {} : { drawingXml }),
  });
  const document: Document = {
    package: {
      document: {
        content: [
          chartParagraphWith(chartWith(first)),
          chartParagraphWith(chartWith(last)),
          chartParagraphWith(chartWith()),
        ],
      },
    },
  };
  const session = await createYrsSession({ clientId: 74033 });
  try {
    documentToYrs(session, document);
    const saved = yrsToDocument(session, document);
    expect(chartRunXml(saved.package.document.content[0] as Paragraph)).toContain('cx="111"');
    expect(chartRunXml(saved.package.document.content[1] as Paragraph)).toContain('cx="999"');
    expect(chartRunXml(saved.package.document.content[2] as Paragraph)).toContain('cx="111"');
    expect(saved.warnings ?? []).toEqual([]);
  } finally {
    session.destroy();
  }
});

it('drops an unrecoverable chart inside tracked changes instead of an empty wrapper', async () => {
  const chart: Chart = { type: 'chart', chartType: 'column', rId: 'rIdNope', series: [] };
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: 'paragraph',
            paraId: '00000001',
            content: [
              {
                type: 'insertion',
                info: { id: 1, author: 'Ada' },
                content: [{ type: 'run', content: [{ type: 'chart', chart }] }],
              },
            ],
          },
        ],
      },
    },
  };
  const session = await createYrsSession({ clientId: 74034 });
  try {
    documentToYrs(session, document);
    const saved = yrsToDocument(session, document);
    const paragraph = saved.package.document.content[0];
    expect(paragraph?.type).toBe('paragraph');
    if (paragraph?.type === 'paragraph') expect(paragraph.content).toEqual([]);
    expect(saved.warnings?.join('\n')).toContain('rIdNope');
  } finally {
    session.destroy();
  }
});

it('warns on every dropped chart run with a malformed payload', async () => {
  const document: Document = {
    package: {
      document: {
        content: [
          { type: 'paragraph', paraId: '00000001', content: [{ type: 'run', content: [{ type: 'text', text: 'AB' }] }] },
        ],
      },
    },
  };
  const session = await createYrsSession({ clientId: 74035 });
  try {
    documentToYrs(session, document);
    session.applyRawOps('body', [
      { op: 'insertEmbed', index: 1, kind: 'chart', payload: {} },
      { op: 'insertEmbed', index: 2, kind: 'chart', payload: { chartJson: 'not json' } },
      { op: 'insertEmbed', index: 3, kind: 'chart', payload: { chartJson: JSON.stringify({ type: 'chart' }) } },
    ]);
    const saved = yrsToDocument(session, document);
    const paragraph = saved.package.document.content[0] as Paragraph;
    expect(JSON.stringify(paragraph)).not.toContain('"type":"chart"');
    expect(saved.warnings).toHaveLength(3);
    expect(saved.warnings?.join('\n')).toContain('chart');
  } finally {
    session.destroy();
  }
});

it('recovers legacy chart placements per story when relationship ids collide', async () => {
  const bodyDrawing = CHART_DRAWING.replace('Chart 1', 'Body chart');
  const headerDrawing = CHART_DRAWING.replace('Chart 1', 'Header chart');
  const document: Document = {
    originalBuffer: fixture().buffer as ArrayBuffer,
    package: {
      document: {
        content: [
          chartParagraphWith({
            type: 'chart',
            chartType: 'column',
            rId: 'rId5',
            path: 'word/charts/chart1.xml',
            series: [],
            drawingXml: bodyDrawing,
          }),
        ],
      },
      headers: new Map([
        [
          'rId5',
          {
            type: 'header',
            hdrFtrType: 'default',
            content: [
              chartParagraphWith({
                type: 'chart',
                chartType: 'column',
                rId: 'rId5',
                path: 'word/charts/chart2.xml',
                series: [],
                drawingXml: headerDrawing,
              }),
            ],
          },
        ],
      ]),
    },
  };
  const session = await createYrsSession({ clientId: 74032 });
  try {
    documentToYrs(session, document);
    for (const storyId of ['body', 'hf:rId5']) {
      let index = 0;
      for (const segment of session.storySegments(storyId)) {
        if (segment.kind === 'embed' && segment.embedKind === 'chart') {
          const stored = JSON.parse(segment.payload.chartJson as string) as Chart;
          delete stored.drawingXml;
          session.applyRawOps(storyId, [
            { op: 'setEmbedAttr', index, key: 'chartJson', value: JSON.stringify(stored) },
          ]);
        }
        index += segment.kind === 'text' ? segment.text.length : 1;
      }
    }
    const saved = yrsToDocument(session, document);
    expect(chartRunXml(saved.package.document.content[0] as Paragraph)).toBe(bodyDrawing);
    const header = saved.package.headers?.get('rId5');
    if (!header) throw new Error('expected the header part to survive the save');
    expect(chartRunXml(header.content[0] as Paragraph)).toBe(headerDrawing);
    expect(saved.warnings ?? []).toEqual([]);
  } finally {
    session.destroy();
  }
});
