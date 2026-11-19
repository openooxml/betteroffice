import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';

import {
  findDocxContentControls,
  listDocxContentControls
} from '../../packages/docx/src/docx/contentControls';
import { parseDocx } from '../../packages/docx/src/docx/parser';
import {
  exportDocxMarkdown,
  exportDocxStructured
} from '../../packages/docx/src/docx/structuredExport';
import type {
  BlockContent,
  HeaderFooter,
  Paragraph,
  Run,
  Section,
  Table
} from '../../packages/docx/src/types/content';
import type { Document } from '../../packages/docx/src/types/document';
import { preloadEditWasm } from '../../packages/docx/src/wasm/edit';
import { engineKind } from '../../packages/docx/src/wasm/engineEdit';
import { preloadOpcWasm, unzipContainer } from '../../packages/docx/src/wasm/opc';
import { preloadParseWasm } from '../../packages/docx/src/wasm/parse';
import {
  createYrsSession,
  saveYrsDocx,
  type DocxContentControlQuery,
  type DocxContentControlsOptions,
  type DocxExportOptions,
  type YrsReplaceReceipt,
  type YrsSession
} from '../../packages/docx/src/yrs';
import { restoreBrowser, stubBrowser } from '../../packages/docx/src/yrs/renderPngTestSupport';
import type { Format } from './harness';

const EDIT_CLIENT_ID = 1;

function wasmBytes(path: string): Uint8Array {
  return new Uint8Array(
    readFileSync(new URL(`../../packages/docx/src/wasm/generated/${path}`, import.meta.url))
  );
}

function runText(run: Run, separators: boolean): string {
  let text = '';
  for (const content of run.content) {
    if (content.type === 'text') text += content.text;
    else if (separators && content.type === 'tab') text += '\t';
    else if (separators && content.type === 'break') {
      text += content.breakType === 'page' ? '\u000c' : '\n';
    }
  }
  return text;
}

function paragraphRuns(paragraph: Paragraph): [Run, boolean][] {
  const runs: [Run, boolean][] = [];
  for (const node of paragraph.content) {
    if (node.type === 'run') runs.push([node, true]);
    else if (node.type === 'hyperlink') {
      for (const child of node.children) if (child.type === 'run') runs.push([child, false]);
    } else if (node.type === 'simpleField') {
      for (const child of node.content) if (child.type === 'run') runs.push([child, false]);
    } else if (node.type === 'complexField') {
      for (const child of node.fieldResult) runs.push([child, false]);
    }
  }
  return runs;
}

function runResult(run: Run, separators: boolean) {
  const formatting = run.formatting;
  return {
    text: runText(run, separators),
    bold: formatting?.bold ?? null,
    italic: formatting?.italic ?? null,
    underline: formatting?.underline?.style ?? null,
    strike: formatting?.strike ?? null,
    fontSize: formatting?.fontSize === undefined ? null : formatting.fontSize / 2,
    fontFamily: formatting?.fontFamily?.ascii ?? formatting?.fontFamily?.hAnsi ?? null,
    color: formatting?.color?.rgb ?? null,
    styleId: formatting?.styleId ?? null
  };
}

function paragraphResult(paragraph: Paragraph) {
  const runs = paragraphRuns(paragraph).map(([run, separators]) => runResult(run, separators));
  return {
    id: paragraph.paraId ?? null,
    style: paragraph.formatting?.styleId ?? null,
    alignment: paragraph.formatting?.alignment ?? null,
    runs,
    text: runs.map(({ text }) => text).join('')
  };
}

function collectParagraphs(blocks: BlockContent[], output: Paragraph[] = []): Paragraph[] {
  for (const block of blocks) {
    if (block.type === 'paragraph') output.push(block);
    else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) collectParagraphs(cell.content, output);
      }
    } else if (block.type === 'blockSdt') collectParagraphs(block.content, output);
  }
  return output;
}

function storyParagraphs(blocks: BlockContent[], output: Paragraph[] = []): Paragraph[] {
  for (const block of blocks) {
    if (block.type === 'paragraph') output.push(block);
    else if (block.type === 'blockSdt') storyParagraphs(block.content, output);
  }
  return output;
}

function collectTables(blocks: BlockContent[], output: Table[] = []): Table[] {
  for (const block of blocks) {
    if (block.type === 'table') {
      output.push(block);
      for (const row of block.rows) {
        for (const cell of row.cells) collectTables(cell.content, output);
      }
    } else if (block.type === 'blockSdt') collectTables(block.content, output);
  }
  return output;
}

function joinText(parts: string[]): string {
  return parts.filter((part) => part !== '').join('\n');
}

function paragraphsOf(blocks: BlockContent[]) {
  return collectParagraphs(blocks).map(paragraphResult);
}

function tableResult(table: Table): unknown {
  const rows = table.rows.map((row) => {
    const cells = row.cells.map((cell) => {
      const paragraphs = paragraphsOf(cell.content);
      return {
        paragraphs,
        tables: collectTables(cell.content).map(tableResult),
        text: joinText(paragraphs.map(({ text }) => text))
      };
    });
    return { cells, text: cells.map(({ text }) => text).join('\t') };
  });
  return {
    rows,
    columnWidths: table.columnWidths ?? null,
    text: joinText(rows.map(({ text }) => text))
  };
}

function runHighlights(paragraph: Paragraph) {
  return {
    id: paragraph.paraId ?? null,
    highlights: paragraphRuns(paragraph).map(([run]) => run.formatting?.highlight ?? null)
  };
}

function cellMerges(table: Table) {
  return {
    rows: table.rows.map((row) =>
      row.cells.map((cell) => ({
        gridSpan: cell.formatting?.gridSpan ?? null,
        vMerge: cell.formatting?.vMerge ?? null
      }))
    )
  };
}

function sectionResult(section: Section) {
  const { properties } = section;
  const paragraphs = paragraphsOf(section.content);
  return {
    id: section.id ?? properties.sectionId ?? null,
    pageWidth: properties.pageWidth ?? null,
    pageHeight: properties.pageHeight ?? null,
    orientation: properties.orientation ?? null,
    marginTop: properties.marginTop ?? null,
    marginBottom: properties.marginBottom ?? null,
    marginLeft: properties.marginLeft ?? null,
    marginRight: properties.marginRight ?? null,
    paragraphs,
    text: joinText(paragraphs.map(({ text }) => text))
  };
}

function headerFooterResults(entries: Map<string, HeaderFooter> | undefined) {
  return [...(entries ?? new Map<string, HeaderFooter>())].map(([relId, story]) => {
    const paragraphs = paragraphsOf(story.content);
    return {
      relId,
      kind: story.hdrFtrType,
      paragraphs,
      text: joinText(paragraphs.map(({ text }) => text))
    };
  });
}

function receiptResult(receipt: YrsReplaceReceipt) {
  return {
    story: receipt.range?.story ?? null,
    start: receipt.range?.start.offset ?? null,
    end: receipt.range?.end.offset ?? null,
    revisionIds: receipt.revisionId === null ? [] : [receipt.revisionId]
  };
}

function pngSize(bytes: Uint8Array): [number, number] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return [view.getUint32(16), view.getUint32(20)];
}

function partsResult(bytes: Uint8Array) {
  const parts = Object.entries(unzipContainer(bytes))
    .sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0)
    .map(([name, data]) => ({
      name,
      sha256: createHash('sha256').update(data).digest('hex'),
      byteLength: data.byteLength
    }));
  return { parts };
}

export const docx: Format = {
  engineKind,
  versionPattern: /(?!)/,
  saveRecord: partsResult,
  hashesScript: true,
  async open(fixture) {
    await preloadEditWasm(wasmBytes('edit/docx_edit_bg.wasm'));
    await preloadOpcWasm(wasmBytes('opc/ooxml_opc_bg.wasm'));
    await preloadParseWasm(wasmBytes('parse/docx_parse_bg.wasm'));
    const parse = async (bytes: Uint8Array) =>
      parseDocx(Uint8Array.from(bytes).buffer, { preloadFonts: false });
    let bytes = Uint8Array.from(fixture);
    let document: Document = await parse(bytes);
    let session: YrsSession | undefined;
    let saved: Uint8Array | undefined;
    let editedModel: Document | undefined;
    const close = () => {
      session?.destroy();
      session = undefined;
    };
    const reopen = async (next: Uint8Array) => {
      close();
      bytes = Uint8Array.from(next);
      document = await parse(bytes);
      saved = undefined;
      editedModel = undefined;
      return null;
    };
    const read = () => {
      if (session !== undefined) throw new Error('A read after an edit needs openSaved first');
      return document;
    };
    const edited = async () => {
      if (session === undefined) {
        session = await createYrsSession({ clientId: EDIT_CLIENT_ID });
        session.seedFromDocx(Uint8Array.from(bytes), {
          generation: createHash('sha256').update(bytes).digest('hex').slice(0, 32)
        });
      }
      return session;
    };
    const bodyParagraphs = (editing: YrsSession) =>
      editing.paragraphIdentities().paragraphs.filter((entry) => entry.session?.story === 'body');
    const replace = (editing: YrsSession, key: string, text: string) => {
      const span = editing.paragraphSpans('body').find((entry) => entry.paraId === key);
      if (!span) throw new Error(`No body paragraph span for ${key}`);
      const range = {
        story: 'body',
        start: { paraId: key, offset: 0 },
        end: { paraId: key, offset: span.length }
      };
      return receiptResult(editing.replaceRange(range, text));
    };
    const save = async () => {
      const result = await saveYrsDocx(await edited(), { updateModifiedDate: false });
      saved = result.bytes;
      return result;
    };
    const paragraphs = () => collectParagraphs(read().package.document.content);
    return {
      operations: {
        open: (args) =>
          reopen(readFileSync(new URL(`../../${String(args.fixture)}`, import.meta.url))),
        openSaved: () => {
          if (saved === undefined) throw new Error('openSaved needs a save first');
          return reopen(saved);
        },
        text: () => joinText(paragraphs().map((paragraph) => paragraphResult(paragraph).text)),
        structure: () => {
          const pkg = read().package;
          return {
            bodyParagraphs: paragraphs().length,
            bodyTables: collectTables(pkg.document.content).length,
            sections: pkg.document.sections?.length ?? 0,
            headers: pkg.headers?.size ?? 0,
            footers: pkg.footers?.size ?? 0,
            footnotes: pkg.footnotes?.length ?? 0,
            endnotes: pkg.endnotes?.length ?? 0
          };
        },
        paragraphIds: () => paragraphs().map((paragraph) => paragraph.paraId ?? null),
        paragraphs: () => paragraphs().map(paragraphResult),
        runHighlights: () => paragraphs().map(runHighlights),
        paragraph: (args) => {
          const all = paragraphs();
          if (typeof args.index === 'number') {
            const found = all[args.index];
            if (!found) throw new Error(`No paragraph at index ${args.index}`);
            return paragraphResult(found);
          }
          const matches = all.filter((paragraph) => paragraph.paraId === args.id);
          return matches.length === 1 ? paragraphResult(matches[0]!) : null;
        },
        tables: () => collectTables(read().package.document.content).map(tableResult),
        cellMerges: () => collectTables(read().package.document.content).map(cellMerges),
        sections: () => (read().package.document.sections ?? []).map(sectionResult),
        headers: () => headerFooterResults(read().package.headers),
        footers: () => headerFooterResults(read().package.footers),
        warnings: () => read().warnings ?? [],
        listContentControls: (args) => {
          read();
          return listDocxContentControls(bytes, args as DocxContentControlsOptions);
        },
        findContentControls: (args) => {
          read();
          return findDocxContentControls(
            bytes,
            args.query as DocxContentControlQuery,
            args.options as DocxContentControlsOptions
          );
        },
        exportStructured: (args) => {
          read();
          return exportDocxStructured(bytes, args as unknown as DocxExportOptions);
        },
        exportMarkdown: (args) => {
          read();
          return exportDocxMarkdown(bytes, args as unknown as DocxExportOptions);
        },
        renderPng: async (args) => {
          read();
          const rendering = await createYrsSession();
          const calls = engineKind === 'wasm' ? stubBrowser() : undefined;
          try {
            rendering.openDocx(Uint8Array.from(bytes), true);
            let skippedPictures = 0;
            const png = await rendering.renderPng(
              { page: Number(args.page) },
              {
                onDiagnostic: (diagnostic) => {
                  if (diagnostic.code === 'picture-unavailable') {
                    skippedPictures += diagnostic.pictures.length;
                  }
                }
              }
            );
            const [width, height] = calls === undefined ? pngSize(png) : calls.sizes[0]!;
            return { width, height, skippedPictures };
          } finally {
            if (calls !== undefined) restoreBrowser();
            rendering.destroy();
          }
        },
        replaceParagraphText: async (args) => {
          const editing = await edited();
          const identity = bodyParagraphs(editing).find(
            (entry) => entry.ooxmlParaId?.toUpperCase() === String(args.paraId)
          );
          if (!identity?.session) {
            throw new Error(`No body paragraph saves with ${String(args.paraId)}`);
          }
          return replace(editing, identity.session.paraId, String(args.text));
        },
        replaceParagraphTextAt: async (args) => {
          const editing = await edited();
          const identity = bodyParagraphs(editing)[Number(args.index)];
          if (!identity?.session) {
            throw new Error(`No body paragraph at index ${Number(args.index)}`);
          }
          return replace(editing, identity.session.paraId, String(args.text));
        },
        readEdited: async (args) => {
          const editing = await edited();
          const story = editing.readParagraphs({ story: 'body', view: 'accepted' });
          if (!story.ok) throw new Error('The body story cannot be read', { cause: story });
          const identity = bodyParagraphs(editing).find(
            (entry) => entry.ooxmlParaId?.toUpperCase() === String(args.paraId)
          );
          const key = identity?.session?.paraId;
          const paragraph = story.paragraphs.find((entry) => entry.paraId === key);
          const { bytes: edits } = await saveYrsDocx(editing, { updateModifiedDate: false });
          const model = await parse(edits);
          editedModel = model;
          const except = new Set(
            ((args.exceptParaIds ?? []) as string[]).map((key) => key.toUpperCase())
          );
          return {
            paragraph: paragraph?.text ?? null,
            story: joinText(
              storyParagraphs(model.package.document.content)
                .filter((entry) => !except.has(entry.paraId?.toUpperCase() ?? ''))
                .map((entry) => paragraphResult(entry).text)
            )
          };
        },
        pageBreakMove: (args) => {
          if (editedModel === undefined) throw new Error('pageBreakMove needs readEdited first');
          const story = storyParagraphs(editedModel.package.document.content);
          return (args.paraIds as string[]).map((key) => {
            const entry = story.find((item) => item.paraId?.toUpperCase() === key.toUpperCase());
            return { id: key, text: entry === undefined ? null : paragraphResult(entry).text };
          });
        },
        save: async () => (await save()).bytes,
        saveReport: async () => {
          const { bytes: reported, keptUnchanged, structureDropped } = await save();
          return {
            keptUnchanged: keptUnchanged ?? null,
            structureDropped: structureDropped ?? null,
            ...partsResult(reported)
          };
        }
      },
      dispose: close
    };
  }
};
