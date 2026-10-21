import { expect, it, spyOn } from 'bun:test';

import * as parseWasm from '../docx/parseWasm';
import * as opcWasm from '../docx/wasm';
import * as generatedParse from '../wasm/generated/parse/docx_parse.js';
import { writeDocumentWithRust } from '../docx/rustSaveFacade';
import type { Document } from '../types/document';
import type { DocxParagraphSavePlan } from './paragraphIdentity';
import type { EditSession } from './wasm/index';
import { wrapSession } from './yrsSessionFacade';

it('loads the declared paragraph patch writer v1 export', async () => {
  await parseWasm.preloadParseWasm();
  const version: number = generatedParse.paragraphPatchWriterVersion();
  expect(version).toBe(1);
  expect(parseWasm.PARAGRAPH_PATCH_WRITER_VERSION).toBe(version);
  expect(parseWasm.paragraphPatchWriterVersionWire()).toBe(version);
});

function plan(changed = [0]): DocxParagraphSavePlan {
  return {
    assignments: [],
    patchedParts: [],
    splicedParts: [{
      part: 'word/header1.xml',
      sha256: 'part',
      paragraphs: [[0, 'p1']],
      changed,
      paragraphPatches: [],
      legacyParagraphs: changed.map((ordinal) => ({
        ordinal,
        code: 'UNSUPPORTED_KEY',
        detail: 'key:unknown',
      })),
    }],
  };
}

function session(value: DocxParagraphSavePlan) {
  return wrapSession({
    paragraph_save_plan: () => JSON.stringify(value),
  } as unknown as EditSession, 71);
}

for (const version of [undefined, 0, 2]) {
  it(`rejects parser writer capability ${version} with the affected part`, () => {
    const capability = spyOn(parseWasm, 'paragraphPatchWriterVersionWire').mockReturnValue(version);
    try {
      expect(() => session(plan()).paragraphSavePlan()).toThrow(
        'DOCX_MINIMAL_PATCH [INVALID_PLAN] word/header1.xml p#?: ' +
        'paragraphPatchWriterVersion must report PARAGRAPH_PATCH_WRITER_VERSION=1'
      );
    } finally {
      capability.mockRestore();
    }
  });
}

it('accepts matching producer and parser capabilities', () => {
  const capability = spyOn(parseWasm, 'paragraphPatchWriterVersionWire').mockReturnValue(1);
  try {
    expect(session(plan()).paragraphSavePlan()).toEqual(plan());
    expect(capability).toHaveBeenCalledTimes(1);
  } finally {
    capability.mockRestore();
  }
});

it('keeps producer capability failures explicit', () => {
  const value = plan();
  delete value.splicedParts![0]!.paragraphPatches;
  expect(() => session(value).paragraphSavePlan()).toThrow('missing producer capability');
});

it('does not require a parser capability for unchanged paragraphs', () => {
  const capability = spyOn(parseWasm, 'paragraphPatchWriterVersionWire').mockReturnValue(undefined);
  try {
    expect(session(plan([])).paragraphSavePlan()).toEqual(plan([]));
    expect(capability).not.toHaveBeenCalled();
  } finally {
    capability.mockRestore();
  }
});

it('defers the numeric check while a parser bundle with the export is still loading', () => {
  const capability = spyOn(parseWasm, 'paragraphPatchWriterVersionWire').mockReturnValue(null);
  try {
    expect(session(plan()).paragraphSavePlan()).toEqual(plan());
  } finally {
    capability.mockRestore();
  }
});

for (const version of [undefined, null, 0, 2]) {
  it(`refuses writing with parser capability ${version} after preload`, async () => {
    const preloadParser = spyOn(parseWasm, 'preloadParseWasm').mockResolvedValue(undefined);
    const preloadOpc = spyOn(opcWasm, 'preloadOpcWasm').mockResolvedValue(undefined);
    const capability = spyOn(parseWasm, 'paragraphPatchWriterVersionWire').mockReturnValue(version);
    const writer = spyOn(parseWasm, 'writeDocxS13WithWarningsWire');
    try {
      await expect(writeDocumentWithRust(
        {} as Document,
        new ArrayBuffer(0),
        {},
        undefined,
        undefined,
        { assignments: [], patchedParts: [], splicedParts: [{
          ...plan().splicedParts![0]!, paragraphs: [0],
        }] }
      )).rejects.toThrow(
        'DOCX_MINIMAL_PATCH [INVALID_PLAN] word/header1.xml p#?: ' +
        'paragraphPatchWriterVersion must report PARAGRAPH_PATCH_WRITER_VERSION=1'
      );
      expect(writer).not.toHaveBeenCalled();
    } finally {
      writer.mockRestore();
      capability.mockRestore();
      preloadOpc.mockRestore();
      preloadParser.mockRestore();
    }
  });
}
