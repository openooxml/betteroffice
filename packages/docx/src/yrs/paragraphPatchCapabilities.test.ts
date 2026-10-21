import { expect, it, spyOn } from 'bun:test';

import * as parseWasm from '../docx/parseWasm';
import type { DocxParagraphSavePlan } from './paragraphIdentity';
import type { EditSession } from './wasm/index';
import { wrapSession } from './yrsSessionFacade';

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
