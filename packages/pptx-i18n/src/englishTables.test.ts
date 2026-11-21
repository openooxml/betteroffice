import { expect, test } from 'bun:test';
import { createT, editorEn, en, menuEn, proposalEn, type LocaleStrings, type TranslationKey } from './index';
import enDefault, { en as enSubpath } from './en';
import { editorEn as editorTable, createT as createEditorT, deepMerge as mergeEditor } from './editor';
import { proposalEn as proposalTable } from './proposals';
import { menuEn as menuTable } from './menus';

const originalProposals = {
  "title": "Agent proposals",
  "close": "Close",
  "empty": "No changes awaiting review.",
  "preview": "Preview changes",
  "previewTitle": "Review proposed changes",
  "before": "Current slide",
  "after": "Proposed slide",
  "accept": "Accept proposal",
  "reject": "Reject",
  "force": "Apply updated proposal",
  "blank": "Empty",
  "target": "Slide {slide} · {name}",
  "targetLabel": "Change",
  "targetMissing": "The target is no longer available",
  "stale": "These objects changed after the proposal was created. Preview the current changes before applying.",
  "stalePreview": "This preview uses the updated slide. Applying will override the changed targets.",
  "geometry": "Position and size (x, y, width, height)",
  "geometryValues": "{before} → {after} pt",
  "layoutChanged": "Layout changed",
  "backgroundChanged": "Background changed",
  "rowInserted": "Row inserted",
  "rowDeleted": "Row deleted",
  "columnInserted": "Column inserted",
  "columnDeleted": "Column deleted",
  "broughtToFront": "Brought to front",
  "sentToBack": "Sent to back",
  "broughtForward": "Brought forward",
  "sentBackward": "Sent backward",
  "zOrderChanged": "Stacking order changed",
  "appearance": "Appearance or text formatting changed",
  "changedAgain": "The target changed again. Review the refreshed preview before applying.",
  "canvasTitle": "Changes on this slide",
  "editSlide": "Edit slide",
  "showDiff": "Show changes",
  "details": "Review details",
  "removed": "Removed",
  "added": "Added",
  "reviewHint": "Review mode · changes apply when accepted",
  "canvasStale": "Targets changed. Open Review details to inspect the updated preview before applying.",
  "previousPosition": "Previous",
  "proposedPosition": "Proposed",
  "loadingDiff": "Loading changes…",
  "slideStrip": "Proposed slide order",
  "previewSlide": "Preview slide {number}",
  "insertedShape": "Inserted shape",
  "deletedShape": "Deleted shape",
  "insertedSlide": "Inserted slide",
  "deletedSlide": "Deleted slide",
  "movedFrom": "from {number}"
};

const staticKeys = ['title', 'canvasTitle', 'showDiff', 'accept', 'reject'];
const movedKeys = Object.keys(originalProposals).filter(key => !staticKeys.includes(key));

test('every English entry point preserves the complete proposal messages and types', () => {
  const strings: LocaleStrings = en;
  const key: TranslationKey = 'proposals.layoutChanged';
  expect(createT(strings)(key)).toBe('Layout changed');
  for (const english of [en, enSubpath, enDefault]) {
    expect(english.proposals).toEqual(originalProposals);
  }
});

test('the default editor table excludes every lazy proposal message', () => {
  expect(Object.keys(editorEn.proposals).sort()).toEqual([...staticKeys].sort());
  expect(Object.keys(proposalEn.proposals).sort()).toEqual([...movedKeys].sort());
  for (const key of movedKeys) {
    expect(editorEn.proposals).not.toHaveProperty(key);
    expect(proposalEn.proposals).toHaveProperty(key, originalProposals[key as keyof typeof originalProposals]);
  }
});

test('the root and en subpath retain the full English table across separate entry modules', () => {
  const fullTable = {
    ...editorTable,
    proposals: { ...editorTable.proposals, ...proposalTable.proposals },
    contextMenu: menuTable.contextMenu,
  };
  for (const english of [en, enSubpath, enDefault]) {
    expect(english).toEqual(fullTable);
    expect(createEditorT(english)('proposals.layoutChanged')).toBe('Layout changed');
  }
  expect(editorEn).toBe(editorTable);
  expect(proposalEn).toBe(proposalTable);
  expect(menuEn).toBe(menuTable);
  expect(mergeEditor(mergeEditor(editorTable, proposalTable), menuTable)).toEqual(fullTable);
});

test('the default editor table leaves context menu names to the lazy menu table', () => {
  expect(editorEn).not.toHaveProperty('contextMenu');
  expect(Object.keys(menuEn)).toEqual(['contextMenu']);
  expect(createT(en)('contextMenu.table')).toBe('Table menu');
  expect(createT(editorEn as LocaleStrings)('commands.cut')).toBe('Cut');
});
