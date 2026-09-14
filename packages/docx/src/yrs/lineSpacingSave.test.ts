import { expect, test } from 'bun:test';
import { paragraphAttrsToFormatting } from './saveFormatting';

test('fresh paragraph spacing keeps explicit zero and other formatting', () => {
  expect(paragraphAttrsToFormatting({
    spaceBefore: 0,
    spaceAfter: 120,
    spaceBeforeLines: 0,
    spaceAfterLines: 0,
    beforeAutospacing: false,
    afterAutospacing: false,
    indentLeft: 720,
    lineSpacing: 360,
  })).toMatchObject({
    spaceBefore: 0,
    spaceAfter: 120,
    spaceBeforeLines: 0,
    spaceAfterLines: 0,
    beforeAutospacing: false,
    afterAutospacing: false,
    indentLeft: 720,
    lineSpacing: 360,
  });
});

test('imported paragraph spacing keeps the updated authored properties', () => {
  const formatting = {
    spaceBefore: 0,
    spaceAfter: 120,
    spaceBeforeLines: 0,
    spaceAfterLines: 0,
    beforeAutospacing: false,
    afterAutospacing: false,
    indentLeft: 720,
  };
  expect(paragraphAttrsToFormatting({ ...formatting, _originalFormatting: formatting }))
    .toMatchObject(formatting);
});
