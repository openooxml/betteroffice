import { expect, it } from 'bun:test';
import type { RawXml } from '../../types/content';
import type { BlockContent } from '../../types/document';
import { serializeBlockContent } from '.';

it.each([
  '<x:block xmlns:x="urn:raw-block"/>',
  '<x:block xmlns:x="urn:raw-block" x:value="a&amp;&lt;&quot;"> before <x:child/> after </x:block>',
])('serializes a standalone raw block without losing markup: %s', (xml) => {
  const raw: RawXml = { type: 'rawXml', xml };
  const block: BlockContent = raw;

  expect(serializeBlockContent(block)).toBe(xml);
});
