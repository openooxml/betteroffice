import { expect, test } from 'bun:test';
import ts from 'typescript';

test('prepared cell ranges are marked internal in the command protocol', async () => {
  const text = await Bun.file(new URL('./residentCommandProtocol.ts', import.meta.url)).text();
  const source = ts.createSourceFile('residentCommandProtocol.ts', text, ts.ScriptTarget.Latest);
  const prepared = source.statements.filter(ts.isInterfaceDeclaration)
    .find((declaration) => declaration.name.text === 'PreparedDocxCommand');
  const cellRanges = prepared?.members.filter(ts.isPropertySignature)
    .find((property) => ts.isIdentifier(property.name) && property.name.text === 'cellRanges');
  expect(cellRanges && ts.getJSDocTags(cellRanges).map(({ tagName }) => tagName.text)).toContain('internal');
});
