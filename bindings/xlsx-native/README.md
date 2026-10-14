# @betteroffice/xlsx-native

Native Rust bindings for the BetterOffice XLSX engine in Node.js, Bun and Deno.
Facade operations return promises and run in call order on one native worker.

```sh
npm install @betteroffice/xlsx-native
```

```js
import { readFile, writeFile } from 'node:fs/promises';
import { openWorkbook } from '@betteroffice/xlsx-native';

const workbook = await openWorkbook(await readFile('input.xlsx'));
await workbook.setNumberFormat(0, 'A1', 'plainText');
await workbook.setNumberFormat(0, 'B1', { type: 'custom', pattern: '0.000' });
await workbook.setStyle(0, 'A1:B1', {
  border: { preset: 'all', style: 'solid', color: '#123456' },
  horizontalAlignment: 'center', textWrapping: 'wrap'
});
await workbook.setStyle(0, 'A1', { clear: ['borders'] });
const proposal = await workbook.propose({
  agentId: 'editor',
  edits: [{ sheet: 0, address: 'C1', input: '12.5', numberFormat: 'currency' }]
});
await workbook.acceptProposal(proposal.id);
await writeFile('output.xlsx', await workbook.save());
```

Prebuilt binaries cover macOS arm64/x64, GNU Linux arm64/x64, and Windows x64.
The package shares its version with `@betteroffice/xlsx`.
CommonJS `require()` is also supported.
