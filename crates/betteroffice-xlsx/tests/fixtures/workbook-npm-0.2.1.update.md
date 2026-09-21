`workbook-npm-0.2.1.update.bin` was produced with published `@betteroffice/xlsx@0.2.1` from `packages/xlsx/test-fixtures/sample.xlsx`:

```js
const workbook = openWorkbook(source, { collaborative: true, clientId: 1602 });
workbook.editCell(0, 42, 0, 'PublishedReleaseState');
const snapshot = workbook.encodeStateAsUpdate();
```

The fixture retains the released package's bootstrap and style identities for migration tests.
