import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'bun:test';

import { openWorkbook } from './index.js';

const fixture = readFileSync(
  new URL('../../crates/ooxml-opc/tests/fixtures/sample.xlsx', import.meta.url)
);

describe('@betteroffice/xlsx-native', () => {
  test('opens, inspects, edits, renders, and saves a workbook', async () => {
    const workbook = await openWorkbook(fixture);

    expect(await workbook.sheetCount).toBeGreaterThan(0);
    const mutationPromise = workbook.set(0, 'a1', 'BetterOffice');
    const cellPromise = workbook.cell(0, 'a1');
    const [mutation, cell] = await Promise.all([mutationPromise, cellPromise]);
    expect(mutation.applied).toBe(true);
    expect(cell.input).toBe('BetterOffice');
    await workbook.set(0, 'b1', '=1+2');
    expect(await workbook.formula(0, 'b1')).toBe('1+2');
    expect(await workbook.value(0, 'b1')).toMatchObject({ kind: 'number', number: 3 });
    expect((await workbook.setStyle(0, 'a1:b1', { bold: true })).applied).toBe(true);

    const proposal = await workbook.propose({
      agentId: 'test-agent',
      edits: [{ sheet: 0, address: 'a2', input: 'proposed' }]
    });
    expect(await workbook.proposals).toHaveLength(1);
    expect((await workbook.acceptProposal(proposal.id)).applied).toBe(true);
    expect((await workbook.cell(0, 'a2')).input).toBe('proposed');
    const rendered = await workbook.renderSheet({ sheet: 0, range: 'a1:b3' });

    expect(rendered.data.subarray(0, 8)).toEqual(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
    expect((await workbook.save()).subarray(0, 2)).toEqual(Buffer.from('PK'));
  });

  test('rejects client IDs outside the JavaScript safe integer range', async () => {
    await expect(openWorkbook(fixture, { clientId: Number.MAX_SAFE_INTEGER + 1 })).rejects.toThrow(
      'clientId must be a positive safe integer'
    );
  });
});
