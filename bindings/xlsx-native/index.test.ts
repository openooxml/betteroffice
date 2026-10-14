import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'bun:test';

import { openWorkbook, type NumberFormatMutation, type StyleInput } from './index.js';

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
    expect((await workbook.setNumberFormat(0, 'b1', 'plainText')).applied).toBe(true);
    expect((await workbook.setNumberFormat(0, 'b1', { type: 'custom', pattern: '0.00' })).applied).toBe(true);
    expect((await workbook.setStyle(0, 'a1:b1', {
      bold: true, border: { preset: 'all', style: 'solid', color: '#123456' },
      horizontalAlignment: 'center', verticalAlignment: 'middle', textWrapping: 'wrap'
    })).applied).toBe(true);
    expect((await workbook.setStyle(0, 'a1', { clear: ['bold', 'borders'] })).applied).toBe(true);

    const proposal = await workbook.propose({
      agentId: 'test-agent',
      edits: [{ sheet: 0, address: 'a2', input: '12.5', numberFormat: 'currency' }]
    });
    expect(await workbook.proposals).toHaveLength(1);
    expect((await workbook.acceptProposal(proposal.id)).applied).toBe(true);
    expect(proposal.edits[0].after).toBe('$12.50');
    expect((await workbook.cell(0, 'a2')).input).toBe('12.5');
    const rendered = await workbook.renderSheet({ sheet: 0, range: 'a1:b3' });

    expect(rendered.data.subarray(0, 8)).toEqual(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
    expect((await workbook.save()).subarray(0, 2)).toEqual(Buffer.from('PK'));
  });

  test('rejects client IDs outside the JavaScript safe integer range', async () => {
    await expect(openWorkbook(fixture, { clientId: Number.MAX_SAFE_INTEGER + 1 })).rejects.toThrow(
      'clientId must be a positive safe integer'
    );
  });

  test.each(['text', 'bogus', 'PLAINTEXT', { type: 'custom' }])('rejects invalid number format %j', async (format) => {
    const workbook = await openWorkbook(fixture);
    const before = await workbook.save();
    await expect((async () => workbook.setNumberFormat(0, 'A1', format as NumberFormatMutation))()).rejects.toThrow();
    expect(await workbook.save()).toEqual(before);
  });

  test.each(['text', 'bogus', { type: 'custom' }])('rejects invalid proposal number format %j', async (numberFormat) => {
    const workbook = await openWorkbook(fixture);
    const before = await workbook.save();
    await expect((async () => workbook.propose({
      agentId: 'test-agent',
      edits: [{ sheet: 0, address: 'A1', input: '12.5', numberFormat: numberFormat as NumberFormatMutation }]
    }))()).rejects.toThrow();
    expect(await workbook.proposals).toEqual([]);
    expect(await workbook.save()).toEqual(before);
  });

  test.each([
    { horizontalAlignment: 'justify' },
    { verticalAlignment: 'center' },
    { textWrapping: 'bogus' },
    { border: { preset: 'bogus' } },
    { border: { style: 'bogus' } },
    { clear: ['bogus'] }
  ])('rejects invalid style %j', async (style) => {
    const workbook = await openWorkbook(fixture);
    const before = await workbook.save();
    await expect((async () => workbook.setStyle(0, 'A1', style as StyleInput))()).rejects.toThrow();
    expect(await workbook.save()).toEqual(before);
  });
});
