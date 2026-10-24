import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, it } from 'bun:test';
import { useState } from 'react';
import { flushSync } from 'react-dom';
import { EditableCombobox } from './EditableCombobox';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');

afterEach(cleanup);
// The globals are process-wide; leaving them installed poisons every later suite.
afterAll(async () => {
  if (GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

describe('EditableCombobox', () => {
  it.each([
    ['Zoom', '50%', '83%', '83%'],
    ['Zoom', '100%', '75%', '75%'],
    ['Font size', '18', ' 24 ', '24'],
    ['Shape adjustment', '17%', '40', '40%'],
  ])(
    'keeps the synchronously applied %s value after Enter and blur (%s to %s)',
    async (label, initial, entered, applied) => {
      const commits: string[] = [];
      function Field() {
        const [value, setValue] = useState(initial);
        return (
          <EditableCombobox
            value={value}
            options={[]}
            label={label}
            onCommit={(draft) => {
              commits.push(draft);
              flushSync(() => setValue(applied));
            }}
          />
        );
      }
      const { getByRole } = render(<Field />);
      const input = getByRole('combobox') as HTMLInputElement;

      act(() => input.focus());
      fireEvent.input(input, { target: { value: entered } });
      fireEvent.keyDown(input, { key: 'Enter' });
      await act(() => new Promise((resolve) => requestAnimationFrame(resolve)));

      expect(commits).toEqual([entered.trim()]);
      expect(input.value).toBe(applied);
    }
  );

  it('keeps the synchronously applied value after a blur commit', async () => {
    const commits: string[] = [];
    function Field() {
      const [value, setValue] = useState('50%');
      return (
        <EditableCombobox
          value={value}
          options={[]}
          label="Zoom"
          onCommit={(draft) => {
            commits.push(draft);
            flushSync(() => setValue(`${Number.parseFloat(draft)}%`));
          }}
        />
      );
    }
    const { getByRole } = render(<Field />);
    const input = getByRole('combobox') as HTMLInputElement;

    act(() => input.focus());
    fireEvent.input(input, { target: { value: '83' } });
    act(() => input.blur());
    await act(() => new Promise((resolve) => requestAnimationFrame(resolve)));

    expect(commits).toEqual(['83']);
    expect(input.value).toBe('83%');
  });

  it('restores the applied value after a rejected commit', () => {
    const { getByRole } = render(<ZoomCombobox />);
    const input = getByRole('combobox') as HTMLInputElement;

    fireEvent.focus(input);
    fireEvent.input(input, { target: { value: 'abc' } });
    fireEvent.keyDown(input, { key: 'Enter' });

    expect(input.value).toBe('100%');
  });

  it('commits valid input and displays the applied value', () => {
    const commits: string[] = [];
    const { getByRole } = render(<ZoomCombobox onCommit={(value) => commits.push(value)} />);
    const input = getByRole('combobox') as HTMLInputElement;

    fireEvent.focus(input);
    fireEvent.input(input, { target: { value: '120' } });
    fireEvent.keyDown(input, { key: 'Enter' });

    expect(commits).toContain('120');
    expect(input.value).toBe('120%');
  });

  it('commits once when Enter is followed by a blur', async () => {
    const commits: string[] = [];
    const { getByRole } = render(<ZoomCombobox onCommit={(value) => commits.push(value)} />);
    const input = getByRole('combobox') as HTMLInputElement;

    fireEvent.focus(input);
    fireEvent.input(input, { target: { value: '120' } });
    fireEvent.keyDown(input, { key: 'Enter' });
    fireEvent.blur(input);
    await act(() => new Promise((resolve) => requestAnimationFrame(resolve)));

    expect(commits).toEqual(['120']);
    expect(input.value).toBe('120%');
  });

  it('commits nothing when an untouched field loses focus', async () => {
    const commits: string[] = [];
    const { getByRole } = render(<ZoomCombobox onCommit={(value) => commits.push(value)} />);
    const input = getByRole('combobox') as HTMLInputElement;

    fireEvent.focus(input);
    fireEvent.blur(input);
    await act(() => new Promise((resolve) => requestAnimationFrame(resolve)));

    expect(commits).toEqual([]);
  });

  it('commits a typed value once when the field loses focus', async () => {
    const commits: string[] = [];
    const { getByRole } = render(<ZoomCombobox onCommit={(value) => commits.push(value)} />);
    const input = getByRole('combobox') as HTMLInputElement;

    fireEvent.focus(input);
    fireEvent.input(input, { target: { value: '150' } });
    fireEvent.blur(input);
    await act(() => new Promise((resolve) => requestAnimationFrame(resolve)));

    expect(commits).toEqual(['150']);
    expect(input.value).toBe('150%');
  });

  it('discards a typed value on Escape', async () => {
    const commits: string[] = [];
    const { getByRole } = render(<ZoomCombobox onCommit={(value) => commits.push(value)} />);
    const input = getByRole('combobox') as HTMLInputElement;

    fireEvent.focus(input);
    fireEvent.input(input, { target: { value: '150' } });
    fireEvent.keyDown(input, { key: 'Escape' });
    fireEvent.blur(input);
    await act(() => new Promise((resolve) => requestAnimationFrame(resolve)));

    expect(commits).toEqual([]);
    expect(input.value).toBe('100%');
  });

  it('does not commit a draft replaced by an external value', async () => {
    const commits: string[] = [];
    const field = (value: string) => (
      <EditableCombobox value={value} options={[]} label="Zoom" onCommit={(draft) => commits.push(draft)} />
    );
    const { getByRole, rerender } = render(field('100%'));
    const input = getByRole('combobox') as HTMLInputElement;

    fireEvent.focus(input);
    fireEvent.input(input, { target: { value: '150' } });
    rerender(field('80%'));
    fireEvent.blur(input);
    await act(() => new Promise((resolve) => requestAnimationFrame(resolve)));

    expect(commits).toEqual([]);
    expect(input.value).toBe('80%');
  });

  it('ignores keys and drops the draft while disabled', async () => {
    const commits: string[] = [];
    const { getByRole } = render(
      <EditableCombobox value="100%" options={[{ value: '50', label: '50%' }]} label="Zoom" disabled
        description="Open a workbook first" onCommit={(draft) => commits.push(draft)} />
    );
    const input = getByRole('combobox') as HTMLInputElement;

    act(() => input.focus());
    fireEvent.input(input, { target: { value: '150' } });
    fireEvent.keyDown(input, { key: 'ArrowDown' });
    expect(input.getAttribute('aria-expanded')).toBe('false');
    fireEvent.keyDown(input, { key: 'Enter' });
    act(() => input.blur());
    await act(() => new Promise((resolve) => requestAnimationFrame(resolve)));

    expect(commits).toEqual([]);
    expect(input.value).toBe('100%');
  });

  it('commits a clicked option once as a pointer commit', async () => {
    const commits: Array<[string, boolean]> = [];
    const { getByRole } = render(
      <EditableCombobox value="100%" options={[{ value: '50', label: '50%' }]} label="Zoom"
        onCommit={(draft, pointer) => commits.push([draft, pointer])} />
    );
    const input = getByRole('combobox') as HTMLInputElement;

    act(() => input.focus());
    fireEvent.input(input, { target: { value: '150' } });
    fireEvent.click(getByRole('option'));
    act(() => input.blur());
    await act(() => new Promise((resolve) => requestAnimationFrame(resolve)));

    expect(commits).toEqual([['50', true]]);
  });
});

function ZoomCombobox({ onCommit }: { onCommit?: (value: string) => void }) {
  const [value, setValue] = useState('100%');

  return (
    <EditableCombobox
      value={value}
      options={[]}
      label="Zoom"
      onCommit={(draft) => {
        onCommit?.(draft);
        const percent = Number.parseFloat(draft.replace('%', ''));
        if (Number.isFinite(percent) && percent >= 25 && percent <= 400) setValue(`${percent}%`);
      }}
    />
  );
}
