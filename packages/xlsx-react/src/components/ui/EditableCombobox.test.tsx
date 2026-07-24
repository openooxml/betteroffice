import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterEach, describe, expect, it } from 'bun:test';
import { useState } from 'react';
import { EditableCombobox } from './EditableCombobox';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { cleanup, fireEvent, render } = await import('@testing-library/react');

afterEach(cleanup);

describe('EditableCombobox', () => {
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
