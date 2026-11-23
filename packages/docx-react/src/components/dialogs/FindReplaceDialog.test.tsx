import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { useState } from 'react';
import { FindReplaceDialog } from './FindReplaceDialog';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('clearing the search text tells the host to drop its results', async () => {
  const onFind = mock(() => null);
  const view = render(
    <FindReplaceDialog
      isOpen onClose={() => {}} onFind={onFind} onFindNext={() => null} onFindPrevious={() => null}
      onReplace={() => false} onReplaceAll={() => 0} initialSearchText="word"
    />
  );
  await act(async () => { await new Promise((resolve) => setTimeout(resolve, 300)); });
  expect(onFind).toHaveBeenLastCalledWith('word', { matchCase: false, matchWholeWord: false });
  const input = view.container.querySelector<HTMLInputElement>('.docx-find-replace-dialog-input')!;
  act(() => { fireEvent.change(input, { target: { value: '' } }); });
  expect(onFind).toHaveBeenLastCalledWith('', { matchCase: false, matchWholeWord: false });
});

for (const method of ['Escape', 'replace Escape', 'close button']) {
  test.each([
    ['editing the query', 'dog', false, false],
    ['clearing the query', '', false, false],
    ['toggling Match case', 'word', true, false],
    ['toggling Whole words', 'word', false, true],
    ['a settled search', 'word', false, false],
  ] as const)(`${method} after %s passes the current criteria and cancels the pending search`, async (_, query, matchCase, matchWholeWord) => {
    const onFind = mock(() => null);
    const onClose = mock(() => {});
    const view = render(
      <FindReplaceDialog
        isOpen onClose={onClose} onFind={onFind} onFindNext={() => null} onFindPrevious={() => null}
        onReplace={() => false} onReplaceAll={() => 0} initialSearchText="word" replaceMode
      />
    );
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 300)); });
    expect(onFind).toHaveBeenLastCalledWith('word', { matchCase: false, matchWholeWord: false });
    act(() => {
      if (query !== 'word') fireEvent.change(view.getByRole('textbox', { name: 'Find text' }), { target: { value: query } });
      if (matchCase) fireEvent.click(view.getByRole('checkbox', { name: 'Match case' }));
      if (matchWholeWord) fireEvent.click(view.getByRole('checkbox', { name: 'Whole words' }));
    });
    const searches = onFind.mock.calls.length;
    act(() => {
      if (method === 'close button') fireEvent.click(view.getByRole('button', { name: 'Close dialog' }));
      else fireEvent.keyDown(view.getByRole('textbox', { name: method === 'replace Escape' ? 'Replace text' : 'Find text' }), { key: 'Escape' });
    });
    expect(onClose).toHaveBeenCalledTimes(1);
    expect(onClose).toHaveBeenCalledWith(query, { matchCase, matchWholeWord });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 300)); });
    expect(onFind).toHaveBeenCalledTimes(searches);
  });
}

const HANDLERS = {
  onClose: () => {}, onFind: () => null, onFindNext: () => null, onFindPrevious: () => null,
  onReplace: () => false, onReplaceAll: () => 0,
};

const focused = () => document.activeElement?.getAttribute('aria-label') ?? document.activeElement?.tagName ?? 'nothing';

test('the Find field has the focus as the dialog opens', () => {
  render(<FindReplaceDialog isOpen {...HANDLERS} />);
  expect(focused()).toBe('Find text');
});

test('opening with search text fills the Find field and selects that text at once', () => {
  const view = render(<FindReplaceDialog isOpen={false} {...HANDLERS} />);
  for (const text of ['word', 'another']) {
    view.rerender(<FindReplaceDialog isOpen {...HANDLERS} initialSearchText={text} />);
    const field = view.getByRole('textbox', { name: 'Find text' }) as HTMLInputElement;
    expect(field.value).toBe(text);
    expect(focused()).toBe('Find text');
    expect([field.selectionStart, field.selectionEnd]).toEqual([0, text.length]);
    view.rerender(<FindReplaceDialog isOpen={false} {...HANDLERS} />);
  }
});

test('opened for replacing, the dialog still starts in the Find field', () => {
  render(<FindReplaceDialog isOpen replaceMode {...HANDLERS} />);
  expect(focused()).toBe('Find text');
});

test('the Replace field has the focus as "+ Replace" shows it', () => {
  const view = render(<FindReplaceDialog isOpen {...HANDLERS} />);
  act(() => { fireEvent.click(view.getByRole('button', { name: '+ Replace' })); });
  expect(focused()).toBe('Replace text');
});

test('letters typed right after the shortcut that opens the dialog go to the Find field, not to the document', () => {
  function Host() {
    const [open, setOpen] = useState(false);
    const [text, setText] = useState('Text');
    return (
      <>
        <textarea
          aria-label="Document" value={text} onChange={(event) => setText(event.target.value)}
          onKeyDown={(event) => { if (event.ctrlKey && event.key === 'f') setOpen(true); }}
        />
        <FindReplaceDialog isOpen={open} {...HANDLERS} />
      </>
    );
  }
  const view = render(<Host />);
  const page = view.getByRole('textbox', { name: 'Document' }) as HTMLTextAreaElement;
  page.focus();
  fireEvent.keyDown(page, { key: 'f', ctrlKey: true });
  const typedInto = document.activeElement as HTMLInputElement | HTMLTextAreaElement;
  fireEvent.change(typedInto, { target: { value: `${typedInto.value}zq` } });
  expect((view.getByRole('textbox', { name: 'Find text' }) as HTMLInputElement).value).toBe('zq');
  expect(page.value).toBe('Text');
});
