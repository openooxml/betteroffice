import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
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
