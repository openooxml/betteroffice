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
