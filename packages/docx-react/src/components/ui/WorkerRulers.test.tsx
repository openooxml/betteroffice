import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { HorizontalRuler } from './HorizontalRuler';
import { VerticalRuler } from './VerticalRuler';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, fireEvent, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

test('worker tab markers include the page margin and retain alignment and leader after a scaled drag', () => {
  const change = mock((_tabs: unknown) => {});
  const start = mock(() => {});
  const end = mock(() => {});
  const { container, getByTitle } = render(<HorizontalRuler editable zoom={2} sectionProps={{ marginLeft: 1440 }}
    tabStops={[{ position: 720, alignment: 'right', leader: 'dot' }]} onTabStopsChange={change} onDragStart={start} onDragEnd={end} />);
  const ruler = container.firstElementChild as HTMLElement;
  ruler.getBoundingClientRect = () => ({ left: 50 } as DOMRect);
  const marker = getByTitle('Right tab at 0.50"');
  expect(marker.style.left).toBe('283px');
  fireEvent.mouseDown(marker, { clientX: 338 });
  fireEvent.mouseMove(document, { clientX: 434 });
  expect(change).not.toHaveBeenCalled();
  fireEvent.mouseUp(document);
  expect(change).toHaveBeenCalledWith([{ position: 1440, alignment: 'right', leader: 'dot' }]);
  expect(start).toHaveBeenCalledTimes(1);
  expect(end).toHaveBeenCalledTimes(1);
});

test('a blank ruler adds a content-relative tab and read-only markers cannot remove it', () => {
  const change = mock((_tabs: unknown) => {});
  const remove = mock(() => {});
  const { container, getByTitle, rerender } = render(<HorizontalRuler editable zoom={2} onTabStopsChange={change} />);
  const ruler = container.firstElementChild as HTMLElement;
  ruler.getBoundingClientRect = () => ({ left: 50 } as DOMRect);
  fireEvent.doubleClick(ruler, { clientX: 434 });
  expect(change).toHaveBeenCalledWith([{ position: 1440, alignment: 'left' }]);
  rerender(<HorizontalRuler editable={false} tabStops={[{ position: 720, alignment: 'left' }]} onTabStopsChange={change} onTabStopRemove={remove} />);
  fireEvent.doubleClick(getByTitle('Left tab at 0.50"'));
  expect(remove).not.toHaveBeenCalled();
});

test('vertical margin drag announces its release after its last authored value', () => {
  const values: Array<number | string> = [];
  const { container, getByRole } = render(<VerticalRuler editable zoom={2} onDragStart={() => values.push('start')}
    onTopMarginChange={value => values.push(value)} onDragEnd={() => values.push('end')} />);
  (container.firstElementChild as HTMLElement).getBoundingClientRect = () => ({ top: 20 } as DOMRect);
  fireEvent.mouseDown(getByRole('slider', { name: 'Top margin' }));
  fireEvent.mouseMove(document, { clientY: 308 });
  fireEvent.mouseUp(document);
  expect(values).toEqual(['start', 2160, 'end']);
});

test('clear and list tab stops are named by their own type', () => {
  const { getByTitle } = render(<HorizontalRuler tabStops={[{ position: 720, alignment: 'clear' }, { position: 1440, alignment: 'num' }]} />);
  expect(getByTitle('Clear tab at 0.50"')).toBeTruthy();
  expect(getByTitle('List tab at 1.00"')).toBeTruthy();
});

test('indent and margin markers name their value in the ruler unit on hover, focus and drag', () => {
  const left = mock((_twips: number) => {});
  const { container, getByRole, queryByRole, rerender } = render(<HorizontalRuler editable unit="cm"
    sectionProps={{ marginLeft: 1418, marginRight: 1134 }} indentLeft={567} onIndentLeftChange={() => {}}
    onLeftMarginChange={left} />);
  const indent = getByRole('slider', { name: 'Left indent' });
  expect(indent.getAttribute('aria-valuetext')).toBe('Left indent: 1.0 cm');
  fireEvent.mouseEnter(indent);
  expect(getByRole('tooltip').textContent).toBe('Left indent: 1.0 cm');
  fireEvent.mouseLeave(indent);
  expect(queryByRole('tooltip')).toBeNull();
  fireEvent.focus(indent);
  expect(getByRole('tooltip').textContent).toBe('Left indent: 1.0 cm');
  fireEvent.blur(indent);
  const ruler = container.firstElementChild as HTMLElement;
  ruler.getBoundingClientRect = () => ({ left: 0, top: 0 } as DOMRect);
  const marginZone = ruler.children[0] as HTMLElement;
  fireEvent.mouseEnter(marginZone);
  expect(getByRole('tooltip').textContent).toBe('Left margin: 2.5 cm');
  fireEvent.mouseDown(marginZone, { clientX: 40 });
  fireEvent.mouseMove(document, { clientX: 96 });
  expect(left).toHaveBeenLastCalledWith(1440);
  expect(getByRole('tooltip').textContent).toBe('Left margin: 2.5 cm');
  fireEvent.mouseMove(document, { clientX: 192 });
  expect(getByRole('tooltip').textContent).toBe('Left margin: 5.1 cm');
  fireEvent.mouseUp(document);
  fireEvent.mouseLeave(marginZone);
  expect(queryByRole('tooltip')).toBeNull();
  rerender(<HorizontalRuler editable showFirstLineIndent hangingIndent firstLineIndent={360} indentLeft={720}
    onIndentLeftChange={() => {}} />);
  fireEvent.mouseEnter(getByRole('slider', { name: 'First line indent' }));
  expect(getByRole('tooltip').textContent).toBe('Hanging indent: 0.25"');
});

test('vertical margin markers name their value on hover and drag, and tab stops use the ruler unit', () => {
  const top = mock((_twips: number) => {});
  const { container, getByRole, getByTitle } = render(<>
    <VerticalRuler editable unit="cm" sectionProps={{ marginTop: 1134, marginBottom: 1440 }} onTopMarginChange={top} />
    <HorizontalRuler unit="cm" tabStops={[{ position: 1134, alignment: 'left' }]} />
  </>);
  const marker = getByRole('slider', { name: 'Top margin' });
  expect(marker.getAttribute('aria-valuetext')).toBe('Top margin: 2.0 cm');
  fireEvent.mouseEnter(marker);
  expect(getByRole('tooltip').textContent).toBe('Top margin: 2.0 cm');
  (container.querySelector('.docx-vertical-ruler') as HTMLElement).getBoundingClientRect = () => ({ left: 0, top: 0 } as DOMRect);
  fireEvent.mouseDown(marker);
  fireEvent.mouseMove(document, { clientY: 192 });
  expect(top).toHaveBeenLastCalledWith(2880);
  expect(getByRole('tooltip').textContent).toBe('Top margin: 5.1 cm');
  fireEvent.mouseUp(document);
  expect(getByTitle('Left tab at 2.0 cm')).toBeTruthy();
});
