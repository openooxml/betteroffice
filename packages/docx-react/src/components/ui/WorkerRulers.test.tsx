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
  const marker = getByTitle('right tab at 0.50"');
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
  fireEvent.doubleClick(getByTitle('left tab at 0.50"'));
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
