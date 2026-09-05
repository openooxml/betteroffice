import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, it } from 'bun:test';
import type { FormattingAction, ShapeFormattingAction } from './Toolbar';
import { LocaleProvider } from '../i18n';
import { Toolbar } from './Toolbar';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const elementPrototype = HTMLElement.prototype;
const clientWidth = Object.getOwnPropertyDescriptor(elementPrototype, 'clientWidth');
// Wide enough for every section to stay out of the overflow menu, so a
// `getByTestId` here reaches the control the toolbar renders inline.
Object.defineProperty(elementPrototype, 'clientWidth', {
  configurable: true,
  get: () => 1_320,
});
const { cleanup, fireEvent, render } = await import('@testing-library/react');

afterEach(cleanup);
// The globals are process-wide; leaving them installed poisons every later suite.
afterAll(async () => {
  if (clientWidth) Object.defineProperty(elementPrototype, 'clientWidth', clientWidth);
  else Reflect.deleteProperty(elementPrototype, 'clientWidth');
  if (GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

describe('Toolbar alignment controls', () => {
  it('emits the alignment the button stands for', () => {
    const actions: FormattingAction[] = [];
    const { getByTestId } = render(
      <LocaleProvider>
        <Toolbar
          textSelectionActive
          currentFormatting={{ align: 'ctr' }}
          onFormat={(action) => actions.push(action)}
        />
      </LocaleProvider>
    );

    fireEvent.click(getByTestId('pptx-align-right'));
    fireEvent.click(getByTestId('pptx-align-justify'));

    expect(actions).toEqual([
      { type: 'align', value: 'r' },
      { type: 'align', value: 'just' },
    ]);
  });

  it('marks the alignment of the current selection', () => {
    const { getByTestId } = render(
      <LocaleProvider>
        <Toolbar textSelectionActive currentFormatting={{ align: 'ctr' }} onFormat={() => {}} />
      </LocaleProvider>
    );

    expect(getByTestId('pptx-align-center').getAttribute('aria-pressed')).toBe('true');
    expect(getByTestId('pptx-align-left').getAttribute('aria-pressed')).toBeNull();
  });

  it('disables the buttons without a text selection', () => {
    const actions: FormattingAction[] = [];
    const { getByTestId } = render(
      <LocaleProvider>
        <Toolbar onFormat={(action) => actions.push(action)} />
      </LocaleProvider>
    );

    fireEvent.click(getByTestId('pptx-align-center'));

    expect(actions).toEqual([]);
  });
});

describe('Toolbar shape controls', () => {
  it('arms a preset shape placement tool', () => {
    const tools: string[] = [];
    const { getByTestId } = render(
      <LocaleProvider>
        <Toolbar onToolChange={(tool) => tools.push(tool)} />
      </LocaleProvider>
    );

    fireEvent.click(getByTestId('pptx-tool-shape'));
    fireEvent.click(getByTestId('pptx-shape-roundRect'));

    expect(tools).toEqual(['shape:roundRect']);
  });

  it('emits fill, border, width, and corner-radius actions', () => {
    const actions: ShapeFormattingAction[] = [];
    const { getByLabelText, getByTestId } = render(
      <LocaleProvider>
        <Toolbar
          shapeSelectionActive
          currentShapeFormatting={{
            geometry: 'roundRect',
            fillColor: '#d9eaf7',
            strokeColor: '#202124',
            strokeWidthPt: 1,
            adjustments: { adj: 0.17 },
          }}
          onShapeFormat={(action) => actions.push(action)}
        />
      </LocaleProvider>
    );

    fireEvent.change(getByTestId('pptx-shape-fill'), { target: { value: '#3367d6' } });
    fireEvent.click(getByLabelText('No fill'));
    fireEvent.change(getByTestId('pptx-shape-border-color'), {
      target: { value: '#ea4335' },
    });
    fireEvent.click(getByTestId('pptx-shape-border-width'));
    fireEvent.click(getByLabelText('3 pt'));
    const radius = getByTestId('pptx-shape-corner-radius');
    fireEvent.focus(radius);
    fireEvent.input(radius, { target: { value: '40%' } });
    fireEvent.keyDown(radius, { key: 'Enter' });

    expect(actions).toContainEqual({ type: 'fillColor', value: '#3367d6' });
    expect(actions).toContainEqual({ type: 'fillColor', value: null });
    expect(actions).toContainEqual({ type: 'strokeColor', value: '#ea4335' });
    expect(actions).toContainEqual({ type: 'strokeWidth', value: 3 });
    expect(actions).toContainEqual({ type: 'adjust', name: 'adj', value: 0.4 });
  });

  it('exposes the primary adjustment for non-round presets', () => {
    const actions: ShapeFormattingAction[] = [];
    const { getByTestId, queryByTestId } = render(
      <LocaleProvider>
        <Toolbar
          shapeSelectionActive
          currentShapeFormatting={{
            geometry: 'parallelogram',
            adjustments: { adj: 0.25 },
          }}
          onShapeFormat={(action) => actions.push(action)}
        />
      </LocaleProvider>
    );

    expect(queryByTestId('pptx-shape-corner-radius')).toBeNull();
    const adjustment = getByTestId('pptx-shape-adjustment');
    expect(adjustment.getAttribute('aria-label')).toBe('Shape adjustment');
    fireEvent.focus(adjustment);
    fireEvent.input(adjustment, { target: { value: '40%' } });
    fireEvent.keyDown(adjustment, { key: 'Enter' });

    expect(actions).toContainEqual({ type: 'adjust', name: 'adj', value: 0.4 });
  });

  it('prefers adj1 for multi-value presets and hides adjustment-less controls', () => {
    const actions: ShapeFormattingAction[] = [];
    const { getByTestId, queryByTestId, rerender } = render(
      <LocaleProvider>
        <Toolbar
          shapeSelectionActive
          currentShapeFormatting={{
            geometry: 'rightArrow',
            adjustments: { adj1: 0.5, adj2: 0.5 },
          }}
          onShapeFormat={(action) => actions.push(action)}
        />
      </LocaleProvider>
    );

    const adjustment = getByTestId('pptx-shape-adjustment');
    fireEvent.focus(adjustment);
    fireEvent.input(adjustment, { target: { value: '30%' } });
    fireEvent.keyDown(adjustment, { key: 'Enter' });
    expect(actions).toContainEqual({ type: 'adjust', name: 'adj1', value: 0.3 });
    rerender(
      <LocaleProvider>
        <Toolbar
          shapeSelectionActive
          currentShapeFormatting={{ geometry: 'rect', adjustments: {} }}
          onShapeFormat={() => {}}
        />
      </LocaleProvider>
    );
    expect(queryByTestId('pptx-shape-adjustment')).toBeNull();
  });
});
