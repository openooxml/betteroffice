import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterEach, describe, expect, it } from 'bun:test';
import type { ShapeFormattingAction } from './Toolbar';
import { LocaleProvider } from '../i18n';
import { Toolbar } from './Toolbar';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
  configurable: true,
  get: () => 1_200,
});
const { cleanup, fireEvent, render } = await import('@testing-library/react');

afterEach(cleanup);

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
});
