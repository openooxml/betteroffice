import { Component, type ReactNode, type SyntheticEvent } from 'react';
import { claimPluginEvent } from '../commands/pluginEvents';
import { XlsxCommandContext } from '../commands/XlsxCommandProvider';
import type { XlsxCommandStore } from '../commands/types';
import { EditorToolbarContext, ToolbarModeContext } from '../components/EditorToolbarContext';
import { FormulaBarContext } from '../components/toolbar/FormulaBar';
import { ToolbarOverflowContext } from '../components/toolbar/overflowRegistry';
import type { XlsxPluginActivation, XlsxPluginHost } from './createXlsxPluginHost';
import type { XlsxPluginErrorPhase } from './types';

interface BoundaryProps {
  onError(error: unknown): void;
  children?: ReactNode;
}

/** Stops a failing contribution at its own subtree and reports it once. */
export class PluginContributionBoundary extends Component<BoundaryProps, { failed: boolean }> {
  override state = { failed: false };

  static getDerivedStateFromError(): { failed: boolean } {
    return { failed: true };
  }

  override componentDidCatch(error: unknown): void {
    this.props.onError(error);
  }

  override render(): ReactNode {
    return this.state.failed ? null : this.props.children;
  }
}

function claim(store: XlsxCommandStore) {
  return (event: SyntheticEvent) => claimPluginEvent(event.nativeEvent, store);
}

/**
 * Renders one plugin contribution: behind its own error boundary and with the plugin's restricted
 * commands in place of the editor's, so public command hooks inside it act for the plugin. The
 * editor's formula bar binding and toolbar internals are withheld, so public parts that need them
 * render nothing there. Its keystrokes and clicks, portals included, never reach the grid: the
 * editor's shortcuts apply there as anywhere in its chrome, and a plugin's text fields keep their
 * own keys.
 */
export function PluginRenderScope({
  host,
  activation,
  phase = 'render',
  children,
}: {
  host: XlsxPluginHost;
  activation: XlsxPluginActivation;
  phase?: XlsxPluginErrorPhase;
  children: ReactNode;
}) {
  const store = host.commandStore(activation);
  const mark = claim(store);
  return (
    <PluginContributionBoundary onError={(error) => host.fail(activation.pluginId, phase, error)}>
      <XlsxCommandContext.Provider value={store}>
        <FormulaBarContext.Provider value={null}>
          <EditorToolbarContext.Provider value={null}>
            <ToolbarModeContext.Provider value={null}>
              <ToolbarOverflowContext.Provider value={null}>
                <div
                  style={{ display: 'contents' }}
                  onKeyDownCapture={mark}
                  onMouseDownCapture={mark}
                  onClickCapture={mark}
                  onDoubleClickCapture={mark}
                >
                  {children}
                </div>
              </ToolbarOverflowContext.Provider>
            </ToolbarModeContext.Provider>
          </EditorToolbarContext.Provider>
        </FormulaBarContext.Provider>
      </XlsxCommandContext.Provider>
    </PluginContributionBoundary>
  );
}
