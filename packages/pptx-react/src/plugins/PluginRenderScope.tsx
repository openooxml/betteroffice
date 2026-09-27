import { Component, type ReactNode } from 'react';
import { claimPluginEvent } from '../commands/pluginEvents';
import { PptxCommandContext } from '../commands/PptxCommandProvider';
import type { PptxPluginActivation, PptxPluginHost } from './createPptxPluginHost';
import type { PptxPluginErrorPhase } from './types';

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

/**
 * Renders one plugin contribution: behind its own error boundary and with the plugin's restricted
 * commands in place of the editor's, so public command hooks inside it act for the plugin. Its
 * keystrokes, portals included, belong to the editor: they never edit the slide's text, and the
 * editor's shortcuts apply there as anywhere in its chrome.
 */
export function PluginRenderScope({
  host,
  activation,
  phase = 'render',
  children,
}: {
  host: PptxPluginHost;
  activation: PptxPluginActivation;
  phase?: PptxPluginErrorPhase;
  children: ReactNode;
}) {
  const store = host.commandStore(activation);
  return (
    <PluginContributionBoundary onError={(error) => host.fail(activation.pluginId, phase, error)}>
      <PptxCommandContext.Provider value={store}>
        <div
          style={{ display: 'contents' }}
          onKeyDownCapture={(event) => claimPluginEvent(event.nativeEvent, store)}
        >
          {children}
        </div>
      </PptxCommandContext.Provider>
    </PluginContributionBoundary>
  );
}
