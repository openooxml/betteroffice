import { useCallback, useSyncExternalStore } from 'react';
import { pptxCommandController } from '../commands/createPptxCommandStore';
import { usePptxCommands } from '../commands/hooks';
import type { PptxPluginCommandId } from '../commands/types';
import { ToolbarCommandButton } from '../components/toolbar/ToolbarCommand';
import { ToolbarGroup } from '../components/ui/ToolbarPrimitives';
import { useTranslation } from '../i18n';

const NONE: readonly PptxPluginCommandId[] = Object.freeze([]);

/**
 * The toolbar commands installed plugins contribute, in plugin and declaration order. The
 * default toolbar includes it; place it in replacement chrome to keep those commands reachable.
 *
 * @experimental The plugin API may change in minor releases.
 */
export function PptxPluginToolbar(): React.ReactElement | null {
  const store = usePptxCommands();
  const controller = pptxCommandController(store);
  const read = useCallback(() => controller?.pluginToolbar() ?? NONE, [controller]);
  const ids = useSyncExternalStore(store.subscribe, read, read);
  const { t } = useTranslation();
  if (ids.length === 0) return null;
  return (
    <ToolbarGroup label={t('plugins.toolbarGroup')}>
      {ids.map((id) => (
        <ToolbarCommandButton key={id} id={id} />
      ))}
    </ToolbarGroup>
  );
}
