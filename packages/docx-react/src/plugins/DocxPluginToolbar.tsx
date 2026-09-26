import { useCallback, useSyncExternalStore } from 'react';
import { docxCommandController } from '../commands/createDocxCommandStore';
import { useDocxCommands } from '../commands/hooks';
import type { DocxPluginCommandId } from '../commands/types';
import { ToolbarCommandButton } from '../components/toolbar/ToolbarCommand';
import { ToolbarGroup } from '../components/toolbar/ToolbarPrimitives';
import { useTranslation } from '../i18n';

const NONE: readonly DocxPluginCommandId[] = Object.freeze([]);

/**
 * The toolbar commands installed plugins contribute, in plugin and declaration order. The
 * default toolbar includes it; place it in replacement chrome to keep those commands reachable.
 *
 * @experimental The plugin API may change in minor releases.
 */
export function DocxPluginToolbar(): React.ReactElement | null {
  const store = useDocxCommands();
  const controller = docxCommandController(store);
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
