import { useTranslation } from '../../i18n';
import { ToolbarCommand } from './ToolbarCommand';
import { ToolbarGroup } from './ToolbarPrimitives';

/** @experimental */
export interface ToolbarReviewControlsProps {
  className?: string;
}

/**
 * Review controls: editing mode, the comments sidebar, tracked-change
 * navigation, and accepting or rejecting the change at the selection.
 */
export function ToolbarReviewControls({ className }: ToolbarReviewControlsProps) {
  const { t } = useTranslation();
  return (
    <ToolbarGroup label={t('commands.review')} className={className}>
      <ToolbarCommand id="editingMode" />
      <ToolbarCommand id="commentsSidebar" />
      <ToolbarCommand id="reviewPrevious" />
      <ToolbarCommand id="reviewNext" />
      <ToolbarCommand id="reviewAccept" />
      <ToolbarCommand id="reviewReject" />
    </ToolbarGroup>
  );
}
