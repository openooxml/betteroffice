import type { ComponentProps } from 'react';
import { rangeRect } from '@betteroffice/xlsx';
import { useTranslation } from '../i18n';
import { expandRangeToMergedCells } from '../viewer/sessionGeometry';
import { PresenceStrip as PresenceStripView, RemoteSelections as RemoteSelectionsView } from './PresenceView';

export { expandRangeToMergedCells } from '../viewer/sessionGeometry';

const geometry = { rangeRect, expandRangeToMergedCells };

export function PresenceStrip(props: Omit<ComponentProps<typeof PresenceStripView>, 't'>) {
  const { t } = useTranslation();
  return <PresenceStripView {...props} t={t} />;
}

export function RemoteSelections(props: Omit<ComponentProps<typeof RemoteSelectionsView>, 'geometry'>) {
  return <RemoteSelectionsView {...props} geometry={geometry} />;
}
