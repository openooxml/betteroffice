import { useState } from 'react';
import { useTranslation } from '../i18n';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export function PeerlessRecovery({ run, fileName }: { run: PeerlessPresentation; fileName?: string }) {
  const { t } = useTranslation();
  const [busy, setBusy] = useState(false);
  return <button type="button" disabled={busy} onClick={() => {
    setBusy(true);
    void run.recover().then((result) => {
      const url = URL.createObjectURL(new Blob([new Uint8Array(result.bytes)], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
      const link = document.createElement('a');
      link.href = url; link.download = fileName ?? 'presentation-recovery.pptx'; link.click(); URL.revokeObjectURL(url);
    }).catch(() => {}).finally(() => setBusy(false));
  }}>{t('worker.recovery')}</button>;
}
