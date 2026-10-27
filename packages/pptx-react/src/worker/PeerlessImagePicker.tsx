import { useRef, type RefObject } from 'react';
import { insertPeerlessPicture, type PeerlessInsertionOrigin } from './usePeerlessInsertion';
import { useTranslation } from '../i18n';

export function PeerlessImagePicker({ input, origin, onError }: {
  input: RefObject<HTMLInputElement | null>; origin: RefObject<PeerlessInsertionOrigin | null>; onError(error: Error): void;
}) {
  const { t } = useTranslation();
  const preparing = useRef(false);
  return <input ref={input} type="file" accept="image/png,image/jpeg,image/gif,image/bmp,image/tiff,image/webp,image/svg+xml"
    aria-label={t('objects.insertPicture')} style={{ display: 'none' }} onChange={(event) => {
      const file = event.currentTarget.files?.[0];
      const target = origin.current;
      origin.current = null;
      event.currentTarget.value = '';
      if (!file || !target || preparing.current) return;
      preparing.current = true;
      void insertPeerlessPicture(target, file).catch((error: unknown) => {
        if (target.run.current) onError(error instanceof Error ? error : new Error(String(error)));
      }).finally(() => { preparing.current = false; });
    }} />;
}
