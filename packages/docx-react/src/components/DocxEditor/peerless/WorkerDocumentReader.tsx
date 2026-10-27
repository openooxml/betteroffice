import { useEffect, useState } from 'react';
import type { DocxSessionParagraphAnchor } from '@betteroffice/docx/yrs';
import type { createPeerlessShellController } from './shellController';
import { useTranslation } from '../../../i18n';

export function WorkerDocumentReader({ controller, version }: {
  controller: ReturnType<typeof createPeerlessShellController>;
  version: number;
}) {
  const { t } = useTranslation();
  const [open, setOpen] = useState(false);
  const [offset, setOffset] = useState(0);
  const [state, setState] = useState({ text: [] as string[], total: 0, busy: false, error: '' });
  useEffect(() => { setOffset(0); }, [version]);
  useEffect(() => {
    if (!open) return;
    let current = true;
    setState({ text: [], total: 0, busy: true, error: '' });
    void (async () => {
      const identities = await controller.ref.getParagraphIdentities();
      const paragraphs = identities.paragraphs.flatMap(({ session }) => session ? [session] : []);
      const groups = new Map<string, DocxSessionParagraphAnchor[]>();
      for (const paragraph of paragraphs.slice(offset, offset + 32)) {
        const group = groups.get(paragraph.story) ?? [];
        group.push(paragraph); groups.set(paragraph.story, group);
      }
      const text: string[] = [];
      let token: string | undefined;
      for (const [story, group] of groups) {
        const result = await controller.ref.readParagraphs({ story, paraIds: group.map(({ paraId }) => paraId), view: 'accepted' });
        if (!current) return;
        if (!result.ok) throw new Error(t('editor.documentTextUnavailable'));
        if (token && token !== result.version) throw new Error(t('editor.documentTextChanged'));
        token = result.version;
        text.push(...result.paragraphs.map(({ text }) => text));
      }
      if (current) setState({ text, total: paragraphs.length, busy: false, error: '' });
    })().catch((error: unknown) => {
      if (current) setState({ text: [], total: 0, busy: false, error: error instanceof Error ? error.message : String(error) });
    });
    return () => { current = false; };
  }, [controller, version, open, offset, t]);
  return <div className="sr-only focus-within:not-sr-only">
    <button type="button" onClick={() => setOpen(!open)} aria-expanded={open}>{t('editor.readDocumentText')}</button>
    {open && <section aria-label={t('editor.documentText')} aria-busy={state.busy}>
      {state.error && <p role="alert">{state.error}</p>}
      {state.text.map((text, index) => <p key={index}>{text}</p>)}
      <button type="button" disabled={state.busy || offset === 0} onClick={() => setOffset(Math.max(0, offset - 32))}>{t('editor.previousText')}</button>
      <button type="button" disabled={state.busy || offset + 32 >= state.total} onClick={() => setOffset(offset + 32)}>{t('editor.nextText')}</button>
    </section>}
  </div>;
}
