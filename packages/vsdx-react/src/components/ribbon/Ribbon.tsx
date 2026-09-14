import { useCallback, useEffect, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, ReactNode } from 'react';
import type { TFunction } from '@betteroffice/vsdx-i18n';
import { CommandMenu } from './CommandMenu';
import { RibbonIcon } from './RibbonIcon';
import { useRibbonCommands } from './commands';
import type { RibbonCommandId } from './commands';

const tabs = ['file', 'home', 'insert', 'design', 'review', 'view', 'help'] as const;
type RibbonTab = (typeof tabs)[number];

type IconName = Parameters<typeof RibbonIcon>[0]['name'];

function CommandButton({ id, icon, label }: { id: RibbonCommandId; icon: IconName; label: string }) {
  const command = useRibbonCommands()[id];
  return <button type="button" disabled={!command.enabled} aria-label={label} aria-pressed={command.active || undefined} title={label} data-command-id={id} onMouseDown={(event) => event.preventDefault()} onClick={() => command.run()} className="vsdx-cmd-btn" style={{ ...styles.button, color: command.enabled ? '#242424' : '#b4b4b4', background: command.active ? '#ebf3fc' : 'transparent', cursor: command.enabled ? 'pointer' : 'default' }}><RibbonIcon name={icon} size={20} /></button>;
}

function ColorButton({ id, icon, label }: { id: RibbonCommandId; icon: IconName; label: string }) {
  const command = useRibbonCommands()[id];
  return <label title={label} className="vsdx-cmd-btn" style={{ ...styles.button, color: command.enabled ? '#242424' : '#b4b4b4', cursor: command.enabled ? 'pointer' : 'default', position: 'relative' }}><RibbonIcon name={icon} size={20} /><span aria-hidden="true" style={{ position: 'absolute', bottom: 4, width: 16, height: 3, borderRadius: 1, background: command.value ?? '#000000' }} /><input type="color" value={command.value ?? '#000000'} disabled={!command.enabled} aria-label={label} data-command-id={id} onChange={(event) => command.run(event.target.value)} style={styles.colorInput} /></label>;
}

function LineFormulaControl({ id, label, icon }: { id: 'lineWeight' | 'linePattern'; label: string; icon: IconName }) {
  const command = useRibbonCommands()[id];
  return <label title={label} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, color: command.enabled ? '#242424' : '#b4b4b4' }}><RibbonIcon name={icon} size={18} /><input key={command.value} aria-label={label} data-command-id={id} disabled={!command.enabled} defaultValue={command.value ?? ''} onBlur={(event) => { const next = event.currentTarget.value; if (next !== (command.value ?? '')) command.run(next); }} style={styles.formulaInput} /></label>;
}

function RibbonRun({ label, children }: { label: string; children?: ReactNode }) {
  return <div role="group" aria-label={label} style={styles.run}>{children}</div>;
}

function Divider() { return <div role="separator" aria-orientation="vertical" style={styles.divider} />; }

function RibbonSplitButton({ defaultId, defaultIcon, entries, label }: { defaultId: RibbonCommandId; defaultIcon: IconName; entries: ReadonlyArray<{ id: RibbonCommandId; icon: IconName }>; label: (id: RibbonCommandId) => string }) {
  const commands = useRibbonCommands();
  const [open, setOpen] = useState(false);
  const [intent, setIntent] = useState<'first' | 'last'>('first');
  const triggerRef = useRef<HTMLButtonElement>(null);
  const [pos, setPos] = useState({ top: 0, left: 0 });
  const fallback = entries[0] ?? { id: defaultId, icon: defaultIcon };
  const current = commands[fallback.id];
  const anyEnabled = entries.some((entry) => commands[entry.id].enabled);
  const close = useCallback(() => setOpen(false), []);
  const closeAndFocus = useCallback(() => { setOpen(false); triggerRef.current?.focus(); }, []);
  const openMenu = useCallback((next: 'first' | 'last') => { setIntent(next); setOpen(true); }, []);
  useEffect(() => {
    if (!open || !triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    setPos({ top: rect.bottom + 2, left: rect.left });
  }, [open]);
  function onTriggerKeyDown(event: KeyboardEvent<HTMLButtonElement>) {
    if (event.key === 'ArrowDown') { event.preventDefault(); if (!open && anyEnabled) openMenu('first'); }
    else if (event.key === 'ArrowUp') { event.preventDefault(); if (!open && anyEnabled) openMenu('last'); }
  }
  return (
    <span style={styles.split}>
      <button type="button" disabled={!current.enabled} aria-label={label(fallback.id)} data-command-id={fallback.id} onMouseDown={(event) => event.preventDefault()} onClick={() => current.run()} className="vsdx-cmd-btn vsdx-split-main" style={{ ...styles.splitMain, color: current.enabled ? '#242424' : '#b4b4b4', cursor: current.enabled ? 'pointer' : 'default' }}><RibbonIcon name={fallback.icon} size={20} /></button>
      <button ref={triggerRef} type="button" disabled={!anyEnabled} aria-label={`${label(fallback.id)} options`} aria-haspopup="menu" aria-expanded={open} data-split-toggle={fallback.id} onMouseDown={(event) => event.preventDefault()} onClick={() => { if (!anyEnabled) return; if (open) close(); else openMenu('first'); }} onKeyDown={onTriggerKeyDown} className="vsdx-cmd-btn" style={{ ...styles.splitChevron, color: anyEnabled ? '#242424' : '#b4b4b4', background: open ? '#ebebeb' : 'transparent', cursor: anyEnabled ? 'pointer' : 'default' }}><svg width={10} height={10} viewBox="0 0 10 10" aria-hidden="true" focusable="false" style={{ display: 'block' }}><path d="m2 3.5 3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" /></svg></button>
      {open && (
        <CommandMenu menuLabel={`${label(fallback.id)} options`} entries={entries} position={pos} anchorRef={triggerRef} initialFocus={intent} label={label} onClose={close} onCloseAndFocus={closeAndFocus} />
      )}
    </span>
  );
}

function EmptyState({ t }: { t: TFunction }) {
  return <div style={styles.empty}><span>{t('ribbon.empty')}</span></div>;
}

function HomePanel({ t }: { t: TFunction }) {
  const label = (id: RibbonCommandId) => t(`ribbon.commands.${id}`);
  return <div style={styles.surface} data-testid="vsdx-ribbon-home-panel">
    <RibbonRun label={t('ribbon.groups.history')}><CommandButton id="undo" icon="undo" label={label('undo')} /><CommandButton id="redo" icon="redo" label={label('redo')} /></RibbonRun>
    <Divider />
    <RibbonRun label={t('ribbon.groups.insert')}><CommandButton id="delete" icon="delete" label={label('delete')} /><CommandButton id="addShape" icon="add" label={label('addShape')} /></RibbonRun>
    <Divider />
    <RibbonRun label={t('ribbon.groups.shape')}><ColorButton id="fillColor" icon="fill" label={label('fillColor')} /><ColorButton id="lineColor" icon="line" label={label('lineColor')} /><LineFormulaControl id="lineWeight" icon="weight" label={label('lineWeight')} /><LineFormulaControl id="linePattern" icon="pattern" label={label('linePattern')} /></RibbonRun>
    <Divider />
    <RibbonRun label={t('ribbon.groups.arrange')}>
      <RibbonSplitButton defaultId="bringToFront" defaultIcon="front" label={label} entries={[{ id: 'bringToFront', icon: 'front' }, { id: 'bringForward', icon: 'forward' }, { id: 'sendBackward', icon: 'backward' }, { id: 'sendToBack', icon: 'back' }]} />
      <RibbonSplitButton defaultId="rotateRight" defaultIcon="rotateRight" label={label} entries={[{ id: 'rotateRight', icon: 'rotateRight' }, { id: 'rotateLeft', icon: 'rotateLeft' }, { id: 'flipHorizontal', icon: 'flipHorizontal' }, { id: 'flipVertical', icon: 'flipVertical' }]} />
    </RibbonRun>
  </div>;
}

export function Ribbon({ t }: { t: TFunction }) {
  const [active, setActive] = useState<RibbonTab>('home');
  const tabRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const select = (next: RibbonTab) => setActive(next);
  const onKeyDown = (event: KeyboardEvent<HTMLButtonElement>, index: number) => {
    let next = index;
    if (event.key === 'ArrowRight') next = (index + 1) % tabs.length;
    else if (event.key === 'ArrowLeft') next = (index + tabs.length - 1) % tabs.length;
    else if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = tabs.length - 1;
    else return;
    event.preventDefault(); select(tabs[next]); tabRefs.current[next]?.focus();
  };
  return <section aria-label={t('ribbon.label')} className="vsdx-ribbon-flat" style={styles.root}>
    <style>{'[data-command-id]:focus-visible,.vsdx-cmd-btn:focus-visible,.vsdx-ribbon-flat [role="tab"]:focus-visible{outline:2px solid #0f6cbd;outline-offset:1px}.vsdx-cmd-btn:not(:disabled):hover{background-color:#f5f5f5}.vsdx-cmd-btn:not(:disabled):active{background-color:#ebebeb}'}</style>
    <div role="tablist" aria-label={t('ribbon.tabsLabel')} style={styles.tabs}>{tabs.map((tab, index) => {
      const selected = active === tab;
      return <button ref={(node) => { tabRefs.current[index] = node; }} key={tab} id={`vsdx-ribbon-tab-${tab}`} type="button" role="tab" aria-selected={selected} aria-controls={`vsdx-ribbon-panel-${tab}`} tabIndex={selected ? 0 : -1} onClick={() => select(tab)} onKeyDown={(event) => onKeyDown(event, index)} style={styles.tab}><span style={{ ...styles.tabLabel, borderBottomColor: selected ? '#0f6cbd' : 'transparent', color: '#242424', fontWeight: selected ? 600 : 400 }}>{t(`ribbon.tabs.${tab}`)}</span></button>;
    })}</div>
    <div id={`vsdx-ribbon-panel-${active}`} role="tabpanel" aria-labelledby={`vsdx-ribbon-tab-${active}`} style={styles.panel}>
      {active === 'home' ? <HomePanel t={t} /> : active === 'file' ? <div style={styles.surface}><RibbonRun label={t('ribbon.groups.file')}><CommandButton id="download" icon="download" label={t('ribbon.commands.download')} /></RibbonRun></div> : active === 'insert' ? <div style={styles.surface}><RibbonRun label={t('ribbon.groups.insert')}><CommandButton id="addShape" icon="add" label={t('ribbon.commands.addShape')} /></RibbonRun></div> : <div style={styles.surface}><EmptyState t={t} /></div>}
    </div>
  </section>;
}

const styles: Record<string, CSSProperties> = {
  root: { flex: '0 0 auto', background: '#ffffff', borderBottom: '1px solid #e0e0e0', fontFamily: "'Segoe UI', ui-sans-serif, system-ui, sans-serif" },
  tabs: { display: 'flex', alignItems: 'stretch', height: 32, padding: '0 8px', gap: 2, borderBottom: '1px solid #edebe9' },
  tab: { height: 32, padding: '0 12px', border: 0, background: 'transparent', fontSize: 13, lineHeight: '32px', cursor: 'pointer' },
  tabLabel: { display: 'inline-block', paddingBottom: 3, borderBottom: '2px solid' },
  panel: { height: 45, overflowX: 'auto', overflowY: 'hidden' },
  surface: { display: 'flex', alignItems: 'center', minWidth: 'max-content', height: 45, boxSizing: 'border-box', padding: '0 8px' },
  run: { display: 'flex', alignItems: 'center', gap: 2 },
  divider: { width: 1, height: 24, flex: '0 0 auto', margin: '0 6px', background: '#e0e0e0' },
  button: { appearance: 'none', display: 'inline-grid', placeItems: 'center', width: 32, height: 32, padding: 0, border: 0, borderRadius: 4, boxSizing: 'border-box' },
  colorInput: { position: 'absolute', inset: 0, opacity: 0, width: '100%', height: '100%', cursor: 'inherit' },
  formulaInput: { width: 56, height: 28, boxSizing: 'border-box', border: '1px solid #d1d1d1', borderRadius: 4, color: 'inherit', background: '#ffffff', fontSize: 12, padding: '0 6px' },
  split: { display: 'inline-flex', alignItems: 'stretch' },
  splitMain: { appearance: 'none', display: 'inline-grid', placeItems: 'center', width: 28, height: 32, padding: 0, border: 0, borderRadius: '4px 0 0 4px', boxSizing: 'border-box' },
  splitChevron: { appearance: 'none', display: 'inline-grid', placeItems: 'center', width: 16, height: 32, padding: 0, border: 0, borderRadius: '0 4px 4px 0', boxSizing: 'border-box' },
  empty: { display: 'flex', alignItems: 'center', height: 32, padding: '0 8px', color: '#424242', fontSize: 12 },
};
