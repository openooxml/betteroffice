import { useCallback, useEffect, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, ReactNode } from 'react';
import type { TFunction } from '@betteroffice/vsdx-i18n';
import { CommandMenu } from './CommandMenu';
import { RibbonIcon } from './RibbonIcon';
import { LINE_PATTERN_VALUES, parseLinePatternInput, parseLineWeightInput, useRibbonCommands } from './commands';
import type { RibbonCommandId } from './commands';

const baseTabs = ['file', 'home', 'insert', 'view'] as const;
type RibbonTab = (typeof baseTabs)[number] | 'shape';
export type ViewToggleKey = 'grid' | 'snap' | 'rulers';
export interface RibbonViewState { grid: boolean; snap: boolean; rulers: boolean; }

type IconName = Parameters<typeof RibbonIcon>[0]['name'];

function CommandButton({ id, icon, label }: { id: RibbonCommandId; icon: IconName; label: string }) {
  const command = useRibbonCommands()[id];
  return <button type="button" disabled={!command.enabled} aria-label={label} aria-pressed={command.active || undefined} title={label} data-command-id={id} onMouseDown={(event) => event.preventDefault()} onClick={() => command.run()} className="vsdx-cmd-btn" style={{ ...styles.button, color: command.enabled ? '#242424' : '#b4b4b4', background: command.active ? '#ebf3fc' : 'transparent', cursor: command.enabled ? 'pointer' : 'default' }}><RibbonIcon name={icon} size={20} /></button>;
}

function ColorButton({ id, icon, label }: { id: RibbonCommandId; icon: IconName; label: string }) {
  const command = useRibbonCommands()[id];
  return <label title={label} className="vsdx-cmd-btn" style={{ ...styles.button, color: command.enabled ? '#242424' : '#b4b4b4', cursor: command.enabled ? 'pointer' : 'default', position: 'relative' }}><RibbonIcon name={icon} size={20} /><span aria-hidden="true" style={{ position: 'absolute', bottom: 4, width: 16, height: 3, borderRadius: 1, background: command.value ?? '#000000' }} /><input type="color" value={command.value ?? '#000000'} disabled={!command.enabled} aria-label={label} title={label} data-command-id={id} onChange={(event) => command.run(event.target.value)} style={styles.colorInput} /></label>;
}

/** Validated weight field. Commits on blur or Enter, reverts on invalid blur. */
function LineWeightControl({ label, icon, t }: { label: string; icon: IconName; t: TFunction }) {
  const command = useRibbonCommands()['lineWeight'];
  const current = command.value ?? '';
  const [draft, setDraft] = useState(current);
  const [invalid, setInvalid] = useState(false);
  useEffect(() => { setDraft(current); setInvalid(false); }, [current]);
  const hint = `${label}: ${t('ribbon.lineWeightHint')}`;
  const commit = (raw: string): boolean => {
    if (raw.trim() === current) { setInvalid(false); return true; }
    if (parseLineWeightInput(raw) === null) return false;
    command.run(raw);
    return true;
  };
  return <label title={invalid ? hint : label} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, color: command.enabled ? '#242424' : '#b4b4b4' }}><RibbonIcon name={icon} size={18} /><input
    aria-label={label}
    title={invalid ? hint : label}
    data-command-id="lineWeight"
    disabled={!command.enabled}
    value={draft}
    aria-invalid={invalid || undefined}
    onChange={(event) => { setDraft(event.target.value); if (invalid) setInvalid(false); }}
    onBlur={(event) => { if (!commit(event.currentTarget.value)) { setDraft(current); setInvalid(false); } }}
    onKeyDown={(event) => {
      if (event.key === 'Enter') { event.preventDefault(); if (!commit(event.currentTarget.value)) setInvalid(true); }
      else if (event.key === 'Escape') { event.preventDefault(); setDraft(current); setInvalid(false); event.currentTarget.blur(); }
    }}
    style={invalid ? { ...styles.formulaInput, border: '1px solid #c42b1c', outline: '1px solid #c42b1c' } : styles.formulaInput}
  /></label>;
}

function linePatternLabel(t: TFunction, value: string): string {
  if (value === '0') return t('ribbon.linePattern.none');
  if (value === '1') return t('ribbon.linePattern.solid');
  return t('ribbon.linePattern.indexed', { index: value });
}

/** Pattern picker. A select cannot hold an out-of-range value, so 999 is unreachable. */
function LinePatternControl({ label, icon, t }: { label: string; icon: IconName; t: TFunction }) {
  const command = useRibbonCommands()['linePattern'];
  const selected = parseLinePatternInput(command.value ?? '');
  const custom = selected === null && command.enabled && command.value ? command.value : null;
  const title = custom ? `${label}: ${custom}` : label;
  return <label title={title} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, color: command.enabled ? '#242424' : '#b4b4b4' }}><RibbonIcon name={icon} size={18} /><select
    aria-label={label}
    title={title}
    data-command-id="linePattern"
    disabled={!command.enabled}
    value={selected ?? ''}
    onChange={(event) => { if (event.target.value !== '') command.run(event.target.value); }}
    onKeyDown={(event) => { if (event.key === 'Escape') event.currentTarget.blur(); }}
    style={styles.formulaInput}
  >
    {selected === null ? <option value="">{custom ? t('ribbon.linePattern.custom', { value: custom }) : t('ribbon.linePattern.unset')}</option> : null}
    {LINE_PATTERN_VALUES.map((value) => <option key={value} value={value}>{linePatternLabel(t, value)}</option>)}
  </select></label>;
}

function RibbonRun({ label, children }: { label: string; children?: ReactNode }) {
  return <div role="group" aria-label={label} style={styles.run}>{children}</div>;
}

function Divider() { return <div role="separator" aria-orientation="vertical" style={styles.divider} />; }

export interface RibbonConnectorToggle { active: boolean; disabled: boolean; onToggle: () => void; }

function ConnectorToggle({ t, connector }: { t: TFunction; connector: RibbonConnectorToggle }) {
  const label = `${t('ribbon.commands.connector')} (Alt+3)`;
  return <button type="button" disabled={connector.disabled} aria-label={label} aria-pressed={connector.active} title={label} onMouseDown={(event) => event.preventDefault()} onClick={() => connector.onToggle()} className="vsdx-cmd-btn" style={{ ...styles.button, color: connector.disabled ? '#b4b4b4' : '#242424', background: connector.active ? '#ebf3fc' : 'transparent', cursor: connector.disabled ? 'default' : 'pointer' }}><RibbonIcon name="connector" size={20} /></button>;
}

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
      <button type="button" disabled={!current.enabled} aria-label={label(fallback.id)} title={label(fallback.id)} data-command-id={fallback.id} onMouseDown={(event) => event.preventDefault()} onClick={() => current.run()} className="vsdx-cmd-btn vsdx-split-main" style={{ ...styles.splitMain, color: current.enabled ? '#242424' : '#b4b4b4', cursor: current.enabled ? 'pointer' : 'default' }}><RibbonIcon name={fallback.icon} size={20} /></button>
      <button ref={triggerRef} type="button" disabled={!anyEnabled} aria-label={`${label(fallback.id)} options`} title={`${label(fallback.id)} options`} aria-haspopup="menu" aria-expanded={open} data-split-toggle={fallback.id} onMouseDown={(event) => event.preventDefault()} onClick={() => { if (!anyEnabled) return; if (open) close(); else openMenu('first'); }} onKeyDown={onTriggerKeyDown} className="vsdx-cmd-btn" style={{ ...styles.splitChevron, color: anyEnabled ? '#242424' : '#b4b4b4', background: open ? '#ebebeb' : 'transparent', cursor: anyEnabled ? 'pointer' : 'default' }}><svg width={10} height={10} viewBox="0 0 10 10" aria-hidden="true" focusable="false" style={{ display: 'block' }}><path d="m2 3.5 3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" /></svg></button>
      {open && (
        <CommandMenu menuLabel={`${label(fallback.id)} options`} entries={entries} position={pos} anchorRef={triggerRef} initialFocus={intent} label={label} onClose={close} onCloseAndFocus={closeAndFocus} />
      )}
    </span>
  );
}

function ArrangeRun({ t }: { t: TFunction }) {
  const label = (id: RibbonCommandId) => t(`ribbon.commands.${id}`);
  return <RibbonRun label={t('ribbon.groups.arrange')}>
    <RibbonSplitButton defaultId="bringToFront" defaultIcon="front" label={label} entries={[{ id: 'bringToFront', icon: 'front' }, { id: 'bringForward', icon: 'forward' }, { id: 'sendBackward', icon: 'backward' }, { id: 'sendToBack', icon: 'back' }]} />
    <RibbonSplitButton defaultId="rotateRight" defaultIcon="rotateRight" label={label} entries={[{ id: 'rotateRight', icon: 'rotateRight' }, { id: 'rotateLeft', icon: 'rotateLeft' }, { id: 'flipHorizontal', icon: 'flipHorizontal' }, { id: 'flipVertical', icon: 'flipVertical' }]} />
  </RibbonRun>;
}

function FormatRun({ t }: { t: TFunction }) {
  const label = (id: RibbonCommandId) => t(`ribbon.commands.${id}`);
  return <RibbonRun label={t('ribbon.groups.shape')}><ColorButton id="fillColor" icon="fill" label={label('fillColor')} /><ColorButton id="lineColor" icon="line" label={label('lineColor')} /><LineWeightControl icon="weight" label={label('lineWeight')} t={t} /><LinePatternControl icon="pattern" label={label('linePattern')} t={t} /></RibbonRun>;
}

function HomePanel({ t }: { t: TFunction }) {
  const label = (id: RibbonCommandId) => t(`ribbon.commands.${id}`);
  return <div style={styles.surface} data-testid="vsdx-ribbon-home-panel">
    <RibbonRun label={t('ribbon.groups.history')}><CommandButton id="undo" icon="undo" label={label('undo')} /><CommandButton id="redo" icon="redo" label={label('redo')} /></RibbonRun>
    <Divider />
    <RibbonRun label={t('ribbon.groups.clipboard')}><CommandButton id="cut" icon="cut" label={label('cut')} /><CommandButton id="copy" icon="copy" label={label('copy')} /><CommandButton id="paste" icon="paste" label={label('paste')} /><CommandButton id="duplicate" icon="duplicate" label={label('duplicate')} /></RibbonRun>
    <Divider />
    <RibbonRun label={t('ribbon.groups.insert')}><CommandButton id="delete" icon="delete" label={label('delete')} /><CommandButton id="addShape" icon="add" label={label('addShape')} /></RibbonRun>
    <Divider />
    <FormatRun t={t} />
    <Divider />
    <ArrangeRun t={t} />
    <Divider />
    <ShowRun t={t} />
  </div>;
}

function ShowRun({ t }: { t: TFunction }) {
  return <RibbonRun label={t('ribbon.groups.view')}><CommandButton id="pageBreaks" icon="pageBreaks" label={t('ribbon.commands.pageBreaks')} /></RibbonRun>;
}

function ShapePanel({ t }: { t: TFunction }) {
  return <div style={styles.surface} data-testid="vsdx-ribbon-shape-panel">
    <FormatRun t={t} />
    <Divider />
    <ArrangeRun t={t} />
  </div>;
}

function ViewToggleButton({ toggleKey, icon, label, pressed, onToggle }: { toggleKey: ViewToggleKey; icon: IconName; label: string; pressed: boolean; onToggle: (key: ViewToggleKey) => void }) {
  return <button type="button" aria-label={label} aria-pressed={pressed} title={label} data-view-toggle={toggleKey} onMouseDown={(event) => event.preventDefault()} onClick={() => onToggle(toggleKey)} className="vsdx-cmd-btn" style={{ ...styles.button, color: '#242424', background: pressed ? '#ebf3fc' : 'transparent', cursor: 'pointer' }}><RibbonIcon name={icon} size={20} /></button>;
}

function ViewPanel({ t, view, onToggleView }: { t: TFunction; view: RibbonViewState; onToggleView: (key: ViewToggleKey) => void }) {
  return <div style={styles.surface} data-testid="vsdx-ribbon-view-panel">
    <RibbonRun label={t('ribbon.groups.view')}>
      <ViewToggleButton toggleKey="grid" icon="grid" label={t('ribbon.commands.showGrid')} pressed={view.grid} onToggle={onToggleView} />
      <ViewToggleButton toggleKey="snap" icon="snap" label={t('ribbon.commands.snapObjects')} pressed={view.snap} onToggle={onToggleView} />
      <ViewToggleButton toggleKey="rulers" icon="ruler" label={t('ribbon.commands.showRulers')} pressed={view.rulers} onToggle={onToggleView} />
    </RibbonRun>
  </div>;
}

export function Ribbon({ t, hasSelection = false, connector, view = { grid: true, snap: true, rulers: true }, onToggleView = () => {} }: { t: TFunction; hasSelection?: boolean; connector?: RibbonConnectorToggle; view?: RibbonViewState; onToggleView?: (key: ViewToggleKey) => void }) {
  const visibleTabs: readonly RibbonTab[] = hasSelection ? [...baseTabs, 'shape'] : baseTabs;
  const [active, setActive] = useState<RibbonTab>(hasSelection ? 'shape' : 'home');
  const [selectionShown, setSelectionShown] = useState(hasSelection);
  const tabRefs = useRef<Array<HTMLButtonElement | null>>([]);
  if (selectionShown !== hasSelection) {
    setSelectionShown(hasSelection);
    if (hasSelection) setActive('shape');
    else if (active === 'shape') setActive('home');
  }
  const select = (next: RibbonTab) => setActive(next);
  const onKeyDown = (event: KeyboardEvent<HTMLButtonElement>, index: number) => {
    let next = index;
    if (event.key === 'ArrowRight') next = (index + 1) % visibleTabs.length;
    else if (event.key === 'ArrowLeft') next = (index + visibleTabs.length - 1) % visibleTabs.length;
    else if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = visibleTabs.length - 1;
    else return;
    event.preventDefault(); select(visibleTabs[next]); tabRefs.current[next]?.focus();
  };
  return <section aria-label={t('ribbon.label')} className="vsdx-ribbon-flat" style={styles.root}>
    <style>{'[data-command-id]:focus-visible,.vsdx-cmd-btn:focus-visible,.vsdx-ribbon-flat [role="tab"]:focus-visible{outline:2px solid #0f6cbd;outline-offset:1px}.vsdx-cmd-btn:not(:disabled):hover{background-color:#f5f5f5}.vsdx-cmd-btn:not(:disabled):active{background-color:#ebebeb}'}</style>
    <div role="tablist" aria-label={t('ribbon.tabsLabel')} style={styles.tabs}>{visibleTabs.map((tab, index) => {
      const selected = active === tab;
      return <button ref={(node) => { tabRefs.current[index] = node; }} key={tab} id={`vsdx-ribbon-tab-${tab}`} type="button" role="tab" aria-selected={selected} aria-controls={`vsdx-ribbon-panel-${tab}`} tabIndex={selected ? 0 : -1} onClick={() => select(tab)} onKeyDown={(event) => onKeyDown(event, index)} style={styles.tab}><span style={{ ...styles.tabLabel, borderBottomColor: selected ? '#0f6cbd' : 'transparent', color: '#242424', fontWeight: selected ? 600 : 400 }}>{t(`ribbon.tabs.${tab}`)}</span></button>;
    })}</div>
    <div id={`vsdx-ribbon-panel-${active}`} role="tabpanel" aria-labelledby={`vsdx-ribbon-tab-${active}`} style={styles.panel}>
      {active === 'home' ? <HomePanel t={t} /> : active === 'shape' ? <ShapePanel t={t} /> : active === 'view' ? <ViewPanel t={t} view={view} onToggleView={onToggleView} /> : active === 'file' ? <div style={styles.surface}><RibbonRun label={t('ribbon.groups.file')}><CommandButton id="download" icon="download" label={t('ribbon.commands.download')} /></RibbonRun></div> : <div style={styles.surface}><RibbonRun label={t('ribbon.groups.insert')}><CommandButton id="addShape" icon="add" label={t('ribbon.commands.addShape')} /></RibbonRun>{connector ? <><Divider /><RibbonRun label={t('ribbon.groups.connector')}><ConnectorToggle t={t} connector={connector} /></RibbonRun></> : null}</div>}
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
};
