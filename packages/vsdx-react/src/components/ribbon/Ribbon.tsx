import { useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, ReactNode } from 'react';
import type { TFunction } from '@betteroffice/vsdx-i18n';
import { RibbonIcon } from './RibbonIcon';
import { useRibbonCommands } from './commands';
import type { RibbonCommandId } from './commands';

const tabs = ['file', 'home', 'insert', 'design', 'review', 'view', 'help'] as const;
type RibbonTab = (typeof tabs)[number];

function CommandButton({ id, icon, label, value, type = 'button' }: { id: RibbonCommandId; icon: Parameters<typeof RibbonIcon>[0]['name']; label: string; value?: string; type?: 'button' | 'color' }) {
  const command = useRibbonCommands()[id];
  if (type === 'color') return <label title={label} style={{ ...styles.button, color: command.enabled ? '#27364a' : '#9aa5b4', cursor: command.enabled ? 'pointer' : 'default', position: 'relative' }}><RibbonIcon name={icon} /><span aria-hidden="true" style={{ position: 'absolute', bottom: 3, width: 17, height: 3, background: command.value ?? '#000000' }} /><input type="color" value={command.value ?? '#000000'} disabled={!command.enabled} aria-label={label} onChange={(event) => command.run(event.target.value)} style={styles.colorInput} /></label>;
  return <button type="button" disabled={!command.enabled} aria-label={label} aria-pressed={command.active || undefined} title={label} onMouseDown={(event) => event.preventDefault()} onClick={() => command.run(value)} style={{ ...styles.button, color: command.enabled ? '#27364a' : '#9aa5b4', background: command.active ? '#dbeafe' : 'transparent', cursor: command.enabled ? 'pointer' : 'default' }}><RibbonIcon name={icon} /></button>;
}

function RibbonGroup({ label, children, empty = false }: { label: string; children?: ReactNode; empty?: boolean }) {
  return <div role="group" aria-label={label} style={styles.group}><div style={styles.controls}>{children}</div><span style={{ ...styles.groupLabel, opacity: empty ? 0.62 : 1 }}>{label}</span></div>;
}

function LineFormulaControl({ id, label, icon }: { id: 'lineWeight' | 'linePattern'; label: string; icon: Parameters<typeof RibbonIcon>[0]['name'] }) {
  const command = useRibbonCommands()[id];
  return <label title={label} style={{ display: 'inline-flex', alignItems: 'center', gap: 3, color: command.enabled ? '#27364a' : '#9aa5b4' }}><RibbonIcon name={icon} size={17} /><input key={command.value} aria-label={label} disabled={!command.enabled} defaultValue={command.value ?? ''} onBlur={(event) => { const next = event.currentTarget.value; if (next !== (command.value ?? '')) command.run(next); }} style={{ width: 52, height: 24, border: '1px solid #b9c3d0', borderRadius: 3, color: 'inherit', background: '#fff', fontSize: 11 }} /></label>;
}

function HomePanel({ t }: { t: TFunction }) {
  return <div style={styles.surface} data-testid="vsdx-ribbon-home-panel">
    <RibbonGroup label={t('ribbon.groups.history')}><CommandButton id="undo" icon="undo" label={t('ribbon.commands.undo')} /><CommandButton id="redo" icon="redo" label={t('ribbon.commands.redo')} /></RibbonGroup>
    <Divider />
    <RibbonGroup label={t('ribbon.groups.clipboard')} empty />
    <Divider />
    <RibbonGroup label={t('ribbon.groups.font')} empty />
    <Divider />
    <RibbonGroup label={t('ribbon.groups.paragraph')} empty />
    <Divider />
    <RibbonGroup label={t('ribbon.groups.insert')}><CommandButton id="addShape" icon="add" label={t('ribbon.commands.addShape')} /></RibbonGroup>
    <Divider />
    <RibbonGroup label={t('ribbon.groups.shape')}><CommandButton id="delete" icon="delete" label={t('ribbon.commands.delete')} /><CommandButton id="fillColor" icon="fill" label={t('ribbon.commands.fillColor')} type="color" /><CommandButton id="lineColor" icon="line" label={t('ribbon.commands.lineColor')} type="color" /><LineFormulaControl id="lineWeight" icon="weight" label={t('ribbon.commands.lineWeight')} /><LineFormulaControl id="linePattern" icon="pattern" label={t('ribbon.commands.linePattern')} /></RibbonGroup>
    <Divider />
    <RibbonGroup label={t('ribbon.groups.arrange')}><CommandButton id="bringToFront" icon="front" label={t('ribbon.commands.bringToFront')} /><CommandButton id="bringForward" icon="forward" label={t('ribbon.commands.bringForward')} /><CommandButton id="sendBackward" icon="backward" label={t('ribbon.commands.sendBackward')} /><CommandButton id="sendToBack" icon="back" label={t('ribbon.commands.sendToBack')} /><CommandButton id="rotateLeft" icon="rotateLeft" label={t('ribbon.commands.rotateLeft')} /><CommandButton id="rotateRight" icon="rotateRight" label={t('ribbon.commands.rotateRight')} /><CommandButton id="flipHorizontal" icon="flipHorizontal" label={t('ribbon.commands.flipHorizontal')} /><CommandButton id="flipVertical" icon="flipVertical" label={t('ribbon.commands.flipVertical')} /></RibbonGroup>
  </div>;
}

function Divider() { return <div role="separator" style={styles.divider} />; }

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
  return <section aria-label={t('ribbon.label')} style={styles.root}>
    <div role="tablist" aria-label={t('ribbon.tabsLabel')} style={styles.tabs}>{tabs.map((tab, index) => <button ref={(node) => { tabRefs.current[index] = node; }} key={tab} id={`vsdx-ribbon-tab-${tab}`} type="button" role="tab" aria-selected={active === tab} aria-controls={`vsdx-ribbon-panel-${tab}`} tabIndex={active === tab ? 0 : -1} onClick={() => select(tab)} onKeyDown={(event) => onKeyDown(event, index)} style={{ ...styles.tab, borderBottomColor: active === tab ? '#2563eb' : 'transparent', color: active === tab ? '#174ea6' : '#425466' }}>{t(`ribbon.tabs.${tab}`)}</button>)}</div>
    <div id={`vsdx-ribbon-panel-${active}`} role="tabpanel" aria-labelledby={`vsdx-ribbon-tab-${active}`} style={styles.panel}>{active === 'home' ? <HomePanel t={t} /> : active === 'file' ? <div style={styles.surface}><RibbonGroup label={t('ribbon.groups.file')}><CommandButton id="download" icon="download" label={t('ribbon.commands.download')} /></RibbonGroup></div> : <div style={styles.surface} />}</div>
  </section>;
}

const styles: Record<string, CSSProperties> = {
  root: { flex: '0 0 auto', background: '#fff', borderBottom: '1px solid #cfd8e3', fontFamily: 'ui-sans-serif, system-ui, sans-serif' },
  tabs: { display: 'flex', alignItems: 'end', height: 34, padding: '0 12px', gap: 2, borderBottom: '1px solid #e4e9f0' },
  tab: { height: 34, padding: '0 12px', border: 0, borderBottom: '2px solid', background: 'transparent', fontSize: 13, fontWeight: 500, cursor: 'pointer' },
  panel: { minHeight: 70, overflowX: 'auto' },
  surface: { display: 'flex', alignItems: 'stretch', minWidth: 'max-content', height: 70, padding: '3px 10px' },
  group: { display: 'flex', flexDirection: 'column', justifyContent: 'space-between', minWidth: 38, padding: '2px 5px 1px', gap: 3 },
  controls: { display: 'flex', alignItems: 'center', minHeight: 35, gap: 1 },
  groupLabel: { alignSelf: 'center', color: '#536579', fontSize: 10, lineHeight: 1, whiteSpace: 'nowrap' },
  divider: { width: 1, height: 57, alignSelf: 'center', background: '#d7dee8' },
  button: { appearance: 'none', display: 'inline-grid', placeItems: 'center', width: 29, height: 30, padding: 0, border: 0, borderRadius: 3, boxSizing: 'border-box' },
  colorInput: { position: 'absolute', inset: 0, opacity: 0, width: '100%', height: '100%', cursor: 'inherit' },
};
