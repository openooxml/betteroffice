import type { CSSProperties } from 'react';

export const notesStyles: Record<'panel' | 'label' | 'textarea', CSSProperties> = {
  panel: {
    flex: '0 0 auto',
    display: 'flex',
    flexDirection: 'column',
    gap: 4,
    padding: '8px 14px 10px',
    background: '#ffffff',
    borderTop: '1px solid #e2e8f0',
  },
  label: {
    fontSize: 11,
    fontWeight: 650,
    color: '#647087',
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
  },
  textarea: {
    width: '100%',
    height: 64,
    resize: 'vertical',
    border: '1px solid #d8dee9',
    borderRadius: 6,
    padding: '6px 8px',
    font: '13px ui-sans-serif, system-ui, sans-serif',
    color: '#172033',
    boxSizing: 'border-box',
    outline: 'none',
  },
};

export const slideActionStyles: Record<'row' | 'button', CSSProperties> = {
  row: { display: 'flex', flexWrap: 'wrap', gap: 6, paddingTop: 8 },
  button: {
    flex: '1 1 auto',
    height: 26,
    whiteSpace: 'nowrap',
    padding: '0 8px',
    border: '1px solid #cbd5e1',
    borderRadius: 5,
    background: '#ffffff',
    color: '#334155',
    font: '12px ui-sans-serif, system-ui, sans-serif',
    cursor: 'pointer',
  },
};
