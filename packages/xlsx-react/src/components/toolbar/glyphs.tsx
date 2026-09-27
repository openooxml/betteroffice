import type {
  BorderPreset,
  BorderStyle,
  HorizontalAlignment,
  VerticalAlignment,
} from '../Toolbar';

export function BorderGlyph({ preset }: { preset: BorderPreset }) {
  const edge = (side: BorderPreset) =>
    preset === side || preset === 'all' || preset === 'outer' ? 2 : 0.6;
  return (
    <svg
      width="20"
      height="20"
      viewBox="0 0 20 20"
      fill="none"
      stroke="currentColor"
      aria-hidden="true"
    >
      {preset !== 'none' && (
        <>
          <path d="M3 3h14" strokeWidth={edge('top')} />
          <path d="M17 3v14" strokeWidth={edge('right')} />
          <path d="M3 17h14" strokeWidth={edge('bottom')} />
          <path d="M3 3v14" strokeWidth={edge('left')} />
          {(preset === 'all' || preset === 'inner' || preset === 'horizontal') && (
            <path d="M3 10h14" strokeWidth="1.5" />
          )}
          {(preset === 'all' || preset === 'inner' || preset === 'vertical') && (
            <path d="M10 3v14" strokeWidth="1.5" />
          )}
        </>
      )}
      {preset === 'none' && <path d="M4 4h12v12H4zM3 17 17 3" strokeWidth="1.5" />}
    </svg>
  );
}

export function HorizontalAlignmentGlyph({ value }: { value: HorizontalAlignment }) {
  const x1 = value === 'left' ? 3 : value === 'center' ? 5 : 7;
  const x2 = value === 'left' ? 15 : value === 'center' ? 17 : 19;
  return (
    <svg
      width="20"
      height="20"
      viewBox="0 0 20 20"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.6"
      aria-hidden="true"
    >
      <path d="M3 4h14M3 10h14M3 16h14" />
      <path d={`M${x1} 7h${x2 - x1}M${x1} 13h${x2 - x1}`} />
    </svg>
  );
}

export function VerticalAlignmentGlyph({ value }: { value: VerticalAlignment }) {
  const y = value === 'top' ? 5 : value === 'middle' ? 10 : 15;
  return (
    <svg
      width="20"
      height="20"
      viewBox="0 0 20 20"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.6"
      aria-hidden="true"
    >
      <path d="M3 3h14M3 17h14" />
      <path d={`M6 ${y}h8`} />
      <path
        d={
          value === 'top'
            ? 'm8 8 2-3 2 3'
            : value === 'bottom'
            ? 'm8 12 2 3 2-3'
            : 'm8 7 2 3 2-3m-4 6 2-3 2 3'
        }
      />
    </svg>
  );
}

export function BorderStyleGlyph({ value }: { value: BorderStyle }) {
  return (
    <svg
      width="44"
      height="12"
      viewBox="0 0 44 12"
      fill="none"
      stroke="currentColor"
      aria-hidden="true"
    >
      {value === 'double' ? (
        <path d="M2 4h40M2 8h40" />
      ) : (
        <path
          d="M2 6h40"
          strokeWidth="2"
          strokeDasharray={value === 'dashed' ? '7 4' : value === 'dotted' ? '1 4' : undefined}
        />
      )}
    </svg>
  );
}

export function ColorSwatch({ color }: { color: string }) {
  return (
    <span
      aria-hidden="true"
      style={{ display: 'inline-block', width: 20, height: 5, borderRadius: 2, background: color }}
    />
  );
}
