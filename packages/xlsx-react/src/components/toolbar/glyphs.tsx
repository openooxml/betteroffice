import type {
  BorderPreset,
  BorderStyle,
  HorizontalAlignment,
  VerticalAlignment,
} from '../Toolbar';
import { ToolbarSymbol, type ToolbarSymbolName } from '../../../../../shared/react-toolbar/icons';

const BORDER_SYMBOLS: Record<BorderPreset, ToolbarSymbolName> = {
  all: 'border_all',
  inner: 'border_inner',
  horizontal: 'border_horizontal',
  vertical: 'border_vertical',
  outer: 'border_outer',
  left: 'border_left',
  top: 'border_top',
  right: 'border_right',
  bottom: 'border_bottom',
  none: 'border_clear',
};

const HORIZONTAL_SYMBOLS: Record<HorizontalAlignment, ToolbarSymbolName> = {
  left: 'format_align_left',
  center: 'format_align_center',
  right: 'format_align_right',
};

const VERTICAL_SYMBOLS: Record<VerticalAlignment, ToolbarSymbolName> = {
  top: 'vertical_align_top',
  middle: 'vertical_align_center',
  bottom: 'vertical_align_bottom',
};

export function BorderGlyph({ preset }: { preset: BorderPreset }) {
  return <ToolbarSymbol name={BORDER_SYMBOLS[preset]} />;
}

export function HorizontalAlignmentGlyph({ value }: { value: HorizontalAlignment }) {
  return <ToolbarSymbol name={HORIZONTAL_SYMBOLS[value]} />;
}

export function VerticalAlignmentGlyph({ value }: { value: VerticalAlignment }) {
  return <ToolbarSymbol name={VERTICAL_SYMBOLS[value]} />;
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
