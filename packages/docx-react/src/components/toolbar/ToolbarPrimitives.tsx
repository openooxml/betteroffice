import { useId, useRef } from 'react';
import type { ReactNode } from 'react';
import { Button } from '../ui/Button';
import { Tooltip } from '../ui/Tooltip';
import { cn } from '../../lib/utils';
import { useOverflowSource } from './overflowRegistry';

export interface ToolbarButtonProps {
  /** Pressed state; omit for buttons that are not toggles. */
  active?: boolean | 'mixed';
  disabled?: boolean;
  /** Why the button is disabled; announced and shown on hover. */
  description?: string;
  /** Tooltip, and the label when `ariaLabel` is omitted. */
  title?: string;
  /** Keyboard shortcut shown in the tooltip and overflow menu, such as `Ctrl+B`. */
  shortcut?: string;
  onClick?: () => void;
  children: ReactNode;
  className?: string;
  ariaLabel?: string;
  /** Label of the overflow-menu entry; defaults to `ariaLabel` or `title`. */
  overflowLabel?: string;
}

export interface ToolbarGroupProps {
  /** Accessible group name, also the heading of its overflow-menu section. */
  label?: string;
  children: ReactNode;
  className?: string;
}

function testIdFor(ariaLabel: string | undefined, title: string | undefined): string | undefined {
  const source =
    ariaLabel?.toLowerCase().replace(/\s+/g, '-') ||
    title
      ?.toLowerCase()
      .replace(/\s+/g, '-')
      // `[^()]` (not `[^)]`) so the run can't span unmatched `(` and
      // backtrack quadratically on a long parenthesis-heavy title.
      .replace(/\([^()]*\)/g, '')
      .trim();
  return source ? `toolbar-${source}` : undefined;
}

/** A toolbar button for host actions; it moves into the overflow menu at narrow widths. */
export function ToolbarButton({
  active,
  disabled = false,
  description,
  title,
  shortcut,
  onClick,
  children,
  className,
  ariaLabel,
  overflowLabel,
}: ToolbarButtonProps) {
  const wrapperRef = useRef<HTMLSpanElement>(null);
  const descriptionId = useId();
  const label = overflowLabel ?? ariaLabel ?? title;
  const described = disabled && description ? description : undefined;
  useOverflowSource(
    wrapperRef,
    label
      ? () => [
          {
            kind: 'item',
            id: descriptionId,
            label,
            shortcut,
            checked: active,
            disabled,
            description: described,
            onSelect: () => onClick?.(),
          },
        ]
      : null
  );

  const button = (
    <Button
      variant="ghost"
      size="icon-sm"
      type="button"
      className={cn(
        // Hover + active states live in editor.css (.oox-toolbar-toggle); see
        // that rule for why they're not Tailwind utilities here.
        'oox-toolbar-toggle text-muted-foreground',
        disabled && 'opacity-30 cursor-not-allowed',
        className
      )}
      data-active={active === true ? 'true' : undefined}
      onMouseDown={(event) => event.preventDefault()}
      onClick={disabled ? undefined : onClick}
      disabled={disabled && !described}
      aria-disabled={described ? true : undefined}
      aria-pressed={active === undefined ? undefined : active}
      aria-label={ariaLabel || title}
      aria-describedby={described ? descriptionId : undefined}
      data-testid={testIdFor(ariaLabel, title)}
    >
      {children}
    </Button>
  );

  const name = title ?? ariaLabel;
  const titled = name && shortcut ? `${name} (${shortcut})` : name;
  const tooltip = [titled, described].filter(Boolean).join(': ');
  return (
    <span ref={wrapperRef} className="inline-flex flex-shrink-0">
      {tooltip ? <Tooltip content={tooltip}>{button}</Tooltip> : button}
      {described && (
        <span id={descriptionId} hidden>
          {described}
        </span>
      )}
    </span>
  );
}

/** A labelled cluster of controls; it overflows whole, and only when each control has a menu entry. */
export function ToolbarGroup({ label, children, className }: ToolbarGroupProps) {
  return (
    <div
      className={cn(
        'flex flex-shrink-0 items-center gap-px px-1.5 border-r border-border/50 last:border-r-0 first:pl-0',
        className
      )}
      role="group"
      aria-label={label}
    >
      {children}
    </div>
  );
}

export function ToolbarSeparator() {
  return (
    <div
      className="w-px h-6 flex-shrink-0 bg-border mx-1.5"
      role="separator"
      aria-orientation="vertical"
    />
  );
}
