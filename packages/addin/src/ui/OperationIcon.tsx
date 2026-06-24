import type { CSSProperties } from 'react';
import type { TimelineOp } from './contract.ts';

// One line-glyph per operation (Fusion-style), drawn on a 20×20 grid with
// `currentColor` so the pane can tint it by step kind. Shape = what you did;
// the caller supplies the colour. Paths kept simple so they stay legible at the
// ~14px the timeline renders them.

const PLUS = 'M15.5 12.5v5M13 15h5';
const MINUS = 'M13 15h5';

/** The path/shape markup for each operation, as a `d`-string list + extra rects. */
const GLYPH: Record<TimelineOp, string[]> = {
  // value
  edit: ['M3 14.2 11 6.2l2.6 2.6L5.6 16.8H3z', 'M11 6.2 13.6 8.8'],
  formula: ['M11.5 4.5h-1.2c-1 0-1.5.6-1.7 1.7L6.2 16', 'M5.5 9h4.2', 'M11.5 10l4 5m0-5-4 5'],
  paste: ['M5 4.5h10v12H5z', 'M8 4.5V3.5h4v1', 'M7.5 9h5M7.5 12h5'],
  clear: ['M4.5 5h11v9h-11z', 'M6.5 7.5l7 4M13.5 7.5l-7 4'],
  // structural
  'insert-row': ['M2.5 8h11v4h-11z', PLUS],
  'delete-row': ['M2.5 8h15v4h-15z', MINUS],
  'insert-col': ['M8 2.5h4v11H8z', PLUS],
  'delete-col': ['M8 2.5h4v15H8z', MINUS],
  'insert-cells': ['M2.5 2.5h9v9h-9z', 'M7 2.5v9M2.5 7h9', PLUS],
  'delete-cells': ['M2.5 4h11v9h-11z', 'M8 4v9M2.5 8.5h11', MINUS],
  // worksheet
  'sheet-add': [
    'M4 3.5h6l3 3v7.5a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1z',
    'M10 3.5V6.5h3',
    PLUS,
  ],
  'sheet-delete': [
    'M4 3.5h6l3 3v9a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1z',
    'M10 3.5V6.5h3',
    MINUS,
  ],
  'sheet-rename': [
    'M4 3.5h6l3 3V11a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1z',
    'M10 3.5V6.5h3',
    'M5.5 16l4-4 1.6 1.6-4 4H5.5z',
  ],
  'sheet-reorder': [
    'M4 4.5h6l3 3V13a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V5.5a1 1 0 0 1 1-1z',
    'M10 4.5V7.5h3',
    'M5 17h9M5 17l1.6-1.6M14 17l-1.6-1.6',
  ],
  // reconciliation
  reconcile: ['M5 8a5 5 0 0 1 9-2.5', 'M15 12a5 5 0 0 1-9 2.5', 'M14 3v3h-3M6 17v-3h3'],
};

export interface OperationIconProps {
  readonly op: TimelineOp;
  readonly size?: number;
  readonly title?: string;
  readonly style?: CSSProperties;
}

export function OperationIcon({ op, size = 15, title, style }: OperationIconProps) {
  return (
    <svg
      viewBox="0 0 20 20"
      width={size}
      height={size}
      fill="none"
      stroke="currentColor"
      strokeWidth={1.6}
      strokeLinecap="round"
      strokeLinejoin="round"
      data-op={op}
      role="img"
      aria-label={title ?? op}
      style={style}
    >
      {title !== undefined ? <title>{title}</title> : null}
      {GLYPH[op].map((d, i) => (
        <path key={i} d={d} />
      ))}
    </svg>
  );
}
