import { render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { OperationIcon } from '../src/ui/OperationIcon.tsx';
import type { TimelineOp } from '../src/ui/contract.ts';

const ALL_OPS: TimelineOp[] = [
  'edit',
  'formula',
  'paste',
  'clear',
  'insert-row',
  'delete-row',
  'insert-col',
  'delete-col',
  'insert-cells',
  'delete-cells',
  'sheet-add',
  'sheet-delete',
  'sheet-rename',
  'sheet-reorder',
  'reconcile',
];

describe('OperationIcon', () => {
  it('renders an accessible glyph for every operation', () => {
    for (const op of ALL_OPS) {
      const { unmount } = render(<OperationIcon op={op} />);
      // Defaults its accessible name to the op id, so every op renders a glyph.
      expect(screen.getByRole('img', { name: op })).toBeTruthy();
      unmount();
    }
  });

  it('exposes an accessible label and uses currentColor so the pane can tint it', () => {
    render(<OperationIcon op="insert-row" title="Inserted row" />);
    const img = screen.getByRole('img', { name: 'Inserted row' });
    expect(img.getAttribute('stroke')).toBe('currentColor');
  });
});
