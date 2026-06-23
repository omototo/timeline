import { describe, it, expect } from 'vitest';
import { diffCell, type CellDiff } from '../src/index.ts';

describe('diffCell', () => {
  it('returns a delta when the value changed', () => {
    const delta = diffCell('A1', '1', '2');
    expect(delta).toEqual<CellDiff>({ address: 'A1', before: '1', after: '2' });
  });

  it('returns null for a no-op edit (no Step)', () => {
    expect(diffCell('A1', '7', '7')).toBeNull();
  });

  it('treats a cleared cell as a change to null', () => {
    expect(diffCell('B2', 'x', null)).toEqual<CellDiff>({
      address: 'B2',
      before: 'x',
      after: null,
    });
  });

  it('treats a newly-filled empty cell as a change from null', () => {
    expect(diffCell('C3', null, 'hi')).toEqual<CellDiff>({
      address: 'C3',
      before: null,
      after: 'hi',
    });
  });
});
