/**
 * Pure cell-diff helper.
 *
 * The smallest real unit of the Value Delta class (ADR-0001): a single observed
 * change at a cell, diffed against the Shadow State. Kept deliberately tiny and
 * pure (no DOM, no Office.js) as the seed of the value-diff path.
 *
 * Note: this `CellDiff` is the minimal `(address, before, after)` triple — NOT
 * the richer `ValueDelta` from the engine interface spec (which is sheet-scoped
 * and carries lossless before/after `CellState`s). They are distinct concepts;
 * see `types.ts` for the spec `ValueDelta`.
 */

/** A single observed change at a cell. */
export interface CellDiff {
  readonly address: string;
  readonly before: string | null;
  readonly after: string | null;
}

/**
 * Diff one cell's observed value against its Shadow State value.
 *
 * Returns a {@link CellDiff} when the value changed, or `null` when it did
 * not — a no-op edit produces no Step.
 */
export function diffCell(
  address: string,
  before: string | null,
  after: string | null,
): CellDiff | null {
  if (before === after) {
    return null;
  }
  return { address, before, after };
}
