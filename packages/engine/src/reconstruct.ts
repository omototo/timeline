/**
 * Reconstruction — forward-replay-only rebuild of Shadow State (ADR-0001, Q6).
 *
 * Navigation reaches any Step N by loading the nearest keyframe ≤ N and
 * replaying the intervening deltas *forward*. Deltas are never inverted (no
 * backward-stepping logic): a `ValueDelta` stores `before` for the inspect/diff
 * UI and lossless capture (ADR-0008), but replay only ever uses `after`.
 *
 * Pure: no Office.js, DOM, or React.
 */
import { ShadowState } from './shadow-state.ts';
import type { ShadowSnapshot } from './shadow-state.ts';
import type { Delta } from './types.ts';

/**
 * Apply one {@link Delta} forward into a {@link ShadowState}, dispatching on
 * its kind. The single choke point every replay path (and live recording)
 * routes through, so forward-apply semantics live in exactly one place.
 *
 * A `reconciliation` delta is a drift-repair record (ADR-0006): it folds its
 * per-sheet cell changes and structural ops forward in the order captured.
 */
export function applyDelta(state: ShadowState, delta: Delta): void {
  switch (delta.kind) {
    case 'value':
      state.apply(delta);
      return;
    case 'structural':
      state.applyStructural(delta);
      return;
    case 'worksheet':
      state.applyWorksheet(delta);
      return;
    case 'reconciliation':
      for (const sheet of delta.perSheet) {
        for (const op of sheet.structural) {
          state.applyStructural({
            kind: 'structural',
            sheetId: sheet.sheetId,
            changeType: op.changeType,
            address: op.address,
            ...(op.shiftDirection !== undefined ? { shiftDirection: op.shiftDirection } : {}),
          });
        }
        if (sheet.cells.length > 0) {
          state.apply({ kind: 'value', sheetId: sheet.sheetId, cells: sheet.cells });
        }
      }
      return;
  }
}

/**
 * Reconstruct the Shadow State at a target stepIndex by forward-replay.
 *
 * Seeds from `keyframe` when present (the nearest keyframe ≤ target), else from
 * an empty state, then applies `deltas` — which the caller has sliced to the
 * inclusive window `(keyframeStepIndex, targetStepIndex]` — forward in order.
 * Never inverts a delta.
 */
export function reconstruct(
  keyframe: ShadowSnapshot | null,
  deltas: readonly Delta[],
): ShadowState {
  const state = keyframe === null ? new ShadowState() : ShadowState.fromSnapshot(keyframe);
  for (const delta of deltas) {
    applyDelta(state, delta);
  }
  return state;
}
