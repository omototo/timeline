import { diffCell, type CellDiff } from '@timeline/engine';

function formatDelta(delta: CellDiff | null): string {
  if (!delta) {
    return 'No change';
  }
  return `${delta.address}: ${delta.before ?? '∅'} → ${delta.after ?? '∅'}`;
}

export interface AppProps {
  /** The delta to display. Defaults to a sample diff so the engine wiring is exercised. */
  readonly delta?: CellDiff | null;
}

/**
 * Trivial task-pane component. Exercises the engine import across the
 * package boundary so the addin↔engine wiring is real, not just declared.
 */
export function App({ delta = diffCell('A1', '1', '2') }: AppProps = {}): React.JSX.Element {
  return (
    <main>
      <h1>Timeline</h1>
      <p>{formatDelta(delta)}</p>
    </main>
  );
}
