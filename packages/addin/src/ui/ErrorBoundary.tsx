import { Component } from 'react';
import type { ErrorInfo, ReactNode } from 'react';

interface ErrorBoundaryProps {
  readonly children: ReactNode;
}

interface ErrorBoundaryState {
  readonly error: Error | null;
}

/** Readable error panel — shared by the boundary and the bootstrap failure path. */
export function TimelineErrorView({ error }: { error: Error }): ReactNode {
  return (
    <section
      aria-label="Timeline error"
      style={{
        padding: 12,
        font: '12px ui-monospace, monospace',
        color: '#fecaca',
        background: '#13161c',
        height: '100%',
        overflow: 'auto',
      }}
    >
      <strong style={{ display: 'block', marginBottom: 8, color: '#f87171' }}>
        Timeline failed to load
      </strong>
      <div style={{ marginBottom: 8 }}>{error.message}</div>
      <pre style={{ whiteSpace: 'pre-wrap', color: '#9aa3af', margin: 0 }}>{error.stack}</pre>
    </section>
  );
}

/**
 * Renders a readable error panel instead of a blank task pane when the tree
 * throws. In a 360px Office pane a white screen is indistinguishable from a
 * hang, so surfacing the message (and stack) in place is the difference between
 * a debuggable failure and a mystery.
 */
export class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
  override state: ErrorBoundaryState = { error: null };

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { error };
  }

  override componentDidCatch(error: Error, info: ErrorInfo): void {
    globalThis.console.error('[timeline] render error', error, info.componentStack);
  }

  override render(): ReactNode {
    const { error } = this.state;
    if (!error) {
      return this.props.children;
    }
    return <TimelineErrorView error={error} />;
  }
}
