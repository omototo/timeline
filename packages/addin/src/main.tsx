import { StrictMode } from 'react';
import type { ReactNode } from 'react';
import type { Root } from 'react-dom/client';
import { createRoot } from 'react-dom/client';
import { App } from './App.tsx';
import {
  getGlobalOffice,
  getOfficeTimelineTheme,
  type OfficeLike,
  type OfficeReadyInfo,
} from './office-theme.ts';
import type { TimelineTheme } from './ui/contract.ts';
import { ErrorBoundary, TimelineErrorView } from './ui/ErrorBoundary.tsx';
import { createRealTimelineDataSource } from './ui/create-real-source.ts';
import { FakeTimelineDataSource } from './ui/data-source.ts';
import type { TimelineDataSource } from './ui/data-source.ts';

/** Live-source wiring can hang on an unresponsive host; bound it so we never sit blank. */
const LIVE_SOURCE_TIMEOUT_MS = 12_000;

function withTimeout<T>(promise: Promise<T>, ms: number): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const handle = setTimeout(() => {
      reject(new Error(`Timed out after ${String(ms)}ms building the live timeline source.`));
    }, ms);
    promise.then(
      (value) => {
        clearTimeout(handle);
        resolve(value);
      },
      (error: unknown) => {
        clearTimeout(handle);
        reject(error instanceof Error ? error : new Error(String(error)));
      },
    );
  });
}

const roots = new WeakMap<HTMLElement, Root>();

interface RenderTimelineOptions {
  readonly container?: HTMLElement | null | undefined;
  readonly source?: TimelineDataSource | undefined;
  readonly theme?: TimelineTheme;
}

interface BootstrapTimelineOptions {
  readonly container?: HTMLElement | null | undefined;
  readonly office?: OfficeLike | undefined;
  readonly source?: TimelineDataSource | undefined;
}

function isExcelHost(info: OfficeReadyInfo, office: OfficeLike): boolean {
  const excelHost = office.HostType?.Excel ?? 'Excel';
  return info.host === excelHost || String(info.host) === 'Excel';
}

function renderTree(container: HTMLElement, tree: ReactNode): Root {
  const root = roots.get(container) ?? createRoot(container);
  roots.set(container, root);
  root.render(<StrictMode>{tree}</StrictMode>);
  return root;
}

export function renderTimelinePane({
  container = document.getElementById('root'),
  source = new FakeTimelineDataSource(),
  theme = 'light',
}: RenderTimelineOptions = {}): Root | undefined {
  if (!container) {
    return undefined;
  }
  return renderTree(
    container,
    <ErrorBoundary>
      <App source={source} theme={theme} />
    </ErrorBoundary>,
  );
}

/**
 * Build the live, Excel-backed source and swap it in. The pane is already
 * painted (with the fake) by the time this runs, so a hang or failure here
 * downgrades to a visible error panel rather than a blank pane — never silence.
 */
async function upgradeToLiveSource(container: HTMLElement, theme: TimelineTheme): Promise<void> {
  try {
    const real = await withTimeout(createRealTimelineDataSource(), LIVE_SOURCE_TIMEOUT_MS);
    if (real) {
      renderTimelinePane({ container, source: real, theme });
    }
  } catch (error) {
    const err = error instanceof Error ? error : new Error(String(error));
    globalThis.console.error('[timeline] live source failed', err);
    renderTree(container, <TimelineErrorView error={err} />);
  }
}

export function unmountTimelinePane(container: HTMLElement | null): void {
  if (!container) {
    return;
  }

  roots.get(container)?.unmount();
  roots.delete(container);
}

export async function bootstrapTimeline({
  container = document.getElementById('root'),
  office = getGlobalOffice(),
  source,
}: BootstrapTimelineOptions = {}): Promise<void> {
  let theme: TimelineTheme = 'light';
  let inExcel = false;
  if (office?.onReady) {
    try {
      const info = await office.onReady();
      inExcel = isExcelHost(info, office);
      theme = inExcel ? getOfficeTimelineTheme(office) : 'light';
    } catch {
      // Fall through with light/non-Excel defaults; still paint the pane.
    }
  }

  // Paint immediately so the pane is never blank, then upgrade to the live,
  // Excel-backed source in the background when one is warranted.
  renderTimelinePane({ container, source: source ?? new FakeTimelineDataSource(), theme });

  if (!source && inExcel && container) {
    void upgradeToLiveSource(container, theme);
  }
}

void bootstrapTimeline();
