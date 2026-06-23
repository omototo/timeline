import { StrictMode } from 'react';
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
import { FakeTimelineDataSource } from './ui/data-source.ts';
import type { TimelineDataSource } from './ui/data-source.ts';

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

export function renderTimelinePane({
  container = document.getElementById('root'),
  source = new FakeTimelineDataSource(),
  theme = 'light',
}: RenderTimelineOptions = {}): Root | undefined {
  if (!container) {
    return undefined;
  }

  const root = roots.get(container) ?? createRoot(container);
  roots.set(container, root);
  root.render(
    <StrictMode>
      <App source={source} theme={theme} />
    </StrictMode>,
  );

  return root;
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
  if (!office?.onReady) {
    renderTimelinePane({ container, source, theme: 'light' });
    return;
  }

  try {
    const info = await office.onReady();
    const theme = getOfficeTimelineTheme(office);

    renderTimelinePane({
      container,
      source,
      theme: isExcelHost(info, office) ? theme : 'light',
    });
  } catch {
    renderTimelinePane({ container, source, theme: 'light' });
  }
}

void bootstrapTimeline();
