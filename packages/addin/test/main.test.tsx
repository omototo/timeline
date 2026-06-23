import { screen } from '@testing-library/react';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { bootstrapTimeline, renderTimelinePane, unmountTimelinePane } from '../src/main.tsx';
import type { OfficeLike } from '../src/office-theme.ts';

let rootElement: HTMLElement | undefined;

function createTaskPaneRoot(): HTMLElement {
  rootElement = document.createElement('div');
  rootElement.id = 'root';
  document.body.append(rootElement);
  return rootElement;
}

afterEach(() => {
  unmountTimelinePane(rootElement ?? null);
  rootElement = undefined;
  document.body.replaceChildren();
  vi.restoreAllMocks();
});

describe('Office bootstrap', () => {
  it('renders the fake-backed pane in plain browser development', async () => {
    const root = createTaskPaneRoot();

    await bootstrapTimeline({ container: root });

    expect(await screen.findByRole('heading', { name: 'Parametric Timeline' })).toBeInTheDocument();
    expect(root.querySelector('.timeline-pane')).toHaveAttribute('data-theme', 'light');
  });

  it('waits for Office.onReady and maps the Office dark theme', async () => {
    const root = createTaskPaneRoot();
    const office: OfficeLike = {
      HostType: { Excel: 'Excel' },
      context: {
        officeTheme: {
          bodyBackgroundColor: '#1F1F1F',
          bodyForegroundColor: '#FFFFFF',
          controlBackgroundColor: '#2B2B2B',
          controlForegroundColor: '#FFFFFF',
          isDarkTheme: true,
        },
      },
      onReady: vi.fn().mockResolvedValue({ host: 'Excel', platform: 'PC' }),
    };

    await bootstrapTimeline({ container: root, office });

    expect(office.onReady).toHaveBeenCalledTimes(1);
    expect(await screen.findByLabelText('Worksheet drill-down')).toBeInTheDocument();
    expect(root.querySelector('.timeline-pane')).toHaveAttribute('data-theme', 'dark');
  });

  it('renders directly when called by the dev server fallback', async () => {
    const root = createTaskPaneRoot();

    renderTimelinePane({ container: root, theme: 'light' });

    expect(await screen.findByText('Present')).toBeInTheDocument();
  });
});
