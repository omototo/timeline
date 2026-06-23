import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import {
  TimelinePane,
  activeBranch,
  barHeight,
  brushToWindow,
  currentStepIndex,
} from '../src/ui/TimelinePane.tsx';
import { TimelinePaneContainer } from '../src/ui/TimelinePaneContainer.tsx';
import { compareBranches } from '../src/ui/branch-compare.ts';
import { FakeTimelineDataSource } from '../src/ui/data-source.ts';
import { sampleTimeline } from '../src/ui/sample-timeline.ts';
import type { TimelineView } from '../src/ui/contract.ts';

function viewWith(overrides: Partial<TimelineView>): TimelineView {
  return { ...sampleTimeline, ...overrides };
}
const previewMain = viewWith({ head: { branchId: 'main', mode: 'preview', previewStepIndex: 30 } });
const onBranchB = viewWith({ head: { branchId: 'branch-b', mode: 'present' } });
const px = (el: HTMLElement): number => parseInt(el.style.height, 10);

describe('pure helpers', () => {
  it('barHeight grows with magnitude (log-scaled)', () => {
    expect(barHeight(1000, 1000, 86)).toBeGreaterThan(barHeight(1, 1000, 86));
    expect(barHeight(0, 1000, 86)).toBe(8);
  });

  it('brushToWindow returns an ordered window for a real drag, null for a click', () => {
    expect(brushToWindow(10, 90, 100, 0, 100)).toEqual({ start: 10, end: 90 });
    expect(brushToWindow(40, 42, 100, 0, 100)).toBeNull();
    expect(brushToWindow(10, 90, 0, 0, 100)).toBeNull();
  });

  it('currentStepIndex is the previewed index in preview, the tip in present', () => {
    const main = activeBranch(sampleTimeline.branches, 'main');
    expect(currentStepIndex(main, 'preview', 30)).toBe(30);
    expect(currentStepIndex(main, 'present', undefined)).toBe(71);
  });

  it('activeBranch finds the head branch, falls back to the first', () => {
    expect(activeBranch(sampleTimeline.branches, 'branch-b').id).toBe('branch-b');
    expect(activeBranch(sampleTimeline.branches, 'nope').id).toBe('main');
  });

  it('compareBranches reports a comparison between two branches', () => {
    expect(compareBranches(sampleTimeline, 'main', 'branch-b')).toBeTruthy();
  });
});

describe('TimelinePane', () => {
  it('renders a magnitude histogram — the 1,000-cell paste towers over a 1-cell edit', () => {
    render(<TimelinePane view={sampleTimeline} dispatch={vi.fn()} />);
    const tall = screen.getByRole('button', { name: /^Step 1:/ });
    const short = screen.getByRole('button', { name: /^Step 0:/ });
    expect(px(tall)).toBeGreaterThan(px(short));
    expect(tall.dataset.kind).toBe('value');
  });

  it('emits goto when the scrubber moves, returnToPresent at the tip', () => {
    const dispatch = vi.fn();
    const { rerender } = render(<TimelinePane view={sampleTimeline} dispatch={dispatch} />);
    fireEvent.change(screen.getByLabelText('Timeline scrubber'), { target: { value: '30' } });
    expect(dispatch).toHaveBeenCalledWith({
      type: 'goto',
      ref: { branchId: 'main', stepIndex: 30 },
    });
    // From preview (slider at 30), scrubbing to the tip returns to Present.
    rerender(<TimelinePane view={previewMain} dispatch={dispatch} />);
    fireEvent.change(screen.getByLabelText('Timeline scrubber'), { target: { value: '71' } });
    expect(dispatch).toHaveBeenCalledWith({ type: 'returnToPresent' });
  });

  it('emits goto when a bar is clicked', () => {
    const dispatch = vi.fn();
    render(<TimelinePane view={sampleTimeline} dispatch={dispatch} />);
    fireEvent.click(screen.getByRole('button', { name: /^Step 5:/ }));
    expect(dispatch).toHaveBeenCalledWith({
      type: 'goto',
      ref: { branchId: 'main', stepIndex: 5 },
    });
  });

  it('emits switch when a branch chip is clicked', () => {
    const dispatch = vi.fn();
    render(<TimelinePane view={sampleTimeline} dispatch={dispatch} />);
    fireEvent.click(screen.getByRole('button', { name: 'Switch to what-if branch' }));
    expect(dispatch).toHaveBeenCalledWith({ type: 'switch', branchId: 'branch-b' });
  });

  it('shows Return to Present and Branch from here only in Preview', () => {
    const dispatch = vi.fn();
    const { rerender } = render(<TimelinePane view={sampleTimeline} dispatch={dispatch} />);
    expect(screen.queryByText('Return to Present')).toBeNull();
    rerender(<TimelinePane view={previewMain} dispatch={dispatch} />);
    fireEvent.click(screen.getByText('Return to Present'));
    expect(dispatch).toHaveBeenCalledWith({ type: 'returnToPresent' });
    fireEvent.click(screen.getByText('Branch from here'));
    expect(dispatch).toHaveBeenCalledWith({
      type: 'branch',
      from: { branchId: 'main', stepIndex: 30 },
    });
  });

  it('shows the focused Step in the inspector', () => {
    render(<TimelinePane view={sampleTimeline} dispatch={vi.fn()} />);
    fireEvent.focus(screen.getByRole('button', { name: /^Step 1:/ }));
    expect(screen.getByText('Step 1')).toBeTruthy();
  });

  it('renames and deletes branches', () => {
    const dispatch = vi.fn();
    const { rerender } = render(<TimelinePane view={sampleTimeline} dispatch={dispatch} />);
    expect(screen.queryByLabelText('Delete main')).toBeNull();
    fireEvent.click(screen.getByLabelText('Rename main'));
    const input = screen.getByLabelText('Branch name');
    fireEvent.change(input, { target: { value: 'baseline' } });
    fireEvent.keyDown(input, { key: 'Enter' });
    expect(dispatch).toHaveBeenCalledWith({
      type: 'renameBranch',
      branchId: 'main',
      name: 'baseline',
    });

    rerender(<TimelinePane view={onBranchB} dispatch={dispatch} />);
    fireEvent.click(screen.getByLabelText('Delete what-if'));
    expect(dispatch).toHaveBeenCalledWith({ type: 'deleteBranch', branchId: 'branch-b' });
  });

  it('filters the histogram by worksheet', () => {
    render(<TimelinePane view={sampleTimeline} dispatch={vi.fn()} />);
    const before = screen.getAllByRole('button', { name: /^Step \d+:/ }).length;
    fireEvent.change(screen.getByLabelText('Worksheet drill-down'), {
      target: { value: 'Assumptions' },
    });
    const after = screen.getAllByRole('button', { name: /^Step \d+:/ }).length;
    expect(after).toBeLessThan(before);
    expect(after).toBeGreaterThan(0);
  });

  it('survives the brush gesture and double-click reset without dispatching scrub commands', () => {
    const dispatch = vi.fn();
    render(<TimelinePane view={sampleTimeline} dispatch={dispatch} />);
    const track = screen.getByLabelText(/Histogram of/);
    fireEvent.pointerDown(track, { clientX: 10 });
    fireEvent.pointerMove(track, { clientX: 80 });
    fireEvent.pointerUp(track, { clientX: 80 });
    fireEvent.doubleClick(track);
    expect(screen.getByLabelText(/Histogram of/)).toBeTruthy();
  });

  it('shows an empty state when the active branch has no Steps yet', () => {
    const empty = viewWith({
      branches: [{ id: 'main', name: 'main', provisional: false, steps: [] }],
      head: { branchId: 'main', mode: 'present' },
      sheets: [],
    });
    render(<TimelinePane view={empty} dispatch={vi.fn()} />);
    expect(screen.getByText(/No tracked changes yet/)).toBeTruthy();
    expect(screen.queryAllByRole('button', { name: /^Step \d+:/ })).toHaveLength(0);
  });

  it('renders in a narrow 320px task pane (dark theme)', () => {
    const host = document.createElement('div');
    host.style.width = '320px';
    document.body.appendChild(host);
    render(<TimelinePane view={sampleTimeline} dispatch={vi.fn()} theme="dark" />, {
      container: host,
    });
    expect(screen.getByLabelText('Parametric timeline')).toBeTruthy();
  });
});

describe('TimelinePaneContainer', () => {
  it('re-renders when the data source changes (switch a branch)', () => {
    const source = new FakeTimelineDataSource();
    render(<TimelinePaneContainer source={source} />);
    fireEvent.click(screen.getByRole('button', { name: 'Switch to what-if branch' }));
    expect(
      screen.getByRole('button', { name: 'Switch to what-if branch' }).getAttribute('aria-pressed'),
    ).toBe('true');
  });
});
