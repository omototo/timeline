import { act, fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { TimelinePane } from '../src/ui/TimelinePane.tsx';
import { TimelinePaneContainer } from '../src/ui/TimelinePaneContainer.tsx';
import { compareBranches } from '../src/ui/branch-compare.ts';
import { FakeTimelineDataSource } from '../src/ui/data-source.ts';
import { sampleTimeline } from '../src/ui/sample-timeline.ts';
import type { TimelineCommand, TimelineView } from '../src/ui/contract.ts';

function renderPane(view: TimelineView = sampleTimeline) {
  const dispatch = vi.fn<(cmd: TimelineCommand) => void>();
  render(<TimelinePane view={view} dispatch={dispatch} />);
  return dispatch;
}

describe('TimelinePane', () => {
  it('computes branch compare divergence from TimelineView', () => {
    const comparison = compareBranches(sampleTimeline, 'main', 'branch-b');

    expect(comparison?.fork).toEqual({ branchId: 'main', stepIndex: 14 });
    expect(comparison?.firstDifferentStepIndex).toBe(0);
    expect(comparison?.leftStats.stepCount).toBe(72);
    expect(comparison?.rightStats.stepCount).toBe(28);
    expect(comparison?.rightStats.totalMagnitude).toBeGreaterThan(0);
  });

  it('maps Step magnitude to histogram bar height', () => {
    renderPane();

    const small = screen.getByTestId('timeline-bar-main-0');
    const large = screen.getByTestId('timeline-bar-main-1');

    expect(Number(large.dataset.magnitude)).toBe(1000);
    expect(Number.parseFloat(large.style.height)).toBeGreaterThan(
      Number.parseFloat(small.style.height),
    );
  });

  it('windows the temporal axis when zoomed and panned', () => {
    renderPane();

    fireEvent.change(screen.getByLabelText('Temporal density'), { target: { value: '2' } });
    fireEvent.change(screen.getByLabelText('Temporal window'), { target: { value: '10' } });

    expect(screen.queryByTestId('timeline-bar-main-0')).not.toBeInTheDocument();
    expect(screen.getByTestId('timeline-bar-main-10')).toBeInTheDocument();
    expect(screen.getByText('Steps 11-22 of 72')).toBeInTheDocument();
  });

  it('emits goto when the playhead scrubber moves', () => {
    const dispatch = renderPane();

    fireEvent.change(screen.getByLabelText('Timeline playhead'), { target: { value: '2' } });

    expect(dispatch).toHaveBeenCalledWith({
      type: 'goto',
      ref: { branchId: 'main', stepIndex: 2 },
    } satisfies TimelineCommand);
  });

  it('moves the playhead with keyboard arrows on Step bars', () => {
    const dispatch = renderPane();

    fireEvent.keyDown(screen.getByTestId('timeline-bar-main-1'), { key: 'ArrowRight' });

    expect(dispatch).toHaveBeenCalledWith({
      type: 'goto',
      ref: { branchId: 'main', stepIndex: 2 },
    } satisfies TimelineCommand);
    expect(screen.getByRole('tooltip')).toHaveTextContent('main Step 2');
  });

  it('shows a tooltip when a bar receives keyboard focus', () => {
    renderPane();

    fireEvent.focus(screen.getByTestId('timeline-bar-main-1'));

    expect(screen.getByRole('tooltip')).toHaveTextContent('magnitude 1,000');
  });

  it('shows the selected Step in the inspector', () => {
    renderPane();

    fireEvent.click(screen.getByTestId('timeline-bar-main-1'));

    expect(screen.getByRole('heading', { name: 'Step 1' })).toBeInTheDocument();
    expect(screen.getByText('paste 1,000 rows')).toBeInTheDocument();
    expect(screen.getByText('Pending inspectStep')).toBeInTheDocument();
  });

  it('renders branch compare details and future diff slot', () => {
    renderPane();

    expect(screen.getByLabelText('Branch compare')).toHaveTextContent('main Step 14');
    expect(
      screen.getByText('Cell-level diff pending engine state integration.'),
    ).toBeInTheDocument();
  });

  it('emits branch from the previewed Step', () => {
    const previewView: TimelineView = {
      ...sampleTimeline,
      head: { branchId: 'main', mode: 'preview', previewStepIndex: 2 },
    };
    const dispatch = renderPane(previewView);

    fireEvent.click(screen.getByRole('button', { name: 'Branch from here' }));

    expect(dispatch).toHaveBeenCalledWith({
      type: 'branch',
      from: { branchId: 'main', stepIndex: 2 },
    } satisfies TimelineCommand);
  });

  it('emits switch when a branch track is clicked', () => {
    const dispatch = renderPane();

    fireEvent.click(screen.getByRole('button', { name: 'Switch to what-if branch' }));

    expect(dispatch).toHaveBeenCalledWith({
      type: 'switch',
      branchId: 'branch-b',
    } satisfies TimelineCommand);
  });

  it('shows Return to Present in Preview mode', () => {
    const previewView: TimelineView = {
      ...sampleTimeline,
      head: { branchId: 'main', mode: 'preview', previewStepIndex: 2 },
    };
    const dispatch = renderPane(previewView);

    fireEvent.click(screen.getByRole('button', { name: 'Return to Present' }));

    expect(dispatch).toHaveBeenCalledWith({ type: 'returnToPresent' } satisfies TimelineCommand);
  });

  it('renames and deletes branches in the fake data source', () => {
    const source = new FakeTimelineDataSource();

    source.dispatch({ type: 'renameBranch', branchId: 'branch-b', name: 'Upside case' });
    expect(source.getView().branches.find((branch) => branch.id === 'branch-b')?.name).toBe(
      'Upside case',
    );

    source.dispatch({ type: 'switch', branchId: 'branch-b' });
    source.dispatch({ type: 'deleteBranch', branchId: 'branch-b' });

    expect(source.getView().branches.some((branch) => branch.id === 'branch-b')).toBe(false);
    expect(source.getView().head).toEqual({ branchId: 'main', mode: 'present' });
  });

  it('re-renders when the data source changes', () => {
    const source = new FakeTimelineDataSource();
    render(<TimelinePaneContainer source={source} />);

    expect(screen.getByText('Present')).toBeInTheDocument();

    act(() => {
      source.dispatch({ type: 'goto', ref: { branchId: 'main', stepIndex: 4 } });
    });

    expect(screen.getByText('Preview')).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Return to Present' })).toBeInTheDocument();
  });

  it('renders in a narrow 320px task-pane container', () => {
    const { container } = render(
      <div style={{ width: '320px' }}>
        <TimelinePane view={sampleTimeline} dispatch={vi.fn()} />
      </div>,
    );

    expect(screen.getByRole('heading', { name: 'Parametric Timeline' })).toBeInTheDocument();
    expect(screen.getByLabelText('Branch compare')).toBeInTheDocument();
    expect(container.querySelector('.timeline-pane__histogram')).toBeInTheDocument();
  });
});
