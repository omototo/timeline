import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { TimelinePane } from '../src/ui/TimelinePane.tsx';
import { sampleTimeline } from '../src/ui/sample-timeline.ts';
import type { TimelineCommand, TimelineView } from '../src/ui/contract.ts';

function renderPane(view: TimelineView = sampleTimeline) {
  const dispatch = vi.fn<(cmd: TimelineCommand) => void>();
  render(<TimelinePane view={view} dispatch={dispatch} />);
  return dispatch;
}

describe('TimelinePane', () => {
  it('maps Step magnitude to histogram bar height', () => {
    renderPane();

    const small = screen.getByTestId('timeline-bar-main-0');
    const large = screen.getByTestId('timeline-bar-main-1');

    expect(Number(large.dataset.magnitude)).toBe(1000);
    expect(Number.parseFloat(large.style.height)).toBeGreaterThan(
      Number.parseFloat(small.style.height),
    );
  });

  it('emits goto when the playhead scrubber moves', () => {
    const dispatch = renderPane();

    fireEvent.change(screen.getByLabelText('Timeline playhead'), { target: { value: '2' } });

    expect(dispatch).toHaveBeenCalledWith({
      type: 'goto',
      ref: { branchId: 'main', stepIndex: 2 },
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
});
