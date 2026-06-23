import { useMemo, useState } from 'react';
import type { ChangeEvent, CSSProperties, KeyboardEvent } from 'react';
import { compareBranches } from './branch-compare.ts';
import type {
  BranchId,
  StepKind,
  StepRef,
  TimelineBranch,
  TimelineCommand,
  TimelinePaneProps,
  TimelineStep,
} from './contract.ts';

const ALL_SHEETS = '__timeline_all_sheets__';

const TEMPORAL_DENSITIES = [
  { label: 'Overview', windowSize: 48, barWidth: 12 },
  { label: 'Balanced', windowSize: 24, barWidth: 22 },
  { label: 'Detail', windowSize: 12, barWidth: 38 },
] as const;

const KIND_META: Record<StepKind, { label: string; color: string; accent: string }> = {
  value: { label: 'Value', color: '#287a63', accent: '#dff2ea' },
  structural: { label: 'Structural', color: '#3f6db6', accent: '#e3ebfb' },
  worksheet: { label: 'Worksheet', color: '#a46b10', accent: '#f8ead0' },
  reconciliation: { label: 'Reconciliation', color: '#b24d5e', accent: '#f8e1e5' },
};

interface VisibleBranch {
  branch: TimelineBranch;
  steps: TimelineStep[];
}

interface TimelinePaneStyle extends CSSProperties {
  '--timeline-bar-width'?: string;
}

interface BarStyle extends CSSProperties {
  '--bar-color'?: string;
  '--bar-accent'?: string;
}

interface MiniBarStyle extends CSSProperties {
  '--mini-color'?: string;
  '--mini-height'?: string;
}

function branchTitle(branch: TimelineBranch): string {
  return branch.name ?? branch.id;
}

function stepRef(branchId: BranchId, step: TimelineStep): StepRef {
  return { branchId, stepIndex: step.index };
}

function refsEqual(left: StepRef | null, right: StepRef | null): boolean {
  if (!left || !right) return false;
  return left.branchId === right.branchId && left.stepIndex === right.stepIndex;
}

function activeBranch(view: TimelinePaneProps['view']): TimelineBranch | undefined {
  return view.branches.find((branch) => branch.id === view.head.branchId) ?? view.branches[0];
}

function presentRef(branch: TimelineBranch | undefined): StepRef | null {
  const tip = branch?.steps.at(-1);
  if (!branch || !tip) return null;
  return { branchId: branch.id, stepIndex: tip.index };
}

function headRef(view: TimelinePaneProps['view']): StepRef | null {
  if (view.head.mode === 'preview') {
    return {
      branchId: view.head.branchId,
      stepIndex: view.head.previewStepIndex ?? 0,
    };
  }
  return presentRef(activeBranch(view));
}

function stepForRef(view: TimelinePaneProps['view'], ref: StepRef | null): TimelineStep | null {
  if (!ref) return null;
  const branch = view.branches.find((candidate) => candidate.id === ref.branchId);
  return branch?.steps.find((step) => step.index === ref.stepIndex) ?? null;
}

function visibleBranches(view: TimelinePaneProps['view'], sheetFilter: string): VisibleBranch[] {
  return view.branches.map((branch) => ({
    branch,
    steps:
      sheetFilter === ALL_SHEETS
        ? branch.steps
        : branch.steps.filter((step) => step.sheetId === sheetFilter),
  }));
}

function maxMagnitude(branches: VisibleBranch[]): number {
  return Math.max(
    1,
    ...branches.flatMap((branch) => branch.steps.map((step) => Math.max(0, step.magnitude))),
  );
}

function maxStepCount(branches: VisibleBranch[]): number {
  return Math.max(0, ...branches.map((branch) => branch.steps.length));
}

function barHeight(step: TimelineStep, max: number): number {
  const ratio = Math.max(0.04, step.magnitude / max);
  return Math.round(148 * ratio);
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function windowLabel(start: number, end: number, total: number): string {
  if (total === 0) return 'No Steps';
  return `Steps ${String(start + 1)}-${String(end)} of ${String(total)}`;
}

function visibleWindow(steps: TimelineStep[], start: number, size: number): TimelineStep[] {
  return steps.slice(start, start + size);
}

function renameDraftValue(branch: TimelineBranch, drafts: Record<BranchId, string>): string {
  return drafts[branch.id] ?? branchTitle(branch);
}

function branchTone(branch: TimelineBranch): string {
  if (branch.id === 'main') return 'Base';
  return branch.provisional ? 'Provisional' : 'Branch';
}

function tooltipText(branch: TimelineBranch, step: TimelineStep): string {
  return `${branchTitle(branch)} Step ${String(step.index)}: ${KIND_META[step.kind].label}, magnitude ${step.magnitude.toLocaleString()}, ${step.sheetId}. ${step.label ?? 'No label'}`;
}

function compareSelectValue(
  branches: TimelineBranch[],
  preferred: BranchId,
  fallbackIndex: number,
): BranchId {
  if (branches.some((branch) => branch.id === preferred)) return preferred;
  return branches[fallbackIndex]?.id ?? branches[0]?.id ?? '';
}

export function TimelinePane({
  view,
  dispatch,
  theme = 'light',
}: TimelinePaneProps): React.JSX.Element {
  const [sheetFilter, setSheetFilter] = useState<string>(ALL_SHEETS);
  const [densityIndex, setDensityIndex] = useState(1);
  const [windowStart, setWindowStart] = useState(0);
  const [selected, setSelected] = useState<StepRef | null>(null);
  const [tooltip, setTooltip] = useState<StepRef | null>(null);
  const [branchDrafts, setBranchDrafts] = useState<Record<BranchId, string>>({});
  const [showSplits, setShowSplits] = useState(true);
  const [compareLeft, setCompareLeft] = useState<BranchId>('main');
  const [compareRight, setCompareRight] = useState<BranchId>('branch-b');

  const branches = useMemo(() => visibleBranches(view, sheetFilter), [view, sheetFilter]);
  const density = TEMPORAL_DENSITIES[densityIndex] ?? TEMPORAL_DENSITIES[1];
  const totalWindowableSteps = maxStepCount(branches);
  const maxWindowStart = Math.max(0, totalWindowableSteps - density.windowSize);
  const safeWindowStart = clamp(windowStart, 0, maxWindowStart);
  const windowEnd = Math.min(totalWindowableSteps, safeWindowStart + density.windowSize);
  const tallestMagnitude = useMemo(() => maxMagnitude(branches), [branches]);
  const currentHeadRef = headRef(view);
  const currentBranch = activeBranch(view);
  const selectedRef = selected ?? currentHeadRef;
  const selectedStep = stepForRef(view, selectedRef);
  const compareLeftId = compareSelectValue(view.branches, compareLeft, 0);
  const compareRightId = compareSelectValue(view.branches, compareRight, 1);
  const comparison = compareBranches(view, compareLeftId, compareRightId);
  const activeVisibleSteps =
    branches.find((branch) => branch.branch.id === (currentBranch?.id ?? view.head.branchId))
      ?.steps ?? [];
  const activePlayheadIndex = Math.max(
    0,
    activeVisibleSteps.findIndex((step) =>
      refsEqual(stepRef(view.head.branchId, step), currentHeadRef),
    ),
  );
  const previewRef = view.head.mode === 'preview' ? currentHeadRef : null;
  const paneStyle: TimelinePaneStyle = {
    '--timeline-bar-width': `${String(density.barWidth)}px`,
  };

  function send(command: TimelineCommand): void {
    dispatch(command);
  }

  function handleSheetFilterChange(event: ChangeEvent<HTMLSelectElement>): void {
    setSheetFilter(event.target.value);
    setWindowStart(0);
    setSelected(null);
  }

  function handleDensityChange(event: ChangeEvent<HTMLInputElement>): void {
    setDensityIndex(Number(event.target.value));
  }

  function handleWindowChange(event: ChangeEvent<HTMLInputElement>): void {
    setWindowStart(Number(event.target.value));
  }

  function handlePlayheadChange(event: ChangeEvent<HTMLInputElement>): void {
    const index = Number(event.target.value);
    const target = activeVisibleSteps[index];
    if (!currentBranch || !target) return;
    const ref = stepRef(currentBranch.id, target);
    setSelected(ref);
    send({ type: 'goto', ref });
  }

  function handleStepClick(branchId: BranchId, step: TimelineStep): void {
    const ref = stepRef(branchId, step);
    setSelected(ref);
    send({ type: 'goto', ref });
  }

  function moveStep(branchId: BranchId, step: TimelineStep, direction: -1 | 1): void {
    const branch = branches.find((candidate) => candidate.branch.id === branchId);
    const steps = branch?.steps ?? [];
    const index = steps.findIndex((candidate) => candidate.index === step.index);
    const target = steps[clamp(index + direction, 0, Math.max(0, steps.length - 1))];
    if (!target) return;

    const ref = stepRef(branchId, target);
    setSelected(ref);
    setTooltip(ref);
    if (target.index < safeWindowStart) {
      setWindowStart(target.index);
    } else if (target.index >= safeWindowStart + density.windowSize) {
      setWindowStart(Math.max(0, target.index - density.windowSize + 1));
    }
    send({ type: 'goto', ref });
  }

  function handleStepKeyDown(
    event: KeyboardEvent<HTMLButtonElement>,
    branchId: BranchId,
    step: TimelineStep,
  ): void {
    if (event.key === 'ArrowLeft') {
      event.preventDefault();
      moveStep(branchId, step, -1);
    } else if (event.key === 'ArrowRight') {
      event.preventDefault();
      moveStep(branchId, step, 1);
    }
  }

  function setBranchDraft(branchId: BranchId, value: string): void {
    setBranchDrafts((drafts) => ({ ...drafts, [branchId]: value }));
  }

  function commitBranchName(branch: TimelineBranch): void {
    const name = renameDraftValue(branch, branchDrafts).trim();
    if (name.length === 0 || name === branchTitle(branch)) return;
    send({ type: 'renameBranch', branchId: branch.id, name });
  }

  function handleBranchNameKeyDown(
    event: KeyboardEvent<HTMLInputElement>,
    branch: TimelineBranch,
  ): void {
    if (event.key !== 'Enter') return;
    event.currentTarget.blur();
    commitBranchName(branch);
  }

  return (
    <section className="timeline-pane" style={paneStyle} aria-label="Timeline" data-theme={theme}>
      <style>{TIMELINE_PANE_CSS}</style>
      <div className="timeline-pane__sr-status" aria-live="polite" aria-atomic="true">
        {view.head.mode === 'preview'
          ? `Previewing ${view.head.branchId} Step ${String(view.head.previewStepIndex ?? 0)}`
          : `Present on ${view.head.branchId}`}
      </div>

      <header className="timeline-pane__header">
        <div>
          <p className="timeline-pane__eyebrow">Workbook Timeline</p>
          <h1>Parametric Timeline</h1>
        </div>
        <div className="timeline-pane__status" data-mode={view.head.mode}>
          <span>{view.head.mode === 'preview' ? 'Preview' : 'Present'}</span>
          <small>{currentBranch ? branchTitle(currentBranch) : 'No branch'}</small>
        </div>
      </header>

      <div className="timeline-pane__toolbar" role="group" aria-label="Timeline controls">
        <label className="timeline-pane__field">
          <span>Worksheet</span>
          <select
            value={sheetFilter}
            onChange={handleSheetFilterChange}
            aria-label="Worksheet drill-down"
          >
            <option value={ALL_SHEETS}>Whole workbook</option>
            {view.sheets.map((sheetId) => (
              <option key={sheetId} value={sheetId}>
                {sheetId}
              </option>
            ))}
          </select>
        </label>

        <label className="timeline-pane__field">
          <span>Density: {density.label}</span>
          <input
            aria-label="Temporal density"
            type="range"
            min={0}
            max={TEMPORAL_DENSITIES.length - 1}
            value={densityIndex}
            onChange={handleDensityChange}
          />
        </label>

        <label className="timeline-pane__field">
          <span>{windowLabel(safeWindowStart, windowEnd, totalWindowableSteps)}</span>
          <input
            aria-label="Temporal window"
            type="range"
            min={0}
            max={maxWindowStart}
            value={safeWindowStart}
            disabled={maxWindowStart === 0}
            onChange={handleWindowChange}
          />
        </label>

        <label className="timeline-pane__toggle">
          <input
            type="checkbox"
            checked={showSplits}
            onChange={(event) => {
              setShowSplits(event.target.checked);
            }}
          />
          <span>Split tracks</span>
        </label>
      </div>

      <div className="timeline-pane__scrub" role="group" aria-label="Preview controls">
        <label className="timeline-pane__field timeline-pane__field--wide">
          <span>Preview Step</span>
          <input
            aria-label="Timeline playhead"
            type="range"
            min={0}
            max={Math.max(0, activeVisibleSteps.length - 1)}
            value={Math.min(activePlayheadIndex, Math.max(0, activeVisibleSteps.length - 1))}
            disabled={activeVisibleSteps.length === 0}
            onChange={handlePlayheadChange}
          />
        </label>
        {view.head.mode === 'preview' ? (
          <button
            className="timeline-pane__button timeline-pane__button--primary"
            type="button"
            onClick={() => {
              send({ type: 'returnToPresent' });
            }}
          >
            Return to Present
          </button>
        ) : null}
      </div>

      <div className="timeline-pane__legend" aria-label="Step kinds">
        {(Object.keys(KIND_META) as StepKind[]).map((kind) => (
          <span key={kind} className={`timeline-pane__kind timeline-pane__kind--${kind}`}>
            <span aria-hidden="true" />
            {KIND_META[kind].label}
          </span>
        ))}
      </div>

      <div className="timeline-pane__tracks" aria-label="Branch tracks">
        {branches.map(({ branch, steps }) => {
          const windowedSteps = visibleWindow(steps, safeWindowStart, density.windowSize);
          const branchHeadRef = view.head.branchId === branch.id ? currentHeadRef : null;
          const isMain = branch.id === 'main';
          return (
            <article
              className="timeline-pane__track"
              data-active={branch.id === view.head.branchId}
              data-split={showSplits && Boolean(branch.forkedAt)}
              key={branch.id}
            >
              <div className="timeline-pane__branch">
                <button
                  type="button"
                  className="timeline-pane__branch-button"
                  aria-pressed={branch.id === view.head.branchId}
                  aria-label={`Switch to ${branchTitle(branch)} branch`}
                  onClick={() => {
                    send({ type: 'switch', branchId: branch.id });
                  }}
                >
                  <span>{branchTitle(branch)}</span>
                  <small>{branchTone(branch)}</small>
                </button>

                <label className="timeline-pane__field timeline-pane__field--compact">
                  <span>Name</span>
                  <input
                    aria-label={`Rename ${branchTitle(branch)} branch`}
                    type="text"
                    value={renameDraftValue(branch, branchDrafts)}
                    onBlur={() => {
                      commitBranchName(branch);
                    }}
                    onChange={(event) => {
                      setBranchDraft(branch.id, event.target.value);
                    }}
                    onKeyDown={(event) => {
                      handleBranchNameKeyDown(event, branch);
                    }}
                  />
                </label>

                {branch.forkedAt && showSplits ? (
                  <div className="timeline-pane__fork">
                    <span aria-hidden="true" />
                    <p>
                      From {branch.forkedAt.branchId} Step {String(branch.forkedAt.stepIndex)}
                    </p>
                  </div>
                ) : null}

                {!isMain ? (
                  <button
                    className="timeline-pane__button timeline-pane__button--quiet"
                    type="button"
                    onClick={() => {
                      send({ type: 'deleteBranch', branchId: branch.id });
                    }}
                  >
                    Delete
                  </button>
                ) : null}
              </div>

              <div
                className="timeline-pane__histogram"
                role="group"
                aria-label={`${branchTitle(branch)} histogram, ${windowLabel(
                  safeWindowStart,
                  windowEnd,
                  totalWindowableSteps,
                )}`}
              >
                {windowedSteps.length > 0 ? (
                  windowedSteps.map((step) => {
                    const meta = KIND_META[step.kind];
                    const ref = stepRef(branch.id, step);
                    const isHead = refsEqual(ref, branchHeadRef);
                    const isSelected = refsEqual(ref, selectedRef);
                    const stepIndex = String(step.index);
                    const tooltipId = `timeline-tooltip-${branch.id}-${stepIndex}`;
                    const style: BarStyle = {
                      height: `${String(barHeight(step, tallestMagnitude))}px`,
                      '--bar-color': meta.color,
                      '--bar-accent': meta.accent,
                    };
                    return (
                      <button
                        type="button"
                        className="timeline-pane__bar"
                        data-head={isHead}
                        data-kind={step.kind}
                        data-magnitude={step.magnitude}
                        data-selected={isSelected}
                        data-testid={`timeline-bar-${branch.id}-${stepIndex}`}
                        key={`${branch.id}-${stepIndex}`}
                        onBlur={() => {
                          setTooltip(null);
                        }}
                        onClick={() => {
                          handleStepClick(branch.id, step);
                        }}
                        onFocus={() => {
                          setTooltip(ref);
                        }}
                        onKeyDown={(event) => {
                          handleStepKeyDown(event, branch.id, step);
                        }}
                        onMouseEnter={() => {
                          setTooltip(ref);
                        }}
                        onMouseLeave={() => {
                          setTooltip(null);
                        }}
                        style={style}
                        title={tooltipText(branch, step)}
                        aria-current={isHead ? 'step' : undefined}
                        aria-describedby={refsEqual(tooltip, ref) ? tooltipId : undefined}
                        aria-label={`${branchTitle(branch)} Step ${stepIndex}: ${step.label ?? KIND_META[step.kind].label}`}
                      >
                        <span>{stepIndex}</span>
                        {refsEqual(tooltip, ref) ? (
                          <span className="timeline-pane__tooltip" id={tooltipId} role="tooltip">
                            {tooltipText(branch, step)}
                          </span>
                        ) : null}
                      </button>
                    );
                  })
                ) : (
                  <p className="timeline-pane__empty">No Steps in this window</p>
                )}
              </div>
              <div className="timeline-pane__axis" aria-label={`${branchTitle(branch)} step axis`}>
                <span>Step {String(windowedSteps[0]?.index ?? safeWindowStart)}</span>
                <strong>Tall bars mean larger Deltas</strong>
                <span>
                  Step {String(windowedSteps.at(-1)?.index ?? Math.max(0, windowEnd - 1))}
                </span>
              </div>
            </article>
          );
        })}
      </div>

      <section className="timeline-pane__compare" aria-label="Branch compare">
        <div className="timeline-pane__compare-header">
          <div>
            <p className="timeline-pane__eyebrow">Branch Compare</p>
            <h2>Compare branches</h2>
          </div>
          {currentBranch?.parentBranchId ? (
            <button
              className="timeline-pane__button"
              type="button"
              onClick={() => {
                setCompareLeft(currentBranch.parentBranchId ?? compareLeftId);
                setCompareRight(currentBranch.id);
              }}
            >
              Active vs parent
            </button>
          ) : null}
        </div>
        <div className="timeline-pane__compare-controls">
          <label className="timeline-pane__field">
            <span>Left branch</span>
            <select
              value={compareLeftId}
              onChange={(event) => {
                setCompareLeft(event.target.value);
              }}
              aria-label="Compare left branch"
            >
              {view.branches.map((branch) => (
                <option key={branch.id} value={branch.id}>
                  {branchTitle(branch)}
                </option>
              ))}
            </select>
          </label>
          <label className="timeline-pane__field">
            <span>Right branch</span>
            <select
              value={compareRightId}
              onChange={(event) => {
                setCompareRight(event.target.value);
              }}
              aria-label="Compare right branch"
            >
              {view.branches.map((branch) => (
                <option key={branch.id} value={branch.id}>
                  {branchTitle(branch)}
                </option>
              ))}
            </select>
          </label>
        </div>

        {comparison ? (
          <div className="timeline-pane__compare-body">
            <div className="timeline-pane__compare-summary">
              <p>
                Divergence point:{' '}
                <strong>
                  {comparison.fork
                    ? `${comparison.fork.branchId} Step ${String(comparison.fork.stepIndex)}`
                    : 'No recorded fork'}
                </strong>
              </p>
              <p>
                First differing branch Step:{' '}
                <strong>
                  {comparison.firstDifferentStepIndex === null
                    ? 'None in aligned tracks'
                    : String(comparison.firstDifferentStepIndex)}
                </strong>
              </p>
            </div>
            <div className="timeline-pane__compare-grid">
              {[comparison.left, comparison.right].map((branch) => {
                const stats =
                  branch.id === comparison.left.id ? comparison.leftStats : comparison.rightStats;
                return (
                  <article className="timeline-pane__compare-card" key={branch.id}>
                    <h3>{branchTitle(branch)}</h3>
                    <p>
                      {String(stats.stepCount)} Steps · magnitude{' '}
                      {stats.totalMagnitude.toLocaleString()}
                    </p>
                    <div
                      className="timeline-pane__mini-track"
                      aria-label={`${branchTitle(branch)} compare track`}
                    >
                      {branch.steps.slice(0, 16).map((step) => {
                        const miniStyle: MiniBarStyle = {
                          '--mini-color': KIND_META[step.kind].color,
                          '--mini-height': `${String(clamp(Math.round(step.magnitude / 8), 6, 42))}px`,
                        };
                        return (
                          <span
                            className="timeline-pane__mini-bar"
                            key={`${branch.id}-${String(step.index)}`}
                            style={miniStyle}
                            title={`Step ${String(step.index)} · ${step.magnitude.toLocaleString()}`}
                          />
                        );
                      })}
                    </div>
                  </article>
                );
              })}
            </div>
            <div className="timeline-pane__diff-slot">
              Cell-level diff pending engine state integration.
            </div>
          </div>
        ) : (
          <p className="timeline-pane__empty">Choose two branches to compare.</p>
        )}
      </section>

      <aside className="timeline-pane__inspector" aria-label="Step inspector">
        <div className="timeline-pane__inspector-heading">
          <p className="timeline-pane__eyebrow">Inspector</p>
          <h2>{selectedStep ? `Step ${String(selectedStep.index)}` : 'No Step selected'}</h2>
        </div>
        {selectedStep && selectedRef ? (
          <dl>
            <div>
              <dt>Branch</dt>
              <dd>{selectedRef.branchId}</dd>
            </div>
            <div>
              <dt>Kind</dt>
              <dd>{KIND_META[selectedStep.kind].label}</dd>
            </div>
            <div>
              <dt>Magnitude</dt>
              <dd>{selectedStep.magnitude.toLocaleString()}</dd>
            </div>
            <div>
              <dt>Sheet</dt>
              <dd>{selectedStep.sheetId}</dd>
            </div>
            <div>
              <dt>Summary</dt>
              <dd>{selectedStep.label ?? 'Step metadata pending'}</dd>
            </div>
            <div>
              <dt>Formula metadata</dt>
              <dd>Pending inspectStep</dd>
            </div>
          </dl>
        ) : (
          <p className="timeline-pane__empty">Select a Step to inspect it.</p>
        )}
        {previewRef ? (
          <button
            className="timeline-pane__button"
            type="button"
            onClick={() => {
              send({ type: 'branch', from: previewRef });
            }}
          >
            Branch from here
          </button>
        ) : null}
      </aside>
    </section>
  );
}

const TIMELINE_PANE_CSS = `
.timeline-pane {
  --timeline-bg: #f5f6f4;
  --timeline-surface: #ffffff;
  --timeline-surface-muted: #f7faf8;
  --timeline-text: #1f2522;
  --timeline-muted: #616b65;
  --timeline-border: #d9e0dc;
  --timeline-control-border: #cbd5cf;
  --timeline-primary: #287a63;
  --timeline-primary-soft: #e7f2ee;
  --timeline-danger: #7d2d3a;
  --timeline-focus: #0f6cbd;
  --timeline-shadow: rgba(17, 24, 39, 0.14);
  box-sizing: border-box;
  display: grid;
  gap: 14px;
  min-height: 100vh;
  padding: 18px;
  color: var(--timeline-text);
  background: var(--timeline-bg);
  font-family:
    "Segoe UI", Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
}

.timeline-pane[data-theme="dark"] {
  --timeline-bg: #111615;
  --timeline-surface: #1b2220;
  --timeline-surface-muted: #222b28;
  --timeline-text: #edf4f1;
  --timeline-muted: #a8b8b1;
  --timeline-border: #34423d;
  --timeline-control-border: #52635d;
  --timeline-primary: #61c7a6;
  --timeline-primary-soft: #19382f;
  --timeline-danger: #ff9aaa;
  --timeline-focus: #7bb7ff;
  --timeline-shadow: rgba(0, 0, 0, 0.42);
}

.timeline-pane *,
.timeline-pane *::before,
.timeline-pane *::after {
  box-sizing: border-box;
}

.timeline-pane h1,
.timeline-pane h2,
.timeline-pane p {
  margin: 0;
}

.timeline-pane button:focus-visible,
.timeline-pane input:focus-visible,
.timeline-pane select:focus-visible {
  outline: 3px solid var(--timeline-focus);
  outline-offset: 2px;
}

.timeline-pane__sr-status {
  position: absolute;
  width: 1px;
  height: 1px;
  overflow: hidden;
  clip: rect(0 0 0 0);
  white-space: nowrap;
}

.timeline-pane h1 {
  font-size: clamp(1.35rem, 7vw, 2rem);
  line-height: 1.05;
}

.timeline-pane h2 {
  font-size: 1rem;
  line-height: 1.2;
}

.timeline-pane__header,
.timeline-pane__scrub,
.timeline-pane__legend,
.timeline-pane__inspector {
  display: flex;
  align-items: center;
  gap: 12px;
}

.timeline-pane__header {
  justify-content: space-between;
  min-width: 0;
}

.timeline-pane__eyebrow {
  color: var(--timeline-muted);
  font-size: 0.72rem;
  font-weight: 700;
  letter-spacing: 0;
  text-transform: uppercase;
}

.timeline-pane__status {
  display: grid;
  gap: 2px;
  min-width: 96px;
  padding: 8px 10px;
  border: 1px solid var(--timeline-border);
  border-radius: 8px;
  background: var(--timeline-surface);
  color: var(--timeline-text);
  text-align: right;
}

.timeline-pane__status span {
  font-size: 0.82rem;
  font-weight: 800;
}

.timeline-pane__status small {
  color: var(--timeline-muted);
  font-size: 0.7rem;
  font-weight: 700;
}

.timeline-pane__status[data-mode="preview"] {
  border-color: var(--timeline-focus);
  color: var(--timeline-focus);
}

.timeline-pane__toolbar {
  display: grid;
  grid-template-columns: minmax(138px, 1fr) minmax(132px, 0.75fr) minmax(150px, 0.9fr) auto;
  gap: 10px;
  align-items: end;
  padding: 10px;
  border: 1px solid var(--timeline-border);
  border-radius: 8px;
  background: var(--timeline-surface);
}

.timeline-pane__field {
  display: grid;
  gap: 5px;
  min-width: 0;
  color: var(--timeline-muted);
  font-size: 0.76rem;
  font-weight: 800;
}

.timeline-pane__field--wide {
  flex: 1 1 auto;
}

.timeline-pane__field--compact {
  gap: 4px;
}

.timeline-pane select,
.timeline-pane input[type="text"],
.timeline-pane input[type="range"] {
  width: 100%;
  min-height: 34px;
}

.timeline-pane select,
.timeline-pane input[type="text"] {
  border: 1px solid var(--timeline-control-border);
  border-radius: 7px;
  padding: 6px 8px;
  background: var(--timeline-surface);
  color: var(--timeline-text);
  font: inherit;
}

.timeline-pane input[type="range"] {
  accent-color: var(--timeline-primary);
}

.timeline-pane__toggle {
  display: inline-flex;
  align-items: center;
  gap: 7px;
  min-height: 34px;
  color: var(--timeline-muted);
  font-size: 0.78rem;
  font-weight: 800;
  white-space: nowrap;
}

.timeline-pane__scrub {
  align-items: end;
  padding: 10px;
  border: 1px solid var(--timeline-border);
  border-radius: 8px;
  background: var(--timeline-surface);
}

.timeline-pane__button,
.timeline-pane__branch-button {
  min-height: 34px;
  border: 1px solid var(--timeline-control-border);
  border-radius: 7px;
  background: var(--timeline-surface);
  color: var(--timeline-text);
  font-weight: 800;
  cursor: pointer;
}

.timeline-pane__button {
  flex: 0 0 auto;
  padding: 7px 10px;
}

.timeline-pane__button--primary {
  border-color: var(--timeline-primary);
  background: var(--timeline-primary);
  color: #ffffff;
}

.timeline-pane__button--quiet {
  color: var(--timeline-danger);
}

.timeline-pane__legend {
  flex-wrap: wrap;
}

.timeline-pane__kind {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  color: var(--timeline-muted);
  font-size: 0.76rem;
  font-weight: 800;
}

.timeline-pane__kind span {
  display: inline-block;
  width: 10px;
  height: 10px;
  border-radius: 50%;
}

.timeline-pane__kind--value span {
  background: #287a63;
}

.timeline-pane__kind--structural span {
  background: #3f6db6;
}

.timeline-pane__kind--worksheet span {
  background: #a46b10;
}

.timeline-pane__kind--reconciliation span {
  background: #b24d5e;
}

.timeline-pane__tracks {
  display: grid;
  gap: 10px;
  overflow: visible;
}

.timeline-pane__track {
  display: grid;
  grid-template-columns: minmax(118px, 0.34fr) minmax(0, 1fr);
  gap: 10px;
  min-height: 178px;
  padding: 10px;
  border: 1px solid var(--timeline-border);
  border-radius: 8px;
  background: var(--timeline-surface);
}

.timeline-pane__track[data-active="true"] {
  border-color: var(--timeline-primary);
  box-shadow: inset 3px 0 0 var(--timeline-primary);
}

.timeline-pane__track[data-split="true"] {
  background:
    linear-gradient(90deg, rgba(63, 109, 182, 0.1), transparent 34%),
    var(--timeline-surface);
}

.timeline-pane__branch {
  display: grid;
  align-content: start;
  gap: 8px;
}

.timeline-pane__branch-button {
  width: 100%;
  padding: 8px;
  background: var(--timeline-surface-muted);
  text-align: left;
}

.timeline-pane__branch-button[aria-pressed="true"] {
  border-color: var(--timeline-primary);
  background: var(--timeline-primary-soft);
}

.timeline-pane__branch-button span,
.timeline-pane__branch-button small {
  display: block;
  overflow-wrap: anywhere;
}

.timeline-pane__branch-button small {
  margin-top: 2px;
  color: var(--timeline-muted);
  font-size: 0.68rem;
}

.timeline-pane__fork {
  display: grid;
  grid-template-columns: 12px minmax(0, 1fr);
  gap: 6px;
  color: var(--timeline-muted);
  font-size: 0.72rem;
  line-height: 1.25;
}

.timeline-pane__fork span {
  width: 12px;
  min-height: 34px;
  border-left: 2px solid var(--timeline-control-border);
  border-bottom: 2px solid var(--timeline-control-border);
}

.timeline-pane__histogram {
  display: flex;
  align-items: end;
  gap: 4px;
  min-width: 0;
  min-height: 156px;
  overflow-x: auto;
  overflow-y: visible;
  padding: 10px 6px 2px;
  border-bottom: 1px solid var(--timeline-border);
}

.timeline-pane__bar {
  position: relative;
  flex: 0 0 var(--timeline-bar-width);
  min-width: var(--timeline-bar-width);
  max-width: var(--timeline-bar-width);
  border: 0;
  border-radius: 6px 6px 2px 2px;
  background: var(--bar-color);
  color: #ffffff;
  font-size: 0.62rem;
  font-weight: 900;
  line-height: 1;
  cursor: pointer;
}

.timeline-pane__bar::before {
  content: "";
  position: absolute;
  inset: -4px;
  border: 2px solid transparent;
  border-radius: 8px;
}

.timeline-pane__bar[data-selected="true"]::before {
  border-color: var(--timeline-text);
}

.timeline-pane__bar[data-head="true"] {
  box-shadow: 0 -5px 0 var(--bar-accent);
}

.timeline-pane__bar {
  transition:
    height 160ms ease,
    filter 160ms ease,
    transform 160ms ease;
}

.timeline-pane__bar:hover,
.timeline-pane__bar:focus-visible {
  filter: brightness(1.08);
  transform: translateY(-2px);
}

.timeline-pane__bar span {
  position: absolute;
  right: 3px;
  bottom: 4px;
}

.timeline-pane__tooltip {
  position: absolute;
  z-index: 4;
  right: 0;
  bottom: calc(100% + 8px);
  width: min(220px, 70vw);
  padding: 8px 9px;
  border: 1px solid var(--timeline-control-border);
  border-radius: 7px;
  background: var(--timeline-surface);
  color: var(--timeline-text);
  box-shadow: 0 10px 24px var(--timeline-shadow);
  font-size: 0.73rem;
  font-weight: 700;
  line-height: 1.3;
  text-align: left;
}

.timeline-pane__axis {
  display: grid;
  grid-template-columns: 1fr auto 1fr;
  gap: 8px;
  align-items: center;
  color: var(--timeline-muted);
  font-size: 0.68rem;
  font-weight: 800;
}

.timeline-pane__axis span:last-child {
  text-align: right;
}

.timeline-pane__axis strong {
  color: var(--timeline-text);
  font-size: 0.7rem;
}

.timeline-pane__compare {
  display: grid;
  gap: 12px;
  padding: 12px;
  border: 1px solid var(--timeline-border);
  border-radius: 8px;
  background: var(--timeline-surface);
}

.timeline-pane__compare-header,
.timeline-pane__compare-controls,
.timeline-pane__compare-summary,
.timeline-pane__compare-grid {
  display: grid;
  gap: 10px;
}

.timeline-pane__compare-header {
  grid-template-columns: minmax(0, 1fr) auto;
  align-items: center;
}

.timeline-pane__compare-controls {
  grid-template-columns: repeat(2, minmax(0, 1fr));
}

.timeline-pane__compare-summary {
  grid-template-columns: repeat(2, minmax(0, 1fr));
  color: var(--timeline-muted);
  font-size: 0.8rem;
  font-weight: 800;
}

.timeline-pane__compare-summary strong {
  color: var(--timeline-text);
}

.timeline-pane__compare-grid {
  grid-template-columns: repeat(2, minmax(0, 1fr));
}

.timeline-pane__compare-card {
  display: grid;
  gap: 8px;
  min-width: 0;
  padding: 10px;
  border: 1px solid var(--timeline-border);
  border-radius: 8px;
  background: var(--timeline-surface-muted);
}

.timeline-pane__compare-card h3 {
  margin: 0;
  font-size: 0.95rem;
  overflow-wrap: anywhere;
}

.timeline-pane__compare-card p {
  color: var(--timeline-muted);
  font-size: 0.76rem;
  font-weight: 800;
}

.timeline-pane__mini-track {
  display: flex;
  align-items: end;
  gap: 3px;
  min-height: 48px;
  overflow-x: auto;
  padding-block: 4px;
}

.timeline-pane__mini-bar {
  flex: 0 0 8px;
  height: var(--mini-height);
  border-radius: 4px 4px 1px 1px;
  background: var(--mini-color);
}

.timeline-pane__diff-slot {
  padding: 10px;
  border: 1px dashed var(--timeline-control-border);
  border-radius: 8px;
  color: var(--timeline-muted);
  font-size: 0.8rem;
  font-weight: 800;
}

.timeline-pane__inspector {
  align-items: start;
  justify-content: space-between;
  padding: 12px;
  border: 1px solid var(--timeline-border);
  border-radius: 8px;
  background: var(--timeline-surface);
}

.timeline-pane__inspector-heading {
  flex: 0 0 118px;
}

.timeline-pane__inspector dl {
  display: grid;
  grid-template-columns: repeat(6, minmax(0, 1fr));
  gap: 10px;
  width: 100%;
  margin: 0;
}

.timeline-pane__inspector div {
  min-width: 0;
}

.timeline-pane__inspector dt {
  color: var(--timeline-muted);
  font-size: 0.68rem;
  font-weight: 800;
}

.timeline-pane__inspector dd {
  margin: 2px 0 0;
  color: var(--timeline-text);
  font-size: 0.82rem;
  font-weight: 800;
  overflow-wrap: anywhere;
}

.timeline-pane__empty {
  align-self: center;
  color: var(--timeline-muted);
  font-size: 0.82rem;
  font-weight: 800;
}

@media (max-width: 640px) {
  .timeline-pane {
    padding: 14px;
  }

  .timeline-pane__toolbar,
  .timeline-pane__track,
  .timeline-pane__inspector,
  .timeline-pane__inspector dl {
    grid-template-columns: 1fr;
  }

  .timeline-pane__scrub,
  .timeline-pane__inspector {
    align-items: stretch;
    flex-direction: column;
  }
}

@media (max-width: 360px) {
  .timeline-pane {
    gap: 10px;
    padding: 10px;
  }

  .timeline-pane__header {
    align-items: stretch;
    flex-direction: column;
  }

  .timeline-pane__status {
    width: 100%;
    text-align: left;
  }

  .timeline-pane__toolbar,
  .timeline-pane__compare-header,
  .timeline-pane__compare-controls,
  .timeline-pane__compare-summary,
  .timeline-pane__compare-grid {
    grid-template-columns: 1fr;
  }

  .timeline-pane__histogram {
    min-height: 136px;
  }

  .timeline-pane__axis {
    grid-template-columns: 1fr;
  }

  .timeline-pane__axis span:last-child {
    text-align: left;
  }
}
`;
