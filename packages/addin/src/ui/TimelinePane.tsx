import { useMemo, useState } from 'react';
import type { ChangeEvent, CSSProperties, KeyboardEvent } from 'react';
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

export function TimelinePane({ view, dispatch }: TimelinePaneProps): React.JSX.Element {
  const [sheetFilter, setSheetFilter] = useState<string>(ALL_SHEETS);
  const [densityIndex, setDensityIndex] = useState(1);
  const [windowStart, setWindowStart] = useState(0);
  const [selected, setSelected] = useState<StepRef | null>(null);
  const [branchDrafts, setBranchDrafts] = useState<Record<BranchId, string>>({});
  const [showSplits, setShowSplits] = useState(true);

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
    <section className="timeline-pane" style={paneStyle} aria-label="Timeline">
      <style>{TIMELINE_PANE_CSS}</style>

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

      <div className="timeline-pane__toolbar" aria-label="Timeline controls">
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

      <div className="timeline-pane__scrub">
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

              <div className="timeline-pane__histogram" aria-label={`${branchTitle(branch)} Steps`}>
                {windowedSteps.length > 0 ? (
                  windowedSteps.map((step) => {
                    const meta = KIND_META[step.kind];
                    const ref = stepRef(branch.id, step);
                    const isHead = refsEqual(ref, branchHeadRef);
                    const isSelected = refsEqual(ref, selectedRef);
                    const stepIndex = String(step.index);
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
                        onClick={() => {
                          handleStepClick(branch.id, step);
                        }}
                        style={style}
                        aria-current={isHead ? 'step' : undefined}
                        aria-label={`${branchTitle(branch)} Step ${stepIndex}: ${step.label ?? KIND_META[step.kind].label}`}
                      >
                        <span>{stepIndex}</span>
                      </button>
                    );
                  })
                ) : (
                  <p className="timeline-pane__empty">No Steps in this window</p>
                )}
              </div>
            </article>
          );
        })}
      </div>

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
  box-sizing: border-box;
  display: grid;
  gap: 14px;
  min-height: 100vh;
  padding: 18px;
  color: #1f2522;
  background: #f5f6f4;
  font-family:
    "Segoe UI", Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
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
  color: #616b65;
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
  border: 1px solid #cfd8d2;
  border-radius: 8px;
  background: #ffffff;
  color: #34423b;
  text-align: right;
}

.timeline-pane__status span {
  font-size: 0.82rem;
  font-weight: 800;
}

.timeline-pane__status small {
  color: #616b65;
  font-size: 0.7rem;
  font-weight: 700;
}

.timeline-pane__status[data-mode="preview"] {
  border-color: #3f6db6;
  color: #254b87;
}

.timeline-pane__toolbar {
  display: grid;
  grid-template-columns: minmax(138px, 1fr) minmax(132px, 0.75fr) minmax(150px, 0.9fr) auto;
  gap: 10px;
  align-items: end;
  padding: 10px;
  border: 1px solid #d9e0dc;
  border-radius: 8px;
  background: #ffffff;
}

.timeline-pane__field {
  display: grid;
  gap: 5px;
  min-width: 0;
  color: #46524d;
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
  border: 1px solid #cbd5cf;
  border-radius: 7px;
  padding: 6px 8px;
  background: #ffffff;
  color: #1f2522;
  font: inherit;
}

.timeline-pane input[type="range"] {
  accent-color: #287a63;
}

.timeline-pane__toggle {
  display: inline-flex;
  align-items: center;
  gap: 7px;
  min-height: 34px;
  color: #46524d;
  font-size: 0.78rem;
  font-weight: 800;
  white-space: nowrap;
}

.timeline-pane__scrub {
  align-items: end;
  padding: 10px;
  border: 1px solid #d9e0dc;
  border-radius: 8px;
  background: #ffffff;
}

.timeline-pane__button,
.timeline-pane__branch-button {
  min-height: 34px;
  border: 1px solid #cbd5cf;
  border-radius: 7px;
  background: #ffffff;
  color: #20362b;
  font-weight: 800;
  cursor: pointer;
}

.timeline-pane__button {
  flex: 0 0 auto;
  padding: 7px 10px;
}

.timeline-pane__button--primary {
  border-color: #287a63;
  background: #287a63;
  color: #ffffff;
}

.timeline-pane__button--quiet {
  color: #7d2d3a;
}

.timeline-pane__legend {
  flex-wrap: wrap;
}

.timeline-pane__kind {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  color: #46524d;
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
  overflow: hidden;
}

.timeline-pane__track {
  display: grid;
  grid-template-columns: minmax(118px, 0.34fr) minmax(0, 1fr);
  gap: 10px;
  min-height: 178px;
  padding: 10px;
  border: 1px solid #d9e0dc;
  border-radius: 8px;
  background: #ffffff;
}

.timeline-pane__track[data-active="true"] {
  border-color: #8bb7a7;
  box-shadow: inset 3px 0 0 #287a63;
}

.timeline-pane__track[data-split="true"] {
  background:
    linear-gradient(90deg, rgba(63, 109, 182, 0.1), transparent 34%),
    #ffffff;
}

.timeline-pane__branch {
  display: grid;
  align-content: start;
  gap: 8px;
}

.timeline-pane__branch-button {
  width: 100%;
  padding: 8px;
  background: #f7faf8;
  text-align: left;
}

.timeline-pane__branch-button[aria-pressed="true"] {
  border-color: #287a63;
  background: #e7f2ee;
}

.timeline-pane__branch-button span,
.timeline-pane__branch-button small {
  display: block;
  overflow-wrap: anywhere;
}

.timeline-pane__branch-button small {
  margin-top: 2px;
  color: #616b65;
  font-size: 0.68rem;
}

.timeline-pane__fork {
  display: grid;
  grid-template-columns: 12px minmax(0, 1fr);
  gap: 6px;
  color: #65716b;
  font-size: 0.72rem;
  line-height: 1.25;
}

.timeline-pane__fork span {
  width: 12px;
  min-height: 34px;
  border-left: 2px solid #9eb3aa;
  border-bottom: 2px solid #9eb3aa;
}

.timeline-pane__histogram {
  display: flex;
  align-items: end;
  gap: 4px;
  min-width: 0;
  min-height: 156px;
  overflow: hidden;
  padding: 10px 6px 2px;
  border-bottom: 1px solid #d9e0dc;
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
  border-color: #1f2522;
}

.timeline-pane__bar[data-head="true"] {
  box-shadow: 0 -5px 0 var(--bar-accent);
}

.timeline-pane__bar span {
  position: absolute;
  right: 3px;
  bottom: 4px;
}

.timeline-pane__inspector {
  align-items: start;
  justify-content: space-between;
  padding: 12px;
  border: 1px solid #d9e0dc;
  border-radius: 8px;
  background: #ffffff;
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
  color: #616b65;
  font-size: 0.68rem;
  font-weight: 800;
}

.timeline-pane__inspector dd {
  margin: 2px 0 0;
  color: #1f2522;
  font-size: 0.82rem;
  font-weight: 800;
  overflow-wrap: anywhere;
}

.timeline-pane__empty {
  align-self: center;
  color: #616b65;
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
`;
