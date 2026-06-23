import { useMemo, useState } from 'react';
import type { ChangeEvent, CSSProperties } from 'react';
import type {
  BranchId,
  StepKind,
  StepRef,
  TimelineBranch,
  TimelinePaneProps,
  TimelineStep,
} from './contract.ts';

const ALL_SHEETS = '__timeline_all_sheets__';

const KIND_META: Record<StepKind, { label: string; color: string; accent: string }> = {
  value: { label: 'Value', color: '#27b37e', accent: '#dff8ed' },
  structural: { label: 'Structural', color: '#4f86f7', accent: '#e6efff' },
  worksheet: { label: 'Worksheet', color: '#d89c18', accent: '#fff3d1' },
  reconciliation: { label: 'Reconciliation', color: '#d75f74', accent: '#ffe4e8' },
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
  if (!left || !right) {
    return false;
  }
  return left.branchId === right.branchId && left.stepIndex === right.stepIndex;
}

function activeBranch(view: TimelinePaneProps['view']): TimelineBranch | undefined {
  return view.branches.find((branch) => branch.id === view.head.branchId) ?? view.branches[0];
}

function presentRef(branch: TimelineBranch | undefined): StepRef | null {
  const tip = branch?.steps.at(-1);
  if (!branch || !tip) {
    return null;
  }
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
  if (!ref) {
    return null;
  }
  const branch = view.branches.find((candidate) => candidate.id === ref.branchId);
  return branch?.steps.find((step) => step.index === ref.stepIndex) ?? null;
}

function visibleStepForRef(branches: VisibleBranch[], ref: StepRef | null): TimelineStep | null {
  if (!ref) {
    return null;
  }
  const branch = branches.find((candidate) => candidate.branch.id === ref.branchId);
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

function barHeight(step: TimelineStep, max: number): number {
  const ratio = Math.max(0.04, step.magnitude / max);
  return Math.round(150 * ratio);
}

function kindClassName(kind: StepKind): string {
  return `timeline-pane__kind timeline-pane__kind--${kind}`;
}

function selectedBranchId(ref: StepRef | null, fallback: BranchId): BranchId {
  return ref?.branchId ?? fallback;
}

export function TimelinePane({ view, dispatch }: TimelinePaneProps): React.JSX.Element {
  const [sheetFilter, setSheetFilter] = useState<string>(ALL_SHEETS);
  const [temporalZoom, setTemporalZoom] = useState(2);
  const [selected, setSelected] = useState<StepRef | null>(null);

  const branches = useMemo(() => visibleBranches(view, sheetFilter), [view, sheetFilter]);
  const tallestMagnitude = useMemo(() => maxMagnitude(branches), [branches]);
  const currentHeadRef = headRef(view);
  const currentBranch = activeBranch(view);
  const selectedRef =
    selected && visibleStepForRef(branches, selected)
      ? selected
      : currentHeadRef && visibleStepForRef(branches, currentHeadRef)
        ? currentHeadRef
        : null;
  const selectedStep = stepForRef(view, selectedRef);
  const activeVisibleSteps =
    branches.find((branch) => branch.branch.id === (currentBranch?.id ?? view.head.branchId))
      ?.steps ?? [];
  const playheadIndex = Math.max(
    0,
    activeVisibleSteps.findIndex((step) =>
      refsEqual(stepRef(view.head.branchId, step), currentHeadRef),
    ),
  );
  const previewRef = view.head.mode === 'preview' ? currentHeadRef : null;
  const paneStyle: TimelinePaneStyle = {
    '--timeline-bar-width': `${String(24 + temporalZoom * 10)}px`,
  };

  function handleSheetFilterChange(event: ChangeEvent<HTMLSelectElement>): void {
    setSheetFilter(event.target.value);
    setSelected(null);
  }

  function handleTemporalZoomChange(event: ChangeEvent<HTMLInputElement>): void {
    setTemporalZoom(Number(event.target.value));
  }

  function handlePlayheadChange(event: ChangeEvent<HTMLInputElement>): void {
    const index = Number(event.target.value);
    const target = activeVisibleSteps[index];
    if (!currentBranch || !target) {
      return;
    }
    const ref = stepRef(currentBranch.id, target);
    setSelected(ref);
    dispatch({ type: 'goto', ref });
  }

  function handleStepClick(branchId: BranchId, step: TimelineStep): void {
    const ref = stepRef(branchId, step);
    setSelected(ref);
    dispatch({ type: 'goto', ref });
  }

  return (
    <section className="timeline-pane" style={paneStyle} aria-label="Timeline">
      <style>{TIMELINE_PANE_CSS}</style>

      <header className="timeline-pane__header">
        <div>
          <p className="timeline-pane__eyebrow">Workbook Timeline</p>
          <h1>Timeline</h1>
        </div>
        <div className="timeline-pane__status" data-mode={view.head.mode}>
          {view.head.mode === 'preview' ? 'Preview' : 'Present'}
        </div>
      </header>

      <div className="timeline-pane__controls" aria-label="Timeline controls">
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
          <span>Temporal zoom</span>
          <input
            aria-label="Temporal zoom"
            type="range"
            min={1}
            max={5}
            value={temporalZoom}
            onChange={handleTemporalZoomChange}
          />
        </label>
      </div>

      <div className="timeline-pane__scrub">
        <label className="timeline-pane__field timeline-pane__field--wide">
          <span>Step</span>
          <input
            aria-label="Timeline playhead"
            type="range"
            min={0}
            max={Math.max(0, activeVisibleSteps.length - 1)}
            value={Math.min(playheadIndex, Math.max(0, activeVisibleSteps.length - 1))}
            disabled={activeVisibleSteps.length === 0}
            onChange={handlePlayheadChange}
          />
        </label>
        {view.head.mode === 'preview' ? (
          <button
            className="timeline-pane__button timeline-pane__button--primary"
            type="button"
            onClick={() => {
              dispatch({ type: 'returnToPresent' });
            }}
          >
            Return to Present
          </button>
        ) : null}
      </div>

      <div className="timeline-pane__legend" aria-label="Step kinds">
        {(Object.keys(KIND_META) as StepKind[]).map((kind) => (
          <span key={kind} className={kindClassName(kind)}>
            <span aria-hidden="true" />
            {KIND_META[kind].label}
          </span>
        ))}
      </div>

      <div className="timeline-pane__tracks" aria-label="Branch tracks">
        {branches.map(({ branch, steps }) => {
          const branchHeadRef =
            view.head.branchId === branch.id
              ? currentHeadRef
              : branch.id === currentBranch?.id
                ? presentRef(branch)
                : null;
          return (
            <article
              className="timeline-pane__track"
              data-active={branch.id === view.head.branchId}
              key={branch.id}
            >
              <div className="timeline-pane__branch">
                <button
                  type="button"
                  className="timeline-pane__branch-button"
                  aria-pressed={branch.id === view.head.branchId}
                  aria-label={`Switch to ${branchTitle(branch)} branch`}
                  onClick={() => {
                    dispatch({ type: 'switch', branchId: branch.id });
                  }}
                >
                  <span>{branchTitle(branch)}</span>
                  {branch.provisional ? <small>provisional</small> : null}
                </button>
                {branch.forkedAt ? (
                  <div className="timeline-pane__fork">
                    <span aria-hidden="true" />
                    <p>
                      Forked from {branch.forkedAt.branchId} at Step {branch.forkedAt.stepIndex}
                    </p>
                  </div>
                ) : null}
              </div>

              <div className="timeline-pane__histogram" aria-label={`${branchTitle(branch)} Steps`}>
                {steps.length > 0 ? (
                  steps.map((step) => {
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
                  <p className="timeline-pane__empty">No Steps for this worksheet</p>
                )}
              </div>
            </article>
          );
        })}
      </div>

      <aside className="timeline-pane__inspector" aria-label="Step inspector">
        <div>
          <p className="timeline-pane__eyebrow">Inspector</p>
          <h2>{selectedStep ? `Step ${String(selectedStep.index)}` : 'No Step selected'}</h2>
        </div>
        {selectedStep && selectedRef ? (
          <dl>
            <div>
              <dt>Branch</dt>
              <dd>{selectedBranchId(selectedRef, view.head.branchId)}</dd>
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
          </dl>
        ) : (
          <p className="timeline-pane__empty">Select a Step to inspect it.</p>
        )}
        {previewRef ? (
          <button
            className="timeline-pane__button"
            type="button"
            onClick={() => {
              dispatch({ type: 'branch', from: previewRef });
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
  color: #18211d;
  background: #f7f8f5;
  font-family:
    Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}

.timeline-pane *,
.timeline-pane *::before,
.timeline-pane *::after {
  box-sizing: border-box;
}

.timeline-pane__header,
.timeline-pane__scrub,
.timeline-pane__controls,
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

.timeline-pane h1,
.timeline-pane h2,
.timeline-pane p {
  margin: 0;
}

.timeline-pane h1 {
  font-size: clamp(1.45rem, 8vw, 2.15rem);
  line-height: 1.05;
}

.timeline-pane h2 {
  font-size: 1rem;
  line-height: 1.2;
}

.timeline-pane__eyebrow {
  color: #66736d;
  font-size: 0.72rem;
  font-weight: 700;
  letter-spacing: 0;
  text-transform: uppercase;
}

.timeline-pane__status,
.timeline-pane__button,
.timeline-pane__branch-button,
.timeline-pane select,
.timeline-pane input[type="range"] {
  min-height: 34px;
}

.timeline-pane__status,
.timeline-pane__button,
.timeline-pane__branch-button {
  border: 1px solid #cdd8d1;
  border-radius: 8px;
}

.timeline-pane__status {
  padding: 7px 10px;
  background: #ffffff;
  color: #415047;
  font-size: 0.8rem;
  font-weight: 700;
}

.timeline-pane__status[data-mode="preview"] {
  border-color: #4f86f7;
  color: #244f9f;
}

.timeline-pane__controls {
  display: grid;
  grid-template-columns: minmax(0, 1fr) minmax(150px, 0.65fr);
}

.timeline-pane__field {
  display: grid;
  gap: 5px;
  min-width: 0;
  color: #46534c;
  font-size: 0.77rem;
  font-weight: 700;
}

.timeline-pane__field--wide {
  flex: 1 1 auto;
}

.timeline-pane select {
  width: 100%;
  border: 1px solid #cdd8d1;
  border-radius: 8px;
  padding: 6px 8px;
  background: #ffffff;
  color: #18211d;
}

.timeline-pane input[type="range"] {
  width: 100%;
  accent-color: #1f7a59;
}

.timeline-pane__scrub {
  align-items: end;
}

.timeline-pane__button {
  flex: 0 0 auto;
  padding: 7px 10px;
  background: #ffffff;
  color: #1f372c;
  font-weight: 800;
  cursor: pointer;
}

.timeline-pane__button--primary {
  border-color: #1f7a59;
  background: #1f7a59;
  color: #ffffff;
}

.timeline-pane__legend {
  flex-wrap: wrap;
}

.timeline-pane__kind {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  color: #46534c;
  font-size: 0.76rem;
  font-weight: 700;
}

.timeline-pane__kind span {
  display: inline-block;
  width: 10px;
  height: 10px;
  border-radius: 50%;
}

.timeline-pane__kind--value span {
  background: #27b37e;
}

.timeline-pane__kind--structural span {
  background: #4f86f7;
}

.timeline-pane__kind--worksheet span {
  background: #d89c18;
}

.timeline-pane__kind--reconciliation span {
  background: #d75f74;
}

.timeline-pane__tracks {
  display: grid;
  gap: 10px;
  overflow: hidden;
}

.timeline-pane__track {
  display: grid;
  grid-template-columns: minmax(92px, 0.32fr) minmax(0, 1fr);
  gap: 10px;
  min-height: 178px;
  padding: 10px;
  border: 1px solid #d8e0dc;
  border-radius: 8px;
  background: #ffffff;
}

.timeline-pane__track[data-active="true"] {
  border-color: #7bb89e;
  box-shadow: inset 3px 0 0 #1f7a59;
}

.timeline-pane__branch {
  display: grid;
  align-content: start;
  gap: 8px;
}

.timeline-pane__branch-button {
  width: 100%;
  padding: 8px;
  background: #f4f7f2;
  color: #20362b;
  font-weight: 900;
  text-align: left;
  cursor: pointer;
}

.timeline-pane__branch-button[aria-pressed="true"] {
  border-color: #1f7a59;
  background: #e7f5ee;
}

.timeline-pane__branch-button small {
  display: block;
  color: #66736d;
  font-size: 0.68rem;
}

.timeline-pane__fork {
  display: grid;
  grid-template-columns: 12px minmax(0, 1fr);
  gap: 6px;
  color: #6a746f;
  font-size: 0.72rem;
  line-height: 1.25;
}

.timeline-pane__fork span {
  width: 12px;
  min-height: 34px;
  border-left: 2px solid #9db7aa;
  border-bottom: 2px solid #9db7aa;
}

.timeline-pane__histogram {
  display: flex;
  align-items: end;
  gap: 6px;
  min-width: 0;
  min-height: 156px;
  overflow-x: auto;
  padding: 10px 6px 2px;
  border-bottom: 1px solid #d8e0dc;
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
  font-size: 0.67rem;
  font-weight: 900;
  line-height: 1;
  cursor: pointer;
}

.timeline-pane__bar::before {
  content: "";
  position: absolute;
  inset: -5px;
  border: 2px solid transparent;
  border-radius: 8px;
}

.timeline-pane__bar[data-selected="true"]::before {
  border-color: #18211d;
}

.timeline-pane__bar[data-head="true"] {
  box-shadow: 0 -5px 0 var(--bar-accent);
}

.timeline-pane__bar span {
  position: absolute;
  right: 4px;
  bottom: 4px;
}

.timeline-pane__inspector {
  align-items: start;
  justify-content: space-between;
  padding: 12px;
  border: 1px solid #d8e0dc;
  border-radius: 8px;
  background: #ffffff;
}

.timeline-pane__inspector dl {
  display: grid;
  grid-template-columns: repeat(5, minmax(0, 1fr));
  gap: 10px;
  width: 100%;
  margin: 0;
}

.timeline-pane__inspector div {
  min-width: 0;
}

.timeline-pane__inspector dt {
  color: #66736d;
  font-size: 0.68rem;
  font-weight: 800;
}

.timeline-pane__inspector dd {
  margin: 2px 0 0;
  color: #18211d;
  font-size: 0.82rem;
  font-weight: 800;
  overflow-wrap: anywhere;
}

.timeline-pane__empty {
  align-self: center;
  color: #66736d;
  font-size: 0.82rem;
  font-weight: 700;
}

@media (max-width: 520px) {
  .timeline-pane {
    padding: 14px;
  }

  .timeline-pane__controls,
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
