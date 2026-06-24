import { useState } from 'react';
import type { ChangeEvent, CSSProperties } from 'react';
import type {
  BranchId,
  StepKind,
  TimelineBranch,
  TimelinePaneProps,
  TimelineStep,
  TimelineTheme,
} from './contract.ts';
import { OperationIcon } from './OperationIcon.tsx';

// ---------------------------------------------------------------------------
// Pure helpers (unit-tested directly — jsdom has no layout for pointer math).
// ---------------------------------------------------------------------------

const KIND_COLOR: Record<StepKind, string> = {
  value: '#3ea36b',
  structural: '#3b82f6',
  worksheet: '#e0b341',
  reconciliation: '#e0598b',
};
const KIND_LABEL: Record<StepKind, string> = {
  value: 'Value',
  structural: 'Structural',
  worksheet: 'Worksheet',
  reconciliation: 'Reconciliation',
};
const ALL_SHEETS = '__all__';

export function activeBranch(branches: TimelineBranch[], headBranchId: BranchId): TimelineBranch {
  const found = branches.find((b) => b.id === headBranchId);
  if (found) return found;
  return branches[0] ?? { id: headBranchId, provisional: false, steps: [] };
}

/** The step index the head is on: the tip in Present, the previewed index otherwise. */
export function currentStepIndex(
  branch: TimelineBranch,
  mode: 'present' | 'preview',
  previewIndex: number | undefined,
): number {
  if (mode === 'preview' && previewIndex !== undefined) return previewIndex;
  return branch.steps.at(-1)?.index ?? 0;
}

/** Log-scaled bar height (px): a 1,000-cell paste towers over a 1-cell edit without dwarfing the rest. */
export function barHeight(magnitude: number, maxMagnitude: number, trackPx: number): number {
  const norm = Math.log(Math.max(0, magnitude) + 1) / Math.log(Math.max(1, maxMagnitude) + 1);
  return 8 + Math.round(norm * (trackPx - 8));
}

/** Map a brushed pixel range to a clamped, ordered zoom window. null for a too-small drag (treated as a click). */
export function brushToWindow(
  x0: number,
  x1: number,
  trackWidth: number,
  min: number,
  max: number,
): { start: number; end: number } | null {
  if (trackWidth <= 0 || Math.abs(x1 - x0) < 6) return null;
  const toIndex = (x: number): number =>
    Math.round(min + (Math.max(0, Math.min(trackWidth, x)) / trackWidth) * (max - min));
  const a = toIndex(Math.min(x0, x1));
  const b = toIndex(Math.max(x0, x1));
  return b > a ? { start: a, end: b } : null;
}

const THEME: Record<TimelineTheme, Record<string, string>> = {
  light: {
    bg: '#ffffff',
    ink: '#1f2937',
    muted: '#6b7280',
    line: '#e5e7eb',
    chip: '#f3f4f6',
    accent: '#111827',
  },
  dark: {
    bg: '#13161c',
    ink: '#e5e7eb',
    muted: '#9aa3af',
    line: '#2a2f3a',
    chip: '#1c2128',
    accent: '#e5e7eb',
  },
};

function btn(t: Record<string, string>): CSSProperties {
  return {
    fontSize: 11,
    fontWeight: 600,
    padding: '4px 9px',
    borderRadius: 6,
    cursor: 'pointer',
    border: `1px solid ${t.line ?? '#e5e7eb'}`,
    background: t.chip ?? '#f3f4f6',
    color: t.ink ?? '#1f2937',
  };
}

// ---------------------------------------------------------------------------

export function TimelinePane({ view, dispatch, theme = 'light' }: TimelinePaneProps) {
  const t = THEME[theme];
  const branch = activeBranch(view.branches, view.head.branchId);
  const lastIndex = branch.steps.at(-1)?.index ?? 0;
  const headIndex = currentStepIndex(branch, view.head.mode, view.head.previewStepIndex);
  const inPreview = view.head.mode === 'preview';

  const [hover, setHover] = useState<TimelineStep | null>(null);
  const [renaming, setRenaming] = useState(false);
  const [sheetFilter, setSheetFilter] = useState<string>(ALL_SHEETS);

  const steps = branch.steps.filter((s) => sheetFilter === ALL_SHEETS || s.sheetId === sheetFilter);
  let maxMag = 1;
  for (const s of steps) maxMag = Math.max(maxMag, s.magnitude);
  const sliderValue = Math.max(0, Math.min(lastIndex, headIndex));

  const scrubTo = (index: number): void => {
    if (index >= lastIndex) dispatch({ type: 'returnToPresent' });
    else dispatch({ type: 'goto', ref: { branchId: branch.id, stepIndex: index } });
  };

  const commitRename = (name: string): void => {
    setRenaming(false);
    const trimmed = name.trim();
    if (trimmed) dispatch({ type: 'renameBranch', branchId: branch.id, name: trimmed });
  };

  return (
    <section
      className="timeline-pane"
      data-theme={theme}
      aria-label="Parametric timeline"
      style={{ background: t.bg, color: t.ink, fontSize: 12 }}
    >
      <header
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          marginBottom: 10,
        }}
      >
        <h2 style={{ fontSize: 13, margin: 0, fontWeight: 700 }}>Parametric Timeline</h2>
        <span
          data-mode={view.head.mode}
          style={{
            fontSize: 11,
            fontWeight: 600,
            borderRadius: 999,
            padding: '3px 9px',
            color: inPreview ? '#9a3412' : '#065f46',
            background: inPreview ? '#fff7ed' : '#ecfdf5',
            border: `1px solid ${inPreview ? '#fed7aa' : '#a7f3d0'}`,
          }}
        >
          <span>{inPreview ? `Preview · Step ${String(headIndex)}` : 'Present'}</span> ·{' '}
          {branch.name ?? branch.id}
        </span>
      </header>

      {inPreview ? (
        <div
          role="status"
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            gap: 8,
            marginBottom: 10,
            padding: '7px 10px',
            borderRadius: 8,
            fontSize: 11,
            fontWeight: 600,
            color: '#9a3412',
            background: '#fff7ed',
            border: '1px solid #fdba74',
          }}
        >
          <span>
            Preview — viewing step {String(headIndex)}. Your live sheets are hidden and read-only.
          </span>
          <button
            type="button"
            onClick={() => {
              dispatch({ type: 'returnToPresent' });
            }}
            style={{
              ...btn(t),
              flexShrink: 0,
              color: '#7c2d12',
              background: '#ffedd5',
              border: '1px solid #fdba74',
            }}
          >
            Exit preview
          </button>
        </div>
      ) : null}

      <div
        style={{
          display: 'flex',
          flexWrap: 'wrap',
          gap: 6,
          marginBottom: 10,
          alignItems: 'center',
        }}
      >
        {view.branches.map((b) => (
          <button
            key={b.id}
            type="button"
            aria-pressed={b.id === branch.id}
            aria-label={`Switch to ${b.name ?? b.id} branch`}
            onClick={() => {
              dispatch({ type: 'switch', branchId: b.id });
            }}
            style={{
              fontSize: 11,
              fontWeight: 600,
              padding: '4px 9px',
              borderRadius: 8,
              cursor: 'pointer',
              border: `1px solid ${t.line ?? '#e5e7eb'}`,
              background: b.id === branch.id ? (t.accent ?? '#111') : (t.chip ?? '#f3f4f6'),
              color: b.id === branch.id ? (t.bg ?? '#fff') : (t.ink ?? '#1f2937'),
            }}
          >
            {b.name ?? b.id}
            {b.provisional ? ' •' : ''}
          </button>
        ))}
        <select
          aria-label="Worksheet drill-down"
          value={sheetFilter}
          onChange={(e: ChangeEvent<HTMLSelectElement>) => {
            setSheetFilter(e.target.value);
          }}
          style={{
            marginLeft: 'auto',
            fontSize: 11,
            padding: '3px 6px',
            borderRadius: 6,
            border: `1px solid ${t.line ?? '#e5e7eb'}`,
            background: t.bg,
            color: t.ink,
          }}
        >
          <option value={ALL_SHEETS}>Whole workbook</option>
          {view.sheets.map((s) => (
            <option key={s} value={s}>
              {s}
            </option>
          ))}
        </select>
      </div>

      {/* operation-icon strip (the hero): icon = what, underline = how much */}
      <div
        className="timeline-strip"
        aria-label={`Operations on ${branch.name ?? branch.id}`}
        style={{
          display: 'flex',
          alignItems: 'flex-end',
          gap: 2,
          overflowX: 'auto',
          padding: '8px 2px',
          borderBottom: `1px solid ${t.line ?? '#e5e7eb'}`,
          minHeight: 72,
        }}
      >
        {steps.length === 0 ? (
          <div
            style={{
              flex: 1,
              textAlign: 'center',
              padding: '14px 16px',
              color: t.muted ?? '#6b7280',
              fontSize: 12,
            }}
          >
            No tracked changes yet — edit a cell and it will appear here.
          </div>
        ) : null}
        {steps.map((s) => {
          const ahead = s.index > headIndex;
          const isHead = s.index === headIndex;
          return (
            <button
              key={s.index}
              type="button"
              className="timeline-step"
              data-index={s.index}
              data-kind={s.kind}
              data-op={s.op}
              data-mag={s.magnitude}
              aria-current={isHead}
              aria-label={`Step ${String(s.index)}: ${s.label ?? KIND_LABEL[s.kind]}`}
              onClick={() => {
                scrubTo(s.index);
              }}
              onMouseEnter={() => {
                setHover(s);
              }}
              onFocus={() => {
                setHover(s);
              }}
              onMouseLeave={() => {
                setHover(null);
              }}
              onBlur={() => {
                setHover(null);
              }}
              style={{
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
                gap: 4,
                flex: '0 0 30px',
                padding: '2px 0',
                border: 0,
                borderRadius: 7,
                cursor: 'pointer',
                background: isHead ? `${KIND_COLOR[s.kind]}22` : 'transparent',
              }}
            >
              <span
                aria-hidden
                style={{
                  width: 22,
                  height: 22,
                  display: 'grid',
                  placeItems: 'center',
                  borderRadius: 6,
                  color: KIND_COLOR[s.kind],
                  background: `${KIND_COLOR[s.kind]}22`,
                  border: `1px solid ${KIND_COLOR[s.kind]}59`,
                  opacity: ahead ? 0.32 : 1,
                }}
              >
                <OperationIcon op={s.op} />
              </span>
              <span
                aria-hidden
                style={{
                  width: 16,
                  height: barHeight(s.magnitude, maxMag, 22),
                  borderRadius: '2px 2px 0 0',
                  background: KIND_COLOR[s.kind],
                  opacity: ahead ? 0.22 : 0.55,
                }}
              />
            </button>
          );
        })}
      </div>

      {/* accessible scrubber: keyboard + screen-reader + the test surface for scrubbing */}
      <input
        type="range"
        className="timeline-scrubber"
        aria-label="Timeline scrubber"
        min={0}
        max={lastIndex}
        step={1}
        value={sliderValue}
        onChange={(e: ChangeEvent<HTMLInputElement>) => {
          scrubTo(Number(e.target.value));
        }}
        style={{ width: '100%', marginTop: 6, accentColor: t.accent }}
      />
      <div
        style={{ display: 'flex', justifyContent: 'space-between', fontSize: 10.5, color: t.muted }}
      >
        <span>Step 0</span>
        <span>Present ▸</span>
      </div>

      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 10 }}>
        {inPreview ? (
          <>
            <button
              type="button"
              onClick={() => {
                dispatch({ type: 'returnToPresent' });
              }}
              style={btn(t)}
            >
              Return to Present
            </button>
            <button
              type="button"
              onClick={() => {
                dispatch({ type: 'branch', from: { branchId: branch.id, stepIndex: headIndex } });
              }}
              style={btn(t)}
            >
              Branch from here
            </button>
          </>
        ) : null}
        {renaming ? (
          <input
            aria-label="Branch name"
            defaultValue={branch.name ?? branch.id}
            onBlur={(e) => {
              commitRename(e.currentTarget.value);
            }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') commitRename(e.currentTarget.value);
              if (e.key === 'Escape') setRenaming(false);
            }}
            style={{
              fontSize: 11,
              padding: '3px 6px',
              borderRadius: 6,
              border: `1px solid ${t.line ?? '#e5e7eb'}`,
              background: t.bg,
              color: t.ink,
            }}
          />
        ) : (
          <button
            type="button"
            onClick={() => {
              setRenaming(true);
            }}
            style={btn(t)}
            aria-label={`Rename ${branch.name ?? branch.id}`}
          >
            Rename
          </button>
        )}
        {branch.id !== 'main' ? (
          <button
            type="button"
            onClick={() => {
              dispatch({ type: 'deleteBranch', branchId: branch.id });
            }}
            style={btn(t)}
            aria-label={`Delete ${branch.name ?? branch.id}`}
          >
            Delete branch
          </button>
        ) : null}
      </div>

      <div
        style={{
          display: 'flex',
          flexWrap: 'wrap',
          gap: 10,
          marginTop: 12,
          fontSize: 10.5,
          color: t.muted,
        }}
      >
        {(Object.keys(KIND_LABEL) as StepKind[]).map((k) => (
          <span key={k}>
            <i
              style={{
                display: 'inline-block',
                width: 9,
                height: 9,
                borderRadius: 2,
                marginRight: 4,
                background: KIND_COLOR[k],
                verticalAlign: -1,
              }}
            />
            {KIND_LABEL[k]}
          </span>
        ))}
      </div>

      <div className="timeline-inspector" style={{ marginTop: 12, minHeight: 32, fontSize: 12 }}>
        {hover ? (
          <span>
            <strong>Step {hover.index}</strong>{' '}
            <span style={{ color: t.muted }}>· {hover.label ?? KIND_LABEL[hover.kind]}</span>
          </span>
        ) : (
          <span style={{ color: t.muted }}>
            {inPreview ? `Previewing step ${String(headIndex)}` : 'At present'} — hover a step for
            details
          </span>
        )}
      </div>
    </section>
  );
}
