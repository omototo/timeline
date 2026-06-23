import { useRef, useState } from 'react';
import type { ChangeEvent, CSSProperties, PointerEvent as ReactPointerEvent } from 'react';
import type {
  BranchId,
  StepKind,
  TimelineBranch,
  TimelinePaneProps,
  TimelineStep,
  TimelineTheme,
} from './contract.ts';

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

  const [zoom, setZoom] = useState<{ start: number; end: number }>({ start: 0, end: lastIndex });
  const [hover, setHover] = useState<TimelineStep | null>(null);
  const [renaming, setRenaming] = useState(false);
  const [sheetFilter, setSheetFilter] = useState<string>(ALL_SHEETS);
  const trackRef = useRef<HTMLDivElement>(null);
  const brushX0 = useRef<number | null>(null);
  const [brushRect, setBrushRect] = useState<{ left: number; width: number } | null>(null);

  const vs = Math.max(0, Math.min(zoom.start, lastIndex));
  const ve = Math.max(vs, Math.min(zoom.end, lastIndex));
  const span = Math.max(1, ve - vs);
  const zoomed = vs > 0 || ve < lastIndex;
  const visible = branch.steps.filter(
    (s) =>
      s.index >= vs && s.index <= ve && (sheetFilter === ALL_SHEETS || s.sheetId === sheetFilter),
  );
  let maxMag = 1;
  for (const s of visible) maxMag = Math.max(maxMag, s.magnitude);

  const pctOf = (index: number): number => ((index - vs) / span) * 100;
  const headInWindow = headIndex >= vs && headIndex <= ve;
  const sliderValue = Math.max(vs, Math.min(ve, headIndex));

  const scrubTo = (index: number): void => {
    if (index >= lastIndex) dispatch({ type: 'returnToPresent' });
    else dispatch({ type: 'goto', ref: { branchId: branch.id, stepIndex: index } });
  };
  const resetZoom = (): void => {
    setZoom({ start: 0, end: lastIndex });
  };

  // Grafana brush-zoom over the bars (mouse enhancement; keyboard users scrub via the range + Reset button).
  const onPointerDown = (e: ReactPointerEvent<HTMLDivElement>): void => {
    const rect = trackRef.current?.getBoundingClientRect();
    if (!rect) return;
    brushX0.current = e.clientX - rect.left;
    setBrushRect({ left: brushX0.current, width: 0 });
  };
  const onPointerMove = (e: ReactPointerEvent<HTMLDivElement>): void => {
    const rect = trackRef.current?.getBoundingClientRect();
    if (!rect || brushX0.current === null) return;
    const x = e.clientX - rect.left;
    setBrushRect({ left: Math.min(brushX0.current, x), width: Math.abs(x - brushX0.current) });
  };
  const onPointerUp = (e: ReactPointerEvent<HTMLDivElement>): void => {
    const rect = trackRef.current?.getBoundingClientRect();
    setBrushRect(null);
    if (rect && brushX0.current !== null) {
      const win = brushToWindow(brushX0.current, e.clientX - rect.left, rect.width, vs, ve);
      if (win) setZoom(win);
    }
    brushX0.current = null;
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

      {/* histogram (the hero) — drag across to zoom, double-click to reset */}
      <div
        ref={trackRef}
        className="timeline-bars"
        aria-label={`Histogram of ${branch.name ?? branch.id}, steps ${String(vs)} to ${String(ve)}`}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onDoubleClick={resetZoom}
        style={{
          position: 'relative',
          height: 96,
          borderBottom: `1px solid ${t.line ?? '#e5e7eb'}`,
          touchAction: 'none',
          cursor: 'crosshair',
        }}
      >
        {visible.map((s) => (
          <button
            key={s.index}
            type="button"
            className="timeline-bar"
            data-index={s.index}
            data-kind={s.kind}
            aria-label={`Step ${String(s.index)}: ${KIND_LABEL[s.kind]}, ${String(s.magnitude)} cells${s.label ? `, ${s.label}` : ''}`}
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
              position: 'absolute',
              bottom: 0,
              left: `calc(${String(pctOf(s.index))}% - 5px)`,
              width: 10,
              height: barHeight(s.magnitude, maxMag, 86),
              padding: 0,
              border: 0,
              borderRadius: '3px 3px 0 0',
              background: KIND_COLOR[s.kind],
              opacity: s.index > headIndex ? 0.28 : 1,
              cursor: 'pointer',
            }}
          />
        ))}
        {headInWindow ? (
          <div
            aria-hidden
            style={{
              position: 'absolute',
              top: -4,
              bottom: 0,
              left: `${String(pctOf(headIndex))}%`,
              width: 2,
              marginLeft: -1,
              background: t.accent,
              pointerEvents: 'none',
            }}
          >
            <div
              style={{
                position: 'absolute',
                top: -7,
                left: -6,
                width: 12,
                height: 12,
                borderRadius: '50%',
                background: t.accent,
              }}
            />
          </div>
        ) : null}
        {brushRect ? (
          <div
            aria-hidden
            style={{
              position: 'absolute',
              top: 0,
              bottom: 0,
              left: brushRect.left,
              width: brushRect.width,
              background: 'rgba(147,164,255,0.18)',
              borderLeft: '1px solid #93a4ff',
              borderRight: '1px solid #93a4ff',
              pointerEvents: 'none',
            }}
          />
        ) : null}
      </div>

      {/* accessible scrubber: keyboard + screen-reader + the test surface for scrubbing */}
      <input
        type="range"
        className="timeline-scrubber"
        aria-label="Timeline scrubber"
        min={vs}
        max={ve}
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
        <span>Step {vs}</span>
        <span>
          {zoomed ? `viewing ${String(vs)}–${String(ve)} of ${String(lastIndex)}` : 'Present ▸'}
        </span>
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
        {zoomed ? (
          <button type="button" onClick={resetZoom} style={btn(t)}>
            ⤢ Reset zoom
          </button>
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
            <span style={{ color: t.muted }}>
              · {KIND_LABEL[hover.kind]} · Δ {hover.magnitude.toLocaleString()} cells
            </span>
            {hover.label ? <span style={{ color: t.muted }}> · {hover.label}</span> : null}
          </span>
        ) : (
          <span style={{ color: t.muted }}>
            {inPreview ? `Previewing step ${String(headIndex)}` : 'At present'} — hover a bar for
            details
          </span>
        )}
      </div>
    </section>
  );
}
