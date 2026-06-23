/**
 * Wave 4 — branching, switching, and the attach lifecycle (ADR-0005, ADR-0006,
 * ADR-0013).
 *
 * Drives the engine's `branch`/`switch`/`attach`/`detachToCoauthoring` verbs and
 * asserts on the returned {@link EffectEnvelope}s and pure engine queries — no
 * fakes, no I/O (the functional core returns effect descriptions; ADR-0013).
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { TimelineEngineImpl } from '../src/index.ts';
import type {
  Area,
  BranchMeta,
  CellSlab,
  CellState,
  EffectEnvelope,
  PersistOp,
  PersistedHead,
  ReconcileOp,
  ReconciliationDelta,
  Rect,
  StepRef,
  ValueObservation,
  WorkbookSnapshot,
} from '../src/index.ts';

// --- builders ---------------------------------------------------------------

function cellRect(row: number, col: number): Rect {
  return { startRow: row, startCol: col, rowCount: 1, colCount: 1 };
}

function state(partial: Partial<CellState> & { value: unknown }): CellState {
  return {
    value: partial.value,
    formula: partial.formula ?? null,
    valueType: partial.valueType ?? 'string',
    numberFormat: partial.numberFormat ?? 'General',
  };
}

function slabFromStates(rows: CellState[][]): CellSlab {
  return {
    values: rows.map((r) => r.map((c) => c.value)),
    formulas: rows.map((r) => r.map((c) => c.formula)),
    numberFormats: rows.map((r) => r.map((c) => c.numberFormat)),
    valueTypes: rows.map((r) => r.map((c) => c.valueType)),
  };
}

function valueObs(sheetId: string, area: Area, rows: CellState[][]): ValueObservation {
  return {
    kind: 'value',
    triggerSource: 'thisLocalAddin',
    source: 'local',
    sheetId,
    area,
    after: slabFromStates(rows),
  };
}

function setCell(
  engine: TimelineEngineImpl,
  sheet: string,
  row: number,
  col: number,
  value: string,
): EffectEnvelope {
  return engine.ingest(valueObs(sheet, [cellRect(row, col)], [[state({ value })]]));
}

// --- envelope extractors ----------------------------------------------------

function persistOpsOf<K extends PersistOp['op']>(
  env: EffectEnvelope,
  op: K,
): Extract<PersistOp, { op: K }>[] {
  return (env.persist ?? []).filter((p): p is Extract<PersistOp, { op: K }> => p.op === op);
}

function reconcileOpsOf<K extends ReconcileOp['op']>(
  env: EffectEnvelope,
  op: K,
): Extract<ReconcileOp, { op: K }>[] {
  return (env.reconcile?.ops ?? []).filter(
    (o): o is Extract<ReconcileOp, { op: K }> => o.op === op,
  );
}

const ref = (branchId: string, stepIndex: number): StepRef => ({ branchId, stepIndex });

// ===========================================================================

describe('branch — fork a provisional branch + promote to editable Present', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl({ keyframeStepInterval: 1000, keyframeByteThreshold: 1e9 });
    setCell(engine, 'Sheet1', 0, 0, 'a'); // main step 0
    setCell(engine, 'Sheet1', 0, 0, 'b'); // main step 1
  });

  it('flips HEAD to a new provisional branch in present, emits only setHead (no saveBranch)', () => {
    const env = engine.branch(ref('main', 0));

    const head = engine.head();
    expect(head.mode).toBe('present');
    expect(head.branchId).not.toBe('main');
    expect(head.branchId).toBe('branch-1');

    // Provisional: NOT persisted yet -> setHead only, no saveBranch.
    expect(persistOpsOf(env, 'setHead')).toHaveLength(1);
    expect(persistOpsOf(env, 'saveBranch')).toHaveLength(0);
    expect(env.reconcile).toBeUndefined();

    const meta = engine.branches().find((b) => b.id === 'branch-1');
    expect(meta?.provisional).toBe(true);
    expect(meta?.parentBranchId).toBe('main');
    expect(meta?.forkedAt).toEqual({ branchId: 'main', stepIndex: 0 });
  });

  it('reconstructs the fork-point state as the new branch live Shadow State', () => {
    // Fork at main step 0 (A1 = "a"), not the tip (A1 = "b").
    engine.branch(ref('main', 0));
    expect(engine.readShadow('Sheet1', 0, 0).value).toBe('a');
    expect(engine.tipStepIndex()).toBe(-1); // new branch has no Steps yet
  });

  it('persists (saveBranch) only on the first ingest, and promotes off provisional', () => {
    engine.branch(ref('main', 0));

    const env = setCell(engine, 'Sheet1', 1, 1, 'forked');
    const saves = persistOpsOf(env, 'saveBranch');
    expect(saves).toHaveLength(1);
    expect(saves[0]?.meta.id).toBe('branch-1');
    expect(saves[0]?.meta.provisional).toBe(false);

    // The Step itself was recorded on the new branch.
    expect(engine.tipStepIndex('branch-1')).toBe(0);
    expect(engine.steps('branch-1')).toHaveLength(1);

    // A SECOND ingest does NOT re-emit saveBranch.
    const env2 = setCell(engine, 'Sheet1', 2, 2, 'again');
    expect(persistOpsOf(env2, 'saveBranch')).toHaveLength(0);

    // The branch is no longer provisional.
    expect(engine.branches().find((b) => b.id === 'branch-1')?.provisional).toBe(false);
  });

  it('edits on the branch do not touch the parent main log', () => {
    engine.branch(ref('main', 1));
    setCell(engine, 'Sheet1', 5, 5, 'onlyBranch');
    expect(engine.steps('main')).toHaveLength(2); // unchanged
    expect(engine.steps('branch-1')).toHaveLength(1);
  });

  it('forking from an unregistered non-main branch records that id as the parent', () => {
    // `from.branchId` is a branch the engine has never seen; the fork still
    // records it as the parent verbatim (the shell vouches for the StepRef).
    const env = engine.branch(ref('ghost-branch', 0));
    expect(persistOpsOf(env, 'setHead')).toHaveLength(1);
    const meta = engine.branches().find((b) => b.id === 'branch-1');
    expect(meta?.parentBranchId).toBe('ghost-branch');
  });

  it('mints distinct ids for successive forks', () => {
    engine.branch(ref('main', 0));
    setCell(engine, 'Sheet1', 3, 3, 'keep1'); // promote branch-1 so it survives switch-away
    engine.switch('main');
    engine.branch(ref('main', 1));
    const ids = engine.branches().map((b) => b.id);
    expect(ids).toContain('branch-1');
    expect(ids).toContain('branch-2');
  });
});

describe('switch — non-destructive checkout, NOT a Step', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl({ keyframeStepInterval: 1000, keyframeByteThreshold: 1e9 });
    setCell(engine, 'Sheet1', 0, 0, 'main-a'); // main step 0
    engine.branch(ref('main', 0)); // -> branch-1, A1 = 'main-a'
    setCell(engine, 'Sheet1', 0, 0, 'branch-a'); // branch-1 step 0
    setCell(engine, 'Sheet1', 1, 0, 'branch-b'); // branch-1 step 1
  });

  it('checks out the target tip onto realSheet in formula mode, no appendDelta', () => {
    const env = engine.switch('main');

    expect(env.reconcile?.target).toBe('realSheet');
    // Every cell op is formula mode (live Present, not frozen values).
    const cells = reconcileOpsOf(env, 'setCells');
    for (const op of cells) expect(op.mode).toBe('formula');

    // Ops carry the LOGICAL sheet id (real worksheet), never a preview surface.
    for (const op of cells) expect(op.sheetId).toBe('Sheet1');

    // NAVIGATION, not a Step: no appendDelta anywhere.
    expect(persistOpsOf(env, 'appendDelta')).toHaveLength(0);
    expect(persistOpsOf(env, 'setHead')).toHaveLength(1);

    // HEAD now on main, present.
    expect(engine.head()).toEqual({ branchId: 'main', mode: 'present' });
    // Shadow State reconstructed to main's tip: A1='main-a', A2 cleared.
    expect(engine.readShadow('Sheet1', 0, 0).value).toBe('main-a');
    expect(engine.readShadow('Sheet1', 1, 0).valueType).toBe('empty');
  });

  it('writes the minimal formula-mode diff between the two tips', () => {
    const env = engine.switch('main');
    const cells = reconcileOpsOf(env, 'setCells');
    // branch tip: A1='branch-a', A2='branch-b'; main tip: A1='main-a', A2 empty.
    // Minimal diff: A1 -> 'main-a' (write), A2 -> empty (clear). Two ops.
    expect(cells).toHaveLength(2);
    const at = (c: (typeof cells)[number]): Rect | undefined => c.area[0];
    const a1 = cells.find((c) => at(c)?.startRow === 0 && at(c)?.startCol === 0);
    const a2 = cells.find((c) => at(c)?.startRow === 1 && at(c)?.startCol === 0);
    expect(a1?.slab.values[0]?.[0]).toBe('main-a');
    expect(a2?.slab.values[0]?.[0]).toBe('');
    expect(a2?.slab.valueTypes[0]?.[0]).toBe('empty');
  });

  it('switching back and forth is non-destructive (both branch logs intact)', () => {
    engine.switch('main');
    engine.switch('branch-1');
    expect(engine.steps('main')).toHaveLength(1);
    expect(engine.steps('branch-1')).toHaveLength(2);
    // Back on branch-1 tip.
    expect(engine.readShadow('Sheet1', 0, 0).value).toBe('branch-a');
    expect(engine.readShadow('Sheet1', 1, 0).value).toBe('branch-b');
    expect(engine.head()).toEqual({ branchId: 'branch-1', mode: 'present' });
  });

  it('switching to the current branch is a no-op (empty envelope)', () => {
    const env = engine.switch('branch-1');
    expect(env).toEqual({});
  });

  it('ingest after switch records on the now-current branch', () => {
    engine.switch('main');
    const env = setCell(engine, 'Sheet1', 9, 9, 'onMain');
    expect(persistOpsOf(env, 'appendDelta')[0]?.branchId).toBe('main');
    expect(engine.tipStepIndex('main')).toBe(1);
  });
});

describe('switch — provisional GC on switch-away', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl({ keyframeStepInterval: 1000, keyframeByteThreshold: 1e9 });
    setCell(engine, 'Sheet1', 0, 0, 'a'); // main step 0
  });

  it('discards a zero-Step provisional branch when switching away (deleteBranch)', () => {
    engine.branch(ref('main', 0)); // branch-1, provisional, zero Steps
    expect(engine.hasBranch('branch-1')).toBe(true);

    const env = engine.switch('main');
    // The abandoned fork is GC'd: a deleteBranch op + it is dropped resident.
    const deletes = persistOpsOf(env, 'deleteBranch');
    expect(deletes).toHaveLength(1);
    expect(deletes[0]?.branchId).toBe('branch-1');
    expect(engine.hasBranch('branch-1')).toBe(false);
  });

  it('does NOT GC a provisional branch that has recorded a Step', () => {
    engine.branch(ref('main', 0));
    setCell(engine, 'Sheet1', 1, 1, 'kept'); // promotes branch-1
    const env = engine.switch('main');
    expect(persistOpsOf(env, 'deleteBranch')).toHaveLength(0);
    expect(engine.hasBranch('branch-1')).toBe(true);
  });

  it('does NOT GC a non-provisional (main) branch', () => {
    engine.branch(ref('main', 0));
    setCell(engine, 'Sheet1', 1, 1, 'x'); // branch-1 now real
    engine.switch('main'); // leaving real branch-1: no GC
    const env = engine.switch('branch-1'); // leaving main (non-provisional): no GC
    expect(persistOpsOf(env, 'deleteBranch')).toHaveLength(0);
    expect(engine.hasBranch('main') || engine.steps('main').length > 0).toBe(true);
  });
});

describe('attach — clean resume vs drift (ADR-0006)', () => {
  function snapshot(
    contentHash: string,
    sheets: { sheetId: string; rows: CellState[][] }[],
  ): WorkbookSnapshot {
    return {
      workbookGuid: 'wb-guid',
      contentHash,
      sheets: sheets.map((s) => ({ sheetId: s.sheetId, slab: slabFromStates(s.rows) })),
    };
  }

  it('fresh workbook (no persisted head): seeds the mirror, no Step, empty reconcile', () => {
    const engine = new TimelineEngineImpl();
    const env = engine.attach(
      snapshot('h0', [{ sheetId: 'Sheet1', rows: [[state({ value: 'seed' })]] }]),
      null,
    );
    expect(env.reconcile).toBeUndefined();
    expect(env.persist ?? []).toHaveLength(0);
    // The mirror now holds the observed content so future ingests diff against it.
    expect(engine.readShadow('Sheet1', 0, 0).value).toBe('seed');
    expect(engine.steps('main')).toHaveLength(0);
  });

  it('clean match: restores HEAD, no writes, no Step', () => {
    const engine = new TimelineEngineImpl();
    const persisted: PersistedHead = {
      head: { branchId: 'main', mode: 'present' },
      tipHash: 'tip-123',
    };
    const env = engine.attach(snapshot('tip-123', []), persisted);

    expect(env.reconcile).toBeUndefined();
    // Clean resume just restores HEAD.
    expect(persistOpsOf(env, 'setHead')).toHaveLength(1);
    expect(persistOpsOf(env, 'appendDelta')).toHaveLength(0);
    expect(engine.head()).toEqual({ branchId: 'main', mode: 'present' });
    expect(engine.steps('main')).toHaveLength(0);
  });

  it('clean match restores a persisted non-main branch head', () => {
    const engine = new TimelineEngineImpl();
    const persisted: PersistedHead = {
      head: { branchId: 'branch-7', mode: 'present' },
      tipHash: 'same',
    };
    engine.attach(snapshot('same', []), persisted);
    expect(engine.head()).toEqual({ branchId: 'branch-7', mode: 'present' });
  });

  it('clean match restores a persisted PREVIEW head verbatim (preview index kept)', () => {
    const engine = new TimelineEngineImpl();
    const persisted: PersistedHead = {
      head: { branchId: 'main', mode: 'preview', previewStepIndex: 4 },
      tipHash: 'h',
    };
    engine.attach(snapshot('h', []), persisted);
    expect(engine.head()).toEqual({ branchId: 'main', mode: 'preview', previewStepIndex: 4 });
  });

  it('drift: appends a Reconciliation Step with an itemized per-sheet before/after diff', () => {
    const engine = new TimelineEngineImpl();
    // The engine last witnessed A1='old' on Sheet1.
    setCell(engine, 'Sheet1', 0, 0, 'old');

    // Observe a DRIFTED workbook: A1='new' (changed) and B2='added' (new), with
    // a tip hash that does not match.
    const observed = snapshot('drifted-hash', [
      {
        sheetId: 'Sheet1',
        rows: [
          [state({ value: 'new' }), state({ value: '', valueType: 'empty' })],
          [state({ value: '', valueType: 'empty' }), state({ value: 'added' })],
        ],
      },
    ]);
    const persisted: PersistedHead = {
      head: { branchId: 'main', mode: 'present' },
      tipHash: 'tip-pre-drift',
    };

    const env = engine.attach(observed, persisted);

    // A single Reconciliation Step was appended.
    const append = persistOpsOf(env, 'appendDelta')[0];
    expect(append).toBeDefined();
    const delta = append?.delta as ReconciliationDelta;
    expect(delta.kind).toBe('reconciliation');

    const sheet1 = delta.perSheet.find((s) => s.sheetId === 'Sheet1');
    expect(sheet1).toBeDefined();
    // A1: 'old' -> 'new'; B2: empty -> 'added'. Two itemized cells.
    const a1 = sheet1?.cells.find((c) => c.addr.startRow === 0 && c.addr.startCol === 0);
    const b2 = sheet1?.cells.find((c) => c.addr.startRow === 1 && c.addr.startCol === 1);
    expect(a1?.before.value).toBe('old');
    expect(a1?.after.value).toBe('new');
    expect(b2?.before.valueType).toBe('empty');
    expect(b2?.after.value).toBe('added');

    // The Step advanced history and the mirror is now the observed state.
    expect(engine.steps('main')).toHaveLength(2); // original + reconciliation
    expect(engine.readShadow('Sheet1', 0, 0).value).toBe('new');
    expect(engine.readShadow('Sheet1', 1, 1).value).toBe('added');
  });

  it('drift records a cleared cell (present before, absent after)', () => {
    const engine = new TimelineEngineImpl();
    setCell(engine, 'Sheet1', 0, 0, 'willClear');
    // Observed workbook has Sheet1 entirely empty.
    const observed = snapshot('h', [{ sheetId: 'Sheet1', rows: [] }]);
    const persisted: PersistedHead = {
      head: { branchId: 'main', mode: 'present' },
      tipHash: 'no-match',
    };
    const env = engine.attach(observed, persisted);
    const delta = persistOpsOf(env, 'appendDelta')[0]?.delta as ReconciliationDelta;
    const cell = delta.perSheet[0]?.cells[0];
    expect(cell?.before.value).toBe('willClear');
    expect(cell?.after.valueType).toBe('empty');
    expect(engine.readShadow('Sheet1', 0, 0).valueType).toBe('empty');
  });

  it('clears co-authoring suspension on a clean re-attach', () => {
    const engine = new TimelineEngineImpl();
    engine.detachToCoauthoring();
    expect(engine.isSuspended()).toBe(true);
    engine.attach(snapshot('x', []), { head: { branchId: 'main', mode: 'present' }, tipHash: 'x' });
    expect(engine.isSuspended()).toBe(false);
    // Tracking resumes: ingest records again.
    const env = setCell(engine, 'Sheet1', 0, 0, 'resumed');
    expect(persistOpsOf(env, 'appendDelta')).toHaveLength(1);
  });
});

describe('detachToCoauthoring — suspend tracking (ADR-0006)', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl();
  });

  it('sets suspended-tracking and returns a diagnostic envelope (no Step)', () => {
    const env = engine.detachToCoauthoring();
    expect(env).toEqual({});
    expect(engine.isSuspended()).toBe(true);
    expect(engine.lastDiagnostic()).toEqual({
      code: 'coauthoringSuspended',
      message: 'Co-authoring detected (source: remote); tracking suspended for this session.',
    });
  });

  it('subsequent ingest is a no-op while suspended (+ diagnostic)', () => {
    engine.detachToCoauthoring();
    const env = setCell(engine, 'Sheet1', 0, 0, 'ignored');
    expect(env).toEqual({});
    expect(engine.steps('main')).toHaveLength(0);
    expect(engine.readShadow('Sheet1', 0, 0).valueType).toBe('empty');
    expect(engine.lastDiagnostic()?.code).toBe('ingestSuspended');
  });
});

describe('mode validity — invalid ops are no-ops, never corrupt (state machine)', () => {
  it('ingest in Preview is refused with a diagnostic (no Step)', () => {
    const engine = new TimelineEngineImpl({
      keyframeStepInterval: 1000,
      keyframeByteThreshold: 1e9,
    });
    setCell(engine, 'Sheet1', 0, 0, 'a'); // step 0
    engine.goto(ref('main', 0)); // enter Preview
    const env = setCell(engine, 'Sheet1', 1, 1, 'blocked');
    expect(env).toEqual({});
    expect(engine.lastDiagnostic()?.code).toBe('ingestInPreview');
    expect(engine.steps('main')).toHaveLength(1);
  });

  it('branch from Preview implicitly returns to Present (a fork is a Present op)', () => {
    const engine = new TimelineEngineImpl({
      keyframeStepInterval: 1000,
      keyframeByteThreshold: 1e9,
    });
    setCell(engine, 'Sheet1', 0, 0, 'a'); // step 0
    engine.goto(ref('main', 0));
    expect(engine.head().mode).toBe('preview');

    engine.branch(ref('main', 0));
    expect(engine.head().mode).toBe('present');
    expect(engine.head().branchId).toBe('branch-1');
  });
});

describe('branches() query — the resident branch graph', () => {
  it('returns branch metas in tab order with fork lineage', () => {
    const engine = new TimelineEngineImpl({
      keyframeStepInterval: 1000,
      keyframeByteThreshold: 1e9,
    });
    setCell(engine, 'Sheet1', 0, 0, 'a');
    engine.branch(ref('main', 0));
    setCell(engine, 'Sheet1', 1, 1, 'b'); // promote branch-1

    const metas: BranchMeta[] = engine.branches();
    const main = metas.find((m) => m.id === 'main');
    const b1 = metas.find((m) => m.id === 'branch-1');
    expect(main?.provisional).toBe(false);
    expect(b1?.parentBranchId).toBe('main');
    expect(b1?.provisional).toBe(false);
  });
});
