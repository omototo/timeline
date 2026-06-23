/**
 * `IndexedDbStore` — the IndexedDB-backed `HistoryStore` (ADR-0007).
 *
 * Lives in the addin (host) package: IndexedDB and `navigator` are host
 * globals the pure engine may not touch. Semantics are IDENTICAL to the
 * engine's `InMemoryStore` reference (`packages/engine/src/in-memory-store.ts`)
 * — only the backing store differs:
 *
 * - deltas: one record per (branch, stepIndex) in an append-only object store
 *   keyed by `[branchId, stepIndex]`; `loadDeltas` is the inclusive [from, to]
 *   range read from that branch.
 * - keyframes: one record per (branch, stepIndex); `loadKeyframeAtOrBefore`
 *   returns the highest keyframe whose stepIndex is <= the requested index.
 * - head: a single nullable record under a fixed key.
 * - branches: one record per branch id; `listBranches` is sorted by `order`.
 *
 * IndexedDB is the primary working store (offline-first); `init()` opens the
 * DB and best-effort requests `navigator.storage.persist()` to reduce the
 * chance the host WebView evicts history.
 */
import type { HistoryStore } from '@timeline/engine';
import type { BranchId, BranchMeta, Delta, Head } from '@timeline/engine';

const DB_NAME = 'timeline-history';
const DB_VERSION = 1;

/** Per-workbook database name so two workbooks never share one history. */
export function databaseNameFor(workbookKey: string | null): string {
  if (workbookKey === null || workbookKey === '') {
    return DB_NAME;
  }
  let hash = 0x811c9dc5;
  for (let i = 0; i < workbookKey.length; i += 1) {
    hash ^= workbookKey.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return `${DB_NAME}-${(hash >>> 0).toString(16).padStart(8, '0')}`;
}

const STORE_DELTAS = 'deltas';
const STORE_KEYFRAMES = 'keyframes';
const STORE_HEAD = 'head';
const STORE_BRANCHES = 'branches';

/** The fixed key under which the single HEAD record is stored. */
const HEAD_KEY = 'head';

interface DeltaRecord {
  branchId: BranchId;
  stepIndex: number;
  delta: Delta;
}

interface KeyframeRecord {
  branchId: BranchId;
  stepIndex: number;
  state: unknown;
}

interface HeadRecord {
  key: string;
  head: Head;
}

/** Wraps an IDBRequest in a Promise that settles on success/error. */
function promisifyRequest<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    request.onsuccess = (): void => {
      resolve(request.result);
    };
    request.onerror = (): void => {
      reject(request.error ?? new Error('IndexedDB request failed'));
    };
  });
}

/** Resolves when a transaction completes; rejects on error/abort. */
function awaitTransaction(tx: IDBTransaction): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    tx.oncomplete = (): void => {
      resolve();
    };
    tx.onerror = (): void => {
      reject(tx.error ?? new Error('IndexedDB transaction failed'));
    };
    tx.onabort = (): void => {
      reject(tx.error ?? new Error('IndexedDB transaction aborted'));
    };
  });
}

export class IndexedDbStore implements HistoryStore {
  readonly #factory: IDBFactory;
  readonly #dbName: string;
  #db: IDBDatabase | null = null;

  /**
   * @param factory injectable `IDBFactory` (default `globalThis.indexedDB`) so
   *   tests can pass `fake-indexeddb`.
   * @param dbName the database name (default `timeline-history`); pass a
   *   per-workbook name (see {@link databaseNameFor}) to isolate histories.
   */
  constructor(factory: IDBFactory = globalThis.indexedDB, dbName: string = DB_NAME) {
    this.#factory = factory;
    this.#dbName = dbName;
  }

  /**
   * Opens the database (creating object stores on first run) and best-effort
   * requests persistent storage. Idempotent: a second call is a no-op.
   */
  async init(): Promise<void> {
    if (this.#db !== null) {
      return;
    }
    this.#db = await this.#open();
    await this.#requestPersist();
  }

  #open(): Promise<IDBDatabase> {
    const request = this.#factory.open(this.#dbName, DB_VERSION);
    request.onupgradeneeded = (): void => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_DELTAS)) {
        db.createObjectStore(STORE_DELTAS, { keyPath: ['branchId', 'stepIndex'] });
      }
      if (!db.objectStoreNames.contains(STORE_KEYFRAMES)) {
        db.createObjectStore(STORE_KEYFRAMES, { keyPath: ['branchId', 'stepIndex'] });
      }
      if (!db.objectStoreNames.contains(STORE_HEAD)) {
        db.createObjectStore(STORE_HEAD, { keyPath: 'key' });
      }
      if (!db.objectStoreNames.contains(STORE_BRANCHES)) {
        db.createObjectStore(STORE_BRANCHES, { keyPath: 'id' });
      }
    };
    return promisifyRequest(request);
  }

  /** Best-effort `navigator.storage.persist()`; guards for absence (ADR-0007). */
  async #requestPersist(): Promise<void> {
    // Read through a loose view: `navigator`/`storage` may be absent in some
    // hosts and headless environments, so we cannot assume the DOM lib shape.
    const storage = (
      globalThis as { navigator?: { storage?: { persist?: () => Promise<boolean> } } }
    ).navigator?.storage;
    if (storage !== undefined && typeof storage.persist === 'function') {
      try {
        await storage.persist();
      } catch {
        // Persistence is best-effort; a rejection is non-fatal.
      }
    }
  }

  #requireDb(): IDBDatabase {
    if (this.#db === null) {
      throw new Error('IndexedDbStore not initialised: call init() first.');
    }
    return this.#db;
  }

  async appendDelta(branchId: BranchId, delta: Delta): Promise<void> {
    const db = this.#requireDb();
    const tx = db.transaction(STORE_DELTAS, 'readwrite');
    const store = tx.objectStore(STORE_DELTAS);
    // The next stepIndex is one past the branch's current MAX key. We derive it
    // from a `prev` cursor over this branch's [branchId, *] key range and `put`
    // inside the cursor's success callback — keeping a SINGLE active transaction.
    //
    // The previous code did `await count()` then `put()`: after the await the
    // IDB transaction goes inactive, so the put could throw
    // `TransactionInactiveError` on real IndexedDB (fake-indexeddb is lenient
    // and masked it). Driving
    // the put from the cursor callback (no intervening awaited microtask) avoids
    // that, and deriving from the actual max key — not a separate count — keeps
    // sequential appends correct even if the store ever became non-contiguous.
    const range = IDBKeyRange.bound([branchId, -Infinity], [branchId, Infinity]);
    const cursorRequest = store.openCursor(range, 'prev');
    cursorRequest.onsuccess = (): void => {
      const cursor = cursorRequest.result;
      const lastIndex = cursor === null ? -1 : (cursor.value as DeltaRecord).stepIndex;
      const record: DeltaRecord = { branchId, stepIndex: lastIndex + 1, delta };
      store.put(record);
    };
    await awaitTransaction(tx);
  }

  async writeKeyframe(branchId: BranchId, stepIndex: number, state: unknown): Promise<void> {
    const db = this.#requireDb();
    const tx = db.transaction(STORE_KEYFRAMES, 'readwrite');
    const record: KeyframeRecord = { branchId, stepIndex, state };
    // put() overwrites a keyframe at the same (branch, stepIndex).
    tx.objectStore(STORE_KEYFRAMES).put(record);
    await awaitTransaction(tx);
  }

  async loadKeyframeAtOrBefore(
    branchId: BranchId,
    stepIndex: number,
  ): Promise<{ stepIndex: number; state: unknown } | null> {
    const db = this.#requireDb();
    const tx = db.transaction(STORE_KEYFRAMES, 'readonly');
    const store = tx.objectStore(STORE_KEYFRAMES);
    // Walk keys backward from [branchId, stepIndex]; the first hit is the
    // highest keyframe at or before the requested index for this branch.
    const range = IDBKeyRange.bound([branchId, -Infinity], [branchId, stepIndex]);
    const cursor = await promisifyRequest(store.openCursor(range, 'prev'));
    if (cursor === null) {
      return null;
    }
    const record = cursor.value as KeyframeRecord;
    return { stepIndex: record.stepIndex, state: record.state };
  }

  async listKeyframes(branchId: BranchId): Promise<{ stepIndex: number; state: unknown }[]> {
    const db = this.#requireDb();
    const tx = db.transaction(STORE_KEYFRAMES, 'readonly');
    const store = tx.objectStore(STORE_KEYFRAMES);
    const range = IDBKeyRange.bound([branchId, -Infinity], [branchId, Infinity]);
    const records = (await promisifyRequest(store.getAll(range))) as KeyframeRecord[];
    return records
      .map((r) => ({ stepIndex: r.stepIndex, state: r.state }))
      .sort((a, b) => a.stepIndex - b.stepIndex);
  }

  async loadDeltas(branchId: BranchId, from: number, to: number): Promise<Delta[]> {
    const db = this.#requireDb();
    const tx = db.transaction(STORE_DELTAS, 'readonly');
    const store = tx.objectStore(STORE_DELTAS);
    // Inclusive [from, to]; clamp negative `from` to match InMemoryStore.
    const start = Math.max(0, from);
    if (to < start) {
      return [];
    }
    const range = IDBKeyRange.bound([branchId, start], [branchId, to]);
    const records = (await promisifyRequest(store.getAll(range))) as DeltaRecord[];
    return records.map((r) => r.delta);
  }

  async getHead(): Promise<Head | null> {
    const db = this.#requireDb();
    const tx = db.transaction(STORE_HEAD, 'readonly');
    const record = (await promisifyRequest(tx.objectStore(STORE_HEAD).get(HEAD_KEY))) as
      | HeadRecord
      | undefined;
    return record?.head ?? null;
  }

  async setHead(head: Head): Promise<void> {
    const db = this.#requireDb();
    const tx = db.transaction(STORE_HEAD, 'readwrite');
    const record: HeadRecord = { key: HEAD_KEY, head };
    tx.objectStore(STORE_HEAD).put(record);
    await awaitTransaction(tx);
  }

  async saveBranch(meta: BranchMeta): Promise<void> {
    const db = this.#requireDb();
    const tx = db.transaction(STORE_BRANCHES, 'readwrite');
    tx.objectStore(STORE_BRANCHES).put(meta);
    await awaitTransaction(tx);
  }

  async listBranches(): Promise<BranchMeta[]> {
    const db = this.#requireDb();
    const tx = db.transaction(STORE_BRANCHES, 'readonly');
    const branches = (await promisifyRequest(
      tx.objectStore(STORE_BRANCHES).getAll(),
    )) as BranchMeta[];
    return branches.sort((a, b) => a.order - b.order);
  }

  async getBranch(id: BranchId): Promise<BranchMeta | null> {
    const db = this.#requireDb();
    const tx = db.transaction(STORE_BRANCHES, 'readonly');
    const meta = (await promisifyRequest(tx.objectStore(STORE_BRANCHES).get(id))) as
      | BranchMeta
      | undefined;
    return meta ?? null;
  }

  async deleteBranch(id: BranchId): Promise<void> {
    const db = this.#requireDb();
    const tx = db.transaction([STORE_BRANCHES, STORE_DELTAS, STORE_KEYFRAMES], 'readwrite');
    tx.objectStore(STORE_BRANCHES).delete(id);
    // Drop every delta and keyframe belonging to this branch.
    const branchRange = IDBKeyRange.bound([id, -Infinity], [id, Infinity]);
    tx.objectStore(STORE_DELTAS).delete(branchRange);
    tx.objectStore(STORE_KEYFRAMES).delete(branchRange);
    await awaitTransaction(tx);
  }
}
