// IndexedDB 래퍼 — 실패 시 메모리 Map으로 강등, 워크북 LRU 최대 20개

import { openDB, type IDBPDatabase } from "idb";
import type { ChatMessage } from "@/lib/ai/chat";
import type { Workbook } from "@/types/workbook";

export interface AppSettings {
  theme?: string;
  splitRatio?: number;
  bottomPanelHeight?: number;
  fontSize?: number;
  lastWorkbookId?: string;
  pyodideIndexUrl?: string;
  /** 붙여넣기 시 항상 미리보기 다이얼로그 표시 (기본 false: 5행 이하는 즉시 반영) */
  pasteAlwaysDialog?: boolean;
  /** 붙여넣기 날짜 추론 순서 — 9/2/2026 같은 모호한 표기 해석 (기본 'ymd') */
  dateOrder?: "ymd" | "mdy";
  /** 목차 전용 패널 열림 (기본 false) */
  tocOpen?: boolean;
  /** 실행 참조 표시 (부록 J.3, 기본 true) */
  showRefs?: boolean;
  /** 상단 뷰 전환 — 워크북 | 데이터 예제/분석 (부록 E, 기본 'workbook') */
  view?: "workbook" | "reference";
  /** 참조 뷰의 마지막 활성 탭 (excel | methods | dist | fit) */
  referenceTab?: string;
  /** 데이터 불러오기 옵션 기억 (부록 E R5, 기본 시트·FS 모두 ✓ + xl 블록) */
  dataImport?: { toSheet: boolean; toFs: boolean; makeBlock: "xl" | "pandas" | "none" };
  /** Anthropic API 키 (부록 E R6) — 이 브라우저 IndexedDB 전용.
   *  워크북 JSON·내보내기·git 어디에도 실리지 않는다(워크북 객체와 분리 저장) */
  anthropicApiKey?: string;
  /** 부록 P: 속성 창 열림·고정·너비(px) */
  propsOpen?: boolean;
  propsPinned?: boolean;
  propsWidth?: number;
  /** 부록 P.6: 스프레드시트 작게 보기 (기본 true) */
  gridCompact?: boolean;
  /** 부록 P.6: 그리드:Python 분할(그리드 %, 기본 40 — Python 작업 중심). 구 splitRatio(72 기본)는 쓰지 않는다 */
  gridSplit?: number;
  /** AI 채팅 패널 열림 (부록 G.2, 기본 false) */
  aiChatOpen?: boolean;
  /** AI 채팅 이력 — 로컬 전용, 워크북 파일 미포함. 최대 200개(저장 시 캡) */
  aiChatHistory?: ChatMessage[];
  /** AI 채팅 사용자 지침 (부록 G.3) — 미설정이면 기본 시드 사용 */
  aiChatInstructions?: string;
}

const DB_NAME = "pygrid";
const MAX_WORKBOOKS = 20;

let degraded = false;
/** IndexedDB 접근 실패로 메모리 저장으로 강등되었는지 (새로고침 시 데이터 소실) */
export const isStorageDegraded = (): boolean => degraded;

const mem = {
  workbooks: new Map<string, Workbook>(),
  blobs: new Map<string, Blob>(),
  settings: new Map<string, AppSettings>(),
};

let dbPromise: Promise<IDBPDatabase> | null = null;

async function getDb(): Promise<IDBPDatabase | null> {
  if (degraded) return null;
  try {
    dbPromise ??= openDB(DB_NAME, 1, {
      upgrade(db) {
        db.createObjectStore("workbooks");
        db.createObjectStore("blobs");
        db.createObjectStore("settings");
      },
    });
    return await dbPromise;
  } catch {
    degraded = true;
    return null;
  }
}

async function evictLru(db: IDBPDatabase): Promise<void> {
  // ponytail: getAll 전체 로드 — 워크북 20+개 × 수 MB가 문제 되면 updatedAt 인덱스로 교체
  const all: Workbook[] = await db.getAll("workbooks");
  if (all.length <= MAX_WORKBOOKS) return;
  all.sort((a, b) => (a.updatedAt < b.updatedAt ? -1 : 1));
  for (const wb of all.slice(0, all.length - MAX_WORKBOOKS)) {
    await db.delete("workbooks", wb.id);
    for (const block of wb.pyBlocks) {
      if (block.last?.imageBlobId) await db.delete("blobs", block.last.imageBlobId);
    }
  }
}

export async function getWorkbook(id: string): Promise<Workbook | undefined> {
  const db = await getDb();
  if (db) {
    try {
      return await db.get("workbooks", id);
    } catch {
      degraded = true;
    }
  }
  return mem.workbooks.get(id);
}

export async function putWorkbook(wb: Workbook): Promise<void> {
  const db = await getDb();
  if (db) {
    try {
      await db.put("workbooks", wb, wb.id);
      await evictLru(db);
      return;
    } catch {
      degraded = true;
    }
  }
  mem.workbooks.set(wb.id, wb);
}

export async function deleteWorkbook(id: string): Promise<void> {
  const db = await getDb();
  if (db) {
    try {
      const wb: Workbook | undefined = await db.get("workbooks", id);
      await db.delete("workbooks", id);
      for (const block of wb?.pyBlocks ?? []) {
        if (block.last?.imageBlobId) await db.delete("blobs", block.last.imageBlobId);
      }
      return;
    } catch {
      degraded = true;
    }
  }
  const memWb = mem.workbooks.get(id);
  for (const block of memWb?.pyBlocks ?? []) {
    if (block.last?.imageBlobId) mem.blobs.delete(block.last.imageBlobId);
  }
  mem.workbooks.delete(id);
}

export async function listWorkbooks(): Promise<Workbook[]> {
  const db = await getDb();
  if (db) {
    try {
      const all: Workbook[] = await db.getAll("workbooks");
      return all.sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : -1));
    } catch {
      degraded = true;
    }
  }
  return [...mem.workbooks.values()].sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : -1));
}

export async function getBlob(id: string): Promise<Blob | undefined> {
  const db = await getDb();
  if (db) {
    try {
      return await db.get("blobs", id);
    } catch {
      degraded = true;
    }
  }
  return mem.blobs.get(id);
}

export async function putBlob(id: string, blob: Blob): Promise<void> {
  const db = await getDb();
  if (db) {
    try {
      await db.put("blobs", blob, id);
      return;
    } catch {
      degraded = true;
    }
  }
  mem.blobs.set(id, blob);
}

export async function loadSettings(): Promise<AppSettings | undefined> {
  const db = await getDb();
  if (db) {
    try {
      return await db.get("settings", "app");
    } catch {
      degraded = true;
    }
  }
  return mem.settings.get("app");
}

export async function saveSettings(patch: Partial<AppSettings>): Promise<void> {
  const current = (await loadSettings()) ?? {};
  const next = { ...current, ...patch };
  const db = await getDb();
  if (db) {
    try {
      await db.put("settings", next, "app");
      return;
    } catch {
      degraded = true;
    }
  }
  mem.settings.set("app", next);
}
