// R5: 데이터 불러오기 — cp949 감지·디코드, 시트명 중복 회피, xl 로드 블록 참조 정확성, makeBlock 3모드
import { beforeEach, describe, expect, it } from "vitest";
import {
  buildPandasLoadCode,
  decodeSmart,
  detectTextEncoding,
  importData,
  uniqueSheetName,
} from "@/lib/io/data-import";
import { useWorkbookStore } from "@/lib/grid/model";
import { cellKey } from "@/types/workbook";

const st = () => useWorkbookStore.getState();
const utf8 = (s: string) => new TextEncoder().encode(s);
// "한글" CP949: 한=C7 D1, 글=B1 DB (utf-8 fatal 디코드 실패 바이트열)
const CP949_HANGUL = new Uint8Array([0xc7, 0xd1, 0xb1, 0xdb]);

beforeEach(() => {
  st().newWorkbook();
});

describe("detectTextEncoding / decodeSmart", () => {
  it("UTF-8 BOM → utf-8-sig", () => {
    expect(detectTextEncoding(new Uint8Array([0xef, 0xbb, 0xbf, 0x61]))).toBe("utf-8-sig");
  });

  it("유효한 UTF-8 → null(기본), 한글 포함", () => {
    expect(detectTextEncoding(utf8("이름,값\n가,1"))).toBeNull();
  });

  it("CP949 바이트 → cp949 감지 + euc-kr 디코드", () => {
    expect(detectTextEncoding(CP949_HANGUL)).toBe("cp949");
    expect(decodeSmart(CP949_HANGUL)).toBe("한글");
  });
});

describe("uniqueSheetName", () => {
  it('중복 시 " (2)", " (3)" 부여', () => {
    const taken = new Set(["policy", "policy (2)"]);
    expect(uniqueSheetName("claims", taken)).toBe("claims");
    expect(uniqueSheetName("policy", taken)).toBe("policy (3)");
  });
});

describe("importData", () => {
  it("CSV → 새 시트 추가(기존 시트 유지) + 시트명 중복 회피", async () => {
    st().renameSheet(st().workbook.sheets[0].id, "mini");
    const res = await importData("mini.csv", utf8("a,b\n1,2\n3,4"), {
      toSheet: true,
      toFs: false,
      makeBlock: "none",
    });
    expect(res.sheetNames).toEqual(["mini (2)"]);
    const wb = st().workbook;
    expect(wb.sheets.map((s) => s.name)).toEqual(["mini", "mini (2)"]);
    expect(wb.sheets[1].cells[cellKey(0, 0)]?.v).toBe("a");
    expect(wb.sheets[1].cells[cellKey(2, 1)]?.v).toBe(4);
    expect(st().activeSheetId).toBe(wb.sheets[1].id); // 추가된 시트 활성화
  });

  it("cp949 CSV → euc-kr 디코드된 셀 + pandas 블록에 encoding 힌트", async () => {
    // "이름,값\n한글,1" 을 cp949로: 이름=C0 CC B8 A7, 한글=C7 D1 B1 DB
    const bytes = new Uint8Array([
      0xc0, 0xcc, 0xb8, 0xa7, 0x2c, 0xb0, 0xaa, 0x0a, // 이름,값\n
      0xc7, 0xd1, 0xb1, 0xdb, 0x2c, 0x31, // 한글,1
    ]);
    const res = await importData("k.csv", bytes, {
      toSheet: true,
      toFs: false,
      makeBlock: "pandas",
    });
    expect(res.encoding).toBe("cp949");
    const sheet = st().workbook.sheets.find((s) => s.name === "k")!;
    expect(sheet.cells[cellKey(0, 0)]?.v).toBe("이름");
    expect(sheet.cells[cellKey(1, 0)]?.v).toBe("한글");
    const block = st().workbook.pyBlocks.find((b) => b.id === res.blockId)!;
    expect(block.code).toContain('pd.read_csv("k.csv", encoding="cp949")');
    expect(block.title).toBe("k.csv");
  });

  it("xl 로드 블록의 참조 범위 = 생성 시트의 사용 범위", async () => {
    const res = await importData("mini.csv", utf8("a,b\n1,2\n3,4"), {
      toSheet: true,
      toFs: false,
      makeBlock: "xl",
    });
    const block = st().workbook.pyBlocks.find((b) => b.id === res.blockId)!;
    // 시트 "mini", 사용 범위 A1:B3 (헤더 1행 + 데이터 2행 × 2열)
    expect(block.code).toContain('df = sheet("mini!A1:B3", headers=True)');
    // 블록 앵커는 추가된 데이터 시트(활성)의 빈 열
    expect(block.sheetId).toBe(st().activeSheetId);
  });

  it("시트명에 공백이 있으면 xl 참조는 따옴표 시트명", async () => {
    st().renameSheet(st().workbook.sheets[0].id, "mini");
    const res = await importData("mini.csv", utf8("a\n1"), {
      toSheet: true,
      toFs: false,
      makeBlock: "xl",
    });
    const block = st().workbook.pyBlocks.find((b) => b.id === res.blockId)!;
    expect(block.code).toContain(`df = sheet("'mini (2)'!A1:A2", headers=True)`);
  });

  it("makeBlock 'none' → 블록 없음, 'xl'인데 시트 미생성(toFs도 없음) → 블록 강등 없음", async () => {
    const r1 = await importData("a.csv", utf8("x\n1"), {
      toSheet: true,
      toFs: false,
      makeBlock: "none",
    });
    expect(r1.blockId).toBeNull();
    expect(st().workbook.pyBlocks).toHaveLength(0);

    const r2 = await importData("b.csv", utf8("x\n1"), {
      toSheet: false,
      toFs: false,
      makeBlock: "xl",
    });
    expect(r2.sheetNames).toEqual([]);
    expect(r2.blockId).toBeNull();
    expect(st().workbook.pyBlocks).toHaveLength(0);
  });

  it("pandas 로드 코드 — xlsx는 openpyxl import 포함, csv는 미포함", () => {
    expect(buildPandasLoadCode("claims.xlsx", null)).toContain("import openpyxl");
    expect(buildPandasLoadCode("claims.xlsx", null)).toContain('pd.read_excel("claims.xlsx")');
    expect(buildPandasLoadCode("k.csv", null)).not.toContain("openpyxl");
  });
});
