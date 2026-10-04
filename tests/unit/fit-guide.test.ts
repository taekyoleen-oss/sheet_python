// 부록 H.3: 모델적합 가이드 — 블록 시퀀스 구성(모형별 단계 수·마크다운/코드 교차·제목),
// sheet() 참조·헤더 일치, 분포 선택의 ③ 반영, 한 undo 단계(runFitGuide 선택 모드).
import { beforeEach, describe, expect, it } from "vitest";
import { buildFitGuideBlocks, runFitGuide } from "@/lib/grid/fit-guide";
import { useWorkbookStore } from "@/lib/grid/model";

const st = () => useWorkbookStore.getState();

const SEV = { ref: "청구액!A1:A301", headers: ["청구액"] };
const FREQ = { ref: "건수!A1:B11", headers: ["연도", "건수"] };

describe("buildFitGuideBlocks (순수 코어)", () => {
  it("심도: 제목 + 4단계 [마크다운+코드] = 9블록, 교차 순서·제목", () => {
    const blocks = buildFitGuideBlocks({
      model: "severity",
      dists: ["lognormal", "gamma", "weibull"],
      sev: SEV,
    });
    expect(blocks).toHaveLength(9);
    expect(blocks.map((b) => b.kind)).toEqual([
      "markdown", // 제목
      "markdown", "code", // 1단계
      "markdown", "code", // 2단계
      "markdown", "code", // 3단계
      "markdown", "code", // 4단계
    ]);
    expect(blocks[0].body).toContain("# 모델적합 — 심도");
    for (const n of [1, 2, 3, 4]) {
      expect(blocks.some((b) => b.kind === "markdown" && b.body.includes(`## ${n}단계`))).toBe(true);
    }
    // ③ 비교표는 값 모드 + spill 여유, ②·④ fig는 object 모드
    expect(blocks[6].outputMode).toBe("values");
    expect(blocks[6].reserve).toBeGreaterThan(2);
    expect(blocks[4].outputMode).toBe("object");
    expect(blocks[8].outputMode).toBe("object");
  });

  it("빈도: 9블록 — 과산포 확인·카이제곱 비교", () => {
    const blocks = buildFitGuideBlocks({
      model: "frequency",
      dists: ["poisson", "negbinom"],
      freq: FREQ,
    });
    expect(blocks).toHaveLength(9);
    expect(blocks[0].body).toContain("# 모델적합 — 빈도");
    expect(blocks[4].body).toContain("과산포");
    expect(blocks[6].body).toContain("stats.poisson");
    expect(blocks[6].body).toContain("nbinom");
    expect(blocks[6].body).toContain("chi2");
  });

  it("합성: 13블록 — 5단계(몬테카를로 요약 + S 히스토그램) 포함", () => {
    const blocks = buildFitGuideBlocks({
      model: "compound",
      dists: ["lognormal", "gamma", "poisson", "negbinom"],
      sev: SEV,
      freq: FREQ,
    });
    expect(blocks).toHaveLength(13);
    expect(blocks[0].body).toContain("# 모델적합 — 합성");
    expect(blocks.some((b) => b.kind === "markdown" && b.body.includes("## 5단계"))).toBe(true);
    // ③은 심도·빈도 비교표 코드 2개
    const tables = blocks.filter((b) => b.kind === "code" && b.title.includes("비교표"));
    expect(tables).toHaveLength(2);
    // ⑤: VaR·TVaR 요약(값 모드) + S 분포(fig)
    const mc = blocks.find((b) => b.title.includes("VaR"));
    expect(mc?.outputMode).toBe("values");
    expect(mc?.body).toContain("TVaR");
    expect(blocks[blocks.length - 1].outputMode).toBe("object");
  });

  it("sheet() 참조가 지정 범위·헤더와 일치한다 (값 열 = 마지막 헤더)", () => {
    const blocks = buildFitGuideBlocks({
      model: "severity",
      dists: ["lognormal"],
      sev: { ref: "'내 시트'!B2:C42", headers: ["연도", "손해액"] },
    });
    for (const b of blocks.filter((x) => x.kind === "code")) {
      expect(b.body).toContain(`sheet("'내 시트'!B2:C42", headers=True)`);
      expect(b.body).toContain(`df["손해액"]`);
    }
  });

  it("분포 선택이 ③ 코드에 반영된다 — 선택만 포함, 미선택 제외", () => {
    const blocks = buildFitGuideBlocks({
      model: "severity",
      dists: ["lognormal", "exponential"],
      sev: SEV,
    });
    const table = blocks[6].body;
    expect(table).toContain("stats.lognorm.fit");
    expect(table).toContain("stats.expon.fit");
    expect(table).not.toContain("stats.gamma.fit");
    expect(table).not.toContain("weibull");
  });
});

describe("runFitGuide (선택 범위 모드)", () => {
  beforeEach(() => {
    st().newWorkbook();
    const sid = st().activeSheetId;
    const cells = [{ r: 0, c: 0, cell: { v: "손해액", t: "s" as const } }];
    for (let r = 1; r <= 10; r++) cells.push({ r, c: 0, cell: { v: 100 * r, t: "n" } as never });
    st().setCells(sid, cells);
  });

  it("블록 생성은 한 undo 단계, 워크북 뷰 전환·포커스", async () => {
    st().setView("reference");
    const res = await runFitGuide({
      model: "severity",
      dists: ["lognormal", "gamma"],
      data: { mode: "selection", sevRef: "A1:A11" },
    });
    expect("ids" in res && res.ids).toHaveLength(9);
    expect(st().workbook.pyBlocks).toHaveLength(9);
    expect(st().view).toBe("workbook");
    if ("ids" in res) expect(st().focusBlockId).toBe(res.ids[0]);
    // 시트 접두어가 붙은 실제 범위·헤더 반영
    const sheetName = st().workbook.sheets[0].name;
    const code = st().workbook.pyBlocks.find((b) => b.kind !== "markdown")!.code;
    expect(code).toContain(`sheet("${sheetName}!A1:A11", headers=True)`);
    expect(code).toContain(`df["손해액"]`);
    // 한 undo = 블록 9개가 한 번에 사라진다
    useWorkbookStore.temporal.getState().undo();
    expect(st().workbook.pyBlocks).toHaveLength(0);
  });

  it("헤더가 아닌 첫 행·과다 열·분포 미선택은 오류", async () => {
    const bad1 = await runFitGuide({
      model: "severity",
      dists: ["lognormal"],
      data: { mode: "selection", sevRef: "A2:A11" }, // 첫 행이 숫자
    });
    expect("error" in bad1 && bad1.error).toContain("헤더");

    const bad2 = await runFitGuide({
      model: "severity",
      dists: ["lognormal"],
      data: { mode: "selection", sevRef: "A1:C11" }, // 3열
    });
    expect("error" in bad2 && bad2.error).toContain("1~2열");

    const bad3 = await runFitGuide({
      model: "severity",
      dists: [],
      data: { mode: "selection", sevRef: "A1:A11" },
    });
    expect("error" in bad3 && bad3.error).toContain("분포");
    expect(st().workbook.pyBlocks).toHaveLength(0);
  });
});
