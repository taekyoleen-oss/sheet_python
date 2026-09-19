import { expect, test, type Page } from "@playwright/test";

// 부록 N — 예제 코드를 '불러와서 단계별로 작업하는' 흐름을 실제 데이터로 끝까지 돌린다.
//
// 시나리오: 임금 회귀 워크북을 열고 → 그리드에서 데이터 범위를 선택 → 코드 삽입 팝업에서
// 통계분석 5개 하위 카테고리의 조각을 하나씩 새 블록으로 넣고 → 각각 실행해 'ok'를 확인한다.
// {{range}} 자리표시자가 선택 범위의 xl() 호출로 치환되는지도 함께 본다.
//
// scikit-learn 첫 로드가 있어 상한을 넉넉히 둔다.

/* eslint-disable @typescript-eslint/no-explicit-any */

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as unknown as { __pygridStore?: unknown }).__pygridStore !== "undefined" &&
      (window as unknown as { __pygridReady?: boolean }).__pygridReady === true,
  );
  await page.waitForSelector('[data-testid="data-grid-canvas"]');
}

/** 카테고리별로 하나씩 — 사용자가 위에서 아래로 훑는 순서 그대로 */
const STEPS = [
  { group: "통계분석", snippet: "데이터 개요 — 형태·타입·결측·중복" },
  { group: "전처리 과정", snippet: "범주형 인코딩 — 원-핫(drop_first)" },
  { group: "특성공학", snippet: "설명력 지표 η²(상관비) — 회귀판 IV" },
  {
    group: "데이터 분석 (회귀 모델)",
    snippet: "Lasso — 계수를 0으로 만드는 변수 선택",
  },
  { group: "모델 평가", snippet: "모델 비교표 — 5종 한 표로 정렬" },
];

test("예제 코드 불러오기 — 통계분석 5개 카테고리를 차례로 삽입·실행", async ({ page }) => {
  test.setTimeout(720_000);
  await page.goto("/");
  await waitForApp(page);

  // ── 데이터 준비: 임금 회귀 워크북을 열고 wage 시트 전체를 선택 ──
  await page.getByRole("button", { name: "샘플 워크북", exact: true }).click();
  await page.getByRole("menuitem", { name: "임금 회귀 예측 (5단계)", exact: true }).click();
  await expect
    .poll(() => page.evaluate(() => (window as any).__pygridStore.getState().workbook.title), { timeout: 30_000 })
    .toBe("임금 회귀 예측 — 통계분석 5단계");

  // 기존 16단계는 지우고 빈 상태에서 시작한다 — '내가 직접 조립하는' 흐름을 본다
  await page.evaluate(() => {
    const s = (window as any).__pygridStore.getState();
    for (const b of [...s.workbook.pyBlocks]) s.removePyBlock(b.id);
    s.setSelection({ r0: 0, c0: 0, r1: 534, c1: 10 }); // wage!A1:K535
  });
  await expect.poll(() => page.evaluate(() => (window as any).__pygridStore.getState().workbook.pyBlocks.length)).toBe(0);

  const dialog = page.getByRole("dialog");
  for (const [i, step] of STEPS.entries()) {
    await page.getByRole("button", { name: "코드 삽입" }).click();
    await dialog.getByTestId("snippet-groups").getByRole("button", { name: step.group, exact: true }).click();
    // 스니펫 버튼은 [라벨 + 설명] 두 줄이라 접근명이 합쳐진다 → 라벨로 부분 일치
    const re = new RegExp(step.snippet.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
    await dialog.getByRole("button", { name: re }).click();

    // {{range}} → 그리드 선택 범위의 xl() 호출로 치환 (자동 선택된 유일 후보)
    const preview = dialog.getByTestId("snippet-code-preview");
    const rangeSelect = dialog.getByLabel("자리표시자 {{range}}");
    if (await rangeSelect.isVisible()) {
      await rangeSelect.click();
      await page.getByRole("option", { name: /xl\(/ }).click();
      await expect(preview).toContainText('xl("A1:K535", headers=True)');
    }

    // 기준 블록이 없으면(첫 삽입) 버튼이 "새 블록으로" 하나로 줄어든다
    const below = dialog.getByRole("button", { name: "아래 새 블록으로" });
    await ((await below.count()) ? below : dialog.getByRole("button", { name: "새 블록으로" })).click();
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).__pygridStore.getState().workbook.pyBlocks.length),
      )
      .toBe(i + 1);

    // ── 넣자마자 실행한다 — 실제 작업 순서(넣고 → 돌리고 → 다음을 그 아래에).
    //    다음 블록의 자리는 '빈 셀'을 찾으므로, 먼저 돌려 spill을 만들어 두지 않으면
    //    바로 아래 칸에 놓여 #SPILL! 충돌이 난다.
    const id: string = await page.evaluate(
      () => (window as any).__pygridStore.getState().workbook.pyBlocks.at(-1).id,
    );
    const code: string = await page.evaluate(
      (bid) =>
        (window as any).__pygridStore
          .getState()
          .workbook.pyBlocks.find((b: any) => b.id === bid).code,
      id,
    );
    expect(code, `${step.snippet}: 자리표시자가 남았다`).not.toContain("{{range}}");

    await page.locator(`[data-block-id="${id}"] [aria-label="실행"]`).click();
    await expect
      .poll(
        () =>
          page.evaluate(
            (bid) =>
              (window as any).__pygridStore
                .getState()
                .workbook.pyBlocks.find((b: any) => b.id === bid)?.last?.status ?? null,
            id,
          ),
        { timeout: 300_000, intervals: [2000] },
      )
      .toBe("ok");
  }

  expect(await page.evaluate(() => (window as any).__pygridStore.getState().workbook.pyBlocks.length)).toBe(5);

  // ── 각 단계가 실제로 그 데이터를 계산했는지 '셀에 나온 값'으로 확인 ──
  // (stdout은 설계상 콘솔로 스트리밍되고 블록에는 저장되지 않는다 — 워커 payload.stdout = "".
  //  문자열 출력 검증은 tests/pyodide/example-snippets.test.ts가 맡는다.)
  // 블록 i가 실제로 채운 범위(spillRange)만 읽는다 — 블록들이 세로로 맞닿아 있어
  // '헤더 아래로 값이 이어지는 동안'식으로 읽으면 다음 블록 표까지 딸려 온다.
  const blockTable = (i: number) =>
    page.evaluate((idx) => {
      const s = (window as any).__pygridStore.getState();
      const b = s.workbook.pyBlocks[idx];
      const range = (b.outputs?.[0]?.last ?? b.last)?.spillRange;
      const sh = s.workbook.sheets.find((x: any) => x.id === b.sheetId);
      if (!range || !sh) return [];
      const rows: unknown[][] = [];
      for (let r = range.r0; r <= range.r1; r++) {
        const row: unknown[] = [];
        for (let c = range.c0; c <= range.c1; c++) row.push(sh.cells[`${r}:${c}`]?.v ?? null);
        rows.push(row);
      }
      return rows;
    }, i);

  /** 표에서 header 열의 데이터 값들 (첫 행 = 헤더) */
  const col = (table: unknown[][], header: string): unknown[] => {
    const j = (table[0] ?? []).indexOf(header);
    expect(j, `헤더 '${header}' 없음: ${JSON.stringify(table[0])}`).toBeGreaterThanOrEqual(0);
    return table.slice(1).map((r) => r[j]);
  };

  // ① 데이터 개요 — 원본 11개 열이 그대로 점검표에 오른다
  const t1 = await blockTable(0);
  expect(col(t1, "0비율")).toHaveLength(11);
  expect(col(t1, "열")).toContain("WAGE");

  // ② 원-핫 인코딩 — 범주형이 펼쳐져 더미 열(16개)이 생겼다
  const t2 = await blockTable(1);
  expect(t2[0]).toContain("OCCUPATION_6");
  expect(t2[0]).toHaveLength(16);

  // ③ eta^2 — 범주형 7종이 0~1 값으로 순위가 매겨진다
  const eta = col(await blockTable(2), "eta2").map(Number);
  expect(eta).toHaveLength(7);
  for (const v of eta) {
    expect(v).toBeGreaterThanOrEqual(0);
    expect(v).toBeLessThanOrEqual(1);
  }

  // ④ Lasso — 16개 변수에 선택/제외 표시가 붙고, 일부는 실제로 0이 된다
  const t4 = await blockTable(3);
  expect(col(t4, "선택")).toHaveLength(16);
  expect(col(t4, "선택")).toContain("· (제외)");

  // ⑤ 모델 비교표 — 지정한 5종 회귀 모델이 전부 셀에 있다
  expect(new Set(col(await blockTable(4), "model") as string[])).toEqual(
    new Set(["Ridge", "Lasso", "ElasticNet", "Polynomial(2)", "Log(Ridge)"]),
  );
});
