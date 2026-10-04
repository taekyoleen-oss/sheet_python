import { expect, test, type Page } from "@playwright/test";


// 블록 기본은 '시트에 추가 안 함' — 이 스펙은 샘플·블록 결과를 시트 셀에서 검증하므로 켜고 시작한다
test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    (window as unknown as { __pygridToSheetDefault?: boolean }).__pygridToSheetDefault = true;
  });
});

// 부록 H: 모델적합 가이드 마법사(샘플 claims → 심도 → 로그정규+지수) 실런타임 검증 +
// 보험 예제 워크북(청구 심도 적합) 로드 → 전체 실행 성공.
// 부록 K 데이터 내장 워크북 5종(체인래더 확장판 포함)은 sample-workbooks.spec.ts가 다룬다.

/* eslint-disable @typescript-eslint/no-explicit-any */

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as unknown as { __pygridStore?: unknown }).__pygridStore !==
        "undefined" &&
      (window as unknown as { __pygridReady?: boolean }).__pygridReady === true,
  );
  await page.waitForSelector('[data-testid="data-grid-canvas"]');
  await page.evaluate(() => {
    (window as any).__pygridStore.getState().newWorkbook();
  });
}

/** 모든 시트에서 값이 정확히 v인 셀 개수 */
const countCellsWith = (page: Page, v: string) =>
  page.evaluate((val) => {
    const st = (window as any).__pygridStore.getState();
    let n = 0;
    for (const sh of st.workbook.sheets) {
      for (const c of Object.values(sh.cells) as any[]) if (c.v === val) n++;
    }
    return n;
  }, v);

/** 코드 블록 실행 상태 목록 */
const codeStatuses = (page: Page) =>
  page.evaluate(() =>
    (window as any).__pygridStore
      .getState()
      .workbook.pyBlocks.filter((b: any) => b.kind !== "markdown")
      .map((b: any) => b.last?.status ?? null),
  );

/** 블록 앵커 셀 값 (앵커가 놓인 시트에서) */
const anchorCellValue = (page: Page, blockIndex: number) =>
  page.evaluate((i) => {
    const st = (window as any).__pygridStore.getState();
    const b = st.workbook.pyBlocks[i];
    const sheet = st.workbook.sheets.find((s: any) => s.id === b.sheetId);
    return sheet?.cells[`${b.anchor.r}:${b.anchor.c}`]?.v ?? null;
  }, blockIndex);

test("마법사: 샘플 claims → 심도 → 로그정규+지수 → 블록 9개 → 전체 실행 → 비교표 spill + Figure", async ({
  page,
}) => {
  test.setTimeout(420_000);
  await page.goto("/");
  await waitForApp(page);

  // 진입점: 코드 삽입 팝업의 '모델적합 가이드' 항목
  await page.getByRole("button", { name: "코드 삽입" }).click();
  await page.getByRole("button", { name: "모델적합 가이드…" }).click();
  const dialog = page.locator('[data-testid="fit-guide-dialog"]');
  await expect(dialog).toBeVisible();

  // ① 데이터 — 샘플 claims.xlsx (선택이 없으므로 샘플 모드가 기본)
  await expect(dialog.getByLabel(/claims\.xlsx/)).toBeChecked();
  await dialog.getByRole("button", { name: "다음" }).click();
  // ② 모형 — 심도(기본)
  await expect(dialog.getByLabel(/심도\(개별 손해액\)/)).toBeChecked();
  await dialog.getByRole("button", { name: "다음" }).click();
  // ③ 후보 분포 — 로그정규+지수만
  await dialog.getByLabel("감마").uncheck();
  await dialog.getByLabel("와이블").uncheck();
  await dialog.getByLabel("지수").check();
  await dialog.getByRole("button", { name: /완료/ }).click();

  // 블록 9개(제목 md + 4×[md+code]) 생성 — 샘플 시트(claims)도 추가됨
  await expect
    .poll(
      () => page.evaluate(() => (window as any).__pygridStore.getState().workbook.pyBlocks.length),
      { timeout: 30_000 },
    )
    .toBe(9);
  const summary = await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    return {
      sheets: st.workbook.sheets.map((s: any) => s.name),
      titles: st.workbook.pyBlocks
        .filter((b: any) => b.kind === "markdown")
        .map((b: any) => b.title),
      fitCode: st.workbook.pyBlocks[6].code,
      view: st.view,
    };
  });
  expect(summary.view).toBe("workbook");
  expect(summary.sheets).toContain("claims");
  // 분포 선택 반영 — 로그정규·지수만
  expect(summary.fitCode).toContain("stats.lognorm.fit");
  expect(summary.fitCode).toContain("stats.expon.fit");
  expect(summary.fitCode).not.toContain("stats.gamma.fit");
  // 목차에 1~4단계
  await page.evaluate(() => (window as any).__pygridStore.getState().setTocOpen(true));
  const toc = page.locator('[data-testid="toc-panel"]');
  for (const n of [1, 2, 3, 4]) {
    await expect(toc.getByText(new RegExp(`${n}단계`)).first()).toBeVisible();
  }

  // 전체 실행 → ③ 비교표 spill(AIC 열)과 ④ Figure 카드 (실런타임, scipy 다운로드 포함)
  await page.getByRole("button", { name: "전체 실행", exact: true }).click();
  await expect
    .poll(() => countCellsWith(page, "AIC"), { timeout: 360_000, intervals: [3000] })
    .toBeGreaterThan(0);
  await expect
    .poll(async () => String((await anchorCellValue(page, 8)) ?? ""), {
      timeout: 120_000,
      intervals: [2000],
    })
    .toMatch(/^\[Figure/);
  expect(await codeStatuses(page)).toEqual(["ok", "ok", "ok", "ok"]);
});

test("샘플 예제: 청구 심도 적합 — 로드 직후 전체 실행 성공", async ({ page }) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  await page.getByRole("button", { name: "샘플 워크북", exact: true }).click();
  await page.getByRole("menuitem", { name: "청구 심도 적합", exact: true }).click();
  await expect
    .poll(() => page.evaluate(() => (window as any).__pygridStore.getState().workbook.title))
    .toBe("청구 심도 적합 예제");

  await page.getByRole("button", { name: "전체 실행", exact: true }).click();
  await expect
    .poll(() => countCellsWith(page, "AIC"), { timeout: 240_000, intervals: [3000] })
    .toBeGreaterThan(0);
  await expect
    .poll(async () => String((await anchorCellValue(page, 8)) ?? ""), {
      timeout: 120_000,
      intervals: [2000],
    })
    .toMatch(/^\[Figure/);
  expect(await codeStatuses(page)).toEqual(["ok", "ok", "ok", "ok"]);
});
