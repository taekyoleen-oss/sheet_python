import { expect, test, type Page } from "@playwright/test";


// 새 블록 기본은 '시트로 보내지 않음' — 이 스펙은 spill을 검증하므로 켜고 시작한다
test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    (window as unknown as { __pygridToSheetDefault?: boolean }).__pygridToSheetDefault = true;
  });
});

// G3 핵심: 블록 실행 → spill → 재실행 시 이전 spill 교체 → Ctrl+Z 복원 → 오류 → 객체 카드
// 한 테스트로 묶어 런타임 부트(첫 CDN 로드 ~15초+)를 1회만 치른다.

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as unknown as { __pygridStore?: unknown }).__pygridStore !==
        "undefined" &&
      (window as unknown as { __pygridReady?: boolean }).__pygridReady === true,
  );
  await page.waitForSelector('[data-testid="data-grid-canvas"]');
  await page.evaluate(() => {
    /* eslint-disable @typescript-eslint/no-explicit-any */
    (window as any).__pygridStore.getState().newWorkbook();
  });
}

const cellAt = (page: Page, key: string) =>
  page.evaluate(
    (k) =>
      (window as any).__pygridStore.getState().workbook.sheets[0]?.cells[k] ?? null,
    key,
  );

const srcCellCount = (page: Page) =>
  page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const blockId = st.workbook.pyBlocks[0]?.id;
    if (!blockId) return 0;
    // src는 "<blockId>:<outputId>" (부록 D.1)
    return Object.values(st.workbook.sheets[0].cells).filter(
      (c: any) => typeof c.src === "string" && c.src.split(":")[0] === blockId,
    ).length;
  });

test("G3: 실행→spill→재실행 교체→undo 복원→오류→객체 카드", async ({ page }) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  // 데이터 A1:B4 (헤더 + 숫자 3행)
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    st.setCells(sid, [
      { r: 0, c: 0, cell: { v: "a", t: "s" } },
      { r: 0, c: 1, cell: { v: "b", t: "s" } },
      { r: 1, c: 0, cell: { v: 1, t: "n" } },
      { r: 1, c: 1, cell: { v: 10, t: "n" } },
      { r: 2, c: 0, cell: { v: 2, t: "n" } },
      { r: 2, c: 1, cell: { v: 20, t: "n" } },
      { r: 3, c: 0, cell: { v: 3, t: "n" } },
      { r: 3, c: 1, cell: { v: 30, t: "n" } },
    ]);
    st.setSelection({ r0: 0, c0: 3, r1: 0, c1: 3 }); // D1
  });

  // Ctrl+Shift+P → 블록 생성 + 카드 포커스
  await page.keyboard.press("Control+Shift+P");
  await expect
    .poll(() =>
      page.evaluate(() => (window as any).__pygridStore.getState().workbook.pyBlocks.length),
    )
    .toBe(1);

  const textarea = page.getByLabel("Python 코드");
  await expect(textarea).toBeFocused();

  // 1) describe() → 큰 spill (부트 포함이라 폴링 여유)
  await textarea.fill('df = xl("A1:B4", headers=True)\ndf.describe()');
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(() => srcCellCount(page), { timeout: 150_000, intervals: [1000] })
    .toBeGreaterThan(8);
  // describe = 헤더 1행 + 통계 8행 × (index + a + b) 3열 = 27셀
  const describeCount = await srcCellCount(page);
  expect(describeCount).toBe(27);
  expect((await cellAt(page, "0:4"))?.v).toBe("a"); // E1: 첫 값 열 헤더
  expect((await cellAt(page, "0:5"))?.v).toBe("b"); // F1: 둘째 값 열 헤더
  const deepCell = await cellAt(page, "8:3"); // D9: 마지막 통계 index 라벨
  expect(deepCell?.v).toBe("max");
  expect(deepCell.src).toBeTruthy();

  // 2) mean() → 작은 spill로 교체, 이전 spill 완전 제거, 무관 셀 유지
  await textarea.fill('df = xl("A1:B4", headers=True)\ndf.mean(numeric_only=True)');
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(() => cellAt(page, "8:3"), { timeout: 60_000, intervals: [500] })
    .toBeNull();
  const meanCount = await srcCellCount(page);
  expect(meanCount).toBeGreaterThan(0);
  expect(meanCount).toBeLessThan(describeCount);
  expect(await cellAt(page, "0:0")).toEqual({ v: "a", t: "s" }); // 데이터 셀 무변화

  // 3) Ctrl+Z → 이전(describe) spill 값 복원 (G3)
  await page.keyboard.press("Control+z");
  await expect
    .poll(() => cellAt(page, "8:3"), { timeout: 10_000 })
    .not.toBeNull();
  expect(await srcCellCount(page)).toBe(describeCount);

  // 4) 오류 코드 → 앵커 #PYTHON!
  await textarea.fill("undefined_name");
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(async () => (await cellAt(page, "0:3"))?.v, { timeout: 60_000, intervals: [500] })
    .toBe("#PYTHON!");
  expect((await cellAt(page, "0:3")).t).toBe("e");
  // 한국어 요약이 last에 저장됨
  const summary = await page.evaluate(
    () => (window as any).__pygridStore.getState().workbook.pyBlocks[0].last?.summaryKo,
  );
  expect(summary).toContain("NameError");

  // 5) 객체 모드 → 앵커 카드 라벨
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setBlockOutputMode(st.workbook.pyBlocks[0].id, "object");
  });
  await textarea.fill('xl("A1:B4", headers=True)');
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(async () => String((await cellAt(page, "0:3"))?.v ?? ""), {
      timeout: 60_000,
      intervals: [500],
    })
    .toMatch(/^\[DataFrame/);
});
