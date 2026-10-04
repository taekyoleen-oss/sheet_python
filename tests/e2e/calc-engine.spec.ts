import { expect, test, type Page } from "@playwright/test";


// 새 블록 기본은 '시트로 보내지 않음' — 이 스펙은 spill을 검증하므로 켜고 시작한다
test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    (window as unknown as { __pygridToSheetDefault?: boolean }).__pygridToSheetDefault = true;
  });
});

// G4: 의존성 재계산(입력 셀 수정 → A·B 순 재실행) + 순환 참조 오류
// G5: 무한 루프 중단(KeyboardInterrupt) 후 런타임 유지(변수 보존, 재부트 없음)

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

test("G4: 의존 블록 자동 재계산 + 순환 참조", async ({ page }) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  // 입력 A1=1, A2=2 · 블록 A(C1): A1+A2 · 블록 B(E1): C1×10
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    st.setCells(sid, [
      { r: 0, c: 0, cell: { v: 1, t: "n" } },
      { r: 1, c: 0, cell: { v: 2, t: "n" } },
    ]);
    const a = st.addPyBlock(sid, { r: 0, c: 2 });
    const b = st.addPyBlock(sid, { r: 0, c: 4 });
    st.setBlockCode(a, 'sheet("A1") + sheet("A2")');
    st.setBlockCode(b, 'sheet("C1") * 10');
  });

  await page.getByRole("button", { name: "전체 실행", exact: true }).click();
  await expect
    .poll(async () => (await cellAt(page, "0:2"))?.v, { timeout: 150_000, intervals: [1000] })
    .toBe(3);
  await expect
    .poll(async () => (await cellAt(page, "0:4"))?.v, { timeout: 30_000, intervals: [500] })
    .toBe(30);

  // 입력 셀 수정(자동 모드) → A·B 순 재실행. B=70이 되려면 A의 새 spill(7)을 읽어야 한다
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    st.setCellValue(sid, 0, 0, { v: 5, t: "n" });
    (window as any).__pygridCalc.notifyEdit(
      [{ sheetId: sid, r0: 0, c0: 0, r1: 0, c1: 0 }],
      [],
    );
  });
  await expect
    .poll(async () => (await cellAt(page, "0:2"))?.v, { timeout: 60_000, intervals: [500] })
    .toBe(7);
  await expect
    .poll(async () => (await cellAt(page, "0:4"))?.v, { timeout: 30_000, intervals: [500] })
    .toBe(70);

  // 순환: A가 B의 spill(E1)을 참조 → A↔B 순환 → 둘 다 #PYTHON! 순환 참조
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setBlockCode(st.workbook.pyBlocks[0].id, 'sheet("E1") + 1');
  });
  await page.getByRole("button", { name: "전체 실행", exact: true }).click();
  for (const key of ["0:2", "0:4"]) {
    await expect
      .poll(async () => (await cellAt(page, key))?.v, { timeout: 30_000, intervals: [500] })
      .toBe("#PYTHON!");
  }
  const summaries = await page.evaluate(() =>
    (window as any).__pygridStore
      .getState()
      .workbook.pyBlocks.map((b: any) => b.last?.summaryKo),
  );
  expect(summaries).toEqual(["순환 참조", "순환 참조"]);
});

test("G5: 무한 루프 중단 후 런타임 유지", async ({ page }) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    const id = st.addPyBlock(sid, { r: 4, c: 0 }); // A5
    st.setBlockCode(id, "x = 42\nx");
  });
  await page.getByRole("button", { name: "전체 실행", exact: true }).click();
  await expect
    .poll(async () => (await cellAt(page, "4:0"))?.v, { timeout: 150_000, intervals: [1000] })
    .toBe(42);

  // 무한 루프 실행 → 실행 중 확인 → ■ 중단
  const textarea = page.getByLabel("Python 코드");
  await textarea.fill("while True:\n    pass");
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(
      () =>
        page.evaluate(
          () => Object.keys((window as any).__pygridStore.getState().runningBlocks).length,
        ),
      { timeout: 30_000, intervals: [200] },
    )
    .toBeGreaterThan(0);
  await page.waitForTimeout(1500); // 루프가 실제로 도는 중임을 보장
  await page.getByRole("button", { name: "실행 중단", exact: true }).click();

  // 3초 내 KeyboardInterrupt → #PYTHON!, 재부트 아님
  await expect
    .poll(async () => (await cellAt(page, "4:0"))?.v, { timeout: 10_000, intervals: [300] })
    .toBe("#PYTHON!");
  const summary = await page.evaluate(
    () => (window as any).__pygridStore.getState().workbook.pyBlocks[0].last?.summaryKo,
  );
  expect(summary).toContain("중단");
  await expect(page.getByText(/재설정되어/)).toHaveCount(0); // 재부트 토스트 없음

  // 네임스페이스 유지: x → 42, 1+1 → 2
  await textarea.fill("x");
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(async () => (await cellAt(page, "4:0"))?.v, { timeout: 30_000, intervals: [500] })
    .toBe(42);
  await textarea.fill("1 + 1");
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(async () => (await cellAt(page, "4:0"))?.v, { timeout: 30_000, intervals: [500] })
    .toBe(2);
});
