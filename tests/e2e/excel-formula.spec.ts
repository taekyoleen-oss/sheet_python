import { expect, test, type Page } from "@playwright/test";

// 부록 O.3: 앱 안 복사 → 붙여넣기 시 수식 상대 참조 이동(타일링), Ctrl+D 채우기 — 실제 window 이벤트 경로

/* eslint-disable @typescript-eslint/no-explicit-any */

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as any).__pygridStore !== "undefined" &&
      (window as any).__pygridReady === true,
  );
  await page.waitForSelector('[data-testid="data-grid-canvas"]');
  await page.evaluate(() => (window as any).__pygridStore.getState().newWorkbook());
}

const cellAt = (page: Page, key: string) =>
  page.evaluate(
    (k) => (window as any).__pygridStore.getState().workbook.sheets[0]?.cells[k] ?? null,
    key,
  );

const select = (page: Page, r0: number, c0: number, r1 = r0, c1 = c0) =>
  page.evaluate((rg) => (window as any).__pygridStore.getState().setSelection(rg), { r0, c0, r1, c1 });

test("수식 셀 복사 → 범위 붙여넣기: 상대 참조만 이동하고 값이 계산된다", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);
  // 그리드 포커스 (복사·붙여넣기 핸들러는 텍스트 입력 밖에서만 동작)
  const box = (await page.locator('[data-testid="data-grid-canvas"]').boundingBox())!;
  await page.mouse.click(box.x + 200, box.y + 100);
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setCells(st.workbook.sheets[0].id, [
      { r: 0, c: 0, cell: { v: 2, t: "n" } },
      { r: 1, c: 0, cell: { v: 3, t: "n" } },
      { r: 2, c: 0, cell: { v: 4, t: "n" } },
      { r: 0, c: 1, cell: { v: null, t: "n", fx: "=A1*$A$1" } },
    ]);
  });
  await select(page, 0, 1);
  const text = await page.evaluate(() => {
    const dt = new DataTransfer();
    document.body.dispatchEvent(new ClipboardEvent("copy", { clipboardData: dt, bubbles: true, cancelable: true }));
    return dt.getData("text/plain");
  });
  expect(text.trim()).toBe("4"); // 시스템 클립보드는 값만

  await select(page, 1, 1, 2, 1); // B2:B3
  await page.evaluate((t) => {
    const dt = new DataTransfer();
    dt.setData("text/plain", t);
    document.body.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  }, text);

  await expect.poll(async () => (await cellAt(page, "1:1"))?.fx).toBe("=A2*$A$1");
  expect(await cellAt(page, "1:1")).toMatchObject({ v: 6, t: "n" });
  expect(await cellAt(page, "2:1")).toMatchObject({ fx: "=A3*$A$1", v: 8 });
});

test("Ctrl+D: 첫 행 수식을 아래로 채운다", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);
  const box = (await page.locator('[data-testid="data-grid-canvas"]').boundingBox())!;
  await page.mouse.click(box.x + 200, box.y + 100);
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setCells(st.workbook.sheets[0].id, [
      { r: 0, c: 0, cell: { v: 10, t: "n" } },
      { r: 1, c: 0, cell: { v: 20, t: "n" } },
      { r: 2, c: 0, cell: { v: 30, t: "n" } },
      { r: 0, c: 2, cell: { v: null, t: "n", fx: '=IF(A1>15,"큼","작음")' } },
    ]);
  });
  await select(page, 0, 2, 2, 2); // C1:C3
  await page.keyboard.press("Control+d");

  await expect.poll(async () => (await cellAt(page, "2:2"))?.fx).toBe('=IF(A3>15,"큼","작음")');
  expect(await cellAt(page, "0:2")).toMatchObject({ v: "작음", t: "s" });
  expect(await cellAt(page, "1:2")).toMatchObject({ v: "큼", t: "s" });
});

test("수식 입력줄: 함수 자동완성(Tab) → Enter 확정, 이름 상자로 이름 정의", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setCells(st.workbook.sheets[0].id, [
      { r: 0, c: 0, cell: { v: 4, t: "n" } },
      { r: 1, c: 0, cell: { v: 6, t: "n" } },
    ]);
  });
  // 이름 상자: A1:A2 선택 → "합" 입력 → 이름 정의
  await select(page, 0, 0, 1, 0);
  const nameBox = page.getByLabel("이름 상자");
  await nameBox.click();
  await nameBox.fill("합");
  await nameBox.press("Enter");
  await expect
    .poll(() => page.evaluate(() => (window as any).__pygridStore.getState().workbook.names))
    .toEqual([{ name: "합", ref: "Sheet1!$A$1:$A$2" }]);

  // C1 선택 → 입력줄에 =AVER → Tab으로 AVERAGE( 완성 → 합) 입력 → Enter
  await select(page, 0, 2);
  const bar = page.getByLabel("수식 입력줄");
  await bar.click();
  await bar.pressSequentially("=AVER");
  await expect(page.getByText("AVERAGE(수1, [수2], …) — 평균")).toBeVisible();
  await bar.press("Tab");
  await expect(bar).toHaveValue("=AVERAGE(");
  await bar.pressSequentially("합)");
  await bar.press("Enter");
  await expect.poll(async () => (await cellAt(page, "0:2"))?.v).toBe(5);
  expect((await cellAt(page, "0:2"))?.fx).toBe("=AVERAGE(합)");
});

test("잘라내기 → 붙여넣기: 셀이 이동하고 참조 수식이 따라간다", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);
  const box = (await page.locator('[data-testid="data-grid-canvas"]').boundingBox())!;
  await page.mouse.click(box.x + 200, box.y + 100);
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setCells(st.workbook.sheets[0].id, [
      { r: 0, c: 0, cell: { v: 7, t: "n" } },
      { r: 0, c: 1, cell: { v: null, t: "n", fx: "=A1*3" } },
    ]);
  });
  await select(page, 0, 0);
  const text = await page.evaluate(() => {
    const dt = new DataTransfer();
    document.body.dispatchEvent(new ClipboardEvent("cut", { clipboardData: dt, bubbles: true, cancelable: true }));
    return dt.getData("text/plain");
  });
  await select(page, 4, 3); // D5
  await page.evaluate((t) => {
    const dt = new DataTransfer();
    dt.setData("text/plain", t);
    document.body.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  }, text);
  await expect.poll(async () => (await cellAt(page, "4:3"))?.v).toBe(7);
  expect(await cellAt(page, "0:0")).toBeNull();
  expect(await cellAt(page, "0:1")).toMatchObject({ fx: "=D5*3", v: 21 });
});
