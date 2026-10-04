import { expect, test, type Page } from "@playwright/test";

// 부록 I 미니 수식: 입력→계산→자동 갱신, SUM, 편집 재진입 원문, #DIV/0! hover,
// 수식→Python 자동 재실행(실런타임), .pygrid.json 왕복 fx 보존

/* eslint-disable @typescript-eslint/no-explicit-any */

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as any).__pygridStore !== "undefined" &&
      (window as any).__pygridReady === true,
  );
  await page.waitForSelector('[data-testid="data-grid-canvas"]');
  await page.evaluate(() => (window as any).__pygridStore.getState().newWorkbook());
  // 부록 P.6: 셀 좌표 보정(88×34)은 기본 크기 보기 기준
  await page.evaluate(() => (window as any).__pygridStore.getState().setGridCompact(false));
}

const cellAt = (page: Page, key: string) =>
  page.evaluate(
    (k) => (window as any).__pygridStore.getState().workbook.sheets[0]?.cells[k] ?? null,
    key,
  );

/** 셀에 텍스트 입력: 첫 키로 오버레이를 열고 포커스를 기다린 뒤 나머지 입력 → 값 확인 → Enter */
async function typeAndCommit(page: Page, text: string) {
  await page.keyboard.type(text[0]);
  await page.waitForFunction(() => {
    const ta = document.getElementById("portal")?.querySelector("textarea, input");
    return !!ta && document.activeElement === ta;
  });
  await page.keyboard.type(text.slice(1));
  await expect
    .poll(
      () =>
        page.evaluate(
          () =>
            (document.getElementById("portal")?.querySelector("textarea, input") as
              | HTMLTextAreaElement
              | null)?.value ?? null,
        ),
      { timeout: 5_000 },
    )
    .toBe(text);
  await page.keyboard.press("Enter");
}

/** 캔버스 좌표 보정 (grid-nav와 동일: 기본 열 88px·행 34px, 원점 역산 + 그리드 포커스) */
async function calibrate(page: Page) {
  const box = (await page.locator('[data-testid="data-grid-canvas"]').boundingBox())!;
  await page.mouse.click(box.x + 200, box.y + 100);
  const sel = await page.evaluate(() => (window as any).__pygridStore.getState().selection);
  return { x0: box.x + 200 - sel.c0 * 88, y0: box.y + 100 - sel.r0 * 34 };
}

test("수식 입력 → 계산 값 표시 → 참조 셀 수정 시 자동 갱신 → 편집 재진입 시 원문", async ({
  page,
}) => {
  await page.goto("/");
  await waitForApp(page);
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    st.setCells(sid, [
      { r: 0, c: 0, cell: { v: 3, t: "n" } }, // A1
      { r: 1, c: 1, cell: { v: 4, t: "n" } }, // B2
    ]);
  });
  const { x0, y0 } = await calibrate(page);

  // D1에 =A1+B2*2 입력
  await page.mouse.click(x0 + 3 * 88, y0);
  await typeAndCommit(page, "=A1+B2*2");
  await expect.poll(async () => (await cellAt(page, "0:3"))?.v, { timeout: 5_000 }).toBe(11);
  expect(await cellAt(page, "0:3")).toMatchObject({ t: "n", fx: "=A1+B2*2" });

  // A1 수정 → 자동 갱신
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setCellValue(st.workbook.sheets[0].id, 0, 0, { v: 10, t: "n" });
  });
  await expect.poll(async () => (await cellAt(page, "0:3"))?.v, { timeout: 5_000 }).toBe(18);

  // 편집 재진입 → 오버레이에 수식 원문
  await page.mouse.dblclick(x0 + 3 * 88, y0);
  await expect
    .poll(
      () =>
        page.evaluate(
          () =>
            (document.getElementById("portal")?.querySelector("textarea, input") as
              | HTMLTextAreaElement
              | null)?.value ?? null,
        ),
      { timeout: 5_000 },
    )
    .toBe("=A1+B2*2");
  await page.keyboard.press("Escape");
});

test("=SUM(A1:A5) 집계 + #DIV/0! 표시·hover 한국어 설명", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    st.setCells(
      sid,
      [1, 2, 3, 4, 5].map((v, r) => ({ r, c: 0, cell: { v, t: "n" } })),
    );
  });
  const { x0, y0 } = await calibrate(page);

  await page.mouse.click(x0 + 2 * 88, y0); // C1
  await typeAndCommit(page, "=SUM(A1:A5)");
  await expect.poll(async () => (await cellAt(page, "0:2"))?.v, { timeout: 5_000 }).toBe(15);

  // C2: 0으로 나누기 → 오류 셀 + hover 한국어 툴팁
  await page.mouse.click(x0 + 2 * 88, y0 + 1 * 34);
  await typeAndCommit(page, "=1/0");
  await expect
    .poll(async () => (await cellAt(page, "1:2"))?.v, { timeout: 5_000 })
    .toBe("#DIV/0!");
  expect((await cellAt(page, "1:2"))?.t).toBe("e");

  await page.mouse.move(x0, y0); // 먼저 다른 셀로
  await page.mouse.move(x0 + 2 * 88, y0 + 1 * 34);
  await expect(page.getByText("0으로 나눌 수 없습니다")).toBeVisible({ timeout: 5_000 });
});

test("수식 값을 xl()로 읽는 블록이 수식 갱신 시 자동 재실행된다 (실런타임)", async ({
  page,
}) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  // A1=2 · B1==A1*3(=6) · 블록 D1: xl("B1")+1
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    st.setCellValue(sid, 0, 0, { v: 2, t: "n" });
    st.setCellValue(sid, 0, 1, { v: null, t: "n", fx: "=A1*3" });
    const id = st.addPyBlock(sid, { r: 0, c: 3 });
    st.setBlockCode(id, 'xl("B1") + 1');
  });
  expect((await cellAt(page, "0:1"))?.v).toBe(6); // 수식 즉시 계산

  await page.getByRole("button", { name: "전체 실행", exact: true }).click();
  await expect
    .poll(async () => (await cellAt(page, "0:3"))?.v, { timeout: 150_000, intervals: [1000] })
    .toBe(7);

  // A1 수정 → 수식 B1이 30으로 재계산 → 통지 경로로 블록 자동 재실행 (수동 notifyEdit 없이)
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setCellValue(st.workbook.sheets[0].id, 0, 0, { v: 10, t: "n" });
  });
  await expect.poll(async () => (await cellAt(page, "0:1"))?.v, { timeout: 5_000 }).toBe(30);
  await expect
    .poll(async () => (await cellAt(page, "0:3"))?.v, { timeout: 60_000, intervals: [500] })
    .toBe(31);
});

test(".pygrid.json 저장 → 열기에 fx가 보존되고 다시 계산 가능하다", async ({
  page,
}, testInfo) => {
  await page.goto("/");
  await waitForApp(page);
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    st.setTitle("수식왕복");
    st.setCellValue(sid, 0, 0, { v: 1, t: "n" });
    st.setCellValue(sid, 0, 1, { v: null, t: "n", fx: "=A1+1" });
  });
  expect((await cellAt(page, "0:1"))?.v).toBe(2);

  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "파일" }).click();
  await page.getByRole("menuitem", { name: "저장 (.pygrid.json)" }).click();
  const savedPath = testInfo.outputPath("formula.pygrid.json");
  await (await downloadPromise).saveAs(savedPath);

  await page.evaluate(() => (window as any).__pygridStore.getState().newWorkbook());
  await page.locator('input[aria-label="워크북 파일 열기"]').setInputFiles(savedPath);
  await expect
    .poll(async () => (await cellAt(page, "0:1"))?.fx ?? null, { timeout: 10_000 })
    .toBe("=A1+1");
  expect((await cellAt(page, "0:1"))?.v).toBe(2);

  // 열린 워크북에서도 재계산이 살아 있다
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setCellValue(st.workbook.sheets[0].id, 0, 0, { v: 5, t: "n" });
  });
  await expect.poll(async () => (await cellAt(page, "0:1"))?.v, { timeout: 5_000 }).toBe(6);
});
