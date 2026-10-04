import { expect, test, type Page } from "@playwright/test";

// 부록 P.6: 기본 분할이 시트 40%라, 셀 좌표로 누르는 이 스펙은 넓은 창에서 돌린다
test.use({ viewport: { width: 1600, height: 900 } });

// 엑셀식 그리드 탐색: Ctrl+방향키 / Ctrl+Shift+방향키 / 선택 가장자리 Shift+더블클릭

/* eslint-disable @typescript-eslint/no-explicit-any */

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as any).__pygridStore !== "undefined" &&
      (window as any).__pygridReady === true,
  );
  await page.waitForSelector('[data-testid="data-grid-canvas"]');
}

/** 새 워크북 + A1:C5(r0..4 × c0..2) 데이터 채우기 */
async function setup(page: Page) {
  await page.goto("/");
  await waitForApp(page);
  await page.evaluate(() => {
    const store = (window as any).__pygridStore;
    store.getState().newWorkbook();
    store.getState().setGridCompact(false); // 부록 P.6: 셀 좌표 보정(88×34)은 기본 크기 보기 기준
    const st = store.getState();
    const sid = st.workbook.sheets[0].id;
    for (let r = 0; r < 5; r++)
      for (let c = 0; c < 3; c++) st.setCellValue(sid, r, c, { v: r * 10 + c, t: "n" });
  });
  await page.waitForTimeout(300); // 크기 전환 뒤 격자 다시 그리기를 기다린 다음 좌표를 잰다
}

const selection = (page: Page) =>
  page.evaluate(() => (window as any).__pygridStore.getState().selection);

/** 캔버스 좌표 보정: 임의 지점 클릭 → 선택 셀로 원점 역산 (기본 열 88px·행 34px). 포커스도 그리드로 이동.
 *  이후 (x0 + c*88, y0 + r*34)는 셀 (r,c) 내부의 같은 상대 지점에 떨어진다 */
async function calibrate(page: Page) {
  const box = (await page.locator('[data-testid="data-grid-canvas"]').boundingBox())!;
  await page.mouse.click(box.x + 200, box.y + 100);
  const sel = await selection(page);
  return { x0: box.x + 200 - sel.c0 * 88, y0: box.y + 100 - sel.r0 * 34 };
}

test("Ctrl+방향키는 데이터 끝으로 이동, Ctrl+Shift+방향키는 앵커를 유지하며 확장한다", async ({
  page,
}) => {
  await setup(page);
  const { x0, y0 } = await calibrate(page);

  // A1 클릭 → Ctrl+Shift+ArrowRight → 앵커 A1 유지, C1까지 확장 (A1:C1)
  await page.mouse.click(x0, y0);
  expect(await selection(page)).toEqual({ r0: 0, c0: 0, r1: 0, c1: 0 });
  await page.keyboard.press("Control+Shift+ArrowRight");
  expect(await selection(page)).toEqual({ r0: 0, c0: 0, r1: 0, c1: 2 });

  // 이어서 Ctrl+Shift+ArrowDown → A1:C5 (이동 코너만 데이터 끝으로)
  await page.keyboard.press("Control+Shift+ArrowDown");
  expect(await selection(page)).toEqual({ r0: 0, c0: 0, r1: 4, c1: 2 });

  // Ctrl+ArrowDown → 앵커 A1 기준 연속 데이터 끝 A5 단일 선택
  await page.keyboard.press("Control+ArrowDown");
  expect(await selection(page)).toEqual({ r0: 4, c0: 0, r1: 4, c1: 0 });

  // 다시 Ctrl+ArrowDown → 아래에 데이터 없음 → 그리드 마지막 행
  await page.keyboard.press("Control+ArrowDown");
  const last = await selection(page);
  expect(last.c0).toBe(0);
  expect(last.r0).toBe(last.r1);
  expect(last.r0).toBeGreaterThan(4);
});

test("2셀 이상 선택의 아래 가장자리 Shift+더블클릭 → 아래 데이터 끝까지 확장", async ({
  page,
}) => {
  await setup(page);
  const { x0, y0 } = await calibrate(page);

  // A1 클릭 + Shift+클릭 B2 → 2×2 (A1:B2), 앵커 A1
  await page.mouse.click(x0, y0);
  await page.keyboard.down("Shift");
  await page.mouse.click(x0 + 1 * 88, y0 + 1 * 34);
  expect(await selection(page)).toEqual({ r0: 0, c0: 0, r1: 1, c1: 1 });

  // 아래 가장자리 셀 A2를 Shift+더블클릭: 첫 클릭이 A1:A2로 좁히고, 더블클릭이 A1:A5로 확장
  await page.mouse.dblclick(x0, y0 + 1 * 34);
  await page.keyboard.up("Shift");
  expect(await selection(page)).toEqual({ r0: 0, c0: 0, r1: 4, c1: 0 });

  // 편집기는 열리지 않는다
  expect(
    await page.evaluate(
      () => document.getElementById("portal")?.querySelector("textarea, input") ?? null,
    ),
  ).toBeNull();
});

test("수식어 없는 셀 더블클릭은 편집기를 연다 (회귀)", async ({ page }) => {
  await setup(page);
  const { x0, y0 } = await calibrate(page);

  // B2(값 11) 더블클릭 → 포털에 오버레이 편집기
  await page.mouse.dblclick(x0 + 1 * 88, y0 + 1 * 34);
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
    .toBe("11");
});
