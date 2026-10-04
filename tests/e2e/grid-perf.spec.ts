import { expect, test, type Page } from "@playwright/test";

// window.__pygridStore는 app/page.tsx에서 노출, __pygridReady는 WorkbookShell 복원 완료 신호

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as unknown as { __pygridStore?: unknown }).__pygridStore !==
        "undefined" &&
      (window as unknown as { __pygridReady?: boolean }).__pygridReady === true,
  );
  await page.waitForSelector('[data-testid="data-grid-canvas"]');
}

test("10,000×50 시트 로드 후 스크롤해도 응답성이 유지된다", async ({ page }) => {
  const pageErrors: Error[] = [];
  page.on("pageerror", (e) => pageErrors.push(e));

  await page.goto("/");
  await waitForApp(page);

  const loadMs = await page.evaluate(() => {
    const t0 = performance.now();
    /* eslint-disable @typescript-eslint/no-explicit-any */
    const store = (window as any).__pygridStore;
    const state = store.getState();
    const wb = JSON.parse(JSON.stringify(state.workbook));
    wb.id = crypto.randomUUID();
    wb.title = "성능 테스트";
    const sheet = wb.sheets[0];
    sheet.rowCount = 10000;
    sheet.colCount = 50;
    const cells: Record<string, { v: number; t: string }> = {};
    for (let r = 0; r < 10000; r++) {
      for (let c = 0; c < 50; c++) cells[`${r}:${c}`] = { v: r * 100 + c, t: "n" };
    }
    sheet.cells = cells;
    state.loadWorkbook(wb);
    return performance.now() - t0;
  });
  // 느슨한 상한 — 크래시/멈춤 감지가 목적
  expect(loadMs).toBeLessThan(30_000);

  // 캔버스 위에는 .dvn-scroller가 겹쳐 있어 hover가 막힌다 — 좌표로 직접 이동
  const box = await page.getByTestId("data-grid-canvas").first().boundingBox();
  if (!box) throw new Error("그리드 캔버스를 찾지 못했습니다");
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  const scrollStart = Date.now();
  for (let i = 0; i < 30; i++) {
    await page.mouse.wheel(0, 1200);
  }
  const scrollMs = Date.now() - scrollStart;
  expect(scrollMs).toBeLessThan(20_000);

  // 스크롤 후에도 메인 스레드가 응답하는지 (rAF 왕복)
  const rafMs = await page.evaluate(
    () =>
      new Promise<number>((resolve) => {
        const t = performance.now();
        requestAnimationFrame(() => resolve(performance.now() - t));
      }),
  );
  expect(rafMs).toBeLessThan(2_000);

  // 그리드 상태 확인 — 크래시 없이 10,000행 유지
  const rowCount = await page.evaluate(
    () => (window as any).__pygridStore.getState().workbook.sheets[0].rowCount,
  );
  expect(rowCount).toBe(10000);
  expect(pageErrors).toEqual([]);
});

test("Ctrl+Z / Ctrl+Y로 실행 취소·다시 실행이 된다", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);

  const cellValue = () =>
    page.evaluate(
      () =>
        /* eslint-disable @typescript-eslint/no-explicit-any */
        (window as any).__pygridStore.getState().workbook.sheets[0]?.cells["0:0"]?.v ??
        null,
    );

  await page.evaluate(() => {
    const store = (window as any).__pygridStore;
    store.getState().newWorkbook();
    store.getState().setGridCompact(false); // 부록 P.6: 기본 크기 보기 기준
    const state = store.getState();
    state.setCellValue(state.workbook.sheets[0].id, 0, 0, { v: "이전", t: "s" });
  });
  await page.waitForTimeout(400); // undo 이력 스로틀(300ms) 경계
  await page.evaluate(() => {
    const state = (window as any).__pygridStore.getState();
    state.setCellValue(state.workbook.sheets[0].id, 0, 0, { v: "이후", t: "s" });
  });
  expect(await cellValue()).toBe("이후");

  await page.keyboard.press("Control+z");
  expect(await cellValue()).toBe("이전"); // 실행 취소
  await page.keyboard.press("Control+y");
  expect(await cellValue()).toBe("이후"); // 다시 실행
  await page.keyboard.press("Control+z");
  expect(await cellValue()).toBe("이전");
  await page.keyboard.press("Control+Shift+KeyZ");
  expect(await cellValue()).toBe("이후"); // Ctrl+Shift+Z도 다시 실행
});

test("자동 저장 후 새로고침하면 값이 복원된다", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);

  await page.evaluate(() => {
    /* eslint-disable @typescript-eslint/no-explicit-any */
    const store = (window as any).__pygridStore;
    store.getState().newWorkbook();
    const state = store.getState();
    state.setCellValue(state.workbook.sheets[0].id, 0, 0, { v: "복원 확인", t: "s" });
  });

  // 자동 저장 디바운스 2초 + 여유
  await page.waitForTimeout(3500);
  await page.reload();
  await waitForApp(page);

  await expect
    .poll(
      () =>
        page.evaluate(
          () =>
            (window as any).__pygridStore.getState().workbook.sheets[0]?.cells["0:0"]
              ?.v ?? null,
        ),
      { timeout: 10_000 },
    )
    .toBe("복원 확인");
});

test("셀 클릭 직후 초고속 타이핑에도 첫 글자가 유실되지 않는다 (QA 보정)", async ({ page }) => {
  await page.goto("/");
  await page.waitForFunction(
    () =>
      typeof (window as any).__pygridStore !== "undefined" &&
      (window as any).__pygridReady === true,
  );
  await page.evaluate(() => {
    (window as any).__pygridStore.getState().newWorkbook();
    (window as any).__pygridStore.getState().setGridCompact(false); // 부록 P.6: 88×34 기준
  });
  await page.waitForTimeout(300);
  // 캔버스 좌표 보정: 임의 지점 클릭 → 선택된 셀로 원점 역산 (행 34px·열 88px)
  const box = (await page.locator('[data-testid="data-grid-canvas"]').boundingBox())!;
  await page.mouse.click(box.x + 200, box.y + 100);
  const sel = await page.evaluate(() => (window as any).__pygridStore.getState().selection);
  const x0 = box.x + 200 - sel.c0 * 88;
  const y0 = box.y + 100 - sel.r0 * 34;
  // B2 클릭 즉시 지연 0ms 타이핑 — 오버레이 편집기 마운트 전에 키가 캔버스로 떨어지는 상황
  await page.mouse.click(x0 + 1 * 88, y0 + 1 * 34);
  await page.keyboard.type("hello", { delay: 0 });
  // 편집기 보정(버퍼→전체 문자열)이 반영된 것을 확인한 뒤 Enter — 사용자도 화면을 보고 확정한다
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
    .toBe("hello");
  await page.keyboard.press("Enter");
  await expect
    .poll(
      () =>
        page.evaluate(
          () => (window as any).__pygridStore.getState().workbook.sheets[0].cells["1:1"]?.v ?? null,
        ),
      { timeout: 10_000 },
    )
    .toBe("hello");
});

test("hover 셀 툴바가 카드 헤더(제목·상태 칩)를 덮지 않는다 (QA 보정)", async ({ page }) => {
  await page.goto("/");
  await page.waitForFunction(
    () =>
      typeof (window as any).__pygridStore !== "undefined" &&
      (window as any).__pygridReady === true,
  );
  await page.evaluate(() => {
    (window as any).__pygridStore.getState().newWorkbook();
    // newWorkbook 이후엔 상태를 다시 읽어야 새 시트 id를 얻는다
    const st = (window as any).__pygridStore.getState();
    st.addPyBlock(st.workbook.sheets[0].id, { r: 0, c: 2 });
  });
  const card = page.locator('[data-block-kind="code"]').first();
  await card.hover();
  const toolbar = page.getByTestId("cell-toolbar");
  await expect(toolbar).toBeVisible();
  // 인라인 배치 — 카드 안에 있고(위 카드를 가리지 않음) 제목 입력과 수평으로 겹치지 않는다
  const tb = (await toolbar.boundingBox())!;
  const title = (await card.getByLabel("블록 제목").boundingBox())!;
  const cardBox = (await card.boundingBox())!;
  expect(tb.y).toBeGreaterThanOrEqual(cardBox.y); // 카드 밖(이전 카드 영역)으로 나가지 않는다
  expect(tb.x).toBeGreaterThanOrEqual(title.x + title.width); // 제목을 덮지 않는다
});
