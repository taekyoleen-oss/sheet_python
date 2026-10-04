import { expect, test, type Page } from "@playwright/test";


// 새 블록 기본은 '시트로 보내지 않음' — 이 스펙은 spill을 검증하므로 켜고 시작한다
test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    (window as unknown as { __pygridToSheetDefault?: boolean }).__pygridToSheetDefault = true;
  });
});

// M8: 키보드 전용 블록 추가→실행 흐름(마우스 0회) + 800×600 탭 전환 UI

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

test("키보드 전용: Ctrl+1 → 방향키 → Ctrl+Shift+P → 코드 → Ctrl+Enter → Ctrl+1", async ({ page }) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  // Ctrl+1: 그리드 포커스 + 기본 선택 A1
  await page.keyboard.press("Control+1");
  await expect
    .poll(() =>
      page.evaluate(
        () => document.activeElement?.getAttribute("data-testid") ?? null,
      ),
    )
    .toBe("data-grid-canvas");

  // 방향키로 C3(2,2)까지 이동 — glide 네이티브 키보드
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("ArrowRight");
  await expect
    .poll(() =>
      page.evaluate(() => (window as any).__pygridStore.getState().selection),
    )
    .toEqual({ r0: 2, c0: 2, r1: 2, c1: 2 });

  // Ctrl+Shift+P → 블록 생성 + 편집기 자동 포커스
  await page.keyboard.press("Control+Shift+P");
  await expect
    .poll(() =>
      page.evaluate(() => (window as any).__pygridStore.getState().workbook.pyBlocks.length),
    )
    .toBe(1);
  await expect(page.getByLabel("Python 코드")).toBeFocused();

  // 코드 입력 → Ctrl+Enter 실행 (부트 포함 폴링)
  await page.keyboard.type("1 + 1");
  await page.keyboard.press("Control+Enter");
  await expect
    .poll(
      () =>
        page.evaluate(
          () =>
            (window as any).__pygridStore.getState().workbook.sheets[0].cells["2:2"]?.v ??
            null,
        ),
      { timeout: 150_000, intervals: [1000] },
    )
    .toBe(2);

  // Ctrl+1 복귀 — 편집기(텍스트 입력) 안에서도 동작해야 한다
  await page.keyboard.press("Control+1");
  await expect
    .poll(() =>
      page.evaluate(
        () => document.activeElement?.getAttribute("data-testid") ?? null,
      ),
    )
    .toBe("data-grid-canvas");
});

test("800×600: 그리드↔Python 탭 전환 UI + 속성 창", async ({ page }) => {
  await page.setViewportSize({ width: 800, height: 600 });
  await page.goto("/");
  await waitForApp(page);

  const tabs = page.getByRole("tablist", { name: "화면 전환" });
  await expect(tabs).toBeVisible();
  await expect(page.getByTestId("data-grid-canvas")).toBeVisible();

  // Python 탭 → 패널 표시, 그리드 숨김
  await tabs.getByRole("tab", { name: "Python" }).click();
  await expect(page.getByRole("tab", { name: "목차" })).toBeVisible(); // 좁은 화면: 목차는 네 번째 탭
  await expect(page.getByTestId("data-grid-canvas")).toHaveCount(0);

  // 부록 P.7: 하단 패널 대신 속성 창 — 좁은 화면에서도 오른쪽에서 덮으며 열린다
  await page.getByRole("button", { name: "속성 창 열기 (변수·파일)" }).click();
  await expect(page.getByRole("tab", { name: /진단/ })).toBeVisible();
  await page.getByTestId("properties-panel").getByRole("button", { name: "속성 창 닫기" }).click();

  // 그리드 복귀
  await tabs.getByRole("tab", { name: "그리드" }).click();
  await expect(page.getByTestId("data-grid-canvas")).toBeVisible();
});
