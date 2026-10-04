import { expect, test, type Page } from "@playwright/test";

// 좌우 패널 접기 — 스프레드시트·Python 패널 감추기/펼치기 + 그리드 최소 폭 15%

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as unknown as { __pygridStore?: unknown }).__pygridStore !==
        "undefined" &&
      (window as unknown as { __pygridReady?: boolean }).__pygridReady === true,
  );
  await page.waitForSelector('[data-testid="data-grid-canvas"]');
}

/* eslint-disable @typescript-eslint/no-explicit-any */
const collapseState = (page: Page) =>
  page.evaluate(() => {
    const s = (window as any).__pygridStore.getState();
    return { grid: s.gridCollapsed, python: s.pyCollapsed };
  });

const widthOf = async (page: Page, selector: string) =>
  (await page.locator(selector).boundingBox())?.width ?? 0;

test("스프레드시트 감추기 → Python이 폭 대부분 → 스트립으로 복원", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);

  const mainWidth = await widthOf(page, '[data-testid="workspace"]');
  await page.getByRole("button", { name: "스프레드시트 감추기 (Ctrl+Alt+1)" }).click();

  await expect(page.getByTestId("data-grid-canvas")).toHaveCount(0);
  await expect(page.locator("#grid")).toHaveCount(0);
  const strip = page.getByTestId("strip-grid");
  await expect(strip).toBeVisible();
  await expect(strip).toContainText("시트");

  // Python 패널이 스트립(28px)을 뺀 나머지 전부를 차지한다
  expect(await widthOf(page, "#python")).toBeGreaterThan(mainWidth * 0.9);

  await strip.click();
  await expect(page.getByTestId("data-grid-canvas")).toBeVisible();
  await expect(page.getByTestId("strip-grid")).toHaveCount(0);
  expect(await widthOf(page, "#python")).toBeLessThan(mainWidth * 0.75); // 기본 40:60으로 복원
});

test("Python 감추기(Ctrl+Alt+2) → 그리드가 폭 대부분 → 스트립으로 복원", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);

  const mainWidth = await widthOf(page, '[data-testid="workspace"]');
  await page.keyboard.press("Control+Alt+2");

  await expect(page.locator("#python")).toHaveCount(0);
  const strip = page.getByTestId("strip-python");
  await expect(strip).toBeVisible();
  await expect(strip).toContainText("Python");
  expect(await widthOf(page, "#grid")).toBeGreaterThan(mainWidth * 0.9);

  await strip.click();
  await expect(page.locator("#python")).toBeVisible();
  await expect(page.getByTestId("strip-python")).toHaveCount(0);
});

test("둘 다 접기 시도 → 나중에 접은 쪽만 접히고 반대쪽은 자동으로 펼쳐진다", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);

  await page.getByRole("button", { name: "스프레드시트 감추기 (Ctrl+Alt+1)" }).click();
  expect(await collapseState(page)).toEqual({ grid: true, python: false });

  // 그리드가 접힌 상태에서 Python도 접으면 → 그리드가 자동으로 펼쳐진다 (빈 화면 방지)
  await page.getByRole("button", { name: "Python 패널 감추기 (Ctrl+Alt+2)" }).click();
  expect(await collapseState(page)).toEqual({ grid: false, python: true });
  await expect(page.getByTestId("data-grid-canvas")).toBeVisible();
  await expect(page.getByTestId("strip-grid")).toHaveCount(0);
  await expect(page.getByTestId("strip-python")).toBeVisible();

  // Ctrl+1(그리드 포커스)은 접힌 그리드를 펼친다
  await page.keyboard.press("Control+Alt+1");
  expect(await collapseState(page)).toEqual({ grid: true, python: false });
  await page.keyboard.press("Control+1");
  expect(await collapseState(page)).toEqual({ grid: false, python: false });
});

test("접힘은 세션 한정 — 새로고침하면 기본값(시트·Python 함께 표시)으로 돌아온다", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);

  await page.getByRole("button", { name: "Python 패널 감추기 (Ctrl+Alt+2)" }).click();
  await expect(page.getByTestId("strip-python")).toBeVisible();
  await expect(page.locator("#python")).toHaveCount(0);

  await page.reload();
  await waitForApp(page);
  // 앱의 기본 상태 = 두 패널이 함께 보인다
  await expect(page.locator("#grid")).toHaveCount(1);
  await expect(page.locator("#python")).toHaveCount(1);
  await expect(page.getByTestId("strip-python")).toHaveCount(0);
  expect(await collapseState(page)).toEqual({ grid: false, python: false });
});

test("그리드 최소 폭 15% — 핸들을 끝까지 밀면 기본 40%보다 좁아진다", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);

  const mainWidth = await widthOf(page, '[data-testid="workspace"]');
  const before = await widthOf(page, "#grid");
  expect(before / mainWidth).toBeGreaterThan(0.3); // 부록 P.6 기본 그리드 40%
  expect(before / mainWidth).toBeLessThan(0.5);

  // 구분자 포커스 → Home = 왼쪽 끝까지 (최소 크기까지 축소)
  const handle = page.locator('#grid + [data-slot="resizable-handle"]');
  await handle.focus();
  await page.keyboard.press("Home");

  await expect
    .poll(async () => (await widthOf(page, "#grid")) / mainWidth)
    .toBeLessThan(0.3); // 기본 40%에서 더 줄어든다
  const ratio = (await widthOf(page, "#grid")) / mainWidth;
  expect(ratio).toBeGreaterThan(0.1); // minSize 15%에서 멈춘다
});
