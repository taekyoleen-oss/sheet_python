import { expect, test, type Page } from "@playwright/test";

// 부록 O.5: 모델 결과 → 시트 다이얼로그 — 실런타임(scikit-learn)으로 모델 적합 →
// 모델별 속성 표(현재 값·코드) 확인 → 절편을 셀에 놓기 + 코드를 새 블록으로 보내기

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

test("모델 결과 → 시트: 속성 표에 값·코드, 셀 배치와 새 블록", async ({ page }) => {
  test.setTimeout(400_000);
  await page.goto("/");
  await waitForApp(page);
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setSelection({ r0: 0, c0: 3, r1: 0, c1: 3 });
  });
  await page.keyboard.press("Control+Shift+P");
  const textarea = page.getByLabel("Python 코드");
  await expect(textarea).toBeFocused();
  await textarea.fill(
    [
      "import pandas as pd",
      "from sklearn.linear_model import LinearRegression",
      'df = pd.DataFrame({"x": [1.0, 2, 3, 4], "y": [3.0, 5, 7, 9]})',
      'lr = LinearRegression().fit(df[["x"]], df.y)',
      "lr.score(df[['x']], df.y)",
    ].join("\n"),
  );
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(async () => (await cellAt(page, "0:3"))?.v, { timeout: 300_000, intervals: [1000] })
    .toBe(1);

  await page.getByRole("button", { name: /모델 결과 → 시트/ }).click();
  const table = page.getByTestId("model-members");
  await expect(table).toBeVisible({ timeout: 30_000 });
  // 모델 종류(sklearn 선형)에 맞는 속성과 그 코드·현재 값
  await expect(table.getByText("lr.intercept_", { exact: true })).toBeVisible();
  await expect(table.getByText("1", { exact: true }).first()).toBeVisible(); // 절편 = 1
  await expect(table.getByText("특성 이름")).toBeVisible();

  // 기본 선택(요약표)을 끄고 절편만 셀에 놓는다
  await table.getByLabel("계수표 (절편 + 특성별 계수) 셀에 놓기").uncheck();
  await table.getByLabel("모델 설정 (하이퍼파라미터) 셀에 놓기").uncheck();
  await table.getByLabel("절편 셀에 놓기").check();
  await page.getByRole("button", { name: "시트에 놓기 (1)" }).click();

  await expect
    .poll(
      () =>
        page.evaluate(() => {
          const st = (window as any).__pygridStore.getState();
          const b = st.workbook.pyBlocks[0];
          const o = b.outputs.find((x: any) => x.selection?.variable === "lr.intercept_");
          return o ? st.workbook.sheets[0].cells[`${o.anchor.r}:${o.anchor.c}`]?.v : null;
        }),
      { timeout: 60_000, intervals: [500] },
    )
    .toBeCloseTo(1, 10);

  // 코드를 새 블록으로 → 블록 2개
  await page.getByRole("button", { name: /모델 결과 → 시트/ }).first().click();
  await expect(page.getByTestId("model-members")).toBeVisible({ timeout: 30_000 });
  await page.getByLabel("lr: 특성 수 코드를 새 블록으로").click();
  await expect
    .poll(() =>
      page.evaluate(() =>
        (window as any).__pygridStore.getState().workbook.pyBlocks.map((b: any) => b.code),
      ),
    )
    .toContainEqual("# lr: 특성 수\nlr.n_features_in_");
});
