import { expect, test, type Page } from "@playwright/test";

// 편집기·실행 결과를 일반 파이썬처럼:
//  Tab — 자동완성 수락 / 줄 중간이면 자동완성 열기 / 들여쓰기 4칸, Shift+Tab 내어쓰기
//  런타임 자동완성 — 실행으로 만든 변수, obj. 속성, df[" 열 이름
//  결과 — print만 보이고 None·패키지 로더 안내는 없음, 시트에 없는 블록은 남은 출력 지정 무시,
//         오류는 트레이스백이 먼저

/* eslint-disable @typescript-eslint/no-explicit-any */

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as any).__pygridStore !== "undefined" && (window as any).__pygridReady === true,
  );
  await page.evaluate(() => (window as any).__pygridStore.getState().newWorkbook());
}

const block0 = (page: Page) =>
  page.evaluate(() => (window as any).__pygridStore.getState().workbook.pyBlocks[0] ?? null);

test("Tab 자동완성·들여쓰기 + 일반 파이썬 같은 실행 결과", async ({ page }) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const id = st.addPyBlock(st.activeSheetId, { r: 0, c: 0 });
    st.setBlockCode(
      id,
      'import pandas as pd\nscores = pd.DataFrame({"amount": [0.123456, 2.0], "age": [30, 40]})\nprint(scores.round(2))',
    );
    // 예전에 시트용으로 고른 출력 지정이 남아 있어도(시트에는 없음) 결과에는 마지막 줄만
    st.setBlockOutput(id, { variable: "scores" });
  });
  await page.getByRole("button", { name: "실행", exact: true }).click();
  const result = page.getByTestId("cell-result");
  await expect(result).toContainText("amount", { timeout: 150_000 });
  await expect(result).not.toContainText("None");
  await expect(result).not.toContainText("already loaded");
  await expect(result.locator("table")).toHaveCount(0); // print만 — 표가 따로 붙지 않는다

  // 편집기: 끝으로 가서 새 줄
  const editor = page.getByLabel("Python 코드");
  await editor.click();
  await page.keyboard.press("Control+End");
  await page.keyboard.press("Enter");

  // 변수 이름 자동완성 → Tab 수락
  await page.keyboard.type("scor", { delay: 120 });
  await expect(page.locator(".cm-tooltip-autocomplete")).toContainText("scores", { timeout: 15_000 });
  await page.keyboard.press("Tab");
  // 속성(dir) 자동완성 → Tab 수락
  await page.keyboard.type(".hea", { delay: 120 });
  await expect(page.locator(".cm-tooltip-autocomplete")).toContainText("head", { timeout: 15_000 });
  await page.keyboard.press("Tab");
  await expect.poll(async () => (await editor.textContent()) ?? "").toContain("scores.head");

  // 열 이름(df[") 자동완성
  await page.keyboard.press("Enter");
  await page.keyboard.type('scores["am', { delay: 120 });
  await expect(page.locator(".cm-tooltip-autocomplete")).toContainText("amount", { timeout: 15_000 });
  await page.keyboard.press("Tab");
  await expect.poll(async () => (await editor.textContent()) ?? "").toContain('scores["amount');

  // 줄 맨 앞 Tab = 4칸 들여쓰기, Shift+Tab = 내어쓰기
  await page.keyboard.press("Control+a");
  await page.keyboard.type("if True:");
  await page.keyboard.press("Enter"); // 자동 들여쓰기(4칸)
  await page.keyboard.press("Shift+Tab");
  await page.keyboard.press("Tab");
  await page.keyboard.type("x = 1");
  await expect.poll(async () => (await editor.textContent()) ?? "").toContain("if True:    x = 1");

  // 오류: 트레이스백이 먼저
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setBlockCode(st.workbook.pyBlocks[0].id, 'print("앞")\nundefined_name');
  });
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect(result).toContainText("Traceback", { timeout: 60_000 });
  await expect(result).toContainText("NameError");
  await expect(result).toContainText("앞");
  expect((await block0(page)).last.status).toBe("error");
});
