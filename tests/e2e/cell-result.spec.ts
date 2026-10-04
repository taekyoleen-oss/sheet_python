import { expect, test, type Page } from "@playwright/test";

// 새 블록 기본: 결과는 코드 아래(노트북식)에만 — 시트에 쓰지 않는다. '시트로 보내기'를 켜야 셀에 spill.

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as any).__pygridStore !== "undefined" &&
      (window as any).__pygridReady === true,
  );
  await page.evaluate(() => (window as any).__pygridStore.getState().newWorkbook());
}

const srcCount = (page: Page) =>
  page.evaluate(
    () =>
      Object.values((window as any).__pygridStore.getState().workbook.sheets[0].cells).filter(
        (c: any) => typeof c.src === "string",
      ).length,
  );

test("셀 아래 실행 결과 → 숨기기 → 시트로 보내기 켜면 spill, 끄면 지움", async ({ page }) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);
  await page.evaluate(() =>
    (window as any).__pygridStore.getState().setSelection({ r0: 0, c0: 0, r1: 0, c1: 0 }),
  );
  await page.keyboard.press("Control+Shift+P");
  const textarea = page.getByLabel("Python 코드");
  await expect(textarea).toBeFocused();
  await textarea.fill('print("안녕")\n[1, 2, 3]');
  await page.getByRole("button", { name: "실행", exact: true }).click();

  const result = page.getByTestId("cell-result");
  await expect(result).toContainText("안녕", { timeout: 240_000 });
  await expect(result).toContainText("[1, 2, 3]"); // 마지막 값 미리보기
  expect(await srcCount(page)).toBe(0); // 시트에는 아무것도 쓰지 않는다
  await expect(page.getByTestId("output-list")).toHaveCount(0);
  await expect(page.getByTitle("해당 셀로 이동")).toHaveCount(0);

  await result.getByRole("button", { name: "실행 결과 숨기기" }).click();
  await expect(result).not.toContainText("안녕");
  await result.getByRole("button", { name: "실행 결과 보기" }).click();
  await expect(result).toContainText("안녕");

  // 시트로 보내기 → 출력 설정·해당 셀로 이동이 나타나고 자동 재실행으로 spill
  await page.getByRole("button", { name: "시트로 보내기" }).click();
  await expect(page.getByTestId("output-list")).toBeVisible();
  await expect(page.getByTitle("해당 셀로 이동")).toBeVisible();
  await expect.poll(() => srcCount(page), { timeout: 60_000 }).toBeGreaterThan(0);

  await page.getByRole("button", { name: "시트로 보내기" }).click();
  await expect.poll(() => srcCount(page)).toBe(0);
});
