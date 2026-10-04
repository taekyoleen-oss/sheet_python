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

test("카드 선택(목차·이동) → 접힌 카드도 펼치고 카드 상단을 화면 상단에 맞춘다", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);
  const ids: string[] = await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    const out: string[] = [];
    for (let i = 0; i < 8; i++) {
      const id = st.addPyBlock(sid, { r: i * 3, c: 0 });
      st.setBlockCode(id, Array.from({ length: 8 }, (_, k) => `x${k} = ${k}`).join("\n"));
      out.push(id);
    }
    st.setBlockCollapsed(out[6], true);
    return out;
  });
  await page.evaluate((id) => (window as any).__pygridStore.getState().setFocusBlock(id), ids[6]);

  const card = page.locator(`[data-block-id="${ids[6]}"]`);
  await expect(card.getByLabel("Python 코드")).toBeVisible(); // 펼쳐짐
  await expect
    .poll(() =>
      card.evaluate((el) => {
        let p = el.parentElement;
        while (p && !(p.scrollHeight > p.clientHeight && /auto|scroll/.test(getComputedStyle(p).overflowY)))
          p = p.parentElement;
        return p ? Math.abs(el.getBoundingClientRect().top - p.getBoundingClientRect().top) : -1;
      }),
    )
    .toBeLessThan(12);
});
