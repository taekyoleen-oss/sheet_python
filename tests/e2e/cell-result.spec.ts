import { expect, test, type Page } from "@playwright/test";

// 새 블록 기본: 결과는 코드 아래(노트북식)에만 — 시트에 쓰지 않는다. '시트에 추가'를 눌러야 셀에 spill.

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

test("셀 아래 실행 결과 → 숨기기 → 시트에 추가하면 spill, 출력을 모두 지우면 빠짐", async ({ page }) => {
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
  await expect(page.getByTestId("exec-count")).toHaveCount(0); // 미실행 — 순번 없음
  await page.getByRole("button", { name: "실행", exact: true }).click();

  const result = page.getByTestId("cell-result");
  await expect(result).toContainText("안녕", { timeout: 240_000 });
  await expect(result).toContainText("[1, 2, 3]"); // 마지막 값 미리보기
  await expect(page.getByTestId("exec-count")).toHaveText(/^\[\d+\]$/); // 실행 순번
  expect(await srcCount(page)).toBe(0); // 시트에는 아무것도 쓰지 않는다
  await expect(page.getByTestId("output-list")).toHaveCount(0);
  await expect(page.getByTitle("해당 셀로 이동")).toHaveCount(0);

  await result.getByRole("button", { name: "실행 결과 숨기기" }).click();
  await expect(result).not.toContainText("안녕");
  await result.getByRole("button", { name: "실행 결과 보기" }).click();
  await expect(result).toContainText("안녕");

  // 시트에 추가 → 셀이 정해지지 않은 출력('셀 선택') + 위치 지정 모드. 셀을 고르기 전엔 아무것도 쓰지 않는다
  await page.getByRole("button", { name: "시트에 추가" }).click();
  await expect(page.getByTestId("output-list")).toBeVisible();
  await expect(page.getByRole("button", { name: "출력 1 위치" })).toHaveText("셀 선택");
  await expect(page.getByText("결과를 놓을 셀을 클릭하세요")).toBeVisible();
  await expect(page.getByTitle("해당 셀로 이동")).toHaveCount(0);
  expect(await srcCount(page)).toBe(0);

  // 셀(C2)을 고르고 실행하면 그 자리에 놓인다
  await page.keyboard.press("Escape");
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const b = st.workbook.pyBlocks[0];
    st.setOutputAnchor(b.id, b.outputs[0].id, { r: 1, c: 2 });
  });
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect.poll(() => srcCount(page), { timeout: 60_000 }).toBeGreaterThan(0);
  await expect(page.getByTitle("해당 셀로 이동")).toBeVisible();

  // 다시 누르면 위치 미정 출력이 하나 더
  await page.getByRole("button", { name: "시트에 추가" }).click();
  await expect(page.locator("[data-output-id]")).toHaveCount(2);
  await expect(page.getByRole("button", { name: "출력 2 위치" })).toHaveText("셀 선택");
  await page.keyboard.press("Escape");

  // 출력을 모두 지우면 시트에서 빠진다
  await page.getByRole("button", { name: "출력 2 삭제" }).click();
  await page.getByRole("button", { name: "출력 1 삭제" }).click();
  await expect(page.getByTestId("output-list")).toHaveCount(0);
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
