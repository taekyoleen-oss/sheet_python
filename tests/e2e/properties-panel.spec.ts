import { expect, test, type Page } from "@playwright/test";

// 부록 P: 속성 창 — 새 폴더 → 작업 폴더 지정(os.chdir) → 파일 올리기 → 불러오기(코드 블록) →
// 실행 → 변수 목록 → 변수를 시트에 연결 출력. 고정/겹침 전환.

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

const st = (page: Page) => page.evaluate(() => (window as any).__pygridStore.getState().workbook);

test("속성 창: 작업 폴더 지정 → 파일 불러오기 → 변수 → 시트에 보이기", async ({ page }) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  await page.getByRole("button", { name: "속성 창 열기 (변수·파일)" }).click();
  const panel = page.getByTestId("properties-panel");
  await expect(panel).toBeVisible();
  await expect(panel.getByTestId("props-cwd")).toHaveText("/Users/tklee/Downloads", { timeout: 150_000 });

  // 고정 ↔ 겹침
  const aside = page.getByRole("complementary", { name: "속성 창" });
  await expect(aside).not.toHaveClass(/absolute/);
  await panel.getByRole("button", { name: "화면 위에 겹치기" }).click();
  await expect(aside).toHaveClass(/absolute/);
  await panel.getByRole("button", { name: "화면에 고정" }).click();
  await expect(aside).not.toHaveClass(/absolute/);

  // 새 폴더 data → 열기 → 작업 폴더로 지정
  await panel.getByRole("button", { name: "새 폴더" }).click();
  await panel.getByLabel("새 폴더 이름").fill("data");
  await panel.getByLabel("새 폴더 이름").press("Enter");
  await panel.locator('[data-entry="data"]').click();
  await panel.getByRole("button", { name: "이 폴더를 작업 폴더로" }).click();
  await expect(panel.getByTestId("props-cwd")).toHaveText("/Users/tklee/Downloads/data");
  expect((await st(page)).workDir).toBe("/Users/tklee/Downloads/data");

  // 작업 폴더에 CSV 올리기
  await panel.locator('input[type="file"]').setInputFiles({
    name: "claims.csv",
    mimeType: "text/csv",
    buffer: Buffer.from("age,amount\n30,100\n40,250\n50,400\n"),
  });
  const file = panel.locator('[data-entry="claims.csv"]');
  await expect(file).toBeVisible();

  // 선택 → 불러오기 → 코드 블록 (실행은 사용자가)
  await file.click();
  await panel.getByRole("button", { name: "불러오기", exact: true }).click();
  await expect.poll(async () => (await st(page)).pyBlocks.length).toBe(1);
  const code: string = (await st(page)).pyBlocks[0].code;
  expect(code).toContain('claims = pd.read_csv("claims.csv")');

  await page.getByRole("button", { name: "실행", exact: true }).click();
  const row = panel.getByTestId("props-variables").locator('[data-var="claims"]');
  await expect(row).toBeVisible({ timeout: 90_000 });
  await expect(row).toContainText("3×2");
  // 실행으로 만든 변수만 — 초기화 스크립트의 pd·np는 보이지 않는다
  await expect(panel.getByTestId("props-variables").locator('[data-var="pd"]')).toHaveCount(0);

  // 변수 → 스프레드시트 (연결 출력, 선택 셀 H1)
  await page.evaluate(() =>
    (window as any).__pygridStore.getState().setSelection({ r0: 0, c0: 7, r1: 0, c1: 7 }),
  );
  // 클릭 → 세부 내용 팝업(상위 행·열 형) → 스프레드시트에 보이기 → 연결 출력 확정
  await row.click();
  const detail = page.getByTestId("variable-detail");
  await expect(detail.getByText("int64").first()).toBeVisible({ timeout: 30_000 });
  await expect(detail.getByRole("cell", { name: "250" })).toBeVisible();
  await expect(detail.getByLabel("소수 자리 (전체)")).toHaveCount(0); // 정수 열뿐이면 자리수 선택 없음
  // 요약 통계(float) — 기본 소수 5자리, 전체 2자리, 열(mean)만 0자리 — 보기에만
  await detail.getByRole("tab", { name: "요약 통계" }).click();
  await expect(detail.getByRole("cell", { name: "40.00000" }).first()).toBeVisible({ timeout: 30_000 });
  await detail.getByLabel("소수 자리 (전체)").selectOption("2");
  await expect(detail.getByRole("cell", { name: "40.00" }).first()).toBeVisible();
  await detail.getByLabel("mean 소수 자리").selectOption("0");
  await expect(detail.getByRole("cell", { name: "40", exact: true })).toBeVisible();
  await expect(detail.getByRole("cell", { name: "3.00" }).first()).toBeVisible(); // count 열은 전체(2자리) 그대로
  // 열 머리 오른쪽 선을 끌어 열 너비 조절
  const th = detail.locator("thead th").nth(1);
  const w0 = (await th.boundingBox())!.width;
  const sep = (await detail.getByRole("separator", { name: "열 너비 조절" }).nth(1).boundingBox())!;
  await page.mouse.move(sep.x + sep.width / 2, sep.y + 10);
  await page.mouse.down();
  await page.mouse.move(sep.x + sep.width / 2 + 80, sep.y + 10, { steps: 5 });
  await page.mouse.up();
  await expect.poll(async () => (await th.boundingBox())!.width).toBeGreaterThan(w0 + 60);
  await detail.getByRole("button", { name: "스프레드시트에 보이기" }).click();
  await page.getByRole("button", { name: "시트에 보이기", exact: true }).click();
  await expect
    .poll(
      async () => {
        const wb = await st(page);
        return wb.sheets[0].cells["0:7"]?.v ?? null;
      },
      { timeout: 60_000, intervals: [500] },
    )
    .toBe("age");
  const outputs = (await st(page)).pyBlocks[0].outputs;
  expect(outputs.some((o: any) => o.selection?.variable === "claims")).toBe(true);
});
