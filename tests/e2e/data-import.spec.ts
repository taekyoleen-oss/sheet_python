import fs from "node:fs";
import { expect, test, type Page } from "@playwright/test";


// 새 블록 기본은 '시트로 보내지 않음' — 이 스펙은 spill을 검증하므로 켜고 시작한다
test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    (window as unknown as { __pygridToSheetDefault?: boolean }).__pygridToSheetDefault = true;
  });
});

// R5: 데이터 불러오기 — 샘플 xlsx가 시트 + 워커 FS에 이중 착지하고,
// sheet() 로드 블록과 pandas(pd.read_excel) 로드 블록이 모두 실행돼 spill되는지.

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

/** 파일 메뉴 → 데이터 불러오기 → 샘플 데이터셋 → 항목 클릭 → 옵션 다이얼로그 대기 */
async function openSampleDialog(page: Page, item: RegExp) {
  await page.getByRole("button", { name: "파일", exact: true }).click();
  await page.getByRole("menuitem", { name: "데이터 불러오기" }).hover();
  await page.getByRole("menuitem", { name: "샘플 데이터셋" }).hover();
  await page.getByRole("menuitem", { name: item }).click();
  await expect(page.getByRole("dialog", { name: "데이터 불러오기" })).toBeVisible();
}

/** 첫 블록의 spill 셀 수 (블록이 놓인 시트에서 src 소유 기준) */
const spillCount = (page: Page) =>
  page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const b = st.workbook.pyBlocks[0];
    if (!b) return 0;
    const sheet = st.workbook.sheets.find((s: any) => s.id === b.sheetId);
    if (!sheet) return 0;
    return Object.values(sheet.cells).filter(
      (c: any) => typeof c.src === "string" && c.src.split(":")[0] === b.id,
    ).length;
  });

/** 첫 블록의 실행 결과 — 실패 시 summaryKo가 diff에 보이도록 */
const lastResult = (page: Page) =>
  page.evaluate(() => {
    const l = (window as any).__pygridStore.getState().workbook.pyBlocks[0]?.last;
    return l ? { status: l.status, summaryKo: l.summaryKo ?? null } : null;
  });

test("샘플 policy.xlsx 기본 옵션 → 새 시트 + xl 로드 블록 → 실행 → spill", async ({
  page,
}) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  await openSampleDialog(page, /policy/);
  // 기본 옵션: 시트 ✓ · 워커 FS ✓ · sheet() 참조
  const dialog = page.getByRole("dialog", { name: "데이터 불러오기" });
  await dialog.getByRole("button", { name: "불러오기" }).click();

  // 새 시트 "policy" + 로드 블록(xl 참조) 생성 — dev 서버의 SheetJS 청크 첫 컴파일이 느릴 수 있다
  await expect
    .poll(
      () =>
        page.evaluate(() =>
          (window as any).__pygridStore.getState().workbook.sheets.map((s: any) => s.name),
        ),
      { timeout: 60_000 },
    )
    .toEqual(["Sheet1", "policy"]);
  const code = await page.evaluate(
    () => (window as any).__pygridStore.getState().workbook.pyBlocks[0]?.code ?? "",
  );
  expect(code).toContain('df = sheet("policy!A1:');
  expect(code).toContain("headers=True");

  // 상태 바의 데이터 파일 칩
  await expect(page.getByRole("button", { name: "policy.xlsx" })).toBeVisible();

  // 블록 실행 → df.head() spill (부트 포함이라 폴링 여유)
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(() => lastResult(page), { timeout: 240_000, intervals: [1000] })
    .toEqual({ status: "ok", summaryKo: null });
  expect(await spillCount(page)).toBeGreaterThan(6); // 헤더 1행 + head() 5행 × (index + 열들)
});

// pd.read_excel은 이 Pyodide 배포에 Excel 엔진(openpyxl)이 없어 아직 불가(런타임 소관, parity E7).
// FS 이중 착지 증명은 pd.read_csv로 한다 — '내 파일 추가' 경로 + pandas 모드.
test("내 파일(CSV) pandas 모드 → pd.read_csv가 워커 FS에서 읽혀 spill", async ({
  page,
}, testInfo) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  const csvPath = testInfo.outputPath("mini.csv");
  fs.writeFileSync(csvPath, "제품,수량\n노트,3\n펜,5", "utf-8");
  await page.locator('input[aria-label="데이터 파일 추가"]').setInputFiles(csvPath);

  const dialog = page.getByRole("dialog", { name: "데이터 불러오기" });
  await expect(dialog).toBeVisible();
  await dialog.getByLabel("pandas 코드 (pd.read_*)").check();
  await dialog.getByRole("button", { name: "불러오기" }).click();

  await expect
    .poll(
      () =>
        page.evaluate(
          () => (window as any).__pygridStore.getState().workbook.pyBlocks[0]?.code ?? "",
        ),
      { timeout: 60_000 },
    )
    .toContain('pd.read_csv("mini.csv")');
  // 시트로도 착지했는지 (이중 착지의 나머지 반쪽)
  expect(
    await page.evaluate(() =>
      (window as any).__pygridStore.getState().workbook.sheets.map((s: any) => s.name),
    ),
  ).toEqual(["Sheet1", "mini"]);

  // 실행 — 코드는 시트가 아니라 워커 FS의 mini.csv를 읽는다 (FS 착지 증명)
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(() => lastResult(page), { timeout: 240_000, intervals: [1000] })
    .toEqual({ status: "ok", summaryKo: null });
  expect(await spillCount(page)).toBe(6); // (헤더+2행) × (제품·수량) — auto 인덱스 생략
});
