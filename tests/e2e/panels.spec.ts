import { writeFileSync } from "node:fs";
import { expect, test, type Page } from "@playwright/test";


// 새 블록 기본은 '시트로 보내지 않음' — 이 스펙은 spill을 검증하므로 켜고 시작한다
test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    (window as unknown as { __pygridToSheetDefault?: boolean }).__pygridToSheetDefault = true;
  });
});

// G6: matplotlib 한글 이미지 카드 + 미리보기 탭 렌더 + 값 모드 거부
// + 참조 삽입 바 · 진단 탭 스모크

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

const cellAt = (page: Page, key: string) =>
  page.evaluate(
    (k) =>
      (window as any).__pygridStore.getState().workbook.sheets[0]?.cells[k] ?? null,
    key,
  );

test("G6: matplotlib 한글 이미지 → 카드·미리보기·스크린샷 → 값 모드 거부", async ({ page }) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    const id = st.addPyBlock(sid, { r: 0, c: 0 });
    st.setBlockOutputMode(id, "object");
    st.setBlockCode(
      id,
      'import matplotlib.pyplot as plt\nfig, ax = plt.subplots()\nax.hist([1, 2, 2, 3, 3, 3], bins=3)\nax.set_title("한글 제목")\nfig',
    );
  });

  // 부트 + matplotlib 지연 로드 — 여유 폴링
  await page.getByRole("button", { name: "전체 실행", exact: true }).click();
  await expect
    .poll(async () => String((await cellAt(page, "0:0"))?.v ?? ""), {
      timeout: 240_000,
      intervals: [2000],
    })
    .toMatch(/^\[Figure/);

  // 첫 실행에서 바로 Pretendard가 적용된다 — bootstrap.py _pygrid_mpl_setup이
  // 다운로드된 matplotlib을 사용자 코드 실행 전에 import해 폰트를 선적용하기 때문.

  // 출력 미리보기 탭 → 이미지 렌더 (선택 블록 없으면 최근 실행 블록)
  await page.getByRole("button", { name: "속성 창 열기 (변수·파일)" }).click(); // 부록 P.7: 진단·미리보기는 속성 창 탭
  await page.getByRole("tab", { name: "출력 미리보기" }).click();
  const img = page.getByTestId("props-bottom").getByTestId("preview-image");
  await expect(img).toBeVisible({ timeout: 15_000 });
  await expect(img).toHaveAttribute("src", /^blob:/);

  // 한글 글리프 수동 확인용 — 렌더된 blob의 실제 PNG 바이트를 저장 (G6)
  // 이미지가 실제로 디코드·페인트된 뒤에만 진행 (빈 캡처 방지)
  await expect
    .poll(
      () => img.evaluate((el) => (el as HTMLImageElement).naturalWidth),
      { timeout: 15_000 },
    )
    .toBeGreaterThan(0);
  await img.evaluate((el) => (el as HTMLImageElement).decode());
  const pngB64 = await page.evaluate(async () => {
    const el = document.querySelector('[data-testid="props-bottom"] [data-testid="preview-image"]') as HTMLImageElement;
    const buf = await (await fetch(el.src)).arrayBuffer();
    const bytes = new Uint8Array(buf);
    let s = "";
    for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return btoa(s);
  });
  const png = Buffer.from(pngB64, "base64");
  expect(png.byteLength).toBeGreaterThan(5000); // 실제 차트 PNG인지
  writeFileSync("output/g6-korean-glyph.png", png);

  // 속성 창(부록 P) 변수 목록 — fig/ax 변수 표시 (실행 완료 시 자동 갱신). 창은 위에서 이미 열었다
  await page.getByRole("tab", { name: "변수" }).click();
  await expect(page.getByTestId("props-variables").locator('[data-var="fig"]')).toBeVisible({
    timeout: 15_000,
  });

  // 값 모드 전환 후 재실행 → 이미지는 값으로 펼칠 수 없음
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setBlockOutputMode(st.workbook.pyBlocks[0].id, "values");
  });
  await page.getByRole("button", { name: "전체 실행", exact: true }).click();
  await expect
    .poll(async () => (await cellAt(page, "0:0"))?.v, { timeout: 60_000, intervals: [500] })
    .toBe("#PYTHON!");
  const summary = await page.evaluate(
    () => (window as any).__pygridStore.getState().workbook.pyBlocks[0].last?.summaryKo,
  );
  expect(summary).toContain("이미지는 값으로 펼칠 수 없습니다");
});

test("참조 삽입 바 + 진단 탭 스모크", async ({ page }) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  // 데이터 + 블록 (에디터 자동 포커스 → lastEditorBlockId 설정)
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    st.setCells(sid, [
      { r: 0, c: 0, cell: { v: "x", t: "s" } },
      { r: 0, c: 1, cell: { v: "y", t: "s" } },
      { r: 1, c: 0, cell: { v: 1, t: "n" } },
      { r: 1, c: 1, cell: { v: 10, t: "n" } },
      { r: 2, c: 0, cell: { v: 2, t: "n" } },
      { r: 2, c: 1, cell: { v: 20, t: "n" } },
    ]);
    st.addPyBlock(sid, { r: 0, c: 3 });
  });
  await page.waitForSelector(".cm-content");
  await page.locator(".cm-content").click(); // 편집기 포커스 → lastEditorBlockId

  // 그리드 범위 선택 → 삽입 바 (헤더 휴리스틱: 첫 행 문자열 + 숫자 본문 → headers=True)
  await page.evaluate(() => {
    (window as any).__pygridStore.getState().setSelection({ r0: 0, c0: 0, r1: 2, c1: 1 });
  });
  const insertBar = page.getByRole("button", { name: /^sheet\(.*삽입$/ });
  await expect(insertBar).toBeVisible();
  await expect(insertBar).toContainText('sheet("A1:B3", headers=True)');
  await insertBar.click();

  // 500ms 디바운스 커밋 후 스토어 코드에 반영
  await expect
    .poll(
      () =>
        page.evaluate(
          () => (window as any).__pygridStore.getState().workbook.pyBlocks[0].code,
        ),
      { timeout: 5_000 },
    )
    .toContain('sheet("A1:B3", headers=True)');

  // 코드 수정만으로는 실행되지 않는다 — dirty 표시만, 실행은 ▶
  await page.waitForTimeout(1500);
  expect(
    await page.evaluate(() => (window as any).__pygridStore.getState().workbook.pyBlocks[0].last ?? null),
  ).toBeNull();
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(
      () =>
        page.evaluate(
          () => (window as any).__pygridStore.getState().workbook.pyBlocks[0].last?.status ?? null,
        ),
      { timeout: 150_000, intervals: [1000] },
    )
    .toBe("ok");

  // 오류 실행 → 진단 탭: NameError 표시 + 오류 수 배지
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.setBlockCode(st.workbook.pyBlocks[0].id, "undefined_name");
  });
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(async () => (await cellAt(page, "0:3"))?.v, { timeout: 150_000, intervals: [1000] })
    .toBe("#PYTHON!");

  await page.getByRole("button", { name: "속성 창 열기 (변수·파일)" }).click(); // 부록 P.7: 진단·미리보기는 속성 창 탭
  await page.getByRole("tab", { name: /진단/ }).click();
  await expect(page.getByText(/NameError/).first()).toBeVisible();
  await expect(page.getByRole("tab", { name: /진단/ })).toContainText("1");
});
