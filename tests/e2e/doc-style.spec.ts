import { expect, test, type Page } from "@playwright/test";


// 새 블록 기본은 '시트로 보내지 않음' — 이 스펙은 spill을 검증하므로 켜고 시작한다
test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    (window as unknown as { __pygridToSheetDefault?: boolean }).__pygridToSheetDefault = true;
  });
});

// 부록 J: 마크다운 서식 툴바·이미지 내장 / 셀 서식(굵게·크기) / 실행 참조 표시

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

/** 마크다운 블록 생성 → 편집 textarea 로케이터 */
async function newMarkdownBlock(page: Page) {
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    st.addPyBlock(st.workbook.sheets[0].id, { r: 0, c: 5 }, "markdown");
  });
  const ta = page.getByRole("textbox", { name: "마크다운" });
  await expect(ta).toBeVisible();
  return ta;
}

const blockMd = (page: Page) =>
  page.evaluate(
    () => (window as any).__pygridStore.getState().workbook.pyBlocks[0]?.markdown ?? null,
  );

/** 그리드 클릭으로 textarea blur → 미리보기 전환 (.dvn-scroller가 locator 클릭을 가로채므로 좌표 클릭) */
async function blurToPreview(page: Page) {
  const box = (await page.getByTestId("data-grid-canvas").first().boundingBox())!;
  await page.mouse.click(box.x + 300, box.y + 200);
}

test("마크다운 툴바: 제목·하위 제목·굵게·목록 → 미리보기 렌더 + 목차 갱신", async ({
  page,
}) => {
  await page.goto("/");
  await waitForApp(page);
  const ta = await newMarkdownBlock(page);

  const bar = page.getByTestId("md-toolbar"); // 그리드 툴바의 "굵게"와 구분
  await ta.click();
  await page.keyboard.type("개요");
  await bar.getByRole("button", { name: "제목2" }).click();
  await expect.poll(() => blockMd(page)).toBe("## 개요");

  await bar.getByRole("button", { name: "하위 제목" }).click();
  await expect.poll(() => blockMd(page)).toBe("## 개요\n### ");
  await page.keyboard.type("세부");
  await expect.poll(() => blockMd(page)).toBe("## 개요\n### 세부");

  // 새 줄에 본문 입력 → 선택 → 굵게
  await page.keyboard.type("\n본문");
  await ta.evaluate((el: HTMLTextAreaElement) => {
    el.setSelectionRange(el.value.length - 2, el.value.length);
  });
  await bar.getByRole("button", { name: "굵게" }).click();
  await expect.poll(() => blockMd(page)).toBe("## 개요\n### 세부\n**본문**");

  // 기호 목록 — 커서 줄에 접두어
  await bar.getByRole("button", { name: "기호 목록" }).click();
  await expect.poll(() => blockMd(page)).toBe("## 개요\n### 세부\n- **본문**");

  // 미리보기 (blur) → h2·h3·strong 렌더
  await blurToPreview(page);
  const preview = page.getByTestId("markdown-preview");
  await expect(preview.locator("h2")).toHaveText("개요");
  await expect(preview.locator("h3")).toHaveText("세부");
  await expect(preview.locator("strong")).toHaveText("본문");

  // 목차 서브 자동 갱신
  await page.getByRole("button", { name: "목차 패널 열기" }).click();
  const toc = page.getByTestId("toc-panel");
  await expect(toc.getByRole("button", { name: "개요", exact: true })).toBeVisible();
  await expect(toc.getByRole("button", { name: "세부", exact: true })).toBeVisible();
});

test("이미지 드래그 앤 드롭 → data URI 삽입 → <img> 렌더", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);
  const ta = await newMarkdownBlock(page);
  await ta.click();

  // 1×1 PNG를 DataTransfer로 합성해 drop
  await ta.evaluate((el) => {
    const b64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
    const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    const dt = new DataTransfer();
    dt.items.add(new File([bytes], "tiny.png", { type: "image/png" }));
    el.dispatchEvent(
      new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: dt }),
    );
  });
  await expect
    .poll(() => blockMd(page), { timeout: 5_000 })
    .toContain("![tiny.png](data:image/png;base64,");

  // 미리보기에서 이미지 렌더 (안전한 data:image src)
  await blurToPreview(page);
  const img = page.getByTestId("markdown-preview").locator("img");
  await expect(img).toHaveAttribute("src", /^data:image\/png;base64,/);
  await expect(img).toHaveAttribute("alt", "tiny.png");
});

test("셀 서식: 굵게+크기 적용 → 저장/열기 왕복 보존 → 굵게 토글 해제", async ({
  page,
}, testInfo) => {
  await page.goto("/");
  await waitForApp(page);
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    st.setTitle("서식왕복");
    st.setCellValue(sid, 0, 0, { v: "머리", t: "s" });
    st.setSelection({ r0: 0, c0: 0, r1: 1, c1: 1 });
  });

  await page.getByRole("button", { name: "굵게", exact: true }).click();
  const cellSt = (key: string) =>
    page.evaluate(
      (k) =>
        (window as any).__pygridStore.getState().workbook.sheets[0].cells[k]?.st ?? null,
      key,
    );
  await expect.poll(() => cellSt("0:0")).toEqual({ b: true });
  await expect.poll(() => cellSt("1:1")).toEqual({ b: true }); // 빈 셀에도 적용

  await page.getByLabel("글자 크기").click();
  await page.getByRole("option", { name: "16px" }).click();
  await expect.poll(() => cellSt("0:0")).toEqual({ b: true, fs: 16 });

  // 저장 → 새 워크북 → 열기: st 보존
  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "파일", exact: true }).click();
  await page.getByRole("menuitem", { name: "저장 (.pygrid.json)" }).click();
  const savedPath = testInfo.outputPath("style.pygrid.json");
  await (await downloadPromise).saveAs(savedPath);
  await page.evaluate(() => (window as any).__pygridStore.getState().newWorkbook());
  await page.locator('input[aria-label="워크북 파일 열기"]').setInputFiles(savedPath);
  await expect.poll(() => cellSt("0:0"), { timeout: 10_000 }).toEqual({ b: true, fs: 16 });

  // 선택 전체가 굵으므로 토글은 해제로 동작
  await page.evaluate(() =>
    (window as any).__pygridStore.getState().setSelection({ r0: 0, c0: 0, r1: 0, c1: 0 }),
  );
  await page.getByRole("button", { name: "굵게 해제", exact: true }).click();
  await expect.poll(() => cellSt("0:0")).toEqual({ fs: 16 });
});

test("실행 참조 표시: sheet(\"A1:B3\") 실행 → 범위 기록·tint → 토글 끔 (실런타임)", async ({
  page,
}) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  const blockId = await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    for (let r = 0; r < 3; r++)
      for (let c = 0; c < 2; c++) st.setCellValue(sid, r, c, { v: r + c, t: "n" });
    const id = st.addPyBlock(sid, { r: 0, c: 4 });
    st.setBlockCode(id, 'sheet("A1:B3").sum().sum()');
    return id;
  });

  await page.getByRole("button", { name: "전체 실행", exact: true }).click();
  await expect
    .poll(
      () =>
        page.evaluate(
          () => (window as any).__pygridStore.getState().workbook.sheets[0].cells["0:4"]?.v ?? null,
        ),
      { timeout: 150_000, intervals: [1000] },
    )
    .toBe(9); // 0+1+1+2+2+3

  // 성공 실행이 읽은 참조가 기록된다 (A1:B3)
  await expect
    .poll(
      () =>
        page.evaluate(
          (id) => (window as any).__pygridStore.getState().executedRefs[id] ?? null,
          blockId,
        ),
      { timeout: 10_000 },
    )
    .toEqual([expect.objectContaining({ r0: 0, c0: 0, r1: 2, c1: 1 })]);

  await page.screenshot({ path: "output/j3-refs.png" }); // teal tint 수동 확인용

  // 상태 바 토글 끔 → showRefs=false (drawCell tint 소멸 경로)
  await page.getByRole("button", { name: "참조 표시 켬" }).click();
  await expect
    .poll(() => page.evaluate(() => (window as any).__pygridStore.getState().showRefs))
    .toBe(false);
  await expect(page.getByRole("button", { name: "참조 표시 끔" })).toBeVisible();
});
