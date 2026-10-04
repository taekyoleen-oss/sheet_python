import { expect, test, type Page } from "@playwright/test";


// 새 블록 기본은 '시트로 보내지 않음' — 이 스펙은 spill을 검증하므로 켜고 시작한다
test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    (window as unknown as { __pygridToSheetDefault?: boolean }).__pygridToSheetDefault = true;
  });
});

// 부록 P.6: 기본 분할이 시트 40%라, 셀 좌표로 누르는 이 스펙은 넓은 창에서 돌린다
test.use({ viewport: { width: 1920, height: 1080 } });

// v1.1 노트북 기능: 출력 위치 지정(앵커 재지정) · 출력 선택(변수·열·행) ·
// 마크다운 블록 · 목차 · 블록 접기 · Colab 스타일 셀 툴바(↑↓ 자리 교환).
// 실제 런타임을 쓰므로 한 테스트로 묶어 부트를 1회만 치른다.

/* eslint-disable @typescript-eslint/no-explicit-any */

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as unknown as { __pygridStore?: unknown }).__pygridStore !==
        "undefined" &&
      (window as unknown as { __pygridReady?: boolean }).__pygridReady === true,
  );
  await page.waitForSelector('[data-testid="data-grid-canvas"]');
  await page.evaluate(() => {
    (window as any).__pygridStore.getState().newWorkbook();
    (window as any).__pygridStore.getState().setGridCompact(false); // 부록 P.6: 셀 좌표 보정(88×34)은 기본 크기 보기 기준
  });
  await page.waitForTimeout(300); // 크기 전환 뒤 격자 다시 그리기를 기다린 다음 좌표를 잰다
}

const cellAt = (page: Page, key: string) =>
  page.evaluate(
    (k) => (window as any).__pygridStore.getState().workbook.sheets[0]?.cells[k] ?? null,
    key,
  );

const selection = (page: Page) =>
  page.evaluate(() => (window as any).__pygridStore.getState().selection);

const block0 = (page: Page) =>
  page.evaluate(() => (window as any).__pygridStore.getState().workbook.pyBlocks[0] ?? null);

/** 첫 블록의 spill 셀 수 (src = "<blockId>:<outputId>") */
const srcCount = (page: Page) =>
  page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const id = st.workbook.pyBlocks[0]?.id;
    if (!id) return 0;
    return Object.values(st.workbook.sheets[0].cells).filter(
      (c: any) => typeof c.src === "string" && c.src.split(":")[0] === id,
    ).length;
  });

/** 특정 출력이 쓴 셀 수 */
const outputCellCount = (page: Page, index: number) =>
  page.evaluate((i) => {
    const st = (window as any).__pygridStore.getState();
    const block = st.workbook.pyBlocks[0];
    const out = block?.outputs?.[i];
    if (!out) return -1;
    const tag = `${block.id}:${out.id}`;
    return Object.values(st.workbook.sheets[0].cells).filter((c: any) => c.src === tag).length;
  }, index);

/**
 * 캔버스 그리드 좌표 보정 — 헤더 높이·행 마커 폭을 가정하지 않는다.
 * 아무 셀이나 한 번 클릭해 (행 34px·열 88px 균일) 원점을 역산한다.
 */
async function calibrate(page: Page) {
  const box = (await page.locator('[data-testid="data-grid-canvas"]').boundingBox())!;
  const px = box.x + 200;
  const py = box.y + 100;
  await page.mouse.click(px, py);
  const sel = await selection(page);
  if (!sel) throw new Error("보정 실패: 그리드 선택을 읽지 못했습니다");
  return { x0: px - sel.c0 * 88, y0: py - sel.r0 * 34 };
}

const clickCell = (page: Page, cal: { x0: number; y0: number }, r: number, c: number) =>
  page.mouse.click(cal.x0 + c * 88, cal.y0 + r * 34);

test("노트북 셀: 앵커 재지정 → 출력 선택 → 마크다운·목차 → 접기 → 툴바 ↑ 자리 교환", async ({ page }) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  // 데이터 A1:B6 — 헤더 + 5행
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    st.setCells(sid, [
      { r: 0, c: 0, cell: { v: "a", t: "s" } },
      { r: 0, c: 1, cell: { v: "b", t: "s" } },
      ...[1, 2, 3, 4, 5].flatMap((i) => [
        { r: i, c: 0, cell: { v: i, t: "n" } },
        { r: i, c: 1, cell: { v: i * 10, t: "n" } },
      ]),
    ]);
  });

  const cal = await calibrate(page);
  await page.evaluate(() => {
    (window as any).__pygridStore.getState().setSelection({ r0: 0, c0: 3, r1: 0, c1: 3 }); // D1
  });

  // ── 1) 코드 블록 생성·실행 → D1 spill (6행 × 2열)
  await page.keyboard.press("Control+Shift+P");
  const editor = page.getByLabel("Python 코드");
  await expect(editor).toBeFocused();
  await editor.fill('df = sheet("A1:B6", headers=True)\nwide = df.assign(c=df["a"] * 2)\ndf');
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(() => srcCount(page), { timeout: 150_000, intervals: [1000] })
    .toBe(12);
  expect((await cellAt(page, "0:3"))?.v).toBe("a"); // D1 헤더
  expect((await cellAt(page, "5:4"))?.v).toBe(50); // E6 마지막 값

  // ── 2) 출력 위치 지정(출력 행의 주소 버튼) → F8으로 앵커 이동
  const codeCard = page.locator('[data-block-kind="code"]');
  await codeCard.getByRole("button", { name: "출력 1 위치" }).click();
  await expect(page.getByText("결과를 놓을 셀을 클릭하세요")).toBeVisible();
  await clickCell(page, cal, 7, 5); // F8
  expect((await block0(page)).anchor).toEqual({ r: 7, c: 5 });
  expect(await cellAt(page, "0:3")).toBeNull(); // 옛 앵커 비워짐
  await expect
    .poll(async () => (await cellAt(page, "7:5"))?.v, { timeout: 90_000, intervals: [500] })
    .toBe("a"); // 자동 모드 재실행 → F8에 새 spill
  expect(await srcCount(page)).toBe(12);
  expect(await cellAt(page, "5:4")).toBeNull(); // 옛 spill 전부 제거

  // ── 3) 출력 변수 = wide (6행 × 3열)
  await page.getByLabel("출력 1 변수").click();
  await expect(page.getByRole("option", { name: "wide" })).toBeVisible({ timeout: 30_000 });
  await page.getByRole("option", { name: "wide" }).click();
  await expect
    .poll(() => srcCount(page), { timeout: 90_000, intervals: [500] })
    .toBe(18);
  expect((await cellAt(page, "7:7"))?.v).toBe("c"); // H8: assign으로 생긴 열
  expect((await block0(page)).output).toEqual({ variable: "wide" });

  // ── 4) 열 a만 + 상위 3행 → 4행 × 1열
  await page.getByRole("button", { name: "출력 1 열 선택" }).click();
  await page.getByRole("menuitemcheckbox", { name: "b", exact: true }).click();
  await page.getByRole("menuitemcheckbox", { name: "c", exact: true }).click();
  await page.keyboard.press("Escape");
  await page.getByLabel("출력 1 상위 N행").fill("3");
  await expect
    .poll(() => srcCount(page), { timeout: 90_000, intervals: [500] })
    .toBe(4);
  expect((await block0(page)).output).toEqual({
    variable: "wide",
    columns: ["a"],
    rowLimit: 3,
  });
  expect((await cellAt(page, "7:5"))?.v).toBe("a"); // 헤더
  expect((await cellAt(page, "10:5"))?.v).toBe(3); // 상위 3행의 마지막
  expect(await cellAt(page, "11:5")).toBeNull(); // 4번째 행은 없다
  expect(await cellAt(page, "7:6")).toBeNull(); // b 열은 없다

  // ── 5) 마크다운 블록 — 앵커 셀에 아무것도 쓰지 않는다
  await page.evaluate(() => {
    (window as any).__pygridStore.getState().setSelection({ r0: 0, c0: 9, r1: 0, c1: 9 }); // J1
  });
  await page.getByRole("button", { name: "마크다운 블록 추가" }).click();
  const md = page.getByLabel("마크다운", { exact: true });
  await md.fill(
    "# 분석 개요\n> 인용 메모\n\n## 데이터\n- 하나\n  - 안쪽\n\n### 세부\n본문 **굵게**",
  );
  await md.blur();
  await expect(page.getByRole("heading", { name: "분석 개요" })).toBeVisible();
  await expect(page.getByText("굵게")).toBeVisible();
  // 부록 D.3: 인용(blockquote) + 중첩 목록 렌더
  const preview = page.getByTestId("markdown-preview");
  await expect(preview.locator("blockquote")).toHaveText("인용 메모");
  await expect(preview.locator("ul li ul li")).toHaveText("안쪽");
  expect(await cellAt(page, "0:9")).toBeNull(); // 앵커 셀은 비어 있다
  expect(
    await page.evaluate(
      () => (window as any).__pygridStore.getState().workbook.pyBlocks.length,
    ),
  ).toBe(2);

  // ── 6) 목차 전용 패널 — 툴바 토글로 열고, 계층 들여쓰기·강조·클릭 이동, ✕로 닫기
  await page.getByRole("button", { name: "목차 패널 열기" }).click();
  const toc = page.getByTestId("toc-panel");
  await expect(toc).toBeVisible();
  // 항목 버튼은 exact — hover 액션이 "<이름> 실행"·"<이름> 메뉴"로 함께 잡히기 때문
  const tocEntry = toc.getByRole("button", { name: "분석 개요", exact: true });
  await expect(tocEntry).toBeVisible();
  await expect(toc.getByRole("button", { name: "F8", exact: true })).toBeVisible(); // 코드 블록 = 앵커 주소

  // 헤딩 계층 = 단계별 들여쓰기, 3단계부터 · 접두
  const padOf = (name: string) =>
    toc
      .getByRole("button", { name, exact: true })
      .evaluate((el) => getComputedStyle(el).paddingLeft);
  expect(await padOf("분석 개요")).toBe("6px"); // #
  expect(await padOf("데이터")).toBe("18px"); // ##
  expect(await padOf("세부")).toBe("30px"); // ###
  expect(await padOf("F8")).toBe("42px"); // 코드 블록은 직전 헤딩 아래 한 단계
  await expect(toc.getByRole("button", { name: "세부", exact: true })).toContainText("·");

  await tocEntry.click();
  await expect.poll(() => selection(page)).toMatchObject({ r0: 0, c0: 9 });
  await expect(page.getByRole("heading", { name: "분석 개요" })).toBeVisible();
  // 현재 항목만 primary 좌측 바로 강조
  await expect(toc.locator("li", { hasText: "분석 개요" }).first()).toHaveClass(
    /border-primary/,
  );
  await expect(toc.locator("li", { hasText: "데이터" }).first()).not.toHaveClass(
    /border-primary/,
  );

  await toc.getByRole("button", { name: "목차 패널 닫기" }).click();
  await expect(toc).toHaveCount(0);

  // ── 7) 블록 접기 — 헤더만 남고 편집기·출력 컨트롤·실행 레일은 숨는다
  await codeCard.getByRole("button", { name: "블록 접기" }).click();
  await expect(page.getByLabel("Python 코드")).toBeHidden();
  await expect(codeCard.getByLabel("출력 1 상위 N행")).toBeHidden();
  await expect(codeCard.getByRole("button", { name: "실행", exact: true })).toBeHidden();
  await expect(codeCard.getByRole("button", { name: "블록 펼치기" })).toBeVisible();
  await expect(codeCard.getByLabel("블록 제목")).toBeVisible(); // 헤더는 그대로
  expect((await block0(page)).collapsed).toBe(true);

  await codeCard.getByRole("button", { name: "블록 펼치기" }).click();
  await expect(page.getByLabel("Python 코드")).toBeVisible();

  // ── 8) 셀 툴바 ↑ — 계산 순서 이웃(마크다운 J1)과 자리 교환 + 자동 재실행
  const toolbar = codeCard.getByTestId("cell-toolbar");
  const opacity = () => toolbar.evaluate((el) => getComputedStyle(el).opacity);
  // 숨김 조건 = hover도 focus도 없을 때 (직전 클릭으로 카드 안 버튼이 포커스를 쥐고 있다)
  await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
  await page.getByRole("button", { name: "전체 실행" }).hover(); // 카드에서 마우스를 뗀다
  await expect.poll(opacity).toBe("0");
  await codeCard.hover();
  await expect.poll(opacity).toBe("1"); // hover 시 노출

  const mdBefore = await page.evaluate(
    () =>
      (window as any).__pygridStore
        .getState()
        .workbook.pyBlocks.find((b: any) => b.kind === "markdown").anchor,
  );
  expect(mdBefore).toEqual({ r: 0, c: 9 }); // J1 — 계산 순서상 코드 블록(F8)보다 앞
  await codeCard.getByRole("button", { name: "위로" }).click();

  expect((await block0(page)).anchor).toEqual({ r: 0, c: 9 }); // 코드 → J1
  expect(
    await page.evaluate(
      () =>
        (window as any).__pygridStore
          .getState()
          .workbook.pyBlocks.find((b: any) => b.kind === "markdown").anchor,
    ),
  ).toEqual({ r: 7, c: 5 }); // 마크다운 → F8
  expect(await cellAt(page, "7:5")).toBeNull(); // 옛 자리의 spill 제거
  await expect
    .poll(async () => (await cellAt(page, "0:9"))?.v, { timeout: 90_000, intervals: [500] })
    .toBe("a"); // 자동 모드: 새 앵커에서 재실행
  expect(await srcCount(page)).toBe(4);

  // 이제 코드 블록이 계산 순서상 첫 블록 → ↑ 비활성
  await codeCard.hover();
  await expect(codeCard.getByRole("button", { name: "위로" })).toBeDisabled();
  await expect(codeCard.getByRole("button", { name: "아래로" })).toBeEnabled();
});

test("다중 출력: 한 블록의 두 결과를 서로 다른 셀에 (부록 D.1)", async ({ page }) => {
  test.setTimeout(300_000);
  await page.goto("/");
  await waitForApp(page);

  // 데이터 A1:B6 — 헤더 + 5행 (a=1..5, b=10..50)
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    st.setCells(sid, [
      { r: 0, c: 0, cell: { v: "a", t: "s" } },
      { r: 0, c: 1, cell: { v: "b", t: "s" } },
      ...[1, 2, 3, 4, 5].flatMap((i) => [
        { r: i, c: 0, cell: { v: i, t: "n" } },
        { r: i, c: 1, cell: { v: i * 10, t: "n" } },
      ]),
    ]);
  });

  const cal = await calibrate(page);
  await page.evaluate(() => {
    (window as any).__pygridStore.getState().setSelection({ r0: 0, c0: 3, r1: 0, c1: 3 }); // D1
  });

  // 출력 #1 = 마지막 표현식(df) → D1
  await page.keyboard.press("Control+Shift+P");
  const editor = page.getByLabel("Python 코드");
  await editor.fill(
    'df = sheet("A1:B6", headers=True)\ntotal = int(df["a"].sum())\navg = float(df["b"].mean())\ndf',
  );
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(() => srcCount(page), { timeout: 150_000, intervals: [1000] })
    .toBe(12);
  expect((await cellAt(page, "0:3"))?.v).toBe("a"); // D1
  expect((await cellAt(page, "5:4"))?.v).toBe(50); // E6

  // 출력 #2 추가 → 기본은 블록 옆 빈 셀, 여기서는 G8으로 지정
  const codeCard = page.locator('[data-block-kind="code"]');
  await codeCard.getByRole("button", { name: "시트에 추가" }).click();
  expect(
    await page.evaluate(
      () => (window as any).__pygridStore.getState().workbook.pyBlocks[0].outputs.length,
    ),
  ).toBe(2);
  await codeCard.getByRole("button", { name: "출력 2 위치" }).click();
  await expect(page.getByText("결과를 놓을 셀을 클릭하세요")).toBeVisible();
  await clickCell(page, cal, 7, 6); // G8
  expect((await block0(page)).outputs[1].anchor).toEqual({ r: 7, c: 6 });

  // 출력 #2 변수 = total → 1×1 스칼라. 코드는 한 번만 실행되고 두 영역이 각각 채워진다
  await page.getByLabel("출력 2 변수").click();
  await expect(page.getByRole("option", { name: "total" })).toBeVisible({ timeout: 30_000 });
  await page.getByRole("option", { name: "total" }).click();
  await expect
    .poll(async () => (await cellAt(page, "7:6"))?.v, { timeout: 90_000, intervals: [500] })
    .toBe(15); // 1+2+3+4+5
  expect(await outputCellCount(page, 0)).toBe(12); // 출력 #1 = D1:E6 그대로
  expect(await outputCellCount(page, 1)).toBe(1); // 출력 #2 = G8 한 칸
  expect((await cellAt(page, "0:3"))?.v).toBe("a");
  expect((await cellAt(page, "5:4"))?.v).toBe(50);

  // 출력별 src 태그가 다르다
  const tags = await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const b = st.workbook.pyBlocks[0];
    return {
      out1: st.workbook.sheets[0].cells["0:3"].src,
      out2: st.workbook.sheets[0].cells["7:6"].src,
      ids: b.outputs.map((o: any) => `${b.id}:${o.id}`),
    };
  });
  expect(tags.out1).toBe(tags.ids[0]);
  expect(tags.out2).toBe(tags.ids[1]);

  // 출력 #2의 변수만 바꾸면 그 영역만 바뀐다
  await page.getByLabel("출력 2 변수").click();
  await page.getByRole("option", { name: "avg" }).click();
  await expect
    .poll(async () => (await cellAt(page, "7:6"))?.v, { timeout: 90_000, intervals: [500] })
    .toBe(30); // (10+20+30+40+50)/5
  expect((await cellAt(page, "0:3"))?.v).toBe("a"); // 출력 #1 영역은 그대로
  expect((await cellAt(page, "5:4"))?.v).toBe(50);
  expect(await outputCellCount(page, 0)).toBe(12);
  expect(await outputCellCount(page, 1)).toBe(1);

  // 진단 탭: 출력마다 상태·오류가 따로 보인다. avg를 지워 출력 #2만 실패시킨다
  await editor.fill(
    'df = sheet("A1:B6", headers=True)\ntotal = int(df["a"].sum())\nif "avg" in globals():\n    del avg\ndf',
  );
  await page.getByRole("button", { name: "실행", exact: true }).click();
  await expect
    .poll(async () => (await cellAt(page, "7:6"))?.v, { timeout: 90_000, intervals: [500] })
    .toBe("#PYTHON!");
  expect((await cellAt(page, "0:3"))?.v).toBe("a"); // 출력 #1은 정상 유지

  await page.getByRole("button", { name: "속성 창 열기 (변수·파일)" }).click(); // 부록 P.7: 진단·미리보기는 속성 창 탭
  const diagTab = page.getByRole("tab", { name: /진단/ });
  await diagTab.click();
  const diag = page.getByTestId("diagnostics-tab");
  await expect(diag.getByText("출력 1", { exact: true })).toBeVisible();
  await expect(diag.getByText("avg", { exact: true })).toBeVisible();
  // 한국어 요약(<p>)과 traceback(<pre>)이 함께 보인다 — 요약만 겨냥
  await expect(diag.getByText(/정의되지 않은 이름을 사용했습니다/)).toBeVisible();
  await expect(diag.getByText(/Traceback/)).toBeAttached();
  await expect(diagTab).toContainText("1"); // 오류 배지 = 실패한 출력 수

  // 출력 삭제 → 그 영역만 비워진다
  await codeCard.getByRole("button", { name: "출력 2 삭제" }).click();
  expect(await cellAt(page, "7:6")).toBeNull();
  expect(await outputCellCount(page, 0)).toBe(12);
  // 마지막 출력의 ×는 '시트에서 빼기' — 누르면 셀이 지워지고 Python 결과로만 본다
  await expect(codeCard.getByRole("button", { name: "출력 1 삭제" })).toHaveAttribute(
    "title",
    /시트에서 빼기/,
  );
});
