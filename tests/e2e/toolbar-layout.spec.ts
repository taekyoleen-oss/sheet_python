import { expect, test, type Page } from "@playwright/test";

// 툴바 3단 배치 —
//  1행 헤더: 로고·제목 | 뷰 전환(뷰포트 정중앙) | 런타임 상태
//  2행 셸 바: [스프레드시트 감추기][파일][샘플][최근][?] … [(▶ 전체 실행)][Python 감추기]
//  3행(그리드 패널 안): 시트 편집 툴바 — 스프레드시트를 접으면 함께 사라진다

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as unknown as { __pygridStore?: unknown }).__pygridStore !==
        "undefined" &&
      (window as unknown as { __pygridReady?: boolean }).__pygridReady === true,
  );
  await page.waitForSelector('[data-testid="data-grid-canvas"]');
}

const boxOf = async (page: Page, name: string) => {
  const box = await page.getByRole("button", { name, exact: true }).boundingBox();
  expect(box, `${name} 버튼이 보이지 않는다`).not.toBeNull();
  return box!;
};

test("셸 바: 감추기 → 파일 → 샘플 → 최근 → ? 순서, Python 감추기·속성 창은 우측 끝", async ({
  page,
}) => {
  await page.goto("/");
  await waitForApp(page);

  const shell = page.getByTestId("shell-bar");
  const bar = (await shell.boundingBox())!;

  const gridToggle = await boxOf(page, "스프레드시트 감추기 (Ctrl+Alt+1)");
  const file = await boxOf(page, "파일");
  const sample = await boxOf(page, "샘플 워크북");
  const recent = await boxOf(page, "최근 워크북");
  const help = await boxOf(page, "키보드 단축키 안내");
  const pyToggle = await boxOf(page, "Python 패널 감추기 (Ctrl+Alt+2)");

  // 1) 왼쪽부터 감추기 → 파일 → 샘플 → 최근 → ?
  const xs = [gridToggle, file, sample, recent, help].map((b) => b.x);
  expect(xs).toEqual([...xs].sort((a, b) => a - b));
  expect(gridToggle.x - bar.x).toBeLessThan(48); // 좌측 패딩 + 한 칸 이내

  // 2) Python 패널 감추기 = 우측, 그 오른쪽 끝은 속성 창 토글(부록 P — 창이 오른쪽에 열린다)
  const propsToggle = await boxOf(page, "속성 창 열기 (변수·파일)");
  expect(pyToggle.x).toBeGreaterThan(bar.x + bar.width * 0.7);
  expect(propsToggle.x).toBeGreaterThan(pyToggle.x);
  expect(bar.x + bar.width - (propsToggle.x + propsToggle.width)).toBeLessThan(24);
  // 3) 같은 셸 바 행에서 서로 마주본다
  expect(Math.abs(gridToggle.y - pyToggle.y)).toBeLessThan(2);

  // 셸 바에는 Python 조작도, 시트 편집도 남아 있지 않다
  await expect(shell.getByRole("button", { name: "Python 블록 추가 (Ctrl+Shift+P)" })).toHaveCount(0);
  await expect(shell.getByRole("button", { name: "전체 실행", exact: true })).toHaveCount(0);
  await expect(shell.getByRole("button", { name: "목차 패널 열기" })).toHaveCount(0);
  await expect(shell.getByRole("button", { name: "붙여넣기 옵션 (텍스트로 붙여넣기)" })).toHaveCount(0);
  await expect(shell.getByRole("button", { name: "굵게", exact: true })).toHaveCount(0);
});

test("헤더: 파일 계열은 없고 뷰 전환이 뷰포트 정중앙", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);

  const header = page.locator("header");
  // 파일 계열 4개는 헤더에서 빠졌다
  for (const name of ["파일", "샘플 워크북", "최근 워크북", "키보드 단축키 안내"]) {
    await expect(header.getByRole("button", { name, exact: true })).toHaveCount(0);
  }
  // 남는 것: 로고 · 제목 · 뷰 전환 · 런타임 상태
  await expect(header).toContainText("시트기반 파이썬");
  await expect(header.getByTitle("클릭하여 제목 수정")).toBeVisible();

  const headerBox = (await header.boundingBox())!;
  const sw = page.getByTestId("view-switch");
  const box = (await sw.boundingBox())!;

  // 헤더 높이 안에 들어 있다
  expect(box.y).toBeGreaterThanOrEqual(headerBox.y - 1);
  expect(box.y + box.height).toBeLessThanOrEqual(headerBox.y + headerBox.height + 1);

  // 뷰포트 기준 가로 중앙 ±10%
  const vw = page.viewportSize()!.width;
  expect(Math.abs(box.x + box.width / 2 - vw / 2)).toBeLessThan(vw * 0.1);
});

test("시트 편집 툴바는 그리드 패널 안 — 스프레드시트를 접으면 함께 사라진다", async ({
  page,
}) => {
  await page.goto("/");
  await waitForApp(page);

  const edit = page.getByTestId("sheet-edit-toolbar");
  await expect(edit).toBeVisible();
  // 그리드 패널(#grid)의 자손이고, 그리드 캔버스 바로 위에 있다
  await expect(page.locator("#grid").getByTestId("sheet-edit-toolbar")).toHaveCount(1);
  const bar = (await edit.boundingBox())!;
  const canvas = (await page.getByTestId("data-grid-canvas").boundingBox())!;
  expect(bar.y + bar.height).toBeLessThanOrEqual(canvas.y + 1);

  // 시트 편집 버튼들이 여기에 모여 있다
  for (const name of [
    /붙여넣기 옵션/,
    /^행 삽입/,
    /^열 삭제/,
    /열까지 고정|열 고정 해제/,
    /오름차순 정렬/,
    /^굵게/,
  ]) {
    await expect(edit.getByRole("button", { name })).toHaveCount(1);
  }
  await expect(edit.getByRole("combobox", { name: "글자 크기" })).toBeVisible();

  // 접으면 툴바도 같이 사라진다 → 스트립으로 펼치면 복귀
  await page.getByRole("button", { name: "스프레드시트 감추기 (Ctrl+Alt+1)" }).click();
  await expect(page.getByTestId("sheet-edit-toolbar")).toHaveCount(0);
  await page.getByTestId("strip-grid").click();
  await expect(page.getByTestId("sheet-edit-toolbar")).toBeVisible();
});

test("Python 패널 헤더가 Python 조작을 모두 갖는다 (블록 N 오른쪽)", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);

  const header = page.getByTestId("python-panel-header");
  await expect(header).toBeVisible();
  await expect(header).toContainText("블록");

  for (const name of [
    "Python 블록 추가 (Ctrl+Shift+P)",
    "마크다운 블록 추가",
    "전체 실행",
    "실행 중단",
    "목차 패널 열기",
    "AI 채팅 패널 열기",
  ]) {
    await expect(header.getByRole("button", { name, exact: true })).toBeVisible();
  }
  await expect(header.getByRole("combobox", { name: "계산 모드" })).toBeVisible();

  // "블록 N" 라벨보다 오른쪽에 놓인다
  const label = (await header.getByText("블록", { exact: true }).boundingBox())!;
  const add = (await header
    .getByRole("button", { name: "Python 블록 추가 (Ctrl+Shift+P)", exact: true })
    .boundingBox())!;
  expect(add.x).toBeGreaterThan(label.x);

  // 좁은 패널에서도 넘치지 않는다 — flex-wrap으로 흘러서 가로 스크롤이 생기지 않는다
  const overflow = await header.evaluate((el) => el.scrollWidth - el.clientWidth);
  expect(overflow).toBeLessThanOrEqual(1);
});

test("Python 패널을 접으면 ▶ 전체 실행이 셸 바로 돌아온다", async ({ page }) => {
  await page.goto("/");
  await waitForApp(page);

  const shell = page.getByTestId("shell-bar");
  await expect(shell.getByRole("button", { name: "전체 실행", exact: true })).toHaveCount(0);

  await page.getByRole("button", { name: "Python 패널 감추기 (Ctrl+Alt+2)" }).click();
  await expect(page.getByTestId("python-panel-header")).toHaveCount(0);
  // 접힌 상태에서도 재계산 가능 — 중복 없이 셸 바에 하나만 존재한다
  await expect(page.getByRole("button", { name: "전체 실행", exact: true })).toHaveCount(1);
  await expect(shell.getByRole("button", { name: "전체 실행", exact: true })).toBeVisible();

  // 스트립으로 복원하면 다시 패널 헤더로 돌아간다
  await page.getByTestId("strip-python").click();
  await expect(shell.getByRole("button", { name: "전체 실행", exact: true })).toHaveCount(0);
  await expect(
    page.getByTestId("python-panel-header").getByRole("button", { name: "전체 실행", exact: true }),
  ).toBeVisible();
});

test("800×600 탭 구간: 셸 바 파일 메뉴는 그대로, 시트 편집 툴바는 그리드 탭에 있다", async ({
  page,
}) => {
  await page.setViewportSize({ width: 800, height: 600 });
  await page.goto("/");
  await waitForApp(page);

  const shell = page.getByTestId("shell-bar");
  await expect(shell.getByRole("button", { name: "파일", exact: true })).toBeVisible();
  await expect(shell.getByRole("button", { name: "샘플 워크북", exact: true })).toBeVisible();
  // 접기 버튼은 탭 구간에서 숨는다
  await expect(
    shell.getByRole("button", { name: "스프레드시트 감추기 (Ctrl+Alt+1)" }),
  ).toBeHidden();

  // 그리드 탭에서 보이고, Python 탭으로 가면 사라진다
  await expect(page.getByTestId("sheet-edit-toolbar")).toBeVisible();
  const tabs = page.getByRole("tablist", { name: "화면 전환" });
  await tabs.getByRole("tab", { name: "Python" }).click();
  await expect(page.getByTestId("sheet-edit-toolbar")).toHaveCount(0);

  // 좁은 폭에서도 셸 바·헤더가 가로로 넘치지 않는다
  for (const id of ["shell-bar"]) {
    const overflow = await page.getByTestId(id).evaluate((el) => el.scrollWidth - el.clientWidth);
    expect(overflow, `${id} 가로 넘침`).toBeLessThanOrEqual(1);
  }
  const headerOverflow = await page
    .locator("header")
    .evaluate((el) => el.scrollWidth - el.clientWidth);
  expect(headerOverflow).toBeLessThanOrEqual(1);
});
