import { expect, test, type Page } from "@playwright/test";


// 새 블록 기본은 '시트로 보내지 않음' — 이 스펙은 spill을 검증하므로 켜고 시작한다
test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    (window as unknown as { __pygridToSheetDefault?: boolean }).__pygridToSheetDefault = true;
  });
});

// 부록 L: 워크북 에이전트 채팅 — 도구 호출(read_range) 로그·근거 하이라이트 → 셀 제안 [적용] → Ctrl+Z,
// 블록 ✦ "오류 원인 알려줘" → 채팅 자동 전송(traceback 첨부) → 코드 제안 [현재 블록에 적용].
// 실제 Anthropic API는 호출하지 않는다 (page.route 모킹).

/* eslint-disable @typescript-eslint/no-explicit-any */

type Body = { content: unknown[]; stop_reason: string };

const toolUse = (name: string, input: unknown, text = "확인해 볼게요."): Body => ({
  content: [
    { type: "text", text },
    { type: "tool_use", id: `tu-${name}`, name, input },
  ],
  stop_reason: "tool_use",
});

const finalText = (text: string): Body => ({
  content: [{ type: "text", text }],
  stop_reason: "end_turn",
});

async function waitForApp(page: Page) {
  await page.waitForFunction(
    () =>
      typeof (window as any).__pygridStore !== "undefined" &&
      (window as any).__pygridReady === true,
  );
  await page.waitForSelector('[data-testid="data-grid-canvas"]');
  await page.evaluate(() => (window as any).__pygridStore.getState().newWorkbook());
}

/** 모킹 라우트 + 응답 큐. 큐가 비면 "(끝)" 최종 응답 */
async function mockApi(page: Page) {
  const state = { queue: [] as Body[], bodies: [] as string[] };
  await page.route("https://api.anthropic.com/**", async (route) => {
    state.bodies.push(route.request().postData() ?? "");
    const body = state.queue.shift() ?? finalText("(끝)");
    await route.fulfill({ contentType: "application/json", body: JSON.stringify(body) });
  });
  return state;
}

/** 키 미설정 → 설정 다이얼로그 유도 → 저장 (앱의 유일한 키 진입점) */
async function saveKey(page: Page, panel: ReturnType<Page["getByTestId"]>) {
  await panel.getByLabel("AI 채팅 입력").fill("키 설정용");
  await panel.getByRole("button", { name: "전송" }).click();
  const keyDialog = page.getByRole("dialog", { name: /AI 설정/ });
  await expect(keyDialog).toBeVisible();
  await keyDialog.getByLabel("Anthropic API 키").fill("sk-ant-e2e-agent");
  await keyDialog.getByRole("button", { name: "저장", exact: true }).click();
  await expect(keyDialog).toBeHidden();
  await panel.getByLabel("AI 채팅 입력").fill("");
}

test("도구 호출 → 로그·근거 하이라이트 → 답변 → 셀 제안 [적용] → Ctrl+Z 복원", async ({
  page,
}) => {
  test.setTimeout(180_000);
  await page.goto("/");
  await waitForApp(page);
  const api = await mockApi(page);

  // 데이터 시드 (A1:B4)
  await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    st.setCells(sid, [
      { r: 0, c: 0, cell: { v: "월", t: "s" } },
      { r: 0, c: 1, cell: { v: "손해액", t: "s" } },
      { r: 1, c: 0, cell: { v: "1월", t: "s" } },
      { r: 1, c: 1, cell: { v: 10, t: "n" } },
      { r: 2, c: 0, cell: { v: "2월", t: "s" } },
      { r: 2, c: 1, cell: { v: 20, t: "n" } },
      { r: 3, c: 0, cell: { v: "3월", t: "s" } },
      { r: 3, c: 1, cell: { v: 30, t: "n" } },
    ]);
  });
  await page.waitForTimeout(400); // undo 이력 경계 (300ms 스로틀)

  await page.getByRole("button", { name: "AI 채팅 패널 열기" }).click();
  const panel = page.getByTestId("ai-chat-panel");
  await expect(panel).toBeVisible();
  await saveKey(page, panel);

  // ── ① 값이 필요한 질문 → read_range 도구 호출 → 최종 답변
  api.queue = [
    toolUse("read_range", { ref: "A1:B4" }),
    finalText("손해액 합계는 60입니다."),
  ];
  await panel.getByLabel("AI 채팅 입력").fill("A1:B4 요약해줘");
  await panel.getByRole("button", { name: "전송" }).click();

  const toolLog = panel.getByTestId("tool-log");
  await expect(toolLog).toBeVisible({ timeout: 20_000 });
  await expect(toolLog).toContainText("Sheet1!A1:B4 읽음 · 4행×2열");
  await expect(panel.getByText("손해액 합계는 60입니다.")).toBeVisible();

  // 근거 칩 + 그리드 하이라이트(J.3 계열 transient)
  const chip = panel.getByTestId("ref-chip");
  await expect(chip).toHaveText("Sheet1!A1:B4");
  expect(
    await page.evaluate(() => (window as any).__pygridStore.getState().chatRefs),
  ).toMatchObject([{ r0: 0, c0: 0, r1: 3, c1: 1 }]);
  await chip.click(); // 클릭 시 그 범위로 이동
  expect(
    await page.evaluate(() => (window as any).__pygridStore.getState().selection),
  ).toEqual({ r0: 0, c0: 0, r1: 3, c1: 1 });

  // 요청 검증: 도구 스키마 동봉 + 2번째 요청에 tool_result(실제 값) 포함
  expect(api.bodies).toHaveLength(2);
  const first = JSON.parse(api.bodies[0]);
  expect(first.tools.map((t: any) => t.name)).toEqual([
    "list_sheets",
    "read_range",
    "list_blocks",
    "propose_cells",
    "propose_block",
  ]);
  expect(api.bodies[0]).toContain("앱 규칙이 우선"); // 지침 레이어 + 우선순위
  expect(api.bodies[0]).toContain("보험·계리"); // 기본 지침 시드
  // 값·헤더는 도구 호출 전에 전송되지 않는다 — 첫 메시지는 개요(시트명·범위)뿐
  expect(first.messages[0].content).toContain("[워크북 개요]");
  expect(first.messages[0].content).toContain("Sheet1!A1:B4");
  expect(first.messages[0].content).not.toContain("손해액");
  const second = JSON.parse(api.bodies[1]);
  expect(second.messages).toHaveLength(3);
  expect(JSON.stringify(second.messages[2])).toContain("tool_result");
  expect(JSON.stringify(second.messages[2])).toContain("손해액");

  // ── ② 셀 제안 → 카드 → [적용] → 셀 반영 → Ctrl+Z 복원
  api.queue = [
    toolUse(
      "propose_cells",
      { ref: "D1", values: [["합계", 60]], reason: "손해액 합계를 옆 열에 적었습니다" },
      "합계를 제안합니다.",
    ),
    finalText("D1:E1에 합계를 제안했습니다. [적용]을 눌러 반영하세요."),
  ];
  await panel.getByLabel("AI 채팅 입력").fill("합계를 D1에 적어줘");
  await panel.getByRole("button", { name: "전송" }).click();

  const card = panel.getByTestId("cell-proposal");
  await expect(card).toBeVisible({ timeout: 20_000 });
  await expect(card).toContainText("Sheet1!D1:E1");
  await expect(card).toContainText("2셀");
  await expect(card).toContainText("손해액 합계를 옆 열에 적었습니다");
  // 아직 시트는 그대로 (자동 적용 없음)
  expect(
    await page.evaluate(() => (window as any).__pygridStore.getState().workbook.sheets[0].cells["0:3"]),
  ).toBeUndefined();

  await card.getByRole("button", { name: "적용", exact: true }).click();
  await expect(card).toContainText("적용됨");
  expect(
    await page.evaluate(() => {
      const cells = (window as any).__pygridStore.getState().workbook.sheets[0].cells;
      return [cells["0:3"]?.v, cells["0:4"]?.v];
    }),
  ).toEqual(["합계", 60]);

  // 한 트랜잭션 = 1 undo — 원본 데이터는 남는다
  const gridBox = (await page.getByTestId("data-grid-canvas").first().boundingBox())!;
  await page.mouse.click(gridBox.x + 300, gridBox.y + 250); // 포커스를 그리드로 (.dvn-scroller가 가로챈다)
  await page.keyboard.press("Control+z");
  await expect
    .poll(() =>
      page.evaluate(() => {
        const cells = (window as any).__pygridStore.getState().workbook.sheets[0].cells;
        return [cells["0:3"] ?? null, cells["0:1"]?.v ?? null];
      }),
    )
    .toEqual([null, "손해액"]);
});

test("블록 ✦ 오류 원인 알려줘 → 채팅 자동 전송(traceback 첨부) → 코드 제안 [현재 블록에 적용]", async ({
  page,
}) => {
  test.setTimeout(180_000);
  await page.goto("/");
  await waitForApp(page);
  const api = await mockApi(page);

  // 오류 상태 블록 준비 (실행 없이 결과만 주입)
  const blockId = await page.evaluate(() => {
    const st = (window as any).__pygridStore.getState();
    const sid = st.workbook.sheets[0].id;
    const id = st.addPyBlock(sid, { r: 0, c: 3 });
    st.setBlockCode(id, "df.sum()");
    st.setBlockTitle(id, "합계");
    st.applyBlockResult(id, [], {
      last: {
        status: "error",
        stdout: "",
        stderr: "",
        traceback: "Traceback (most recent call last):\nNameError: name 'df' is not defined",
        summaryKo: "정의되지 않은 이름 'df'",
        durationMs: 3,
        ranAt: new Date().toISOString(),
      },
    });
    return id as string;
  });

  // 키 설정 (채팅이 유일한 진입점)
  await page.getByRole("button", { name: "AI 채팅 패널 열기" }).click();
  const panel = page.getByTestId("ai-chat-panel");
  await saveKey(page, panel);

  // ── ✦ 메뉴 → "오류 원인 알려줘" → 채팅으로 자동 전송
  api.queue = [
    toolUse("list_blocks", {}, "블록을 확인할게요."),
    {
      content: [
        { type: "text", text: "df가 정의되지 않아서 납니다. 먼저 sheet()로 읽어야 합니다." },
        {
          type: "tool_use",
          id: "tu-fix",
          name: "propose_block",
          input: {
            code: 'df = sheet("A1:B4", headers=True)\ndf.sum()',
            title: "합계 (수정)",
            targetBlockId: blockId,
          },
        },
      ],
      stop_reason: "tool_use",
    },
    finalText("수정안을 제안했습니다 — 확인 후 실행하세요."),
  ];
  await page.getByRole("button", { name: "AI 메뉴" }).click();
  await page.getByRole("menuitem", { name: "오류 원인 알려줘" }).click();

  // 첨부(코드 + traceback)가 실린 사용자 메시지가 전송된다
  const userMsg = panel.locator('[data-role="user"]').last();
  await expect(userMsg).toContainText("NameError: name 'df' is not defined", { timeout: 20_000 });
  await expect(userMsg).toContainText("정의되지 않은 이름 'df'");
  await expect(userMsg).toContainText("df.sum()");
  expect(api.bodies[0]).toContain("NameError");
  expect(api.bodies[0]).toContain("초급 눈높이");

  const logs = panel.getByTestId("tool-log");
  await expect(logs).toHaveCount(2); // list_blocks + propose_block
  await expect(logs.first()).toContainText("블록 목록 읽음 · 1개");
  await expect(logs.last()).toContainText("코드 제안 · 합계 (수정)");
  await expect(panel.getByText("df가 정의되지 않아서 납니다.", { exact: false })).toBeVisible();

  // ── 코드 제안 카드 → [현재 블록에 적용] (교체, 자동 실행 없음)
  const card = panel.getByTestId("block-proposal");
  await expect(card).toBeVisible();
  await expect(card).toContainText('df = sheet("A1:B4", headers=True)');
  await expect(card).toContainText("합계 (수정)");
  await expect(card).toContainText("Sheet1!D1 · 합계"); // 대상 블록 앵커 칩
  await card.getByRole("button", { name: "현재 블록에 적용" }).click();
  await expect(card).toContainText("적용됨");

  const block = await page.evaluate(
    () => (window as any).__pygridStore.getState().workbook.pyBlocks[0],
  );
  expect(block.code).toBe('df = sheet("A1:B4", headers=True)\ndf.sum()');
  expect((await page.evaluate(
    () => (window as any).__pygridStore.getState().workbook.pyBlocks.length,
  ))).toBe(1); // 새 블록을 만들지 않고 대상 블록을 교체
});
