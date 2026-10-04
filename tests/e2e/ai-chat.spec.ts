import { expect, test, type Page } from "@playwright/test";

// 부록 G.2·G.3 (+L 통합 후): AI 채팅 패널 — 토글·키 유도·모킹 응답의 코드 카드 → 블록 반영 ·
// 지침 펜스 → 확인 카드 [반영] → 지침 편집에 반영 · 대화 유지(닫았다 열기).
// 실제 Anthropic API는 호출하지 않는다 (page.route 모킹).

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
  });
}

test("AI 채팅: 토글 → 키 유도 → 코드 카드 → 블록 반영 → 지침 펜스 [반영] → 대화 유지", async ({
  page,
}) => {
  test.setTimeout(180_000);
  await page.goto("/");
  await waitForApp(page);

  // ── 툴바 토글로 패널 열기
  await page.getByRole("button", { name: "AI 채팅 패널 열기" }).click();
  const panel = page.getByTestId("ai-chat-panel");
  await expect(panel).toBeVisible();

  // ── 키 미설정: 전송 → AI 설정 다이얼로그 유도 (메시지는 전송되지 않는다)
  const chatInput = panel.getByLabel("AI 채팅 입력");
  await chatInput.fill("손해율 합계 코드 만들어줘");
  await panel.getByRole("button", { name: "전송" }).click();
  const keyDialog = page.getByRole("dialog", { name: /AI 설정/ });
  await expect(keyDialog).toBeVisible();
  await keyDialog.getByLabel("Anthropic API 키").fill("sk-ant-e2e-chat");
  await keyDialog.getByRole("button", { name: "저장", exact: true }).click();
  await expect(keyDialog).toBeHidden();

  // ── 모킹: 텍스트 + python 코드 블록 응답
  let sentBody = "";
  let reply =
    "합계 코드입니다.\n```python\ntotal = sheet(\"A1:A3\").sum()\ntotal\n```";
  await page.route("https://api.anthropic.com/**", async (route) => {
    sentBody = route.request().postData() ?? "";
    await route.fulfill({
      contentType: "application/json",
      body: JSON.stringify({ content: [{ type: "text", text: reply }] }),
    });
  });

  await panel.getByRole("button", { name: "전송" }).click(); // 입력은 그대로 남아 있다
  const codeCard = panel.getByTestId("chat-code-card");
  await expect(codeCard).toBeVisible({ timeout: 15_000 });
  await expect(codeCard).toContainText('total = sheet("A1:A3").sum()');
  // 시스템 프롬프트: 앱 규칙 + 사용자 지침 레이어 + 우선 순위 문구 + 지침 펜스 계약
  expect(sentBody).toContain("시트기반 파이썬");
  expect(sentBody).toContain("앱 규칙이 우선");
  expect(sentBody).toContain("보험·계리"); // 기본 지침 시드
  expect(sentBody).toContain("[워크북 개요]"); // 값 미포함 개요만 (값은 read_range 도구로)

  // ── 코드 카드 → 아래 새 블록으로 (자동 실행 없음)
  await codeCard.getByRole("button", { name: "아래 새 블록으로" }).click();
  await expect
    .poll(() =>
      page.evaluate(() => (window as any).__pygridStore.getState().workbook.pyBlocks.length),
    )
    .toBe(1);
  const block = await page.evaluate(
    () => (window as any).__pygridStore.getState().workbook.pyBlocks[0],
  );
  expect(block.code).toContain('total = sheet("A1:A3").sum()');
  expect(block.last).toBeUndefined();

  // ── 지침 펜스 응답 → 확인 카드 (코드 카드로 렌더되지 않는다) → [반영]
  reply = "지침에 반영하겠습니다.\n```지침\n예제는 3행 이내 데이터로 만든다.\n```";
  await chatInput.fill("앞으로 예제는 짧게. 지침에 반영해줘");
  await panel.getByRole("button", { name: "전송" }).click();
  const proposal = panel.getByTestId("instruction-proposal");
  await expect(proposal).toBeVisible({ timeout: 15_000 });
  await expect(proposal).toContainText("지침에 추가: 예제는 3행 이내 데이터로 만든다.");
  await expect(panel.getByTestId("chat-code-card")).toHaveCount(1); // 지침 펜스는 코드 카드 아님
  await proposal.getByRole("button", { name: "반영", exact: true }).click();
  await expect(proposal).toBeHidden(); // 펜스 제거로 카드도 사라진다

  // 지침 편집 다이얼로그에 기본 시드 + 추가 지침이 함께 보인다
  await panel.getByRole("button", { name: "채팅 지침 편집" }).click();
  const instrArea = page.locator('textarea[aria-label="채팅 지침"]');
  await expect(instrArea).toBeVisible();
  const instrValue = await instrArea.inputValue();
  expect(instrValue).toContain("보험·계리");
  expect(instrValue).toContain("예제는 3행 이내 데이터로 만든다.");
  // Dialog에는 shadcn 기본 X(sr-only "닫기")와 본문 닫기 버튼이 함께 있다 — 본문 버튼을 겨냥
  await page.getByRole("dialog").getByRole("button", { name: "닫기", exact: true }).last().click();

  // ── 대화 유지: ✕로 닫고 툴바로 다시 열어도 이력이 남아 있다 (IndexedDB)
  await panel.getByRole("button", { name: "AI 채팅 패널 닫기" }).click();
  await expect(panel).toHaveCount(0);
  await page.getByRole("button", { name: "AI 채팅 패널 열기" }).click();
  const panel2 = page.getByTestId("ai-chat-panel");
  await expect(panel2).toBeVisible();
  await expect(panel2.getByText("손해율 합계 코드 만들어줘")).toBeVisible();
  await expect(panel2.getByTestId("chat-code-card")).toBeVisible();
  await expect(panel2.getByText("지침에 반영하겠습니다.")).toBeVisible();
  await expect(panel2.getByTestId("instruction-proposal")).toHaveCount(0); // 반영 완료 상태 유지
});
