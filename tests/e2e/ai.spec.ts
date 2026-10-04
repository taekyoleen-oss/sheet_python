import { expect, test, type Page } from "@playwright/test";

// 부록 L.5: 통합 후 AI 진입점 — Python 패널 상단 ✦ 생성 바는 채팅으로 라우팅된다.
// 키 미설정 → 설정 유도 → 키 저장 → (모킹된) 코드 제안 → [아래 새 블록] → 자동 실행 없음.
// 실제 Anthropic API는 호출하지 않는다 (page.route 모킹).

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

test("생성 바 → 채팅 라우팅: 키 미설정 유도 → 키 저장 → 모킹 코드 제안 → 새 블록(자동 실행 없음)", async ({
  page,
}) => {
  test.setTimeout(180_000);
  await page.goto("/");
  await waitForApp(page);

  const genInput = page.getByLabel("AI 코드 요청");
  await expect(genInput).toBeEnabled();

  // 1) 키 미설정: 생성 바 요청 → 채팅 패널이 열리고 전송 시점에 AI 설정 다이얼로그 유도
  await genInput.fill("시트 합계를 구해줘");
  await page.getByRole("button", { name: "채팅으로 요청" }).click();
  const panel = page.getByTestId("ai-chat-panel");
  await expect(panel).toBeVisible();
  const keyDialog = page.getByRole("dialog", { name: /AI 설정/ });
  await expect(keyDialog).toBeVisible();
  await expect(keyDialog.getByText(/IndexedDB.*에만 저장/)).toBeVisible();

  // 2) 키 저장 → 다이얼로그 닫힘
  await keyDialog.getByLabel("Anthropic API 키").fill("sk-ant-e2e-mock");
  await keyDialog.getByRole("button", { name: "저장", exact: true }).click();
  await expect(keyDialog).toBeHidden();

  // 3) API 모킹 → 다시 요청 → 코드 제안 카드
  let sentBody = "";
  let sentAuth = "";
  let turn = 0;
  await page.route("https://api.anthropic.com/**", async (route) => {
    sentAuth = route.request().headers()["x-api-key"] ?? "";
    sentBody = route.request().postData() ?? "";
    const body =
      turn++ === 0
        ? {
            content: [
              { type: "text", text: "합계 코드입니다." },
              {
                type: "tool_use",
                id: "tu1",
                name: "propose_block",
                input: {
                  code: 'df = sheet("A1:B3", headers=True)\ndf.sum()',
                  title: "합계",
                },
              },
            ],
            stop_reason: "tool_use",
          }
        : {
            content: [{ type: "text", text: "제안했습니다 — 확인 후 실행하세요." }],
            stop_reason: "end_turn",
          };
    await route.fulfill({ contentType: "application/json", body: JSON.stringify(body) });
  });

  await genInput.fill("시트 합계를 구해줘");
  await page.getByRole("button", { name: "채팅으로 요청" }).click();

  const card = panel.getByTestId("block-proposal");
  await expect(card).toBeVisible({ timeout: 20_000 });
  await expect(card).toContainText('df = sheet("A1:B3", headers=True)');

  // 4) [아래 새 블록] → 블록 생성, 자동 실행 없음
  await card.getByRole("button", { name: "아래 새 블록" }).click();
  await expect
    .poll(
      () =>
        page.evaluate(() => (window as any).__pygridStore.getState().workbook.pyBlocks.length),
      { timeout: 15_000 },
    )
    .toBe(1);
  const block = await page.evaluate(
    () => (window as any).__pygridStore.getState().workbook.pyBlocks[0],
  );
  expect(block.title).toBe("합계");
  expect(block.code).toContain('df = sheet("A1:B3", headers=True)');
  expect(block.last).toBeUndefined(); // 자동 실행 없음

  // 요청이 저장된 키·이 앱 규칙 프롬프트로 나갔는지 (모킹 라우트에서 캡처)
  expect(sentAuth).toBe("sk-ant-e2e-mock");
  expect(sentBody).toContain("claude-sonnet-4-6");
  expect(sentBody).toContain("시트기반 파이썬");
});
