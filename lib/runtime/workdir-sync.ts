// 부록 P: 작업 폴더 동기화 — 워크북의 workDir를 런타임에 "코드"(os.makedirs + os.chdir)로 적용한다.
// 런타임 FS는 메모리라 재부트하면 초기화되므로, 부트 완료마다 같은 코드를 다시 실행한다.

import { HOME_DIR, workDirCode } from "@/lib/grid/files";
import { useWorkbookStore } from "@/lib/grid/model";
import type { RuntimeClient } from "./client";

/** 작업 폴더 적용 (요청은 런타임 준비 전이면 큐에 들어간다). 실패 메시지 또는 null */
export async function applyWorkDir(client: RuntimeClient, dir: string): Promise<string | null> {
  const res = await client.repl(workDirCode(dir));
  return res.traceback ? res.traceback.trim().split("\n").pop() ?? "작업 폴더 지정 실패" : null;
}

/** 부트 완료·워크북 workDir 변경 시 자동 적용. 해제 함수를 돌려준다 */
export function startWorkDirSync(client: RuntimeClient): () => void {
  let prev = client.getStatus();
  let applied: string | undefined; // 마지막으로 적용한 폴더
  const sync = (force: boolean) => {
    const want = useWorkbookStore.getState().workbook.workDir;
    if (!force && want === applied) return;
    if (want === undefined && applied === undefined) return; // 기본(홈) 그대로
    applied = want;
    void applyWorkDir(client, want ?? HOME_DIR).catch(() => undefined);
  };
  const offStatus = client.on("status", (s) => {
    const booted = s === "ready" && (prev === "loading" || prev === "rebooting" || prev === "idle");
    prev = s;
    if (booted) {
      applied = undefined; // 새 FS — 지정돼 있으면 다시 만든다
      sync(true);
    }
  });
  const offStore = useWorkbookStore.subscribe((st, old) => {
    if (st.workbook.workDir !== old.workbook.workDir) sync(false);
  });
  return () => {
    offStatus();
    offStore();
  };
}
