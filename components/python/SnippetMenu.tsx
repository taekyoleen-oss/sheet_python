"use client";

// 초보자용 코드 스니펫 드롭다운 — {{range}}는 현재 그리드 선택의 sheet() 참조로 치환

import { CaretDown } from "@phosphor-icons/react";
import { toast } from "sonner";
import snippets from "@/data/snippets.json";
import { Button } from "@/components/ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { editorRegistry } from "@/components/python/CodeEditor";
import { useWorkbookStore } from "@/lib/grid/model";
import { xlRefForSelection } from "@/lib/grid/xl-ref";

export default function SnippetMenu() {
  const insert = (code: string) => {
    const st = useWorkbookStore.getState();
    const blockId = st.lastEditorBlockId;
    const target = blockId ? editorRegistry.get(blockId) : undefined;
    if (!target) {
      toast("블록 편집기를 먼저 클릭한 뒤 스니펫을 선택하세요");
      return;
    }
    const block = st.workbook.pyBlocks.find((b) => b.id === blockId);
    const ref = xlRefForSelection(block?.sheetId) ?? 'sheet("A1")';
    target.insert(code.replace(/\{\{range\}\}/g, ref));
  };

  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <Button variant="ghost" size="sm" className="h-6 gap-1 text-xs">
          스니펫 <CaretDown className="size-3" />
        </Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="end" className="w-64">
        {snippets.map((s) => (
          <DropdownMenuItem key={s.name} onClick={() => insert(s.code)}>
            <div>
              <div className="text-xs font-medium">{s.name}</div>
              <div className="text-[11px] text-muted-foreground">{s.description}</div>
            </div>
          </DropdownMenuItem>
        ))}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}
