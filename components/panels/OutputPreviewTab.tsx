"use client";

// 출력 미리보기 탭 — 선택 블록의 last.preview: 표(정렬)·이미지·repr (§2.3.4)

import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { formatA1 } from "@/lib/grid/a1";
import { useWorkbookStore, type CellEdit } from "@/lib/grid/model";
import type { PreviewPayload } from "@/lib/runtime/protocol";
import { getBlob } from "@/lib/storage/db";
import type { Cell } from "@/types/workbook";

type Row = (string | number | boolean | null)[];

const toCell = (v: string | number | boolean | null): Cell | null => {
  if (v === null || v === "") return null;
  if (typeof v === "number") return { v, t: "n" };
  if (typeof v === "boolean") return { v, t: "b" };
  return { v, t: "s" };
};

export function PreviewTable({ preview }: { preview: Extract<PreviewPayload, { kind: "table" }> }) {
  const [sort, setSort] = useState<{ col: number; dir: 1 | -1 } | null>(null);

  const rows = useMemo(() => {
    if (!sort) return preview.rows;
    const { col, dir } = sort;
    return [...preview.rows].sort((a, b) => {
      const x = a[col];
      const y = b[col];
      if (x === null) return 1;
      if (y === null) return -1;
      if (typeof x === "number" && typeof y === "number") return (x - y) * dir;
      return String(x).localeCompare(String(y), "ko") * dir;
    });
  }, [preview.rows, sort]);

  const exportToSheet = () => {
    const edits: CellEdit[] = [];
    preview.columns.forEach((name, c) => edits.push({ r: 0, c, cell: { v: name, t: "s" } }));
    preview.rows.forEach((row, r) =>
      row.forEach((v, c) => edits.push({ r: r + 1, c, cell: toCell(v) })),
    );
    useWorkbookStore.getState().addSheetWithCells(edits);
    toast("새 시트로 내보냈습니다");
  };

  return (
    <div className="flex h-full flex-col">
      <div className="flex items-center gap-2 px-3 py-1.5 text-xs text-muted-foreground">
        <span className="font-mono">
          {preview.shape[0]}×{preview.shape[1]}
          {preview.rows.length < preview.shape[0] ? ` (상위 ${preview.rows.length}행)` : ""}
        </span>
        <Button variant="ghost" size="sm" className="ml-auto h-6 text-xs" onClick={exportToSheet}>
          새 시트로 내보내기
        </Button>
      </div>
      <div className="min-h-0 flex-1 overflow-auto px-3 pb-2">
        <table className="w-full border-collapse text-xs">
          <thead className="sticky top-0 bg-muted">
            <tr>
              {preview.columns.map((name, c) => (
                <th
                  key={c}
                  onClick={() =>
                    setSort((prev) =>
                      prev?.col === c ? { col: c, dir: prev.dir === 1 ? -1 : 1 } : { col: c, dir: 1 },
                    )
                  }
                  className="cursor-pointer border-b border-r p-1 text-left font-medium hover:bg-accent"
                  title={`${preview.dtypes[c] ?? ""} — 클릭 정렬`}
                >
                  {name}
                  {sort?.col === c ? (sort.dir === 1 ? " ▲" : " ▼") : ""}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, r) => (
              <tr key={r}>
                {row.map((v, c) => (
                  <td
                    key={c}
                    className={`border-b border-r p-1 ${typeof v === "number" ? "text-right tabular-nums" : ""}`}
                  >
                    {v === null ? "" : String(v)}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

export function PreviewImage({ blobId }: { blobId: string }) {
  const [url, setUrl] = useState<string | null>(null);

  useEffect(() => {
    let objectUrl: string | null = null;
    void getBlob(blobId).then((blob) => {
      if (blob) {
        objectUrl = URL.createObjectURL(blob);
        setUrl(objectUrl);
      }
    });
    return () => {
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [blobId]);

  const copyPng = async () => {
    const blob = await getBlob(blobId);
    if (!blob) return;
    try {
      await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
      toast("PNG를 클립보드에 복사했습니다");
    } catch {
      // 폴백: 다운로드
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "pygrid-figure.png";
      a.click();
      URL.revokeObjectURL(a.href);
      toast("클립보드 복사가 막혀 PNG 파일로 내려받았습니다");
    }
  };

  if (!url) {
    return <p className="px-3 py-6 text-xs text-muted-foreground">이미지를 불러오는 중…</p>;
  }
  return (
    <div className="flex h-full flex-col">
      <div className="flex px-3 py-1.5">
        <Button variant="ghost" size="sm" className="ml-auto h-6 text-xs" onClick={() => void copyPng()}>
          PNG 복사
        </Button>
      </div>
      <div className="min-h-0 flex-1 overflow-auto px-3 pb-2">
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src={url} alt="블록 출력 이미지" data-testid="preview-image" className="max-w-full" />
      </div>
    </div>
  );
}

export default function OutputPreviewTab() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const selectedBlockId = useWorkbookStore((s) => s.selectedBlockId);

  // 선택 블록 없으면 가장 최근 실행 블록
  const block = useMemo(() => {
    const withLast = workbook.pyBlocks.filter((b) => b.last);
    return (
      workbook.pyBlocks.find((b) => b.id === selectedBlockId) ??
      withLast.sort((a, b) => (a.last!.ranAt < b.last!.ranAt ? 1 : -1))[0]
    );
  }, [workbook.pyBlocks, selectedBlockId]);

  if (!block?.last) {
    return (
      <p className="px-3 py-6 text-center text-xs text-muted-foreground">
        블록을 실행하면 결과 미리보기가 표시됩니다
      </p>
    );
  }

  const sheetName = workbook.sheets.find((s) => s.id === block.sheetId)?.name ?? "?";
  const anchorLabel = `${sheetName}!${formatA1({
    r0: block.anchor.r,
    c0: block.anchor.c,
    r1: block.anchor.r,
    c1: block.anchor.c,
  })}`;
  const preview = block.last.preview as PreviewPayload | undefined;

  return (
    <div className="flex h-full flex-col">
      <div className="px-3 pt-1.5 font-mono text-[11px] text-muted-foreground">
        {anchorLabel} · {block.last.kind ?? "?"}
      </div>
      <div className="min-h-0 flex-1">
        {block.last.imageBlobId ? (
          <PreviewImage blobId={block.last.imageBlobId} />
        ) : preview?.kind === "table" ? (
          <PreviewTable preview={preview} />
        ) : preview?.kind === "repr" ? (
          <pre className="overflow-auto px-3 py-2 font-mono text-xs leading-5">
            {preview.repr}
          </pre>
        ) : (
          <p className="px-3 py-6 text-xs text-muted-foreground">
            미리보기가 없습니다 (마지막 실행: {block.last.status})
          </p>
        )}
      </div>
    </div>
  );
}
