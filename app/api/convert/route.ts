import { NextResponse } from "next/server";
import { mdToDocxBuffer } from "@/lib/mdToDocx";
import { normalizeDocxFilename } from "@/lib/markdownFilename";

export const runtime = "nodejs";

const MAX_MARKDOWN_SIZE = 2_000_000;

type ConvertPayload = {
  markdown: string;
  filename?: string;
};

export async function POST(req: Request) {
  let payload: unknown;
  try {
    payload = await req.json();
  } catch {
    return badRequest("请求体不是合法 JSON");
  }

  const parsed = parsePayload(payload);
  if (!parsed.ok) {
    return badRequest(parsed.message);
  }

  const filename = normalizeDocxFilename(parsed.value.filename);

  try {
    const docBuffer = await mdToDocxBuffer(parsed.value.markdown);
    return new NextResponse(new Uint8Array(docBuffer), {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": contentDisposition(filename),
        "Cache-Control": "no-store",
        "X-Content-Type-Options": "nosniff"
      }
    });
  } catch {
    return new NextResponse("文档生成失败，请检查 Markdown 内容后重试。", { status: 500 });
  }
}

function parsePayload(input: unknown): { ok: true; value: ConvertPayload } | { ok: false; message: string } {
  if (!input || typeof input !== "object") {
    return { ok: false, message: "请求体必须是对象" };
  }

  const maybe = input as Record<string, unknown>;
  if (typeof maybe.markdown !== "string") {
    return { ok: false, message: "`markdown` 必须是字符串" };
  }

  if (maybe.markdown.length === 0) {
    return { ok: false, message: "Markdown 不能为空" };
  }

  if (maybe.markdown.length > MAX_MARKDOWN_SIZE) {
    return { ok: false, message: `Markdown 过大，最大支持 ${MAX_MARKDOWN_SIZE.toLocaleString()} 字符` };
  }

  if (typeof maybe.filename !== "undefined" && typeof maybe.filename !== "string") {
    return { ok: false, message: "`filename` 必须是字符串" };
  }

  return {
    ok: true,
    value: {
      markdown: maybe.markdown,
      filename: typeof maybe.filename === "string" ? maybe.filename : undefined
    }
  };
}

function badRequest(message: string): NextResponse {
  return new NextResponse(message, { status: 400 });
}

function contentDisposition(name: string): string {
  const asciiFallback = toAsciiFilename(name) || "md2doc.docx";
  const encoded = encodeRFC5987(name);
  return `attachment; filename="${asciiFallback}"; filename*=UTF-8''${encoded}`;
}

function toAsciiFilename(name: string): string {
  return name
    .replace(/[^\x20-\x7E]/g, "_")
    .replace(/["\\]/g, "_")
    .replace(/[\r\n]/g, "");
}

function encodeRFC5987(value: string): string {
  return encodeURIComponent(value).replace(/['()*]/g, (char) => `%${char.charCodeAt(0).toString(16).toUpperCase()}`);
}
