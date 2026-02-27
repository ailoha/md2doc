import { NextResponse } from "next/server";
import { mdToDocxBuffer } from "@/lib/mdToDocx";

export const runtime = "nodejs";

export async function POST(req: Request) {
  let payload: unknown;
  try {
    payload = await req.json();
  } catch {
    return new NextResponse("Invalid JSON body", { status: 400 });
  }

  const markdown = typeof (payload as any)?.markdown === "string" ? (payload as any).markdown : "";
  const filename = typeof (payload as any)?.filename === "string" ? (payload as any).filename : "output.docx";

  const buf = await mdToDocxBuffer(markdown);

  return new NextResponse(buf, {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": `attachment; filename="${sanitizeFilename(filename)}"`
    }
  });
}

function sanitizeFilename(name: string) {
  const base = name.trim().replace(/[/\\?%*:|"<>]/g, "_");
  return base.toLowerCase().endsWith(".docx") ? base : `${base}.docx`;
}