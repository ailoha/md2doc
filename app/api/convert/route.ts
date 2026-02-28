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
  const body = new Uint8Array(buf); // Buffer -> Uint8Array for NextResponse BodyInit

  return new NextResponse(body, {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": contentDisposition(filename)
    }
  });
}

function sanitizeFilename(name: string) {
  const base = name.trim().replace(/[/\\?%*:|"<>]/g, "_");
  return base.toLowerCase().endsWith(".docx") ? base : `${base}.docx`;
}

function contentDisposition(name: string) {
  const safe = sanitizeFilename(name);
  // ASCII fallback for `filename=` to avoid ByteString errors in some runtimes
  const asciiFallback = toAsciiFilename(safe) || "md2doc.docx";
  const encoded = encodeRFC5987ValueChars(safe);
  return `attachment; filename="${asciiFallback}"; filename*=UTF-8''${encoded}`;
}

function toAsciiFilename(name: string) {
  // Replace non-ASCII chars with underscore for the `filename=` parameter.
  return name.replace(/[^\x20-\x7E]/g, "_");
}

function encodeRFC5987ValueChars(str: string) {
  // RFC 5987 encoding for HTTP header parameters
  return encodeURIComponent(str)
    .replace(/['()]/g, escape)
    .replace(/\*/g, "%2A")
    .replace(/%(7C|60|5E)/g, (m) => m.toLowerCase());
}