"use client";

import { useMemo, useState } from "react";

function getH1Title(markdown: string) {
  const lines = markdown.split(/\r?\n/);
  for (const line of lines) {
    // Match first level-1 heading only: "# " but not "## "
    if (line.startsWith("# ") && !line.startsWith("## ")) {
      const title = line.slice(2).trim();
      if (title) return title;
    }
  }
  return "";
}

function sanitizeFilenameBase(name: string) {
  // Remove characters illegal in filenames on common OSes and trim length.
  const cleaned = name
    .replace(/[/\\?%*:|"<>]/g, "_")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned.length > 80 ? cleaned.slice(0, 80).trim() : cleaned;
}

export default function Page() {
  const [md, setMd] = useState<string>(() => `# 标题（一级）

## 二级标题

### 三级标题

正文示例：这是中文正文。English words and 12345 digits should use 宋体 (ascii/hAnsi).

- 无序列表项 1
- 无序列表项 2

段落二：继续正文内容。
`);

  const stats = useMemo(() => {
    const len = md.length;
    const lines = md.split(/\r?\n/).length;
    return { len, lines };
  }, [md]);

  async function convert() {
    const title = sanitizeFilenameBase(getH1Title(md));
    const downloadName = title ? `${title}.docx` : "md2doc.docx";

    const res = await fetch("/api/convert", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ markdown: md, filename: downloadName })
    });

    if (!res.ok) {
      const text = await res.text().catch(() => "");
      alert(`转换失败：${res.status}\n${text}`);
      return;
    }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = downloadName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  return (
    <main style={{ maxWidth: 980, margin: "24px auto", padding: 16 }}>
      <h1 style={{ margin: "0 0 12px 0", fontSize: 22 }}>Markdown → DOCX（固定预设）</h1>

      <div style={{ display: "flex", gap: 12, alignItems: "center", marginBottom: 10, fontSize: 13, opacity: 0.8 }}>
        <span>字符：{stats.len}</span>
        <span>行数：{stats.lines}</span>
      </div>

      <textarea
        value={md}
        onChange={(e) => setMd(e.target.value)}
        spellCheck={false}
        style={{
          width: "100%",
          height: 520,
          fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace",
          fontSize: 13,
          lineHeight: 1.4,
          padding: 12,
          borderRadius: 10,
          border: "1px solid rgba(0,0,0,0.15)"
        }}
      />

      <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
        <button
          onClick={convert}
          style={{
            padding: "10px 14px",
            borderRadius: 10,
            border: "1px solid rgba(0,0,0,0.2)",
            background: "white",
            cursor: "pointer"
          }}
        >
          转换并下载 DOCX
        </button>

        <a
          href="https://vercel.com/new"
          target="_blank"
          rel="noreferrer"
          style={{ alignSelf: "center", fontSize: 13 }}
        >
          部署到 Vercel
        </a>
      </div>

      <section style={{ marginTop: 16, fontSize: 13, opacity: 0.85 }}>
        <div>说明：</div>
        <ul style={{ marginTop: 6 }}>
          <li>字体为 Word 渲染端决定；打开文档的电脑需安装方正字体，否则会回退。</li>
          <li>当前支持：# / ## / ###、普通段落、无序列表、代码块（等宽）。</li>
        </ul>
      </section>
    </main>
  );
}