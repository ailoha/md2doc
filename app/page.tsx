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
      <div
        style={{
          display: "flex",
          alignItems: "flex-start",
          justifyContent: "space-between",
          marginBottom: 12
        }}
      >
        <div>
          <h1 style={{ margin: 0, fontSize: 22, lineHeight: "28px" }}>Markdown → DOCX（固定预设）</h1>
          <div style={{ marginTop: 6, fontSize: 13, fontWeight: 700, opacity: 0.85 }}>
            本工具主要用于把 .md 格式的法规文本转换为预设格式的 .docx 文件
          </div>
        </div>

        <a
          href="https://github.com/ailoha/md2doc"
          target="_blank"
          rel="noreferrer"
          aria-label="GitHub repository"
          title="GitHub 仓库"
          style={{
            width: 28,
            height: 28,
            display: "inline-flex",
            alignItems: "center",
            justifyContent: "center",
            opacity: 0.8,
            color: "#000",
            textDecoration: "none"
          }}
        >
          <svg viewBox="0 0 16 16" width="22" height="22" aria-hidden="true" focusable="false">
            <path
              fill="currentColor"
              d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8Z"
            />
          </svg>
        </a>
      </div>

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
          boxSizing: "border-box",
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
          转换并下载 .docx
        </button>
      </div>

      <section style={{ marginTop: 16, fontSize: 13, opacity: 0.85 }}>
        <div>预设格式：</div>
        <ul style={{ marginTop: 6 }}>
          <li>文件名：默认取第一条一级标题（以“# ”开头）作为“标题.docx”；如无一级标题则为 md2doc.docx。</li>
          <li>页边距：上 3.2cm、下 2.6cm、左 2.8cm、右 2.6cm。</li>
          <li>缩进：左右 0；首行缩进 2 字符。</li>
          <li>行距：固定值 30 磅。</li>
          <li>标题：# 方正小标宋_GBK 二号 居中；## 方正黑体_GBK 三号 居中；### 方正楷体_GBK 三号 居中；标题不加粗；# / ## 段前段后间距各 1 行。</li>
          <li>正文：方正仿宋_GBK 三号 左对齐；原 Markdown 中 **加粗** 的文字会加粗显示。</li>
          <li>英文/数字：按宋体渲染（ascii/hAnsi）。</li>
          <li>页脚：默认居中页码（1, 2, 3…），宋体四号。</li>
        </ul>
      </section>
    </main>
  );
}