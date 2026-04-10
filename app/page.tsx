"use client";

import { useMemo, useState } from "react";
import { buildFilenameFromMarkdown } from "@/lib/markdownFilename";
import type { DocumentStyle } from "@/lib/mdToDocx";
import styles from "./page.module.css";

const DEFAULT_MARKDOWN = `# 标题（一级）

## 二级标题

### 三级标题

正文示例：这是中文正文。English words and 12345 digits should use 宋体 (ascii/hAnsi).

- 无序列表项 1
- 无序列表项 2

段落二：继续正文内容。
`;

export default function Page() {
  const [markdown, setMarkdown] = useState<string>(DEFAULT_MARKDOWN);
  const [docStyle, setDocStyle] = useState<DocumentStyle>("一般公文");
  const [isConverting, setIsConverting] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string>("");

  const stats = useMemo(() => {
    return {
      chars: markdown.length,
      lines: markdown.split(/\r?\n/).length
    };
  }, [markdown]);

  async function convert() {
    if (isConverting) {
      return;
    }

    const trimmed = markdown.trim();
    if (!trimmed) {
      setErrorMessage("请输入 Markdown 内容后再转换。");
      return;
    }

    setIsConverting(true);
    setErrorMessage("");

    const downloadName = buildFilenameFromMarkdown(trimmed);

    try {
      const response = await fetch("/api/convert", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ markdown: trimmed, filename: downloadName, style: docStyle })
      });

      if (!response.ok) {
        const detail = await response.text().catch(() => "");
        setErrorMessage(detail || `转换失败（HTTP ${response.status}）`);
        return;
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = downloadName;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      URL.revokeObjectURL(url);
    } catch {
      setErrorMessage("网络异常，文档转换失败，请稍后再试。");
    } finally {
      setIsConverting(false);
    }
  }

  return (
    <main className={styles.page}>
      <header className={styles.header}>
        <div>
          <h1 className={styles.title}>Markdown → DOCX（固定预设）</h1>
          <p className={styles.subtitle}>本工具用于将 Markdown 法规文本转换为预设格式的 Word 文档。</p>
        </div>
        <a
          className={styles.repoLink}
          href="https://github.com/ailoha/md2doc"
          target="_blank"
          rel="noreferrer"
          aria-label="GitHub repository"
          title="GitHub 仓库"
        >
          <svg viewBox="0 0 16 16" width="22" height="22" aria-hidden="true" focusable="false">
            <path
              fill="currentColor"
              d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8Z"
            />
          </svg>
        </a>
      </header>

      <div className={styles.stats}>
        <span>字符：{stats.chars}</span>
        <span>行数：{stats.lines}</span>
      </div>

      <div className={styles.styleSelector}>
        <span className={styles.styleSelectorLabel}>文档样式：</span>
        {(["制度文件", "一般公文"] as const).map((s) => (
          <button
            key={s}
            className={`${styles.styleBtn} ${docStyle === s ? styles.styleBtnActive : ""}`}
            onClick={() => setDocStyle(s)}
            type="button"
          >
            {s}
          </button>
        ))}
      </div>

      <textarea
        className={styles.editor}
        value={markdown}
        onChange={(event) => setMarkdown(event.target.value)}
        spellCheck={false}
      />

      <div className={styles.actions}>
        <button className={styles.convertBtn} onClick={convert} disabled={isConverting}>
          {isConverting ? "转换中..." : "转换并下载 .docx"}
        </button>
      </div>

      {errorMessage ? <p className={styles.error}>{errorMessage}</p> : null}

      <section className={styles.preset}>
        <h2 className={styles.presetTitle}>预设格式</h2>
        <ul className={styles.presetList}>
          <li>文件名：默认使用第一条一级标题（`# `）作为 `标题.docx`，无一级标题时为 `md2doc.docx`。</li>
          <li>页边距：上 34.58mm、下 32.58mm、左 28mm、右 26mm。</li>
          <li>缩进：左右 0；首行缩进 2 字符。</li>
          <li>行距：固定值 30 磅。</li>
          <li>标题：# 方正小标宋_GBK 二号居中；## 方正黑体_GBK 三号居中；### 方正楷体_GBK 三号居中。</li>
          <li>正文：方正仿宋_GBK 三号左对齐；Markdown 中 `**加粗**` 会保留加粗。</li>
          <li>英文/数字：按宋体（ascii/hAnsi）渲染。</li>
          <li>页脚：默认居中页码（1,2,3...），宋体四号。</li>
        </ul>
      </section>
    </main>
  );
}
