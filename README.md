# md2doc

将 Markdown 文本转换为固定公文排版预设的 `.docx` 文件（Next.js App Router）。

## 功能

- 前端粘贴 Markdown，一键下载 `.docx`
- 默认从第一条一级标题（`# `）生成文件名
- 固定页面预设：
  - 页边距：上 3.2cm、下 2.6cm、左 2.8cm、右 2.6cm
  - 首行缩进：2 字符
  - 行距：固定值 30 磅
  - 标题字体：方正小标宋/黑体/楷体（按级别）
  - 正文字体：方正仿宋，英文数字按宋体（ascii/hAnsi）
  - 页脚：居中页码，宋体四号
- 保留 Markdown 行内 `**加粗**`
- 支持有序/无序列表输出到 Word 编号

## 技术栈

- Next.js 16（App Router）
- React 19
- TypeScript
- `docx`（生成 Word）
- `remark-parse` + `unified`（解析 Markdown）

## 目录结构

```text
app/
  api/convert/route.ts     # 转换接口（输入校验、响应头、错误处理）
  globals.css              # 全局样式
  layout.tsx               # 根布局
  page.module.css          # 页面样式
  page.tsx                 # 前端编辑和下载交互
lib/
  markdownFilename.ts      # 标题提取与文件名规范化（前后端共用）
  mdToDocx.ts              # Markdown AST -> DOCX 段落转换核心
eslint.config.mjs          # ESLint 9 Flat Config
```

## 本地开发

```bash
npm install
npm run dev
```

打开 [http://localhost:3000](http://localhost:3000)。

## 质量检查

```bash
npm run lint
npm run build
```

## 接口说明

- `POST /api/convert`
- 请求体：

```json
{
  "markdown": "# 标题\n正文",
  "filename": "可选，自定义文件名.docx"
}
```

- 约束：
  - `markdown` 必填，且最大 2,000,000 字符
  - `filename` 可选，服务端会统一清洗非法字符并补全 `.docx`

## 已做优化（本次）

- 修复列表节点重复遍历导致的潜在重复段落问题
- 转换核心改为按块级节点处理，减少无效递归
- 增加有序/无序列表编号支持
- API 增加输入校验、体积限制、转换异常兜底
- 前后端文件名逻辑收敛为单一实现（避免重复与不一致）
- 页面样式从内联迁移到 CSS Module + 全局样式
- 前端增加转换中状态与错误提示
- ESLint 适配 Next 16 / ESLint 9（Flat Config）
