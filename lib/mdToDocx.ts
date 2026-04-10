import {
  AlignmentType,
  Document,
  Footer,
  HeadingLevel,
  LevelFormat,
  LineRuleType,
  Packer,
  PageNumber,
  Paragraph,
  TextRun
} from "docx";
import { unified } from "unified";
import remarkParse from "remark-parse";
import JSZip from "jszip";

type MarkdownNode = {
  type: string;
  children?: MarkdownNode[];
  value?: string;
  depth?: number;
  ordered?: boolean;
};

const cmToTwips = (cm: number) => Math.round((cm / 2.54) * 1440);
const ptToHalfPoint = (pt: number) => pt * 2;
const ptToTwips = (pt: number) => pt * 20;

const SIZE_2 = ptToHalfPoint(22);
const SIZE_3 = ptToHalfPoint(16);
const SIZE_4 = ptToHalfPoint(14);

const PRESET = {
  margins: {
    top: cmToTwips(3.458),
    bottom: cmToTwips(3.258),
    left: cmToTwips(2.8),
    right: cmToTwips(2.6)
  },
  indent: {
    left: 0,
    right: 0,
    firstLine: 2 * 240
  },
  line: {
    twips: ptToTwips(30)
  },
  fonts: {
    h1: { eastAsia: "方正小标宋_GBK", ascii: "宋体" },
    h2: { eastAsia: "方正黑体_GBK", ascii: "宋体" },
    h3: { eastAsia: "方正楷体_GBK", ascii: "宋体" },
    body: { eastAsia: "方正仿宋_GBK", ascii: "宋体" },
    code: { eastAsia: "方正仿宋_GBK", ascii: "Consolas" }
  }
} as const;

const URL_LIKE_HINT_RE = /https?:\/\/|www\.|@/;
const HTTP_LINK_RE = /https?:\/\/[^\s]+/g;
const WWW_LINK_RE = /www\.[^\s]+/g;
const EMAIL_RE = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;
const AUTO_LINK_BREAK_RE = /[.:/?#&=%_\-]/g;
const EMAIL_BREAK_RE = /[@.]/g;
const ZERO_WIDTH_SPACE = "\u200B";

const numbering = {
  config: [
    {
      reference: "bullet",
      levels: [
        {
          level: 0,
          format: LevelFormat.BULLET,
          text: "•",
          alignment: AlignmentType.LEFT
        },
        {
          level: 1,
          format: LevelFormat.BULLET,
          text: "◦",
          alignment: AlignmentType.LEFT
        },
        {
          level: 2,
          format: LevelFormat.BULLET,
          text: "▪",
          alignment: AlignmentType.LEFT
        }
      ]
    },
    {
      reference: "decimal",
      levels: [
        {
          level: 0,
          format: LevelFormat.DECIMAL,
          text: "%1.",
          alignment: AlignmentType.LEFT
        },
        {
          level: 1,
          format: LevelFormat.DECIMAL,
          text: "%2.",
          alignment: AlignmentType.LEFT
        },
        {
          level: 2,
          format: LevelFormat.DECIMAL,
          text: "%3.",
          alignment: AlignmentType.LEFT
        }
      ]
    }
  ]
};

function getChildren(node: MarkdownNode | undefined): MarkdownNode[] {
  return Array.isArray(node?.children) ? node!.children! : [];
}

function getNodeText(node: MarkdownNode | undefined): string {
  return typeof node?.value === "string" ? node.value : "";
}

function breakAutoLinks(text: string): string {
  if (!text || !URL_LIKE_HINT_RE.test(text)) {
    return text;
  }

  let output = text;
  output = output.replace(HTTP_LINK_RE, (value) => value.replace(AUTO_LINK_BREAK_RE, (ch) => `${ch}${ZERO_WIDTH_SPACE}`));
  output = output.replace(WWW_LINK_RE, (value) => value.replace(AUTO_LINK_BREAK_RE, (ch) => `${ch}${ZERO_WIDTH_SPACE}`));
  output = output.replace(EMAIL_RE, (value) => value.replace(EMAIL_BREAK_RE, (ch) => `${ch}${ZERO_WIDTH_SPACE}`));
  return output;
}

function commonSpacing() {
  return {
    line: PRESET.line.twips,
    lineRule: LineRuleType.EXACT,
    before: 0,
    after: 0
  };
}

function makeRun(params: {
  text: string;
  eastAsia: string;
  ascii: string;
  size: number;
  bold?: boolean;
}): TextRun {
  return new TextRun({
    text: breakAutoLinks(params.text),
    bold: params.bold,
    size: params.size,
    color: "000000",
    font: {
      eastAsia: params.eastAsia,
      ascii: params.ascii,
      hAnsi: params.ascii
    }
  });
}

function extractPlainText(nodes: MarkdownNode[]): string {
  let output = "";
  for (const node of nodes) {
    if (node.type === "text" || node.type === "inlineCode") {
      output += getNodeText(node);
      continue;
    }

    const children = getChildren(node);
    if (children.length > 0) {
      output += extractPlainText(children);
    }
  }

  return output;
}

function markdownInlinesToRuns(nodes: MarkdownNode[], inheritedBold = false): TextRun[] {
  const runs: TextRun[] = [];
  const font = PRESET.fonts.body;

  for (const node of nodes) {
    if (node.type === "text" || node.type === "inlineCode") {
      const text = getNodeText(node);
      if (text.length > 0) {
        runs.push(
          makeRun({
            text,
            eastAsia: font.eastAsia,
            ascii: font.ascii,
            size: SIZE_3,
            bold: inheritedBold
          })
        );
      }
      continue;
    }

    const children = getChildren(node);
    if (children.length === 0) {
      continue;
    }

    if (node.type === "strong") {
      runs.push(...markdownInlinesToRuns(children, true));
      continue;
    }

    runs.push(...markdownInlinesToRuns(children, inheritedBold));
  }

  return runs;
}

function bodyParagraphFromRuns(runs: TextRun[]): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    children: runs,
    spacing: commonSpacing(),
    indent: {
      left: PRESET.indent.left,
      right: PRESET.indent.right,
      firstLine: PRESET.indent.firstLine
    }
  });
}

function paragraphFromInlineNodes(nodes: MarkdownNode[]): Paragraph {
  const runs = markdownInlinesToRuns(nodes);
  if (runs.length > 0) {
    return bodyParagraphFromRuns(runs);
  }

  return bodyParagraphFromRuns([
    makeRun({
      text: extractPlainText(nodes),
      eastAsia: PRESET.fonts.body.eastAsia,
      ascii: PRESET.fonts.body.ascii,
      size: SIZE_3
    })
  ]);
}

function headingParagraph(text: string, level: 1 | 2 | 3): Paragraph {
  const headingLevel =
    level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3;
  const font = level === 1 ? PRESET.fonts.h1 : level === 2 ? PRESET.fonts.h2 : PRESET.fonts.h3;
  const size = level === 1 ? SIZE_2 : SIZE_3;
  const beforeAfter = level <= 2 ? ptToTwips(12) : 0;

  return new Paragraph({
    heading: headingLevel,
    alignment: AlignmentType.CENTER,
    children: [
      makeRun({
        text,
        eastAsia: font.eastAsia,
        ascii: font.ascii,
        size
      })
    ],
    spacing: {
      ...commonSpacing(),
      before: beforeAfter,
      after: beforeAfter
    },
    indent: { left: 0, right: 0, firstLine: 0 }
  });
}

function codeBlockParagraphs(value: string): Paragraph[] {
  if (!value.trim()) {
    return [];
  }

  const lines = value.replace(/\r\n/g, "\n").split("\n");
  while (lines.length > 1 && lines[lines.length - 1] === "") {
    lines.pop();
  }

  return lines.map(
    (line) =>
      new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [
          makeRun({
            text: line,
            eastAsia: PRESET.fonts.code.eastAsia,
            ascii: PRESET.fonts.code.ascii,
            size: SIZE_3
          })
        ],
        spacing: commonSpacing(),
        indent: { left: 0, right: 0, firstLine: 0 }
      })
  );
}

function toListLevel(level: number): number {
  return Math.min(Math.max(level, 0), 2);
}

function listItemParagraphFromInlineNodes(nodes: MarkdownNode[], ordered: boolean, level: number): Paragraph {
  const runs = markdownInlinesToRuns(nodes);
  const fallback = extractPlainText(nodes);
  const safeLevel = toListLevel(level);

  return new Paragraph({
    alignment: AlignmentType.LEFT,
    numbering: {
      reference: ordered ? "decimal" : "bullet",
      level: safeLevel
    },
    children:
      runs.length > 0
        ? runs
        : [
            makeRun({
              text: fallback,
              eastAsia: PRESET.fonts.body.eastAsia,
              ascii: PRESET.fonts.body.ascii,
              size: SIZE_3
            })
          ],
    spacing: commonSpacing(),
    indent: { left: 0, right: 0, firstLine: 0 }
  });
}

function appendList(listNode: MarkdownNode, paragraphs: Paragraph[], level = 0): void {
  const ordered = !!listNode.ordered;
  const items = getChildren(listNode);

  for (const item of items) {
    if (item.type !== "listItem") {
      continue;
    }

    const itemChildren = getChildren(item);
    const firstParagraph = itemChildren.find((child) => child.type === "paragraph");

    if (firstParagraph) {
      const inlineNodes = getChildren(firstParagraph);
      if (extractPlainText(inlineNodes).trim()) {
        paragraphs.push(listItemParagraphFromInlineNodes(inlineNodes, ordered, level));
      }
    }

    for (const child of itemChildren) {
      if (child === firstParagraph) {
        continue;
      }

      if (child.type === "list") {
        appendList(child, paragraphs, level + 1);
        continue;
      }

      if (child.type === "paragraph") {
        const inlineNodes = getChildren(child);
        if (extractPlainText(inlineNodes).trim()) {
          paragraphs.push(paragraphFromInlineNodes(inlineNodes));
        }
        continue;
      }

      if (child.type === "code") {
        paragraphs.push(...codeBlockParagraphs(getNodeText(child)));
      }
    }
  }
}

function appendBlock(
  node: MarkdownNode,
  paragraphs: Paragraph[],
  headingFn: (text: string, level: 1 | 2 | 3) => Paragraph
): void {
  switch (node.type) {
    case "heading": {
      const text = extractPlainText(getChildren(node)).trim();
      if (!text) {
        return;
      }

      const depth = Math.min(3, Math.max(1, Number(node.depth ?? 1))) as 1 | 2 | 3;
      paragraphs.push(headingFn(text, depth));
      return;
    }

    case "paragraph": {
      const inlineNodes = getChildren(node);
      if (extractPlainText(inlineNodes).trim()) {
        paragraphs.push(paragraphFromInlineNodes(inlineNodes));
      }
      return;
    }

    case "list": {
      appendList(node, paragraphs);
      return;
    }

    case "code": {
      paragraphs.push(...codeBlockParagraphs(getNodeText(node)));
      return;
    }

    case "blockquote": {
      for (const child of getChildren(node)) {
        appendBlock(child, paragraphs, headingFn);
      }
      return;
    }

    default:
      return;
  }
}

function footer(): Footer {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            children: [PageNumber.CURRENT],
            size: SIZE_4,
            bold: false,
            color: "000000",
            font: { eastAsia: "宋体", ascii: "宋体", hAnsi: "宋体" }
          })
        ],
        spacing: commonSpacing()
      })
    ]
  });
}

export type DocumentStyle = "制度文件" | "一般公文";

function headingParagraphGeneral(text: string, level: 1 | 2 | 3): Paragraph {
  const headingLevel =
    level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3;
  const font = level === 1 ? PRESET.fonts.h1 : level === 2 ? PRESET.fonts.h2 : PRESET.fonts.h3;
  const size = level === 1 ? SIZE_2 : SIZE_3;

  if (level === 1) {
    return new Paragraph({
      heading: headingLevel,
      alignment: AlignmentType.CENTER,
      children: [
        makeRun({ text, eastAsia: font.eastAsia, ascii: font.ascii, size })
      ],
      spacing: commonSpacing(),
      indent: { left: 0, right: 0, firstLine: 0 }
    });
  }

  return new Paragraph({
    heading: headingLevel,
    alignment: AlignmentType.LEFT,
    children: [
      makeRun({ text, eastAsia: font.eastAsia, ascii: font.ascii, size })
    ],
    spacing: commonSpacing(),
    indent: {
      left: PRESET.indent.left,
      right: PRESET.indent.right,
      firstLine: PRESET.indent.firstLine
    }
  });
}

export async function mdToDocxBuffer(markdown: string, style: DocumentStyle = "制度文件"): Promise<Buffer> {
  const tree = unified().use(remarkParse).parse(markdown) as MarkdownNode;
  const paragraphs: Paragraph[] = [];
  const emptyParagraph = bodyParagraphFromRuns([
    makeRun({
      text: "",
      eastAsia: PRESET.fonts.body.eastAsia,
      ascii: PRESET.fonts.body.ascii,
      size: SIZE_3
    })
  ]);

  const headingFn = style === "一般公文" ? headingParagraphGeneral : headingParagraph;

  for (const node of getChildren(tree)) {
    appendBlock(node, paragraphs, headingFn);
  }

  const doc = new Document({
    numbering,
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: PRESET.margins.top,
              bottom: PRESET.margins.bottom,
              left: PRESET.margins.left,
              right: PRESET.margins.right
            }
          }
        },
        footers: {
          default: footer()
        },
        children: paragraphs.length > 0 ? paragraphs : [emptyParagraph]
      }
    ]
  });

  const buf = await Packer.toBuffer(doc);

  // The docx library only supports w:firstLine (twips), which Word displays as
  // centimeters. To get Word to show "2 字符" we need w:firstLineChars="200"
  // (hundredths of a character). Post-process the DOCX XML to patch this.
  const zip = await JSZip.loadAsync(buf);
  const docXml = await zip.file("word/document.xml")!.async("string");
  const patched = docXml.replace(
    /w:firstLine="480"/g,
    'w:firstLineChars="200" w:firstLine="480"'
  );
  zip.file("word/document.xml", patched);
  return zip.generateAsync({ type: "nodebuffer" });
}
