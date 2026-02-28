import {
  AlignmentType,
  Document,
  HeadingLevel,
  LineRuleType,
  Packer,
  Paragraph,
  TextRun,
  LevelFormat
} from "docx";
import { unified } from "unified";
import remarkParse from "remark-parse";

const CM = (cm: number) => Math.round((cm / 2.54) * 1440); // cm -> twips
const PT_TO_HALF = (pt: number) => pt * 2; // pt -> half-points (docx size)
const LINE_PT_EXACT = (pt: number) => pt * 20; // pt -> twips

// Common mapping: 二号=22pt, 三号=16pt
const SIZE_2 = PT_TO_HALF(22);
const SIZE_3 = PT_TO_HALF(16);

const preset = {
  margins: {
    top: CM(3.2),
    bottom: CM(2.6),
    left: CM(2.8),
    right: CM(2.6)
  },
  indent: {
    left: 0,
    right: 0,
    firstLine: 2 * 240 // 首行缩进2字符（经验值：1字符≈240 twips）
  },
  line: {
    twips: LINE_PT_EXACT(30) // 固定值 30磅
  },
  fonts: {
    // Headings
    h1: { eastAsia: "方正小标宋_GBK", ascii: "宋体" },
    h2: { eastAsia: "方正黑体_GBK", ascii: "宋体" },
    h3: { eastAsia: "方正楷体_GBK", ascii: "宋体" },
    // Body
    body: { eastAsia: "方正仿宋_GBK", ascii: "宋体" },
    code: { eastAsia: "方正仿宋_GBK", ascii: "Consolas" }
  }
};

function breakAutoLinks(s: string) {
  // Prevent Word from auto-detecting URLs/emails and applying hyperlink (blue) styling.
  const ZWSP = "\u200B"; // zero-width space

  // http(s)://...
  s = s.replace(/https?:\/\/[^\s]+/g, (m) => m.replace(/[.:/?#&=%_\-]/g, (ch) => ch + ZWSP));
  // www....
  s = s.replace(/www\.[^\s]+/g, (m) => m.replace(/[.:/?#&=%_\-]/g, (ch) => ch + ZWSP));
  // email
  s = s.replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, (m) => m.replace(/[@.]/g, (ch) => ch + ZWSP));

  return s;
}

function makeRun(args: {
  text: string;
  eastAsia: string;
  ascii: string;
  size: number;
  bold?: boolean;
}) {
  return new TextRun({
    text: breakAutoLinks(args.text),
    bold: args.bold,
    size: args.size,
    color: "000000",
    font: {
      eastAsia: args.eastAsia,
      ascii: args.ascii,
      hAnsi: args.ascii
    }
  });
}

function commonSpacing() {
  return {
    line: preset.line.twips,
    lineRule: LineRuleType.EXACT,
    before: 0,
    after: 0
  };
}

function heading(text: string, level: 1 | 2 | 3) {
  const font = level === 1 ? preset.fonts.h1 : level === 2 ? preset.fonts.h2 : preset.fonts.h3;
  const size = level === 1 ? SIZE_2 : SIZE_3;
  const headingLevel =
    level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3;

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
    spacing: commonSpacing(),
    indent: { left: 0, right: 0, firstLine: 0 }
  });
}

function bodyParagraph(text: string) {
  const font = preset.fonts.body;
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    children: [
      makeRun({
        text,
        eastAsia: font.eastAsia,
        ascii: font.ascii,
        size: SIZE_3
      })
    ],
    spacing: commonSpacing(),
    indent: {
      left: preset.indent.left,
      right: preset.indent.right,
      firstLine: preset.indent.firstLine
    }
  });
}

function codeBlock(text: string) {
  const font = preset.fonts.code;
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    children: [
      makeRun({
        text,
        eastAsia: font.eastAsia,
        ascii: font.ascii,
        size: SIZE_3
      })
    ],
    spacing: commonSpacing(),
    indent: { left: 0, right: 0, firstLine: 0 }
  });
}

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
        }
      ]
    }
  ]
};

function bulletItem(text: string) {
  const font = preset.fonts.body;
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    numbering: { reference: "bullet", level: 0 },
    children: [
      makeRun({
        text,
        eastAsia: font.eastAsia,
        ascii: font.ascii,
        size: SIZE_3
      })
    ],
    spacing: commonSpacing(),
    indent: { left: 0, right: 0, firstLine: 0 }
  });
}

function extractPlainText(children: any[]): string {
  if (!Array.isArray(children)) return "";
  let out = "";
  for (const c of children) {
    if (!c) continue;
    if (c.type === "text" && typeof c.value === "string") out += c.value;
    else if (c.type === "inlineCode" && typeof c.value === "string") out += c.value;
    else if (Array.isArray(c.children)) out += extractPlainText(c.children);
  }
  return out;
}

export async function mdToDocxBuffer(markdown: string) {
  const tree = unified().use(remarkParse).parse(markdown) as any;

  const paragraphs: Paragraph[] = [];

  const visit = (node: any) => {
    if (!node || typeof node !== "object") return;

    switch (node.type) {
      case "heading": {
        const text = extractPlainText(node.children);
        const depth = Math.min(3, Math.max(1, Number(node.depth || 1))) as 1 | 2 | 3;
        if (text.trim()) paragraphs.push(heading(text.trim(), depth));
        break;
      }
      case "paragraph": {
        const text = extractPlainText(node.children);
        if (text.trim()) paragraphs.push(bodyParagraph(text.trim()));
        break;
      }
      case "list": {
        // 当前实现：统一作为无序列表；后续如需有序列表可扩展 numbering
        const items = Array.isArray(node.children) ? node.children : [];
        for (const item of items) {
          // remark 的 listItem 里通常 children[0] 是 paragraph
          const first = Array.isArray(item.children) ? item.children[0] : null;
          const text =
            first?.type === "paragraph" ? extractPlainText(first.children) : extractPlainText(item.children || []);
          if (text.trim()) paragraphs.push(bulletItem(text.trim()));
        }
        break;
      }
      case "code": {
        const text = typeof node.value === "string" ? node.value : "";
        if (text.trim()) paragraphs.push(codeBlock(text));
        break;
      }
      default:
        break;
    }

    if (Array.isArray(node.children)) {
      for (const c of node.children) visit(c);
    }
  };

  visit(tree);

  const doc = new Document({
    numbering,
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: preset.margins.top,
              bottom: preset.margins.bottom,
              left: preset.margins.left,
              right: preset.margins.right
            }
          }
        },
        children: paragraphs.length ? paragraphs : [bodyParagraph("")]
      }
    ]
  });

  return await Packer.toBuffer(doc);
}