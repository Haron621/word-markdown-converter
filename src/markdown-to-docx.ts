/// <reference lib="dom" />

import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  LevelFormat,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkGfm from 'remark-gfm';

type MdNode = {
  type: string;
  value?: string;
  url?: string;
  lang?: string;
  depth?: number;
  ordered?: boolean;
  children?: MdNode[];
};

type InlineStyle = {
  bold?: boolean;
  italics?: boolean;
  strike?: boolean;
};

const headingLevels = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
} as const;

function plainText(node: MdNode): string {
  if (node.value) return node.value;
  return (node.children ?? []).map(plainText).join('');
}

function inlineRuns(
  nodes: MdNode[] = [],
  style: InlineStyle = {},
): Array<TextRun | ExternalHyperlink> {
  const result: Array<TextRun | ExternalHyperlink> = [];
  for (const node of nodes) {
    if (node.type === 'text') {
      result.push(new TextRun({ text: node.value ?? '', ...style }));
    } else if (node.type === 'strong') {
      result.push(...inlineRuns(node.children, { ...style, bold: true }));
    } else if (node.type === 'emphasis') {
      result.push(...inlineRuns(node.children, { ...style, italics: true }));
    } else if (node.type === 'delete') {
      result.push(...inlineRuns(node.children, { ...style, strike: true }));
    } else if (node.type === 'inlineCode') {
      result.push(
        new TextRun({
          text: node.value ?? '',
          font: 'Consolas',
          shading: { type: ShadingType.CLEAR, fill: 'EDEDED' },
          ...style,
        }),
      );
    } else if (node.type === 'break') {
      result.push(new TextRun({ text: '', break: 1 }));
    } else if (node.type === 'link') {
      const linkedRuns = inlineRuns(node.children, style).filter(
        (child): child is TextRun => child instanceof TextRun,
      );
      result.push(
        new ExternalHyperlink({
          link: node.url ?? '',
          children: linkedRuns.length
            ? linkedRuns
            : [new TextRun({ text: node.url ?? '', style: 'Hyperlink' })],
        }),
      );
    } else if (node.type === 'image') {
      result.push(
        new TextRun({ text: `[Изображение: ${plainText(node) || node.url || ''}]`, italics: true }),
      );
    } else {
      result.push(...inlineRuns(node.children, style));
    }
  }
  return result;
}

function listBlocks(node: MdNode, level = 0): Paragraph[] {
  const blocks: Paragraph[] = [];
  for (const item of node.children ?? []) {
    let firstParagraph = true;
    for (const child of item.children ?? []) {
      if (child.type === 'list') {
        blocks.push(...listBlocks(child, level + 1));
        continue;
      }
      const paragraphOptions = node.ordered
        ? { numbering: { reference: 'markdown-numbering', level } }
        : { bullet: { level } };
      blocks.push(
        new Paragraph({
          ...paragraphOptions,
          children: inlineRuns(child.children ?? [child]),
          spacing: { after: firstParagraph ? 80 : 40 },
        }),
      );
      firstParagraph = false;
    }
  }
  return blocks;
}

function tableBlock(node: MdNode): Table {
  const rows = (node.children ?? []).map(
    (row, rowIndex) =>
      new TableRow({
        ...(rowIndex === 0 ? { tableHeader: true } : {}),
        children: (row.children ?? []).map(
          (cell) =>
            new TableCell({
              shading:
                rowIndex === 0
                  ? { type: ShadingType.CLEAR, fill: 'E8EEF7' }
                  : undefined,
              children: [
                new Paragraph({
                  children: inlineRuns(cell.children),
                }),
              ],
            }),
        ),
      }),
  );
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: 'AAB4C3' },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: 'AAB4C3' },
      left: { style: BorderStyle.SINGLE, size: 1, color: 'AAB4C3' },
      right: { style: BorderStyle.SINGLE, size: 1, color: 'AAB4C3' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'D8DEE8' },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'D8DEE8' },
    },
    rows,
  });
}

function blocks(nodes: MdNode[] = []): Array<Paragraph | Table> {
  const result: Array<Paragraph | Table> = [];
  for (const node of nodes) {
    if (node.type === 'heading') {
      result.push(
        new Paragraph({
          heading: headingLevels[(node.depth ?? 1) as keyof typeof headingLevels],
          children: inlineRuns(node.children),
        }),
      );
    } else if (node.type === 'paragraph') {
      result.push(
        new Paragraph({ children: inlineRuns(node.children), spacing: { after: 140 } }),
      );
    } else if (node.type === 'list') {
      result.push(...listBlocks(node));
    } else if (node.type === 'blockquote') {
      for (const child of node.children ?? []) {
        result.push(
          new Paragraph({
            children: inlineRuns(child.children ?? [child], { italics: true }),
            indent: { left: 540 },
            border: {
              left: { style: BorderStyle.SINGLE, size: 14, color: '7F8C9F' },
            },
            spacing: { after: 120 },
          }),
        );
      }
    } else if (node.type === 'code') {
      const lines = (node.value ?? '').split('\n');
      result.push(
        new Paragraph({
          children: lines.flatMap((line, index) => [
            new TextRun({ text: line, font: 'Consolas', size: 19 }),
            ...(index < lines.length - 1 ? [new TextRun({ text: '', break: 1 })] : []),
          ]),
          shading: { type: ShadingType.CLEAR, fill: 'F1F3F5' },
          spacing: { before: 80, after: 140 },
        }),
      );
    } else if (node.type === 'table') {
      result.push(tableBlock(node));
    } else if (node.type === 'thematicBreak') {
      result.push(
        new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'AAB4C3' } },
          spacing: { after: 160 },
        }),
      );
    } else if (node.children) {
      result.push(...blocks(node.children));
    }
  }
  return result;
}

export async function convertMarkdownToBlob(markdown: string): Promise<Blob> {
  const tree = unified().use(remarkParse).use(remarkGfm).parse(markdown) as MdNode;
  const numberingLevels = Array.from({ length: 9 }, (_, level) => ({
    level,
    format: LevelFormat.DECIMAL,
    text: `%${level + 1}.`,
    alignment: AlignmentType.START,
    style: { paragraph: { indent: { left: 720 + level * 360, hanging: 360 } } },
  }));
  const document = new Document({
    numbering: {
      config: [{ reference: 'markdown-numbering', levels: numberingLevels }],
    },
    sections: [{ properties: {}, children: blocks(tree.children) }],
  });
  return Packer.toBlob(document);
}

function download(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function initMarkdownToWord(): void {
  const input = document.getElementById('md-file') as HTMLInputElement | null;
  const editor = document.getElementById('md-editor') as HTMLTextAreaElement | null;
  const button = document.getElementById('md-to-docx-button') as HTMLButtonElement | null;
  const status = document.getElementById('md-to-docx-status');
  if (!input || !editor || !button || !status) return;

  let filename = 'document.docx';
  const message = (key: string, fallback: string): string =>
    status.dataset[key] ?? fallback;
  input.addEventListener('change', async () => {
    const file = input.files?.[0];
    if (!file) return;
    editor.value = await file.text();
    filename = `${file.name.replace(/\.md$/i, '') || 'document'}.docx`;
    status.textContent = `${message('loaded', 'Loaded')} ${file.name}`;
  });

  button.addEventListener('click', async () => {
    if (!editor.value.trim()) {
      status.textContent = message('empty', 'Paste Markdown or choose a .md file.');
      return;
    }
    button.disabled = true;
    status.textContent = message('creating', 'Creating Word document…');
    try {
      download(await convertMarkdownToBlob(editor.value), filename);
      status.textContent = `${message('ready', 'Ready:')} ${filename}`;
    } catch (error) {
      console.error(error);
      status.textContent = message('error', 'Could not create the Word document.');
    } finally {
      button.disabled = false;
    }
  });
}

if (typeof document !== 'undefined') {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initMarkdownToWord);
  } else {
    initMarkdownToWord();
  }
}
