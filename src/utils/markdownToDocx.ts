import {
  Document,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  PageBreak,
  Packer,
  AlignmentType,
  Header,
  Footer,
  Math
} from 'docx';
import { marked } from 'marked';
import { DocxAdapter } from '@marciorvneto/texpipe';
import * as docx from 'docx';

// ── Adapter for LaTeX → docx math objects ────────────────────
const adapter = new DocxAdapter(docx);

// ── LaTeX math preprocessor ───────────────────────────────────
// Replace $...$ / $$...$$ with unique text markers that marked
// treats as plain text.  The markers are later resolved by
// resolveTokens() and rendered as native Word equations.
const MM = '⍂';

function preprocessMath(src: string): string {
  src = src.replace(/\$\$([\s\S]+?)\$\$/g, (_, content) =>
    `${MM}BM${MM}${content.trim()}${MM}/BM${MM}`
  );
  src = src.replace(/\$([^$\n]+?)\$/g, (_, content) =>
    `${MM}IM${MM}${content.trim()}${MM}/IM${MM}`
  );
  return src;
}

// ── Resolve inline-math markers in a text string ──────────────
const IM_RE = new RegExp(`${MM}IM${MM}(.*?)${MM}/IM${MM}`, 'g');

function resolveInlineText(text: string): { type: string; text: string }[] {
  if (!text.includes(MM)) return [{ type: 'text', text }];
  const parts: { type: string; text: string }[] = [];
  let last = 0, m: RegExpExecArray | null;
  IM_RE.lastIndex = 0;
  while ((m = IM_RE.exec(text)) !== null) {
    if (m.index > last) parts.push({ type: 'text', text: text.slice(last, m.index) });
    parts.push({ type: 'math', text: m[1] });
    last = m.index + m[0].length;
  }
  if (last < text.length) parts.push({ type: 'text', text: text.slice(last) });
  return parts;
}

const BM_RE = new RegExp(`^${MM}BM${MM}(.*?)${MM}/BM${MM}$`);

function resolveTokens(tokens: any[]): any[] {
  const out: any[] = [];
  for (const tok of tokens) {
    if (tok.type === 'paragraph') {
      const raw = tok.text || '';
      const bm = raw.match(BM_RE);
      if (bm) {
        out.push({ type: 'mathBlock', text: bm[1] });
        continue;
      }
    }
    if (tok.tokens) {
      const flat: any[] = [];
      for (const child of tok.tokens) {
        if (child.type === 'text') {
          flat.push(...resolveInlineText(child.text).map(p =>
            p.type === 'math' ? { type: 'math', text: p.text, tokens: [] } : p
          ));
        } else {
          flat.push(child);
        }
      }
      tok.tokens = flat;
    }
    out.push(tok);
  }
  return out;
}

function cleanMarkers(text: string): string {
  return text.replace(new RegExp(`${MM}(?:IM|BM)${MM}.*?${MM}/(?:IM|BM)${MM}`, 'g'), '$1');
}

// Supported Style / Theme Config
export interface DocxStyleConfig {
  fontBody: string;
  fontHeading: string;
  primaryColor: string;
  accentColor: string;
  fontSizeBody: number;
  fontSizeH1: number;
  fontSizeH2: number;
  fontSizeH3: number;
  spacingLineHeight: number;
  marginSize: 'normal' | 'narrow' | 'wide';
  orientation: 'portrait' | 'landscape';
  includeToc: boolean;
  pageNumbers: boolean;
  headerText: string;
  footerText: string;
  titlePage: boolean;
  titlePageTitle: string;
  titlePageSubtitle: string;
  titlePageAuthor: string;
  titlePageOrg: string;
  titlePageDate: string;
}

export const THEME_PRESETS: Record<string, {
  name: string;
  primaryColor: string;
  accentColor: string;
  fontHeading: string;
  fontBody: string;
}> = {
  slate: {
    name: 'Slate Professional',
    primaryColor: '1E293B',
    accentColor: '0284C7',
    fontHeading: 'Arial',
    fontBody: 'Calibri'
  },
  classic: {
    name: 'Classic Editorial',
    primaryColor: '111827',
    accentColor: '7C2D12',
    fontHeading: 'Georgia',
    fontBody: 'Times New Roman'
  },
  teal: {
    name: 'Teal Modern',
    primaryColor: '0F766E',
    accentColor: '0D9488',
    fontHeading: 'Trebuchet MS',
    fontBody: 'Arial'
  },
  warm: {
    name: 'Creative Auburn',
    primaryColor: '78350F',
    accentColor: '9A3412',
    fontHeading: 'Garamond',
    fontBody: 'Garamond'
  },
  monochrome: {
    name: 'Stealth Mono',
    primaryColor: '000000',
    accentColor: '4B5563',
    fontHeading: 'Courier New',
    fontBody: 'Courier New'
  }
};

function parseInlineRuns(tokens: any[], config: DocxStyleConfig, overrides: any = {}): (TextRun | Math)[] {
  const runs: (TextRun | Math)[] = [];
  if (!tokens) return [];

  for (const token of tokens) {
    switch (token.type) {
      case 'strong':
        runs.push(...parseInlineRuns(token.tokens, config, { ...overrides, bold: true }));
        break;
      case 'em':
        runs.push(...parseInlineRuns(token.tokens, config, { ...overrides, italics: true }));
        break;
      case 'del':
        runs.push(...parseInlineRuns(token.tokens, config, { ...overrides, strike: true }));
        break;
      case 'paragraph':
        runs.push(...parseInlineRuns(token.tokens, config, overrides));
        break;
      case 'codespan':
        runs.push(new TextRun({
          text: token.text,
          font: 'Courier New',
          size: (config.fontSizeBody - 1) * 2,
          color: '9D174D',
          shading: { fill: 'F3F4F6' },
          ...overrides
        }));
        break;
      case 'link':
        runs.push(new TextRun({
          text: `${token.text}`,
          font: config.fontBody,
          size: config.fontSizeBody * 2,
          color: config.accentColor,
          underline: {},
          ...overrides
        }));
        runs.push(new TextRun({
          text: ` (${token.href})`,
          font: 'Courier New',
          size: (config.fontSizeBody - 2) * 2,
          color: '6B7280',
          ...overrides
        }));
        break;
      case 'br':
        runs.push(new TextRun({ text: '\n' }));
        break;
      case 'math':
        try {
          const mr = adapter.toMathRun(token.text);
          runs.push(new Math({ children: [mr] }));
        } catch {
          runs.push(new TextRun({
            text: token.text,
            font: 'Cambria Math',
            size: config.fontSizeBody * 2,
            italics: true,
            color: '7C3AED',
            shading: { fill: 'F5F3FF' },
            ...overrides
          }));
        }
        break;
      case 'text':
      default:
        runs.push(new TextRun({
          text: token.text,
          font: config.fontBody,
          size: config.fontSizeBody * 2,
          color: '1F2937',
          ...overrides
        }));
        break;
    }
  }

  return runs;
}

export async function convertMarkdownToDocx(markdown: string, config: DocxStyleConfig): Promise<Blob> {
  const processed = preprocessMath(markdown);
  const rawTokens = marked.lexer(processed);
  const tokens = resolveTokens(rawTokens);

  const children: any[] = [];

  // 1. Cover / Title Page
  if (config.titlePage) {
    children.push(new Paragraph({ spacing: { before: 2880 } }));

    children.push(new Paragraph({
      border: {
        bottom: { color: config.primaryColor, style: BorderStyle.SINGLE, size: 24 }
      },
      spacing: { after: 240 }
    }));

    children.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [
        new TextRun({
          text: config.titlePageTitle || 'Untitled Document',
          font: config.fontHeading,
          size: (config.fontSizeH1 + 8) * 2,
          bold: true,
          color: config.primaryColor
        })
      ],
      spacing: { after: 120 }
    }));

    if (config.titlePageSubtitle) {
      children.push(new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [
          new TextRun({
            text: config.titlePageSubtitle,
            font: config.fontBody,
            size: (config.fontSizeBody + 3) * 2,
            italics: true,
            color: '4B5563'
          })
        ],
        spacing: { after: 2880 }
      }));
    } else {
      children.push(new Paragraph({ spacing: { before: 2160 } }));
    }

    if (config.titlePageAuthor || config.titlePageOrg || config.titlePageDate) {
      if (config.titlePageAuthor) {
        children.push(new Paragraph({
          children: [
            new TextRun({ text: 'PREPARED BY: ', font: config.fontBody, size: 18, bold: true, color: '6B7280' }),
            new TextRun({ text: config.titlePageAuthor, font: config.fontBody, size: 20, bold: true, color: '1F2937' })
          ],
          spacing: { after: 60 }
        }));
      }
      if (config.titlePageOrg) {
        children.push(new Paragraph({
          children: [
            new TextRun({ text: 'ORGANIZATION: ', font: config.fontBody, size: 18, bold: true, color: '6B7280' }),
            new TextRun({ text: config.titlePageOrg, font: config.fontBody, size: 20, color: '374151' })
          ],
          spacing: { after: 60 }
        }));
      }
      if (config.titlePageDate) {
        children.push(new Paragraph({
          children: [
            new TextRun({ text: 'DATE: ', font: config.fontBody, size: 18, bold: true, color: '6B7280' }),
            new TextRun({ text: config.titlePageDate, font: config.fontBody, size: 20, color: '374151' })
          ],
          spacing: { after: 60 }
        }));
      }
    }

    children.push(new Paragraph({
      border: {
        top: { color: config.accentColor, style: BorderStyle.SINGLE, size: 8 }
      },
      spacing: { before: 480 }
    }));

    children.push(new Paragraph({ children: [new PageBreak()] }));
  }

  // 2. TOC
  if (config.includeToc) {
    children.push(new Paragraph({
      children: [
        new TextRun({
          text: 'Table of Contents',
          font: config.fontHeading,
          size: config.fontSizeH1 * 2,
          bold: true,
          color: config.primaryColor
        })
      ],
      spacing: { before: 240, after: 180 },
      keepNext: true
    }));

    children.push(new Paragraph({
      children: [
        new TextRun({
          text: 'Note: To update the interactive Table of Contents in Microsoft Word, open this document, right-click here, and choose "Update Field" (or press F9).',
          font: config.fontBody,
          size: (config.fontSizeBody - 1.5) * 2,
          italics: true,
          color: '6B7280'
        })
      ],
      spacing: { after: 240 }
    }));

    const headings = tokens.filter(t => t.type === 'heading' && t.depth <= 2);
    if (headings.length > 0) {
      headings.forEach((h: any) => {
        const hText = cleanMarkers(h.text || '');
        const levelIndent = h.depth === 2 ? 360 : 0;
        children.push(new Paragraph({
          indent: { left: levelIndent + 240, hanging: 240 },
          children: [
            new TextRun({
              text: h.depth === 1 ? '■  ' : '•  ',
              font: config.fontHeading,
              size: config.fontSizeBody * 2,
              color: h.depth === 1 ? config.primaryColor : config.accentColor,
              bold: h.depth === 1
            }),
            new TextRun({
              text: hText,
              font: config.fontBody,
              size: config.fontSizeBody * 2,
              bold: h.depth === 1,
              color: h.depth === 1 ? '111827' : '4B5563'
            }),
            new TextRun({
              text: ' . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .',
              font: 'Courier New',
              size: (config.fontSizeBody - 2) * 2,
              color: 'D1D5DB'
            })
          ],
          spacing: { after: 60 }
        }));
      });
      children.push(new Paragraph({ spacing: { after: 360 } }));
    } else {
      children.push(new Paragraph({
        children: [
          new TextRun({
            text: '[TOC is empty: Add headings in your markdown to generate entries here]',
            font: config.fontBody,
            size: config.fontSizeBody * 2,
            italics: true,
            color: '9CA3AF'
          })
        ],
        spacing: { after: 360 }
      }));
    }

    children.push(new Paragraph({
      border: {
        bottom: { color: 'E5E7EB', style: BorderStyle.SINGLE, size: 4 }
      },
      spacing: { after: 360 }
    }));
  }

  // 3. Process all elements
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];

    switch (token.type) {
      case 'heading': {
        const depth = token.depth;
        let pSize = config.fontSizeBody * 2;
        let beforeSpace = 180;
        let afterSpace = 100;
        let underLineStyle = {};

        if (depth === 1) {
          pSize = config.fontSizeH1 * 2;
          beforeSpace = 360;
          afterSpace = 120;
          underLineStyle = { color: config.accentColor };
        } else if (depth === 2) {
          pSize = config.fontSizeH2 * 2;
          beforeSpace = 240;
          afterSpace = 90;
        } else if (depth === 3) {
          pSize = config.fontSizeH3 * 2;
          beforeSpace = 180;
          afterSpace = 75;
        } else {
          pSize = (config.fontSizeBody + 2) * 2;
          beforeSpace = 140;
          afterSpace = 60;
        }

        children.push(new Paragraph({
          keepNext: true,
          spacing: { before: beforeSpace, after: afterSpace },
          children: [
            new TextRun({
              text: cleanMarkers(token.text),
              font: config.fontHeading,
              size: pSize,
              bold: true,
              color: config.primaryColor,
              ...underLineStyle
            })
          ]
        }));
        break;
      }

      case 'paragraph': {
        const inlineRuns = parseInlineRuns(token.tokens || [{ type: 'text', text: token.text }], config);
        children.push(new Paragraph({
          children: inlineRuns,
          spacing: {
            line: Math.round(config.spacingLineHeight * 240),
            after: 110
          }
        }));
        break;
      }

      case 'blockquote': {
        const qTokens = token.tokens || [{ type: 'text', text: token.text }];
        children.push(new Paragraph({
          indent: { left: 720 },
          border: {
            left: {
              color: config.accentColor,
              style: BorderStyle.SINGLE,
              size: 24,
              space: 12
            }
          },
          children: parseInlineRuns(qTokens, config, { italics: true, color: '4B5563' }),
          spacing: { before: 120, after: 120, line: Math.round(config.spacingLineHeight * 240) }
        }));
        break;
      }

      case 'code': {
        const codeLines = token.text.split('\n');
        children.push(new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: 'E5E7EB' },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E5E7EB' },
            left: { style: BorderStyle.SINGLE, size: 4, color: 'E5E7EB' },
            right: { style: BorderStyle.SINGLE, size: 4, color: 'E5E7EB' },
          },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  shading: { fill: 'F9FAFB' },
                  margins: { top: 120, bottom: 120, left: 160, right: 160 },
                  children: codeLines.map((line: string) => (
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: line,
                          font: 'Courier New',
                          size: 19,
                          color: '2563EB'
                        })
                      ],
                      spacing: { before: 0, after: 0 }
                    })
                  ))
                })
              ]
            })
          ]
        }));
        children.push(new Paragraph({ spacing: { after: 120 } }));
        break;
      }

      case 'list': {
        const isOrdered = token.ordered;
        token.items.forEach((item: any, idx: number) => {
          const reparsed = marked.lexer(preprocessMath(item.text || ''));
          const resolved = resolveTokens(reparsed);
          const inlineTokens = resolved.length > 0 && resolved[0].type === 'paragraph' && resolved[0].tokens
            ? resolved[0].tokens
            : [{ type: 'text', text: item.text || '' }];
          const runs = parseInlineRuns(inlineTokens, config);

          if (isOrdered) {
            children.push(new Paragraph({
              indent: { left: 432, hanging: 288 },
              children: [
                new TextRun({
                  text: `${idx + 1}.\t`,
                  font: config.fontBody,
                  size: config.fontSizeBody * 2,
                  bold: true,
                  color: config.primaryColor
                }),
                ...runs
              ],
              spacing: { after: 60 }
            }));
          } else {
            children.push(new Paragraph({
              children: runs,
              bullet: { level: 0 },
              spacing: { after: 60 }
            }));
          }
        });
        children.push(new Paragraph({ spacing: { after: 80 } }));
        break;
      }

      case 'table': {
        children.push(new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 6, color: 'CBD5E1' },
            bottom: { style: BorderStyle.SINGLE, size: 8, color: '94A3B8' },
            left: { style: BorderStyle.SINGLE, size: 4, color: 'E2E8F0' },
            right: { style: BorderStyle.SINGLE, size: 4, color: 'E2E8F0' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: 'F1F5F9' },
            insideVertical: { style: BorderStyle.SINGLE, size: 4, color: 'F1F5F9' }
          },
          rows: [
            new TableRow({
              tableHeader: true,
              children: token.header.map((col: any) => (
                new TableCell({
                  shading: { fill: config.primaryColor },
                  margins: { top: 110, bottom: 110, left: 140, right: 140 },
                  children: [
                    new Paragraph({
                      children: parseInlineRuns(col.tokens, config, { bold: true, color: 'FFFFFF' }),
                      spacing: { before: 0, after: 0 }
                    })
                  ]
                })
              ))
            }),
            ...token.rows.map((row: any[], rIdx: number) => (
              new TableRow({
                children: row.map((cell: any) => (
                  new TableCell({
                    shading: rIdx % 2 === 1 ? { fill: 'F8FAFC' } : undefined,
                    margins: { top: 110, bottom: 110, left: 140, right: 140 },
                    children: [
                      new Paragraph({
                        children: parseInlineRuns(cell.tokens, config),
                        spacing: { before: 0, after: 0 }
                      })
                    ]
                  })
                ))
              })
            ))
          ]
        }));
        children.push(new Paragraph({ spacing: { after: 140 } }));
        break;
      }

      case 'hr': {
        children.push(new Paragraph({
          border: {
            bottom: { color: 'D1D5DB', style: BorderStyle.SINGLE, size: 8 }
          },
          spacing: { before: 180, after: 180 }
        }));
        break;
      }

      case 'mathBlock': {
        try {
          const para = adapter.toParagraph(token.text);
          children.push(para);
        } catch {
          children.push(new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 4, color: 'DDD6FE' },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: 'DDD6FE' },
              left: { style: BorderStyle.SINGLE, size: 4, color: 'DDD6FE' },
              right: { style: BorderStyle.SINGLE, size: 4, color: 'DDD6FE' },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    shading: { fill: 'F5F3FF' },
                    margins: { top: 120, bottom: 120, left: 160, right: 160 },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({
                            text: token.text,
                            font: 'Cambria Math',
                            size: 22,
                            italics: true,
                            color: '5B21B6',
                          })
                        ],
                        spacing: { before: 0, after: 0 }
                      })
                    ]
                  })
                ]
              })
            ]
          }));
          children.push(new Paragraph({ spacing: { after: 120 } }));
        }
        break;
      }

      default:
        break;
    }
  }

  // 4. Margins
  let docMargins = { top: 1440, bottom: 1440, left: 1440, right: 1440 };
  if (config.marginSize === 'narrow') {
    docMargins = { top: 720, bottom: 720, left: 720, right: 720 };
  } else if (config.marginSize === 'wide') {
    docMargins = { top: 2160, bottom: 2160, left: 2160, right: 2160 };
  }

  // 5. Build document
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: docMargins,
            size: { orientation: config.orientation }
          }
        },
        headers: config.headerText ? {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: config.headerText,
                    font: config.fontBody,
                    size: 16,
                    color: '9CA3AF'
                  })
                ],
                spacing: { after: 120 }
              })
            ]
          })
        } : undefined,
        footers: (config.footerText || config.pageNumbers) ? {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  ...(config.footerText ? [
                    new TextRun({ text: config.footerText, font: config.fontBody, size: 16, color: '9CA3AF' })
                  ] : []),
                  ...(config.footerText && config.pageNumbers ? [
                    new TextRun({ text: '  |  ', font: config.fontBody, size: 16, color: 'D1D5DB' })
                  ] : []),
                  ...(config.pageNumbers ? [
                    new TextRun({ text: 'Page ', font: config.fontBody, size: 16, color: '9CA3AF' }),
                  ] : [])
                ],
                spacing: { before: 120 }
              })
            ]
          })
        } : undefined,
        children: children
      }
    ]
  });

  return await Packer.toBlob(doc);
}
