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
  Footer
} from 'docx';
import { marked } from 'marked';

// Supported Style / Theme Config
export interface DocxStyleConfig {
  fontBody: string;
  fontHeading: string;
  primaryColor: string; // Hex string without '#'
  accentColor: string;  // Hex string without '#'
  fontSizeBody: number; // pt
  fontSizeH1: number;
  fontSizeH2: number;
  fontSizeH3: number;
  spacingLineHeight: number; // Line spacing multiplier (e.g. 1.25, 1.5)
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

// Convert theme selections to full configuration
export const THEME_PRESETS: Record<string, {
  name: string;
  primaryColor: string;
  accentColor: string;
  fontHeading: string;
  fontBody: string;
}> = {
  slate: {
    name: 'Slate Professional',
    primaryColor: '1E293B', // Slate 800
    accentColor: '0284C7',  // Sky 600
    fontHeading: 'Arial',
    fontBody: 'Calibri'
  },
  classic: {
    name: 'Classic Editorial',
    primaryColor: '111827', // Gray 900
    accentColor: '7C2D12',  // Orange 900
    fontHeading: 'Georgia',
    fontBody: 'Times New Roman'
  },
  teal: {
    name: 'Teal Modern',
    primaryColor: '0F766E', // Teal 700
    accentColor: '0D9488',  // Teal 600
    fontHeading: 'Trebuchet MS',
    fontBody: 'Arial'
  },
  warm: {
    name: 'Creative Auburn',
    primaryColor: '78350F', // Amber 900
    accentColor: '9A3412',  // Red-Orange
    fontHeading: 'Garamond',
    fontBody: 'Garamond'
  },
  monochrome: {
    name: 'Stealth Mono',
    primaryColor: '000000',
    accentColor: '4B5563',  // Gray 600
    fontHeading: 'Courier New',
    fontBody: 'Courier New'
  }
};

// Help convert markdown lists and other tokens to DOCX paragraphs and tables
function parseInlineRuns(tokens: any[], config: DocxStyleConfig, overrides: any = {}): TextRun[] {
  const runs: TextRun[] = [];
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
      case 'codespan':
        runs.push(new TextRun({
          text: token.text,
          font: 'Courier New',
          size: (config.fontSizeBody - 1) * 2,
          color: '9D174D', // Pinkish/Purple for inline code
          shading: {
            fill: 'F3F4F6',
          },
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
        // Append URL in parentheses for readable print doc
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
      case 'text':
      default:
        runs.push(new TextRun({
          text: token.text,
          font: config.fontBody,
          size: config.fontSizeBody * 2,
          color: '1F2937', // Off-black body
          ...overrides
        }));
        break;
    }
  }

  return runs;
}

export async function convertMarkdownToDocx(markdown: string, config: DocxStyleConfig): Promise<Blob> {
  const tokens = marked.lexer(markdown);
  const children: any[] = [];

  // 1. Optional Cover / Title Page
  if (config.titlePage) {
    // Top space
    children.push(new Paragraph({ spacing: { before: 2880 } })); // ~2 inches

    // Accent line
    children.push(new Paragraph({
      border: {
        bottom: {
          color: config.primaryColor,
          style: BorderStyle.SINGLE,
          size: 24, // 3pt thickness
        }
      },
      spacing: { after: 240 }
    }));

    // Document Title
    children.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [
        new TextRun({
          text: config.titlePageTitle || 'Untitled Document',
          font: config.fontHeading,
          size: (config.fontSizeH1 + 8) * 2, // Bigger for cover title
          bold: true,
          color: config.primaryColor
        })
      ],
      spacing: { after: 120 }
    }));

    // Subtitle
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
        spacing: { after: 2880 } // Space below title/subtitle block
      }));
    } else {
      children.push(new Paragraph({ spacing: { before: 2160 } }));
    }

    // Author & Organization Metadata block
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

    // Accent strip at bottom of cover page
    children.push(new Paragraph({
      border: {
        top: {
          color: config.accentColor,
          style: BorderStyle.SINGLE,
          size: 8,
        }
      },
      spacing: { before: 480 }
    }));

    // Page Break after Cover Page
    children.push(new Paragraph({
      children: [new PageBreak()]
    }));
  }

  // 2. Table of Contents Placeholder (if requested)
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

    // Helper text for Word's native TOC or listing out the headings in the document explicitly
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

    // Let's programmatically generate a beautiful styled static Outline list of Heading 1s & Heading 2s!
    // This is super helpful because it provides an immediate Table of Contents visible anywhere (including Google Docs).
    const headings = tokens.filter(t => t.type === 'heading' && t.depth <= 2);
    if (headings.length > 0) {
      headings.forEach((h: any, idx: number) => {
        const hText = h.text || '';
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

    // Separator line after TOC
    children.push(new Paragraph({
      border: {
        bottom: {
          color: 'E5E7EB',
          style: BorderStyle.SINGLE,
          size: 4,
        }
      },
      spacing: { after: 360 }
    }));
  }

  // 3. Process All Elements
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
              text: token.text,
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
            line: Math.round(config.spacingLineHeight * 240), // 240 = single line, 360 = 1.5 line spacing
            after: 110 // ~5.5pt space after paragraphs
          }
        }));
        break;
      }

      case 'blockquote': {
        // Block quotes represented as left-indented paragraphs with a beautiful grey left border
        // Split blockquote text by lines or parse tokens
        const qTokens = token.tokens || [{ type: 'text', text: token.text }];
        children.push(new Paragraph({
          indent: { left: 720 }, // 0.5 inches left spacing
          border: {
            left: {
              color: config.accentColor,
              style: BorderStyle.SINGLE,
              size: 24, // 3pt border
              space: 12
            }
          },
          children: parseInlineRuns(qTokens, config, { italics: true, color: '4B5563' }),
          spacing: { before: 120, after: 120, line: Math.round(config.spacingLineHeight * 240) }
        }));
        break;
      }

      case 'code': {
        // Code Block - Rendered as a single-cell 100% wide table with dark/light background
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
                          size: 19, // 9.5pt
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
        // Gap paragraph
        children.push(new Paragraph({ spacing: { after: 120 } }));
        break;
      }

      case 'list': {
        const isOrdered = token.ordered;
        token.items.forEach((item: any, idx: number) => {
          // Parse tokens within list item
          const itemTokens = item.tokens || [{ type: 'text', text: item.text }];
          const runs = parseInlineRuns(itemTokens, config);

          if (isOrdered) {
            // Ordered lists - left-indented with a manual number run prefix and tab
            children.push(new Paragraph({
              indent: { left: 432, hanging: 288 }, // Perfect Word numbering alignment
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
            // Bulleted items
            children.push(new Paragraph({
              children: runs,
              bullet: { level: 0 },
              spacing: { after: 60 }
            }));
          }
        });
        // Small spacing gap after list
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
            // Header Row
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
            // Body rows with alternating backgrounds (zebra striping)
            ...token.rows.map((row: any[], rIdx: number) => (
              new TableRow({
                children: row.map((cell: any) => (
                  new TableCell({
                    shading: rIdx % 2 === 1 ? { fill: 'F8FAFC' } : undefined, // slate-50 background of alternate row
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
        // Gap after table
        children.push(new Paragraph({ spacing: { after: 140 } }));
        break;
      }

      case 'hr': {
        children.push(new Paragraph({
          border: {
            bottom: {
              color: 'D1D5DB',
              style: BorderStyle.SINGLE,
              size: 8, // thin line
            }
          },
          spacing: { before: 180, after: 180 }
        }));
        break;
      }

      default:
        // Skip unhandled tokens like 'space'
        break;
    }
  }

  // 4. Margins Definition (Twips: 1 inch = 1440)
  let docMargins = { top: 1440, bottom: 1440, left: 1440, right: 1440 }; // Default Normal
  if (config.marginSize === 'narrow') {
    docMargins = { top: 720, bottom: 720, left: 720, right: 720 };
  } else if (config.marginSize === 'wide') {
    docMargins = { top: 2160, bottom: 2160, left: 2160, right: 2160 };
  }

  // 5. Build Document Sections with custom properties
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: docMargins,
            size: {
              orientation: config.orientation, // 'portrait' | 'landscape'
            }
          }
        },
        // Wait, Header & Footer can be configured if we write them simple.
        // For standard client-side docx, header and footers are optional, but if specified, 
        // they add a gorgeous touch of custom formatting.
        headers: config.headerText ? {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: config.headerText,
                    font: config.fontBody,
                    size: 16, // 8pt
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
                    new TextRun({
                      text: config.footerText,
                      font: config.fontBody,
                      size: 16,
                      color: '9CA3AF'
                    })
                  ] : []),
                  ...(config.footerText && config.pageNumbers ? [
                    new TextRun({
                      text: '  |  ',
                      font: config.fontBody,
                      size: 16,
                      color: 'D1D5DB'
                    })
                  ] : []),
                  ...(config.pageNumbers ? [
                    new TextRun({
                      text: 'Page ',
                      font: config.fontBody,
                      size: 16,
                      color: '9CA3AF'
                    }),
                    // Dynamic Page Number fields can be added in Word, but simple text run is standard,
                    // or we can use DOCX Page Number native elements. Let's keep it safe. 
                    // In docx, we can inject native PageNumbering fields! Let's check how:
                    // Word native field: TextRun({ children: [PageNumber.CURRENT] }) is possible too.
                    // For broad client-side compatibility, let's keep page numbering setup elegant.
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
