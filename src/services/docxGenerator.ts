import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeightRule,
  HorizontalPositionAlign,
  HorizontalPositionRelativeFrom,
  ImageRun,
  Packer,
  PageBreak,
  PageNumber,
  PageOrientation,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  VerticalPositionAlign,
  VerticalPositionRelativeFrom,
  WidthType,
} from 'docx';
import { saveAs } from 'file-saver';
import { repository } from '../storage/db';
import {
  ExportRecord,
  PRO_TYPOGRAPHY_CONTRAST_DEFAULTS,
  resolveSheetFooterText,
  resolveSheetHeaderTitle,
  resolveWorkbookPalette,
  WorkbookEntry,
  WorkbookProject,
  WorkbookSettings,
} from '../types/workbook';
import {
  canInlineMultiCharStrokes,
  getPageDimensionsMm,
  getStrokeColumnsForWidth,
  PaginatedSheet,
  paginateWorkbookEntries,
  renderGridGuidelinePngBytes,
} from './layoutPagination';
import { renderStrokeStagePngBytes } from './strokeEngine';

export interface DocxProgress {
  current: number;
  total: number;
  message: string;
}

const MM_TO_TWIP = 56.6929133858;
const MM_TO_PX = 96 / 25.4; // 96dpi logical pixels used by docx ImageRun transformation

function mmToTwip(mm: number): number {
  return Math.round(mm * MM_TO_TWIP);
}

function mmToPx(mm: number): number {
  return Math.round(mm * MM_TO_PX);
}

function cleanHex(hex: string, fallback = '1E1916'): string {
  const stripped = (hex || '').replace('#', '').trim();
  return /^[0-9a-fA-F]{6}$/.test(stripped) ? stripped.toUpperCase() : fallback;
}

function getNoBorders() {
  return {
    top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    insideVertical: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  };
}

function hanziFont(settings: WorkbookSettings) {
  return {
    name: settings.wordFontName,
    eastAsia: settings.wordFontName,
    ascii: settings.wordFontName,
    hAnsi: settings.wordFontName,
  };
}

function arabicFont() {
  return {
    name: 'Arial',
    cs: 'Arial',
    ascii: 'Arial',
    hAnsi: 'Arial',
  };
}

function buildCoverPageBlocks(
  workbook: WorkbookProject,
  entries: WorkbookEntry[],
  sheets: PaginatedSheet[],
  settings: WorkbookSettings
): Array<Paragraph | Table> {
  const palette = resolveWorkbookPalette(settings);
  const accent = cleanHex(palette.accentColor, '6B1D2F');
  const primaryInk = cleanHex(palette.primaryInk, '1E1916');
  const secondaryInk = cleanHex(palette.secondaryInk, '5E544B');
  const subtleInk = cleanHex(palette.subtleInk, '8C8176');
  const ruleHex = palette.isBw ? '000000' : 'CFC4B3';
  const specimenChars = entries
    .slice(0, 6)
    .map((e) => e.text)
    .join('   ');
  const totalStrokes = entries.reduce((sum, e) => sum + e.totalStrokeCount, 0);
  const totalPracticeBoxes = sheets.reduce((sum, s) => sum + s.totalBoxesOnPage, 0);
  const boxesPerCharCounts = Array.from(new Set(sheets.map((s) => s.boxesPerRow)));
  const boxesPerCharLabel =
    boxesPerCharCounts.length === 1
      ? `${boxesPerCharCounts[0]}/char`
      : `${Math.min(...boxesPerCharCounts)}–${Math.max(...boxesPerCharCounts)}/char`;

  return [
    new Paragraph({
      spacing: { before: 1200, after: 180 },
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: 'CHINESE CHARACTER WRITING WORKBOOK',
          font: 'Calibri',
          size: 20,
          bold: true,
          color: accent,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 220 },
      border: {
        top: { style: BorderStyle.SINGLE, size: 12, color: accent, space: 12 },
        bottom: { style: BorderStyle.SINGLE, size: 6, color: ruleHex, space: 12 },
      },
      children: [
        new TextRun({
          text: workbook.title,
          font: 'Georgia',
          size: 62,
          bold: true,
          color: primaryInk,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 140, after: 440 },
      children: [
        new TextRun({
          text: workbook.subtitle || 'Stroke-by-Stroke Guided Calligraphy & Practice Sheets',
          font: 'Georgia',
          italics: true,
          size: 25,
          color: secondaryInk,
        }),
      ],
    }),
    ...(specimenChars
      ? [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 260, after: 480 },
            children: [
              new TextRun({
                text: specimenChars,
                font: hanziFont(settings),
                size: 70,
                color: accent,
              }),
            ],
          }),
        ]
      : []),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 340, after: 120 },
      children: [
        new TextRun({
          text: `${entries.length} Vocabulary ${entries.length === 1 ? 'Entry' : 'Entries'}   ·   ${totalPracticeBoxes} Practice Boxes (${boxesPerCharLabel})   ·   ${totalStrokes} Total Strokes`,
          font: 'Calibri',
          size: 21,
          color: secondaryInk,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 140 },
      children: [
        new TextRun({
          text: `${settings.authorName || 'Hanzi Book Studio Press'}  ·  ${settings.pageSize} ${settings.orientation === 'landscape' ? 'Landscape' : 'Portrait'}  ·  ${new Date().toLocaleDateString()}`,
          font: 'Calibri',
          size: 19,
          color: subtleInk,
        }),
      ],
    }),
    new Paragraph({
      children: [new PageBreak()],
    }),
  ];
}

function buildIndexBlocks(
  entries: WorkbookEntry[],
  sheets: PaginatedSheet[],
  settings: WorkbookSettings,
  contentWidthTwip: number
): Array<Paragraph | Table> {
  const palette = resolveWorkbookPalette(settings);
  const accent = cleanHex(palette.accentColor, '6B1D2F');
  const primaryInk = cleanHex(palette.primaryInk, '1E1916');
  const indexHeaderFill = cleanHex(palette.indexHeaderBg, 'F5EFE4');

  const entryPageMap = new Map<string, number>();
  const entryBoxesPerRowMap = new Map<string, number>();
  for (const sheet of sheets) {
    for (const e of sheet.entries) {
      entryPageMap.set(e.id, sheet.pageNumber);
      entryBoxesPerRowMap.set(e.id, sheet.boxesPerRow);
    }
  }

  const showArabic = settings.showArabicTranslation !== false;
  const colWidths = showArabic
    ? [
        Math.round(contentWidthTwip * 0.05),
        Math.round(contentWidthTwip * 0.12),
        Math.round(contentWidthTwip * 0.14),
        Math.round(contentWidthTwip * 0.29),
        Math.round(contentWidthTwip * 0.19),
        Math.round(contentWidthTwip * 0.07),
        Math.round(contentWidthTwip * 0.07),
        Math.round(contentWidthTwip * 0.07),
      ]
    : [
        Math.round(contentWidthTwip * 0.06),
        Math.round(contentWidthTwip * 0.14),
        Math.round(contentWidthTwip * 0.17),
        Math.round(contentWidthTwip * 0.39),
        Math.round(contentWidthTwip * 0.08),
        Math.round(contentWidthTwip * 0.08),
        Math.round(contentWidthTwip * 0.08),
      ];

  const headerLabels = showArabic
    ? ['#', 'Vocabulary', 'Pinyin', 'English Meaning', 'Arabic (العربية)', 'Boxes', 'Strokes', 'Sheet']
    : ['#', 'Vocabulary', 'Pinyin', 'English Meaning', 'Boxes', 'Strokes', 'Sheet'];

  const arabicColIdx = showArabic ? 4 : -1;
  const numericStartIdx = showArabic ? 5 : 4;

  const headerRow = new TableRow({
    tableHeader: true,
    children: headerLabels.map(
      (label, i) =>
        new TableCell({
          width: { size: colWidths[i], type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: indexHeaderFill, color: 'auto' },
          margins: { top: 70, bottom: 70, left: 90, right: 90 },
          children: [
            new Paragraph({
              alignment:
                i === 0 || i >= numericStartIdx
                  ? AlignmentType.CENTER
                  : i === arabicColIdx
                  ? AlignmentType.RIGHT
                  : AlignmentType.LEFT,
              children: [
                new TextRun({
                  text: label,
                  font: i === arabicColIdx ? arabicFont() : 'Calibri',
                  size: 18,
                  bold: true,
                  color: accent,
                }),
              ],
            }),
          ],
        })
    ),
  });

  const dataRows = entries.map((entry, idx) => {
    const boxes =
      Math.max(1, entry.charResults.length) * (entryBoxesPerRowMap.get(entry.id) ?? 10);
    const sheetNumber = entryPageMap.get(entry.id) ?? 1;
    const cellsData = showArabic
      ? [
          String(idx + 1),
          entry.text,
          entry.pinyin || '—',
          entry.meaning || '—',
          entry.arabicMeaning?.trim() || '—',
          String(boxes),
          String(entry.totalStrokeCount),
          String(sheetNumber),
        ]
      : [
          String(idx + 1),
          entry.text,
          entry.pinyin || '—',
          entry.meaning || '—',
          String(boxes),
          String(entry.totalStrokeCount),
          String(sheetNumber),
        ];

    return new TableRow({
      children: cellsData.map(
        (val, i) =>
          new TableCell({
            width: { size: colWidths[i], type: WidthType.DXA },
            margins: { top: 55, bottom: 55, left: 90, right: 90 },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                bidirectional: i === arabicColIdx,
                alignment:
                  i === 0 || i >= numericStartIdx
                    ? AlignmentType.CENTER
                    : i === arabicColIdx
                    ? AlignmentType.RIGHT
                    : AlignmentType.LEFT,
                children: [
                  new TextRun({
                    text: val,
                    font:
                      i === 1
                        ? hanziFont(settings)
                        : i === arabicColIdx
                        ? arabicFont()
                        : 'Calibri',
                    size: i === 1 ? 24 : 18,
                    bold: i === 1 || i === 2,
                    rightToLeft: i === arabicColIdx && val !== '—',
                    color: i === 2 ? accent : primaryInk,
                  }),
                ],
              }),
            ],
          })
      ),
    });
  });

  return [
    new Paragraph({
      spacing: { before: 100, after: 160 },
      border: {
        bottom: { style: BorderStyle.SINGLE, size: 12, color: accent, space: 6 },
      },
      children: [
        new TextRun({
          text: 'Character Index',
          font: 'Georgia',
          size: 34,
          bold: true,
          color: primaryInk,
        }),
      ],
    }),
    new Table({
      width: { size: contentWidthTwip, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      borders: {
        top: { style: BorderStyle.SINGLE, size: 6, color: accent },
        bottom: {
          style: BorderStyle.SINGLE,
          size: 6,
          color: palette.isBw ? '000000' : 'CFC4B3',
        },
        left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        insideHorizontal: {
          style: BorderStyle.SINGLE,
          size: 4,
          color: palette.isBw ? '999999' : 'E5DDD1',
        },
        insideVertical: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      },
      rows: [headerRow, ...dataRows],
    }),
    new Paragraph({
      children: [new PageBreak()],
    }),
  ];
}

async function buildEntryBandBlocks(
  entry: WorkbookEntry,
  sheet: PaginatedSheet,
  settings: WorkbookSettings,
  contentWidthTwip: number,
  usableWidthMm: number,
  guidelinePngBytes: Uint8Array,
  isFirstOnPage: boolean
): Promise<Array<Paragraph | Table>> {
  const palette = resolveWorkbookPalette(settings);
  const accent = cleanHex(palette.accentColor, '6B1D2F');
  const primaryInk = cleanHex(palette.primaryInk, '1E1916');
  const secondaryInk = cleanHex(palette.secondaryInk, '5E544B');
  const dividerHex = cleanHex(palette.dividerColor, 'E5DDD1');
  const gridBg = cleanHex(palette.gridBgColor, 'F5EFE4');
  const gridBorderHex = cleanHex(palette.gridBorderColor, '9E8D7A');
  const traceColor = cleanHex(palette.traceGlyphColor, 'CFC6B8');
  const isCompact = sheet.density === 'compact';
  const isMultiChar = entry.charResults.length > 1;
  const isPortrait = settings.orientation === 'portrait';

  const leftRatio = isPortrait ? 0.35 : 0.28;
  const leftColTwip = Math.round(contentWidthTwip * leftRatio);
  const rightColTwip = contentWidthTwip - leftColTwip;

  const strokeBreakdown = isMultiChar
    ? `${entry.totalStrokeCount} str (${entry.charResults.map((c) => c.strokeCount).join('+')})`
    : `${entry.totalStrokeCount} ${entry.totalStrokeCount === 1 ? 'stroke' : 'strokes'}`;

  // LEFT CELL: Vocabulary Info (Word + Pinyin + Strokes + English Meaning + Optional Arabic)
  const mainHanziBasePx =
    settings.mainHanziSizePx ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.mainHanziSizePx;
  const metaBasePx = settings.metaFontSizePx ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.metaFontSizePx;
  const multiScale =
    entry.charResults.length >= 4
      ? 0.56
      : entry.charResults.length === 3
      ? 0.68
      : isMultiChar
      ? 0.8
      : 1;
  const heroHalfPt = Math.max(28, Math.round(mainHanziBasePx * 1.45 * multiScale));
  const pinyinHalfPt = Math.max(16, Math.round(metaBasePx * 1.7));
  const englishHalfPt = Math.max(13, Math.round(metaBasePx * 1.32));
  const strokeHalfPt = Math.max(12, Math.round(metaBasePx * 1.2));
  const arabicHalfPt = Math.max(14, Math.round(metaBasePx * 1.42));

  const showArabicRow =
    settings.showArabicTranslation !== false && Boolean(entry.arabicMeaning?.trim());

  const leftParagraphs: Paragraph[] = [
    new Paragraph({
      spacing: { before: 6, after: 12 },
      children: [
        new TextRun({
          text: entry.text,
          font: hanziFont(settings),
          size: heroHalfPt,
          color: primaryInk,
        }),
        new TextRun({
          text: `  ${entry.pinyin || '—'}`,
          font: 'Calibri',
          size: pinyinHalfPt,
          bold: true,
          color: accent,
        }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: showArabicRow ? 8 : 10 },
      children: [
        new TextRun({
          text: `${entry.meaning || 'No English meaning'}  ·  `,
          font: 'Calibri',
          size: englishHalfPt,
          color: secondaryInk,
        }),
        new TextRun({
          text: strokeBreakdown,
          font: 'Calibri',
          size: strokeHalfPt,
          bold: true,
          color: primaryInk,
        }),
      ],
    }),
    ...(showArabicRow
      ? [
          new Paragraph({
            spacing: { before: 0, after: 8 },
            alignment: AlignmentType.LEFT,
            children: [
              new TextRun({
                text: entry.arabicMeaning!.trim(),
                font: arabicFont(),
                size: arabicHalfPt,
                rightToLeft: true,
                color: primaryInk,
              }),
            ],
          }),
        ]
      : []),
  ];

  // RIGHT CELL: Stroke-Order Diagrams directly beside vocabulary info
  const stagePx = Math.max(24, mmToPx(sheet.stageSizeMm));
  const maxCols = getStrokeColumnsForWidth(usableWidthMm, sheet.stageSizeMm, isPortrait);
  const inlineMulti =
    sheet.inlineMultiCharStrokes &&
    canInlineMultiCharStrokes(entry, usableWidthMm, sheet.stageSizeMm, isPortrait);
  const rightChildren: Array<Paragraph | Table> = [];

  if (inlineMulti) {
    const inlineRuns: Array<TextRun | ImageRun> = [];
    for (let cIdx = 0; cIdx < entry.charResults.length; cIdx++) {
      const charRes = entry.charResults[cIdx];
      if (cIdx > 0) {
        inlineRuns.push(new TextRun({ text: '   ', size: 12 }));
      }
      inlineRuns.push(
        new TextRun({
          text: `${charRes.char} `,
          font: hanziFont(settings),
          size: 19,
          bold: true,
          color: '1E1916',
        })
      );

      if (charRes.status === 'ready' && charRes.strokeData) {
        const strokeCount = charRes.strokeData.strokeCount;
        for (let s = 0; s < strokeCount; s++) {
          const pngBytes = await renderStrokeStagePngBytes(charRes.strokeData, s, {
            pixelSize: Math.max(180, stagePx * 4),
            newStrokeColor: palette.newStrokeColor,
            completedStrokeColor: palette.completedStrokeColor,
            borderColor: palette.strokeTileBorderColor,
            guideLineColor: palette.strokeGuideLineColor,
            numberColor: palette.strokeNumberColor,
            backgroundColor: '#FFFFFF',
          });
          inlineRuns.push(
            new ImageRun({
              type: 'png',
              data: pngBytes,
              transformation: { width: stagePx, height: stagePx },
            })
          );
          if (s < strokeCount - 1) {
            inlineRuns.push(new TextRun({ text: ' ', size: 5 }));
          }
        }
      }
    }

    if (inlineRuns.length > 0) {
      rightChildren.push(
        new Paragraph({
          spacing: { before: 10, after: 15 },
          children: inlineRuns,
        })
      );
    }
  } else {
    for (let cIdx = 0; cIdx < entry.charResults.length; cIdx++) {
      const charRes = entry.charResults[cIdx];

      if (charRes.status !== 'ready' || !charRes.strokeData) {
        rightChildren.push(
          new Paragraph({
            spacing: { before: 10, after: 20 },
            children: [
              new TextRun({
                text: `Stroke data unavailable for “${charRes.char}”.`,
                font: 'Calibri',
                italics: true,
                size: 16,
                color: palette.isBw ? '000000' : 'A51D24',
              }),
            ],
          })
        );
        continue;
      }

      const strokeCount = charRes.strokeData.strokeCount;
      const stageRuns: Array<TextRun | ImageRun> = [];

      if (isMultiChar) {
        stageRuns.push(
          new TextRun({
            text: `${charRes.char} `,
            font: hanziFont(settings),
            size: isCompact ? 19 : 24,
            bold: true,
            color: primaryInk,
          })
        );
      }

      for (let s = 0; s < strokeCount; s++) {
        const pngBytes = await renderStrokeStagePngBytes(charRes.strokeData, s, {
          pixelSize: Math.max(180, stagePx * 4),
          newStrokeColor: palette.newStrokeColor,
          completedStrokeColor: palette.completedStrokeColor,
          borderColor: palette.strokeTileBorderColor,
          guideLineColor: palette.strokeGuideLineColor,
          numberColor: palette.strokeNumberColor,
          backgroundColor: '#FFFFFF',
        });

        stageRuns.push(
          new ImageRun({
            type: 'png',
            data: pngBytes,
            transformation: {
              width: stagePx,
              height: stagePx,
            },
          })
        );

        if (s < strokeCount - 1) {
          stageRuns.push(
            new TextRun({
              text: ' ',
              size: 6,
            })
          );
        }

        if ((s + 1) % maxCols === 0 && s < strokeCount - 1) {
          rightChildren.push(
            new Paragraph({
              spacing: { before: 10, after: 20 },
              children: stageRuns.splice(0, stageRuns.length),
            })
          );
        }
      }

      if (stageRuns.length > 0) {
        rightChildren.push(
          new Paragraph({
            spacing: { before: cIdx === 0 ? 10 : 20, after: 15 },
            children: stageRuns,
          })
        );
      }
    }
  }

  if (rightChildren.length === 0) {
    rightChildren.push(new Paragraph({}));
  }

  const upperTable = new Table({
    width: { size: contentWidthTwip, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    borders: {
      top: isFirstOnPage
        ? { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }
        : { style: BorderStyle.SINGLE, size: 4, color: dividerHex },
      bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: dividerHex },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: leftColTwip, type: WidthType.DXA },
            margins: { top: isFirstOnPage ? 15 : 55, bottom: 25, left: 20, right: 110 },
            verticalAlign: VerticalAlign.CENTER,
            children: leftParagraphs,
          }),
          new TableCell({
            width: { size: rightColTwip, type: WidthType.DXA },
            margins: { top: isFirstOnPage ? 15 : 55, bottom: 25, left: 110, right: 20 },
            verticalAlign: VerticalAlign.CENTER,
            children: rightChildren,
          }),
        ],
      }),
    ],
  });

  // LOWER BAND: Independent full-width editable writing table for EACH character in the entry
  const labelColTwip = mmToTwip(7.4);
  const safeContentTwip = Math.max(1000, contentWidthTwip - 16);
  const maxBoxTwip = Math.floor((safeContentTwip - labelColTwip) / Math.max(1, sheet.boxesPerRow));
  const boxSideTwip = Math.min(maxBoxTwip, mmToTwip(sheet.boxSizeMm));
  const boxSidePx = Math.max(32, Math.round(boxSideTwip / 15));
  const guidePx = Math.max(28, boxSidePx - 2);
  const traceFontHalfPt = Math.max(28, Math.round(boxSidePx * 0.68 * 1.45));

  const outerBorderSize =
    settings.gridBorderThickness === 'thin'
      ? 4
      : settings.gridBorderThickness === 'bold'
      ? 12
      : 8;
  const innerBorderSize =
    settings.gridBorderThickness === 'thin'
      ? 4
      : settings.gridBorderThickness === 'bold'
      ? 10
      : 6;

  const blocks: Array<Paragraph | Table> = [upperTable];

  for (let rowIdx = 0; rowIdx < entry.charResults.length; rowIdx++) {
    const ch = entry.charResults[rowIdx].char;

    const labelCell = new TableCell({
      width: { size: labelColTwip, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 0, bottom: 0, left: 0, right: 35 },
      borders: {
        top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        right: { style: BorderStyle.SINGLE, size: outerBorderSize, color: gridBorderHex },
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: ch,
              font: hanziFont(settings),
              size: isCompact ? 21 : 25,
              color: palette.isBw ? '000000' : '5E544B',
            }),
          ],
        }),
      ],
    });

    const boxCells: TableCell[] = Array.from({ length: sheet.boxesPerRow }).map((_, boxIdx) => {
      const isTrace = boxIdx < sheet.traceBoxes;

      const guidelineAnchor = new ImageRun({
        type: 'png',
        data: guidelinePngBytes,
        transformation: {
          width: guidePx,
          height: guidePx,
        },
        floating: {
          behindDocument: true,
          layoutInCell: true,
          allowOverlap: true,
          horizontalPosition: {
            relative: HorizontalPositionRelativeFrom.COLUMN,
            align: HorizontalPositionAlign.CENTER,
          },
          verticalPosition: {
            relative: VerticalPositionRelativeFrom.PARAGRAPH,
            align: VerticalPositionAlign.CENTER,
          },
        },
      });

      const editableRun = new TextRun({
        text: isTrace ? ch : '',
        font: hanziFont(settings),
        size: traceFontHalfPt,
        color: isTrace ? traceColor : primaryInk,
      });

      return new TableCell({
        width: { size: boxSideTwip, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: gridBg, color: 'auto' },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 0 },
            children: [guidelineAnchor, editableRun],
          }),
        ],
      });
    });

    const practiceTable = new Table({
      width: { size: labelColTwip + boxSideTwip * sheet.boxesPerRow, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      borders: {
        top: { style: BorderStyle.SINGLE, size: outerBorderSize, color: gridBorderHex },
        bottom: { style: BorderStyle.SINGLE, size: outerBorderSize, color: gridBorderHex },
        left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        right: { style: BorderStyle.SINGLE, size: outerBorderSize, color: gridBorderHex },
        insideHorizontal: { style: BorderStyle.SINGLE, size: innerBorderSize, color: gridBorderHex },
        insideVertical: { style: BorderStyle.SINGLE, size: innerBorderSize, color: gridBorderHex },
      },
      rows: [
        new TableRow({
          height: { value: boxSideTwip, rule: HeightRule.EXACT },
          children: [labelCell, ...boxCells],
        }),
      ],
    });

    blocks.push(practiceTable);

    blocks.push(
      new Paragraph({
        spacing: {
          before: 0,
          after:
            rowIdx < entry.charResults.length - 1
              ? 35
              : Math.round(sheet.entryGapMm * 13),
        },
        children: [],
      })
    );
  }

  return blocks;
}

export async function generateWorkbookDocx(
  workbook: WorkbookProject,
  selectedEntries?: WorkbookEntry[],
  onProgress?: (progress: DocxProgress) => void
): Promise<ExportRecord> {
  const entries = (
    selectedEntries && selectedEntries.length > 0 ? selectedEntries : workbook.entries
  ).filter((e) => e.charResults.length > 0);

  if (entries.length === 0) {
    throw new Error('No entries selected for Word document export.');
  }

  const { settings } = workbook;
  const palette = resolveWorkbookPalette(settings);
  const dims = getPageDimensionsMm(settings);
  const marginTwip = mmToTwip(settings.pageMarginMm);
  const pageWidthTwip = mmToTwip(dims.widthMm);
  const pageHeightTwip = mmToTwip(dims.heightMm);
  const contentWidthTwip = pageWidthTwip - marginTwip * 2;
  const accent = cleanHex(palette.accentColor, '6B1D2F');
  const subtleInk = cleanHex(palette.subtleInk, '8C8176');
  const secondaryInk = cleanHex(palette.secondaryInk, '5E544B');

  const sheets = paginateWorkbookEntries(entries, settings);
  const guidelinePngBytes = await renderGridGuidelinePngBytes(
    settings.gridStyle,
    palette.gridLineColor,
    180,
    palette.gridDiagonalColor
  );

  const primaryInk = cleanHex(palette.primaryInk, '1E1916');
  const dividerHex = cleanHex(palette.dividerColor, 'E5DDD1');

  const buildStudentInfoBar = (): Table => {
    const nameColTwip = Math.round(contentWidthTwip * 0.65);
    const dateColTwip = contentWidthTwip - nameColTwip;
    return new Table({
      width: { size: contentWidthTwip, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      borders: {
        ...getNoBorders(),
        bottom: { style: BorderStyle.SINGLE, size: 4, color: dividerHex },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: nameColTwip, type: WidthType.DXA },
              margins: { top: 20, bottom: 55, left: 0, right: 120 },
              children: [
                new Paragraph({
                  spacing: { before: 0, after: 0 },
                  children: [
                    new TextRun({
                      text: 'Name: ____________________________________',
                      font: 'Calibri',
                      size: 18,
                      bold: true,
                      color: primaryInk,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              width: { size: dateColTwip, type: WidthType.DXA },
              margins: { top: 20, bottom: 55, left: 120, right: 0 },
              children: [
                new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  spacing: { before: 0, after: 0 },
                  children: [
                    new TextRun({
                      text: 'Date: _______________________',
                      font: 'Calibri',
                      size: 18,
                      bold: true,
                      color: primaryInk,
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    });
  };

  const bodyChildren: Array<Paragraph | Table> = [];

  if (settings.includeCoverPage) {
    onProgress?.({ current: 0, total: entries.length, message: 'Composing cover page…' });
    bodyChildren.push(...buildCoverPageBlocks(workbook, entries, sheets, settings));
  }

  if (settings.includeTableOfContents) {
    onProgress?.({ current: 0, total: entries.length, message: 'Building character index…' });
    bodyChildren.push(...buildIndexBlocks(entries, sheets, settings, contentWidthTwip));
  }

  let processedEntries = 0;

  for (let sIdx = 0; sIdx < sheets.length; sIdx++) {
    const sheet = sheets[sIdx];

    if (settings.showStudentInfoBar) {
      bodyChildren.push(buildStudentInfoBar());
    }

    for (let eIdx = 0; eIdx < sheet.entries.length; eIdx++) {
      const entry = sheet.entries[eIdx];
      processedEntries++;

      onProgress?.({
        current: processedEntries,
        total: entries.length,
        message: `Building sheet ${sheet.pageNumber}/${sheets.length} — “${entry.text}” (${processedEntries}/${entries.length})…`,
      });

      const entryBlocks = await buildEntryBandBlocks(
        entry,
        sheet,
        settings,
        contentWidthTwip,
        dims.usableWidthMm,
        guidelinePngBytes,
        eIdx === 0
      );
      bodyChildren.push(...entryBlocks);

      if (processedEntries % 4 === 0) {
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
    }

    if (sIdx < sheets.length - 1) {
      bodyChildren.push(new Paragraph({ children: [new PageBreak()] }));
    }
  }

  onProgress?.({
    current: entries.length,
    total: entries.length,
    message: 'Packaging editable Microsoft Word (.docx) workbook…',
  });

  const minBoxes = sheets.length > 0 ? Math.min(...sheets.map((s) => s.boxesPerRow)) : 10;
  const maxBoxes = sheets.length > 0 ? Math.max(...sheets.map((s) => s.boxesPerRow)) : 10;
  const traceCount =
    sheets[0]?.traceBoxes ??
    (typeof settings.traceBoxCount === 'number' ? settings.traceBoxCount : 2);
  const headerBoxesSummary =
    minBoxes === maxBoxes
      ? `${minBoxes} Boxes/Char (${traceCount} Trace · ${Math.max(0, minBoxes - traceCount)} Empty)`
      : `${minBoxes}–${maxBoxes} Boxes/Char (${traceCount} Trace)`;
  const headerTitle = resolveSheetHeaderTitle(workbook.title, settings);
  const footerInstruction = resolveSheetFooterText(settings, traceCount, maxBoxes);

  const doc = new Document({
    creator: settings.authorName || 'Hanzi Book Studio',
    title: workbook.title,
    description: workbook.subtitle || 'Chinese Character Writing Workbook',
    sections: [
      {
        properties: {
          page: {
            size: {
              width: pageWidthTwip,
              height: pageHeightTwip,
              orientation:
                settings.orientation === 'landscape'
                  ? PageOrientation.LANDSCAPE
                  : PageOrientation.PORTRAIT,
            },
            margin: {
              top: marginTwip,
              right: marginTwip,
              bottom: marginTwip,
              left: marginTwip,
              header: Math.round(marginTwip * 0.45),
              footer: Math.round(marginTwip * 0.45),
            },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                border: {
                  bottom: { style: BorderStyle.SINGLE, size: 6, color: accent, space: 3 },
                },
                children: [
                  new TextRun({
                    text: `${headerTitle}   ·   ${settings.pageSize} ${settings.orientation === 'landscape' ? 'Landscape' : 'Portrait'}   ·   ${headerBoxesSummary}`,
                    font: 'Calibri',
                    size: 15,
                    color: subtleInk,
                  }),
                ],
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: `${footerInstruction}  ·  Page `,
                    font: 'Calibri',
                    size: 15,
                    color: subtleInk,
                  }),
                  new TextRun({
                    children: [PageNumber.CURRENT],
                    font: 'Calibri',
                    size: 15,
                    color: secondaryInk,
                  }),
                  new TextRun({
                    text: ' of ',
                    font: 'Calibri',
                    size: 15,
                    color: subtleInk,
                  }),
                  new TextRun({
                    children: [PageNumber.TOTAL_PAGES],
                    font: 'Calibri',
                    size: 15,
                    color: secondaryInk,
                  }),
                ],
              }),
            ],
          }),
        },
        children: bodyChildren,
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const safeSlug =
    workbook.title
      .toLowerCase()
      .replace(/[^a-z0-9\u4e00-\u9fff]+/gi, '-')
      .replace(/^-|-$/g, '') || 'hanzi-workbook';
  const fileName = `${safeSlug}-${settings.pageSize.toLowerCase()}-${settings.orientation}${
    palette.isBw ? '-bw' : ''
  }.docx`;

  saveAs(blob, fileName);

  const exportRecord: ExportRecord = {
    id: `exp-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
    workbookId: workbook.id,
    workbookTitle: workbook.title,
    format: 'docx',
    fileName,
    entryCount: entries.length,
    pageSize: settings.pageSize,
    orientation: settings.orientation,
    exportedAt: Date.now(),
  };

  await repository.saveExport(exportRecord);
  return exportRecord;
}
