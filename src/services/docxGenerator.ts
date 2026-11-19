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
  WorkbookEntry,
  WorkbookProject,
  WorkbookSettings,
} from '../types/workbook';
import {
  getPageDimensionsMm,
  getStrokeColumnsForWidth,
  MANDATORY_EMPTY_BOXES,
  MANDATORY_TOTAL_BOXES,
  MANDATORY_TRACE_BOXES,
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

function buildCoverPageBlocks(
  workbook: WorkbookProject,
  entries: WorkbookEntry[],
  settings: WorkbookSettings
): Array<Paragraph | Table> {
  const accent = cleanHex(settings.accentColor, '6B1D2F');
  const specimenChars = entries
    .slice(0, 6)
    .map((e) => e.text)
    .join('   ');
  const totalStrokes = entries.reduce((sum, e) => sum + e.totalStrokeCount, 0);
  const totalPracticeBoxes = entries.reduce(
    (sum, e) => sum + Math.max(1, e.charResults.length) * MANDATORY_TOTAL_BOXES,
    0
  );

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
        bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CFC4B3', space: 12 },
      },
      children: [
        new TextRun({
          text: workbook.title,
          font: 'Georgia',
          size: 62,
          bold: true,
          color: '1E1916',
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
          color: '5E544B',
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
          text: `${entries.length} Vocabulary ${entries.length === 1 ? 'Entry' : 'Entries'}   ·   ${totalPracticeBoxes} Practice Boxes (10/char)   ·   ${totalStrokes} Total Strokes`,
          font: 'Calibri',
          size: 21,
          color: '5E544B',
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
          color: '8C8176',
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
  const accent = cleanHex(settings.accentColor, '6B1D2F');

  const entryPageMap = new Map<string, number>();
  for (const sheet of sheets) {
    for (const e of sheet.entries) {
      entryPageMap.set(e.id, sheet.pageNumber);
    }
  }

  const colWidths = [
    Math.round(contentWidthTwip * 0.06),
    Math.round(contentWidthTwip * 0.14),
    Math.round(contentWidthTwip * 0.17),
    Math.round(contentWidthTwip * 0.39),
    Math.round(contentWidthTwip * 0.08),
    Math.round(contentWidthTwip * 0.08),
    Math.round(contentWidthTwip * 0.08),
  ];

  const headerLabels = ['#', 'Vocabulary', 'Pinyin', 'English Meaning', 'Boxes', 'Strokes', 'Sheet'];
  const headerRow = new TableRow({
    tableHeader: true,
    children: headerLabels.map(
      (label, i) =>
        new TableCell({
          width: { size: colWidths[i], type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: 'F5EFE4', color: 'auto' },
          margins: { top: 70, bottom: 70, left: 90, right: 90 },
          children: [
            new Paragraph({
              alignment: i === 0 || i >= 4 ? AlignmentType.CENTER : AlignmentType.LEFT,
              children: [
                new TextRun({
                  text: label,
                  font: 'Calibri',
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
    const boxes = Math.max(1, entry.charResults.length) * MANDATORY_TOTAL_BOXES;
    const sheetNumber = entryPageMap.get(entry.id) ?? 1;
    const cellsData = [
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
                alignment: i === 0 || i >= 4 ? AlignmentType.CENTER : AlignmentType.LEFT,
                children: [
                  new TextRun({
                    text: val,
                    font: i === 1 ? hanziFont(settings) : 'Calibri',
                    size: i === 1 ? 24 : 18,
                    bold: i === 1 || i === 2,
                    color: i === 2 ? accent : '1E1916',
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
          color: '1E1916',
        }),
      ],
    }),
    new Table({
      width: { size: contentWidthTwip, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      borders: {
        top: { style: BorderStyle.SINGLE, size: 6, color: accent },
        bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CFC4B3' },
        left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: 'E5DDD1' },
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
  const accent = cleanHex(settings.accentColor, '6B1D2F');
  const gridBg = cleanHex(settings.gridBgColor, 'F5EFE4');
  const traceColor = cleanHex(settings.traceGlyphColor, 'CFC6B8');
  const isCompact = sheet.density === 'compact';
  const isMultiChar = entry.charResults.length > 1;
  const isPortrait = settings.orientation === 'portrait';

  const leftRatio = isPortrait ? 0.34 : 0.27;
  const leftColTwip = Math.round(contentWidthTwip * leftRatio);
  const rightColTwip = contentWidthTwip - leftColTwip;

  const strokeBreakdown = isMultiChar
    ? `${entry.totalStrokeCount} str (${entry.charResults.map((c) => c.strokeCount).join('+')})`
    : `${entry.totalStrokeCount} ${entry.totalStrokeCount === 1 ? 'stroke' : 'strokes'}`;

  // LEFT CELL: Vocabulary Info (Word + Pinyin + Strokes + English Meaning)
  const heroHalfPt = isCompact ? (isMultiChar ? 38 : 48) : isMultiChar ? 48 : 64;
  const leftParagraphs: Paragraph[] = [
    new Paragraph({
      spacing: { before: 10, after: 20 },
      children: [
        new TextRun({
          text: entry.text,
          font: hanziFont(settings),
          size: heroHalfPt,
          color: '1E1916',
        }),
        new TextRun({
          text: `  ${entry.pinyin || '—'}`,
          font: 'Calibri',
          size: isCompact ? 22 : 26,
          bold: true,
          color: accent,
        }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: 10 },
      children: [
        new TextRun({
          text: `${entry.meaning || 'No English meaning'}  ·  `,
          font: 'Calibri',
          size: isCompact ? 16 : 18,
          color: '5E544B',
        }),
        new TextRun({
          text: strokeBreakdown,
          font: 'Calibri',
          size: isCompact ? 15 : 17,
          bold: true,
          color: '1E1916',
        }),
      ],
    }),
  ];

  // RIGHT CELL: Stroke-Order Diagrams directly beside vocabulary info
  const stagePx = Math.max(26, mmToPx(sheet.stageSizeMm));
  const maxCols = getStrokeColumnsForWidth(usableWidthMm, sheet.stageSizeMm, isPortrait);
  const rightChildren: Array<Paragraph | Table> = [];

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
              color: 'A51D24',
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
          size: isCompact ? 20 : 24,
          color: '1E1916',
        })
      );
    }

    for (let s = 0; s < strokeCount; s++) {
      const pngBytes = await renderStrokeStagePngBytes(charRes.strokeData, s, {
        pixelSize: Math.max(120, stagePx * 3),
        newStrokeColor: settings.newStrokeColor,
        completedStrokeColor: settings.completedStrokeColor,
        borderColor: '#DDD3C5',
        backgroundColor: '#FDFBF8',
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

      // Tiny horizontal spacer between stroke stage tiles
      if (s < strokeCount - 1) {
        stageRuns.push(
          new TextRun({
            text: ' ',
            size: 6,
          })
        );
      }

      // Wrap gracefully if character exceeds maxCols per line
      if ((s + 1) % maxCols === 0 && s < strokeCount - 1) {
        rightChildren.push(
          new Paragraph({
            spacing: { before: 10, after: 25 },
            children: stageRuns.splice(0, stageRuns.length),
          })
        );
      }
    }

    if (stageRuns.length > 0) {
      rightChildren.push(
        new Paragraph({
          spacing: { before: cIdx === 0 ? 10 : 25, after: 20 },
          children: stageRuns,
        })
      );
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
        : { style: BorderStyle.SINGLE, size: 4, color: 'E5DDD1' },
      bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: 'E5DDD1' },
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

  // LOWER BAND: Independent 10-box editable writing table for EACH character in the entry
  // Boxes 1–2: Light-gray editable Chinese character tracing
  // Boxes 3–10: Empty editable handwriting boxes with traditional Chinese guidelines behind text
  const labelColTwip = mmToTwip(7.5);
  const maxBoxTwip = Math.floor((contentWidthTwip - labelColTwip) / MANDATORY_TOTAL_BOXES);
  const boxSideTwip = Math.min(maxBoxTwip, mmToTwip(sheet.boxSizeMm));
  const boxSidePx = Math.max(36, Math.round(boxSideTwip / 15));
  const guidePx = Math.max(32, boxSidePx - 2);
  const traceFontHalfPt = Math.max(32, Math.round(boxSidePx * 0.68 * 1.45));

  const blocks: Array<Paragraph | Table> = [upperTable];

  for (let rowIdx = 0; rowIdx < entry.charResults.length; rowIdx++) {
    const ch = entry.charResults[rowIdx].char;

    const labelCell = new TableCell({
      width: { size: labelColTwip, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 0, bottom: 0, left: 0, right: 40 },
      borders: {
        top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        right: { style: BorderStyle.SINGLE, size: 8, color: '9E8D7A' },
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: ch,
              font: hanziFont(settings),
              size: isCompact ? 22 : 26,
              color: '5E544B',
            }),
          ],
        }),
      ],
    });

    const boxCells: TableCell[] = Array.from({ length: MANDATORY_TOTAL_BOXES }).map((_, boxIdx) => {
      const isTrace = boxIdx < MANDATORY_TRACE_BOXES;

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
        color: isTrace ? traceColor : '1E1916',
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
      width: { size: labelColTwip + boxSideTwip * MANDATORY_TOTAL_BOXES, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      borders: {
        top: { style: BorderStyle.SINGLE, size: 8, color: '9E8D7A' },
        bottom: { style: BorderStyle.SINGLE, size: 8, color: '9E8D7A' },
        left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        right: { style: BorderStyle.SINGLE, size: 8, color: '9E8D7A' },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 6, color: '9E8D7A' },
        insideVertical: { style: BorderStyle.SINGLE, size: 6, color: '9E8D7A' },
      },
      rows: [
        new TableRow({
          height: { value: boxSideTwip, rule: HeightRule.EXACT },
          children: [labelCell, ...boxCells],
        }),
      ],
    });

    blocks.push(practiceTable);

    // Compact vertical spacer between rows of a multi-character word or after the entry
    blocks.push(
      new Paragraph({
        spacing: {
          before: 0,
          after:
            rowIdx < entry.charResults.length - 1
              ? 35
              : Math.round(sheet.entryGapMm * 14),
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
  const dims = getPageDimensionsMm(settings);
  const marginTwip = mmToTwip(settings.pageMarginMm);
  const pageWidthTwip = mmToTwip(dims.widthMm);
  const pageHeightTwip = mmToTwip(dims.heightMm);
  const contentWidthTwip = pageWidthTwip - marginTwip * 2;
  const accent = cleanHex(settings.accentColor, '6B1D2F');

  const sheets = paginateWorkbookEntries(entries, settings);
  const guidelinePngBytes = await renderGridGuidelinePngBytes(
    settings.gridStyle,
    settings.gridLineColor,
    180
  );

  const bodyChildren: Array<Paragraph | Table> = [];

  if (settings.includeCoverPage) {
    onProgress?.({ current: 0, total: entries.length, message: 'Composing cover page…' });
    bodyChildren.push(...buildCoverPageBlocks(workbook, entries, settings));
  }

  if (settings.includeTableOfContents) {
    onProgress?.({ current: 0, total: entries.length, message: 'Building character index…' });
    bodyChildren.push(...buildIndexBlocks(entries, sheets, settings, contentWidthTwip));
  }

  let processedEntries = 0;

  for (let sIdx = 0; sIdx < sheets.length; sIdx++) {
    const sheet = sheets[sIdx];

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
                    text: `${workbook.title}   ·   ${settings.pageSize} ${settings.orientation === 'landscape' ? 'Landscape' : 'Portrait'}   ·   10 Boxes/Char (${MANDATORY_TRACE_BOXES} Trace · ${MANDATORY_EMPTY_BOXES} Empty)`,
                    font: 'Calibri',
                    size: 15,
                    color: '8C8176',
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
                    text: `${settings.authorName || 'Hanzi Book Studio'}  ·  Trace boxes 1–2, write from memory in boxes 3–10  ·  Page `,
                    font: 'Calibri',
                    size: 15,
                    color: '8C8176',
                  }),
                  new TextRun({
                    children: [PageNumber.CURRENT],
                    font: 'Calibri',
                    size: 15,
                    color: '5E544B',
                  }),
                  new TextRun({
                    text: ' of ',
                    font: 'Calibri',
                    size: 15,
                    color: '8C8176',
                  }),
                  new TextRun({
                    children: [PageNumber.TOTAL_PAGES],
                    font: 'Calibri',
                    size: 15,
                    color: '5E544B',
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
  const fileName = `${safeSlug}-${settings.pageSize.toLowerCase()}-${settings.orientation}.docx`;

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
