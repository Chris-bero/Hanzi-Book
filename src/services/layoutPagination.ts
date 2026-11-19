import { GridStyle, WorkbookEntry, WorkbookSettings } from '../types/workbook';

/**
 * Mandatory workbook practice box specification:
 * Every individual Chinese character in every vocabulary entry MUST receive
 * an independent row of exactly 10 writing practice boxes:
 * - Boxes 1–2: Light-gray character tracing
 * - Boxes 3–10: Empty handwriting boxes with traditional Chinese guidelines
 */
export const MANDATORY_TOTAL_BOXES = 10;
export const MANDATORY_TRACE_BOXES = 2;
export const MANDATORY_EMPTY_BOXES = 8;
export const DEFAULT_TARGET_ENTRIES_PER_PAGE = 5;

export type PageDensity = 'spacious' | 'balanced' | 'compact';

export interface PageDimensionsMm {
  widthMm: number;
  heightMm: number;
  usableWidthMm: number;
  usableHeightMm: number;
}

export interface PaginatedSheet {
  pageNumber: number;
  entries: WorkbookEntry[];
  totalPracticeRows: number;
  density: PageDensity;
  boxSizeMm: number;
  boxGapMm: number;
  stageSizeMm: number;
  entryGapMm: number;
  inlineMultiCharStrokes: boolean;
}

export function getPageDimensionsMm(settings: WorkbookSettings): PageDimensionsMm {
  const base = settings.pageSize === 'A3' ? { w: 297, h: 420 } : { w: 210, h: 297 };
  const widthMm = settings.orientation === 'landscape' ? base.h : base.w;
  const heightMm = settings.orientation === 'landscape' ? base.w : base.h;
  const margin = Math.max(8, Math.min(24, settings.pageMarginMm));
  // Reserve 16mm total for running header, footer, and clean safety padding
  const headerFooterReserveMm = 16;

  return {
    widthMm,
    heightMm,
    usableWidthMm: widthMm - margin * 2,
    usableHeightMm: heightMm - margin * 2 - headerFooterReserveMm,
  };
}

/**
 * Computes the maximum number of stroke-order tiles that fit horizontally in the
 * right-hand stroke-order area beside the vocabulary info column.
 */
export function getStrokeColumnsForWidth(
  usableWidthMm: number,
  stageSizeMm: number,
  isPortrait: boolean
): number {
  const infoRatio = isPortrait ? 0.34 : 0.27;
  const strokeAreaMm = usableWidthMm * (1 - infoRatio) - 8;
  const tileStepMm = stageSizeMm + 1.1;
  return Math.max(6, Math.min(20, Math.floor(strokeAreaMm / tileStepMm)));
}

/**
 * Determines whether a multi-character entry's stroke diagrams (e.g. 你 7 + 好 6 = 13)
 * can sit side-by-side on a single clean horizontal line without wrapping.
 */
export function canInlineMultiCharStrokes(
  entry: WorkbookEntry,
  usableWidthMm: number,
  stageSizeMm: number,
  isPortrait: boolean
): boolean {
  if (entry.charResults.length <= 1) return false;
  const maxCols = getStrokeColumnsForWidth(usableWidthMm, stageSizeMm, isPortrait);
  const totalStrokes = entry.charResults.reduce((sum, c) => sum + Math.max(1, c.strokeCount || 1), 0);
  // Account for ~1 tile width per character label/separator
  return totalStrokes + entry.charResults.length <= maxCols;
}

function measureEntryHeightMm(
  entry: WorkbookEntry,
  usableWidthMm: number,
  stageSizeMm: number,
  boxSizeMm: number,
  isPortrait: boolean,
  allowInlineMulti: boolean
): number {
  const charCount = Math.max(1, entry.charResults.length);
  const maxCols = getStrokeColumnsForWidth(usableWidthMm, stageSizeMm, isPortrait);

  let strokeRowsTotal = 0;
  if (allowInlineMulti && canInlineMultiCharStrokes(entry, usableWidthMm, stageSizeMm, isPortrait)) {
    strokeRowsTotal = 1;
  } else {
    for (const ch of entry.charResults) {
      const strokes = Math.max(1, ch.strokeCount || 1);
      strokeRowsTotal += Math.ceil(strokes / maxCols);
    }
  }

  // Upper band height: at least enough for the character + pinyin + meaning, or all stroke rows
  const minInfoHeightMm = stageSizeMm < 10 ? 9.6 : 12.2;
  const strokesHeightMm = strokeRowsTotal * stageSizeMm + Math.max(0, strokeRowsTotal - 1) * 1.1;
  const upperBandMm = Math.max(minInfoHeightMm, strokesHeightMm);

  // Lower band height: 1 independent row of 10 practice boxes per character in the entry
  const multiRowGapMm = 1.4;
  const upperToPracticeGapMm = 1.8;
  const practiceBandMm =
    upperToPracticeGapMm + charCount * boxSizeMm + Math.max(0, charCount - 1) * multiRowGapMm;

  return upperBandMm + practiceBandMm;
}

/**
 * Estimates minimum readable height (mm) required for a vocabulary entry so
 * auto-pagination never crowds or clips entries at the bottom of a page.
 */
export function estimateEntryHeightMm(
  entry: WorkbookEntry,
  dims: PageDimensionsMm,
  targetPerPage: number,
  isPortrait: boolean
): number {
  const compactMode = targetPerPage >= 4;
  const minStageMm = compactMode ? 8.4 : 11.0;
  const minBoxMm = compactMode ? 12.6 : 16.0;
  const minDividerGapMm = compactMode ? 2.6 : 4.5;

  return (
    measureEntryHeightMm(entry, dims.usableWidthMm, minStageMm, minBoxMm, isPortrait, compactMode) +
    minDividerGapMm
  );
}

/**
 * Shared auto-pagination & adaptive sizing engine used identically by the Live
 * Preview, editable Microsoft Word (.docx) generator, and PDF generator.
 *
 * Guarantees that every page's entries + 10-box practice rows fit strictly inside
 * usableHeightMm with zero bottom clipping while targeting 5 entries per A4 Landscape page.
 */
export function paginateWorkbookEntries(
  entries: WorkbookEntry[],
  settings: WorkbookSettings
): PaginatedSheet[] {
  const validEntries = entries.filter((e) => e.charResults.length > 0);
  if (validEntries.length === 0) return [];

  const dims = getPageDimensionsMm(settings);
  const isPortrait = settings.orientation === 'portrait';
  const maxEntriesPerPage = Math.max(1, Math.min(6, settings.charactersPerPage));

  const pagesRaw: WorkbookEntry[][] = [];
  let currentEntries: WorkbookEntry[] = [];
  let currentHeightMm = 0;

  for (const entry of validEntries) {
    const neededMm = estimateEntryHeightMm(entry, dims, maxEntriesPerPage, isPortrait);

    const wouldExceedHeight =
      currentEntries.length > 0 && currentHeightMm + neededMm > dims.usableHeightMm;
    const wouldExceedMaxCount = currentEntries.length >= maxEntriesPerPage;

    if (wouldExceedHeight || wouldExceedMaxCount) {
      pagesRaw.push(currentEntries);
      currentEntries = [entry];
      currentHeightMm = neededMm;
    } else {
      currentEntries.push(entry);
      currentHeightMm += neededMm;
    }
  }

  if (currentEntries.length > 0) {
    pagesRaw.push(currentEntries);
  }

  return pagesRaw.map((pageEntries, idx) => {
    const totalPracticeRows = pageEntries.reduce(
      (sum, entry) => sum + Math.max(1, entry.charResults.length),
      0
    );

    const density: PageDensity =
      totalPracticeRows >= 5 || pageEntries.length >= 4
        ? 'compact'
        : totalPracticeRows >= 3 || pageEntries.length >= 2
        ? 'balanced'
        : 'spacious';

    const inlineMultiCharStrokes = density === 'compact';

    // Maximum box side that fits 10 equal squares + left character label + gaps horizontally
    const labelReserveMm = 9.5;
    const desiredBoxGapMm = isPortrait ? 1.4 : 2.4;
    const maxBoxWidthMm =
      (dims.usableWidthMm - labelReserveMm - desiredBoxGapMm * (MANDATORY_TOTAL_BOXES - 1)) /
      MANDATORY_TOTAL_BOXES;

    let stageSizeMm =
      density === 'compact' ? 9.0 : density === 'balanced' ? 11.4 : 13.4;
    let boxSizeMm = Math.min(
      maxBoxWidthMm,
      density === 'compact' ? 14.4 : density === 'balanced' ? 18.0 : 21.5
    );
    let entryGapMm = density === 'compact' ? 3.2 : density === 'balanced' ? 5.0 : 7.0;

    // Exact iterative fit check: guarantee total page height <= usableHeightMm - 4mm buffer
    const maxAllowedBodyMm = dims.usableHeightMm - 4;
    const computeTotalPageHeight = () =>
      pageEntries.reduce(
        (sum, e) =>
          sum +
          measureEntryHeightMm(
            e,
            dims.usableWidthMm,
            stageSizeMm,
            boxSizeMm,
            isPortrait,
            inlineMultiCharStrokes
          ),
        0
      ) +
      Math.max(0, pageEntries.length - 1) * entryGapMm;

    let totalHeight = computeTotalPageHeight();
    if (totalHeight > maxAllowedBodyMm) {
      // Scale boxSizeMm, stageSizeMm, and entryGapMm proportionally so nothing ever clips
      const ratio = maxAllowedBodyMm / totalHeight;
      boxSizeMm = Math.max(10.5, Number((boxSizeMm * ratio).toFixed(2)));
      stageSizeMm = Math.max(7.2, Number((stageSizeMm * ratio).toFixed(2)));
      entryGapMm = Math.max(1.8, Number((entryGapMm * ratio).toFixed(2)));
      totalHeight = computeTotalPageHeight();

      // Final safeguard pass if rounding slightly exceeded
      while (totalHeight > maxAllowedBodyMm && boxSizeMm > 9.5) {
        boxSizeMm = Number((boxSizeMm - 0.4).toFixed(2));
        totalHeight = computeTotalPageHeight();
      }
    }

    // Spread horizontal gap slightly on wide landscape pages so the 10 boxes feel balanced
    const totalBoxesWidth =
      labelReserveMm + boxSizeMm * MANDATORY_TOTAL_BOXES + desiredBoxGapMm * (MANDATORY_TOTAL_BOXES - 1);
    const boxGapMm =
      !isPortrait && totalBoxesWidth < dims.usableWidthMm - 35
        ? Math.min(4.2, desiredBoxGapMm + 1.2)
        : desiredBoxGapMm;

    return {
      pageNumber: idx + 1,
      entries: pageEntries,
      totalPracticeRows,
      density,
      boxSizeMm,
      boxGapMm,
      stageSizeMm,
      entryGapMm,
      inlineMultiCharStrokes,
    };
  });
}

/**
 * Generates a transparent HiDPI PNG of traditional Chinese writing guidelines
 * (Mi-Zi-Ge, Tian-Zi-Ge, or Cross) to anchor behind text inside editable
 * Microsoft Word table cells so all 10 boxes stay 100% editable.
 */
export async function renderGridGuidelinePngBytes(
  gridStyle: GridStyle,
  gridLineColor: string,
  pixelSize = 180
): Promise<Uint8Array> {
  const canvas = document.createElement('canvas');
  canvas.width = pixelSize;
  canvas.height = pixelSize;
  const ctx = canvas.getContext('2d');

  if (ctx) {
    ctx.clearRect(0, 0, pixelSize, pixelSize);
    const dash = Math.max(4, Math.round(pixelSize * 0.035));

    // Horizontal and vertical center lines
    ctx.save();
    ctx.strokeStyle = gridLineColor || '#C9BCA9';
    ctx.lineWidth = Math.max(1.2, pixelSize * 0.01);
    if (gridStyle !== 'cross') {
      ctx.setLineDash([dash, dash]);
    }
    ctx.beginPath();
    ctx.moveTo(pixelSize / 2, 0);
    ctx.lineTo(pixelSize / 2, pixelSize);
    ctx.moveTo(0, pixelSize / 2);
    ctx.lineTo(pixelSize, pixelSize / 2);
    ctx.stroke();

    // Diagonal Mi-Zi-Ge lines
    if (gridStyle === 'mi') {
      ctx.globalAlpha = 0.65;
      ctx.lineWidth = Math.max(1, pixelSize * 0.008);
      ctx.beginPath();
      ctx.moveTo(0, 0);
      ctx.lineTo(pixelSize, pixelSize);
      ctx.moveTo(pixelSize, 0);
      ctx.lineTo(0, pixelSize);
      ctx.stroke();
    }
    ctx.restore();
  }

  const blob = await new Promise<Blob | null>((resolve) =>
    canvas.toBlob((b) => resolve(b), 'image/png')
  );
  if (!blob) {
    throw new Error('Failed to generate writing grid guideline graphic.');
  }
  return new Uint8Array(await blob.arrayBuffer());
}
