import { GridStyle, WorkbookEntry, WorkbookSettings } from '../types/workbook';

/**
 * Practice box baseline specification:
 * - Minimum 10 boxes per individual Chinese character in every vocabulary entry.
 * - Boxes 1–2: Light-gray character tracing.
 * - Remaining boxes (3..N): Empty handwriting boxes with traditional Chinese guidelines.
 * - In 'maximize' mode (default), every character row expands edge-to-edge across the
 *   printable page width (e.g. 17–20 boxes/row on A4 Landscape) to maximize total
 *   practice boxes per page with zero wasted horizontal whitespace.
 */
export const MANDATORY_MIN_BOXES = 10;
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
  boxesPerRow: number;
  traceBoxes: number;
  emptyBoxes: number;
  totalBoxesOnPage: number;
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
  const studentBarReserveMm = settings.showStudentInfoBar ? 7.5 : 0;
  const headerFooterReserveMm = 15 + studentBarReserveMm;

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
  const infoRatio = isPortrait ? 0.35 : 0.28;
  const strokeAreaMm = usableWidthMm * (1 - infoRatio) - 8;
  const tileStepMm = stageSizeMm + 1.1;
  return Math.max(6, Math.min(22, Math.floor(strokeAreaMm / tileStepMm)));
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
  return totalStrokes + entry.charResults.length <= maxCols;
}

function getDensityTargets(settings: WorkbookSettings, density: PageDensity) {
  const boxPref = settings.boxDensity ?? 'standard';
  if (boxPref === 'compact') {
    return {
      boxMm: density === 'spacious' ? 14.2 : 12.2,
      stageMm: density === 'spacious' ? 10.2 : 8.2,
      gapMm: 0.9,
    };
  }
  if (boxPref === 'large') {
    return {
      boxMm: density === 'compact' ? 15.2 : 17.5,
      stageMm: density === 'compact' ? 9.4 : 11.4,
      gapMm: 1.3,
    };
  }
  // 'standard' — crisp readable squares that pack 17–18 boxes across A4 Landscape
  return {
    boxMm: density === 'compact' ? 13.8 : density === 'balanced' ? 15.6 : 17.2,
    stageMm: density === 'compact' ? 8.8 : density === 'balanced' ? 10.4 : 12.0,
    gapMm: 1.0,
  };
}

function measureEntryHeightMm(
  entry: WorkbookEntry,
  usableWidthMm: number,
  stageSizeMm: number,
  boxSizeMm: number,
  isPortrait: boolean,
  allowInlineMulti: boolean,
  settings?: WorkbookSettings
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

  const mainHanziPx = settings?.mainHanziSizePx ?? 34;
  const metaPx = settings?.metaFontSizePx ?? 13;
  const hasArabic =
    settings?.showArabicTranslation !== false && Boolean(entry.arabicMeaning?.trim());
  const hanziHeightMm = mainHanziPx / 3.3;
  const metaStackMm = (metaPx * (hasArabic ? 3.55 : 2.55)) / 3.3;
  const baseInfoHeightMm = stageSizeMm < 10 ? 9.4 : 11.6;
  const minInfoHeightMm = Math.max(baseInfoHeightMm, hanziHeightMm, metaStackMm);

  const strokesHeightMm = strokeRowsTotal * stageSizeMm + Math.max(0, strokeRowsTotal - 1) * 1.0;
  const upperBandMm = Math.max(minInfoHeightMm, strokesHeightMm);

  const multiRowGapMm = 1.3;
  const upperToPracticeGapMm = 1.6;
  const practiceBandMm =
    upperToPracticeGapMm + charCount * boxSizeMm + Math.max(0, charCount - 1) * multiRowGapMm;

  return upperBandMm + practiceBandMm;
}

export function estimateEntryHeightMm(
  entry: WorkbookEntry,
  dims: PageDimensionsMm,
  settings: WorkbookSettings,
  isPortrait: boolean
): number {
  const compactMode = settings.charactersPerPage >= 4 || settings.boxDensity === 'compact';
  const targets = getDensityTargets(settings, compactMode ? 'compact' : 'balanced');
  const minDividerGapMm = compactMode ? 2.4 : 4.2;

  return (
    measureEntryHeightMm(
      entry,
      dims.usableWidthMm,
      targets.stageMm,
      targets.boxMm,
      isPortrait,
      compactMode,
      settings
    ) + minDividerGapMm
  );
}

/**
 * Shared auto-pagination & full-width box maximization engine used identically by
 * Live Preview, editable Microsoft Word (.docx), and PDF exports.
 */
export function paginateWorkbookEntries(
  entries: WorkbookEntry[],
  settings: WorkbookSettings
): PaginatedSheet[] {
  const validEntries = entries.filter((e) => e.charResults.length > 0);
  if (validEntries.length === 0) return [];

  const dims = getPageDimensionsMm(settings);
  const isPortrait = settings.orientation === 'portrait';
  const maxEntriesPerPage = Math.max(1, Math.min(7, settings.charactersPerPage));
  const fillMode = settings.boxFillMode ?? 'maximize';
  const rawTraceCount =
    typeof settings.traceBoxCount === 'number' ? settings.traceBoxCount : MANDATORY_TRACE_BOXES;
  const traceBoxes = Math.max(0, Math.min(4, Math.round(rawTraceCount)));

  const pagesRaw: WorkbookEntry[][] = [];
  let currentEntries: WorkbookEntry[] = [];
  let currentHeightMm = 0;

  for (const entry of validEntries) {
    const neededMm = estimateEntryHeightMm(entry, dims, settings, isPortrait);

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
    const targets = getDensityTargets(settings, density);

    // Reserve exact label width + safe right-edge margin so boxes never clip horizontally
    const labelReserveMm = 8.2;
    const rightSafetyMm = 1.4;
    const boxGapMm = targets.gapMm;
    const availBoxStripMm = Math.max(60, dims.usableWidthMm - labelReserveMm - rightSafetyMm);

    let stageSizeMm = targets.stageMm;
    let boxSizeMm = targets.boxMm;
    let entryGapMm = density === 'compact' ? 2.8 : density === 'balanced' ? 4.4 : 6.4;

    // Ensure vertical fit strictly inside usableHeightMm - 4mm safety margin
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
            inlineMultiCharStrokes,
            settings
          ),
        0
      ) +
      Math.max(0, pageEntries.length - 1) * entryGapMm;

    let totalHeight = computeTotalPageHeight();
    if (totalHeight > maxAllowedBodyMm) {
      const ratio = maxAllowedBodyMm / totalHeight;
      boxSizeMm = Math.max(10.4, Number((boxSizeMm * ratio).toFixed(2)));
      stageSizeMm = Math.max(7.2, Number((stageSizeMm * ratio).toFixed(2)));
      entryGapMm = Math.max(1.8, Number((entryGapMm * ratio).toFixed(2)));
      totalHeight = computeTotalPageHeight();

      while (totalHeight > maxAllowedBodyMm && boxSizeMm > 9.4) {
        boxSizeMm = Number((boxSizeMm - 0.35).toFixed(2));
        totalHeight = computeTotalPageHeight();
      }
    }

    // Compute horizontal box count across the full printable width without right-edge overflow
    let boxesPerRow: number;
    if (fillMode === 'fixed10') {
      boxesPerRow = MANDATORY_MIN_BOXES;
      const maxBoxFor10 =
        (availBoxStripMm - boxGapMm * (boxesPerRow - 1)) / boxesPerRow;
      boxSizeMm = Math.min(boxSizeMm, Math.floor(maxBoxFor10 * 100) / 100);
    } else {
      const maxFit = Math.floor((availBoxStripMm + boxGapMm) / (boxSizeMm + boxGapMm));
      boxesPerRow = Math.max(MANDATORY_MIN_BOXES, maxFit);

      const exactFitBoxMm = (availBoxStripMm - boxGapMm * (boxesPerRow - 1)) / boxesPerRow;
      if (exactFitBoxMm < boxSizeMm) {
        boxSizeMm = Math.floor(exactFitBoxMm * 100) / 100;
      } else {
        const candidateBoxMm = Math.floor(Math.min(exactFitBoxMm, boxSizeMm + 0.5) * 100) / 100;
        const prevBox = boxSizeMm;
        boxSizeMm = candidateBoxMm;
        if (computeTotalPageHeight() > maxAllowedBodyMm) {
          boxSizeMm = prevBox;
        }
      }
    }

    const emptyBoxes = Math.max(0, boxesPerRow - traceBoxes);

    return {
      pageNumber: idx + 1,
      entries: pageEntries,
      totalPracticeRows,
      boxesPerRow,
      traceBoxes,
      emptyBoxes,
      totalBoxesOnPage: totalPracticeRows * boxesPerRow,
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
 * Microsoft Word table cells so all boxes stay 100% editable.
 */
export async function renderGridGuidelinePngBytes(
  gridStyle: GridStyle,
  gridLineColor: string,
  pixelSize = 180,
  gridDiagonalColor?: string
): Promise<Uint8Array> {
  const canvas = document.createElement('canvas');
  canvas.width = pixelSize;
  canvas.height = pixelSize;
  const ctx = canvas.getContext('2d');

  if (ctx) {
    ctx.clearRect(0, 0, pixelSize, pixelSize);
    const dash = Math.max(4, Math.round(pixelSize * 0.035));

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

    if (gridStyle === 'mi') {
      ctx.strokeStyle = gridDiagonalColor || gridLineColor || '#C9BCA9';
      ctx.globalAlpha = gridDiagonalColor ? 1 : 0.65;
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
