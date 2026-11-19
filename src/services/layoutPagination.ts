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
  stageSizeMm: number;
  entryGapMm: number;
}

export function getPageDimensionsMm(settings: WorkbookSettings): PageDimensionsMm {
  const base = settings.pageSize === 'A3' ? { w: 297, h: 420 } : { w: 210, h: 297 };
  const widthMm = settings.orientation === 'landscape' ? base.h : base.w;
  const heightMm = settings.orientation === 'landscape' ? base.w : base.h;
  const margin = Math.max(8, Math.min(24, settings.pageMarginMm));
  const headerFooterReserveMm = 14;

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
  const strokeAreaMm = usableWidthMm * (1 - infoRatio) - 6;
  const tileStepMm = stageSizeMm + 1.2;
  return Math.max(6, Math.min(16, Math.floor(strokeAreaMm / tileStepMm)));
}

/**
 * Estimates minimum readable height (mm) required for a single vocabulary entry
 * including its vocabulary header, stroke-order diagrams beside it, and 10-box
 * practice rows for each individual character in the entry.
 */
export function estimateEntryHeightMm(
  entry: WorkbookEntry,
  dims: PageDimensionsMm,
  targetPerPage: number,
  isPortrait: boolean
): number {
  const charCount = Math.max(1, entry.charResults.length);
  const compactMode = targetPerPage >= 4;

  const minStageMm = compactMode ? 8.8 : 11.8;
  const cols = getStrokeColumnsForWidth(dims.usableWidthMm, minStageMm, isPortrait);

  let strokeRowsTotal = 0;
  for (const ch of entry.charResults) {
    const strokes = Math.max(1, ch.strokeCount || 1);
    strokeRowsTotal += Math.ceil(strokes / cols);
  }

  // Upper band: vocabulary info on left + stroke-order rows on right
  const upperMinMm = compactMode
    ? Math.max(10.2, strokeRowsTotal * (minStageMm + 1.0) + (charCount > 1 ? 1.5 : 0))
    : Math.max(17.0, strokeRowsTotal * (minStageMm + 1.6) + (charCount > 1 ? 3.0 : 0));

  // Lower band: 1 independent row of 10 practice boxes per character in the entry
  const minBoxMm = compactMode ? 13.4 : 17.2;
  const practiceBlockMm = charCount * minBoxMm + Math.max(0, charCount - 1) * 1.4 + 1.8;

  // Vertical padding & entry divider breathing room
  const verticalChromeMm = compactMode ? 2.4 : 5.5;

  return upperMinMm + practiceBlockMm + verticalChromeMm;
}

/**
 * Shared auto-pagination engine used identically by the Live Preview, editable
 * Microsoft Word (.docx) generator, and PDF generator.
 *
 * Maintains the target number of entries per page (default 5 on A4 Landscape)
 * whenever readable, and automatically wraps to the next page when multi-character
 * words or high-stroke characters require extra vertical height.
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

  // Compute adaptive spacing, box size, and stroke stage size per page
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

    // Maximum box side that still allows 10 equal boxes + optional left character label
    const maxBoxWidthMm = (dims.usableWidthMm - 18) / MANDATORY_TOTAL_BOXES - 1.2;

    const targetBoxMm =
      density === 'compact'
        ? totalPracticeRows >= 6
          ? 13.8
          : 15.2
        : density === 'balanced'
        ? 18.0
        : Math.min(24.5, settings.gridSizePx * 0.34);

    const boxSizeMm = Math.max(11.5, Math.min(maxBoxWidthMm, targetBoxMm));

    const stageSizeMm =
      density === 'compact'
        ? totalPracticeRows >= 6
          ? 9.2
          : 10.4
        : density === 'balanced'
        ? 12.2
        : Math.max(12.5, Math.min(16.5, settings.strokeStageSizePx * 0.28));

    const entryGapMm = density === 'compact' ? 2.4 : density === 'balanced' ? 4.5 : 7.0;

    return {
      pageNumber: idx + 1,
      entries: pageEntries,
      totalPracticeRows,
      density,
      boxSizeMm,
      stageSizeMm,
      entryGapMm,
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
