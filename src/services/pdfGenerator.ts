import { jsPDF } from 'jspdf';
import { repository } from '../storage/db';
import {
  ExportRecord,
  WorkbookEntry,
  WorkbookProject,
  WorkbookSettings,
} from '../types/workbook';
import {
  canInlineMultiCharStrokes,
  getPageDimensionsMm,
  getStrokeColumnsForWidth,
  MANDATORY_TOTAL_BOXES,
  MANDATORY_TRACE_BOXES,
  PaginatedSheet,
  paginateWorkbookEntries,
} from './layoutPagination';
import { renderStrokeStageCanvas } from './strokeEngine';

export interface PdfProgress {
  current: number;
  total: number;
  message: string;
}

function drawWritingBox(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  size: number,
  char: string | undefined,
  isTrace: boolean,
  settings: WorkbookSettings
) {
  ctx.fillStyle = settings.gridBgColor;
  ctx.fillRect(x, y, size, size);

  ctx.save();
  ctx.strokeStyle = settings.gridLineColor;
  ctx.lineWidth = 1.1;
  if (settings.gridStyle !== 'cross') {
    ctx.setLineDash([4, 4]);
  }
  ctx.beginPath();
  ctx.moveTo(x + size / 2, y);
  ctx.lineTo(x + size / 2, y + size);
  ctx.moveTo(x, y + size / 2);
  ctx.lineTo(x + size, y + size / 2);
  ctx.stroke();

  if (settings.gridStyle === 'mi') {
    ctx.strokeStyle = '#DDD2C0';
    ctx.lineWidth = 0.85;
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x + size, y + size);
    ctx.moveTo(x + size, y);
    ctx.lineTo(x, y + size);
    ctx.stroke();
  }
  ctx.restore();

  ctx.strokeStyle = '#9E8D7A';
  ctx.lineWidth = 1.4;
  ctx.strokeRect(x, y, size, size);

  if (isTrace && char) {
    ctx.save();
    ctx.fillStyle = settings.traceGlyphColor;
    ctx.font = `400 ${Math.round(size * 0.7)}px ${settings.fontFamily}`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(char, x + size / 2, y + size / 2 + size * 0.04);
    ctx.restore();
  }
}

function renderCoverCanvas(
  workbook: WorkbookProject,
  entries: WorkbookEntry[],
  widthPx: number,
  heightPx: number
): HTMLCanvasElement {
  const { settings } = workbook;
  const canvas = document.createElement('canvas');
  canvas.width = widthPx;
  canvas.height = heightPx;
  const ctx = canvas.getContext('2d')!;

  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, widthPx, heightPx);

  const inset = Math.round(widthPx * 0.055);
  ctx.strokeStyle = '#E5DDD1';
  ctx.lineWidth = 2;
  ctx.strokeRect(inset, inset, widthPx - inset * 2, heightPx - inset * 2);

  ctx.textAlign = 'center';
  ctx.fillStyle = settings.accentColor;
  ctx.font = `600 ${Math.round(widthPx * 0.014)}px "Plus Jakarta Sans", sans-serif`;
  ctx.fillText('CHINESE CHARACTER WRITING WORKBOOK', widthPx / 2, heightPx * 0.25);

  ctx.fillStyle = '#1E1916';
  ctx.font = `700 ${Math.round(widthPx * 0.045)}px "Cormorant Garamond", Georgia, serif`;
  ctx.fillText(workbook.title, widthPx / 2, heightPx * 0.35);

  ctx.fillStyle = settings.accentColor;
  ctx.fillRect(widthPx * 0.38, heightPx * 0.39, widthPx * 0.24, 4);

  ctx.fillStyle = '#5E544B';
  ctx.font = `400 ${Math.round(widthPx * 0.018)}px "Plus Jakarta Sans", sans-serif`;
  ctx.fillText(
    workbook.subtitle || 'Guided Stroke-Order & Handwriting Practice Sheets',
    widthPx / 2,
    heightPx * 0.45
  );

  const specimen = entries
    .slice(0, 6)
    .map((e) => e.text)
    .join('   ');
  if (specimen) {
    ctx.fillStyle = settings.accentColor;
    ctx.font = `400 ${Math.round(widthPx * 0.052)}px ${settings.fontFamily}`;
    ctx.fillText(specimen, widthPx / 2, heightPx * 0.62);
  }

  const totalStrokes = entries.reduce((sum, e) => sum + e.totalStrokeCount, 0);
  const totalBoxes = entries.reduce(
    (sum, e) => sum + Math.max(1, e.charResults.length) * MANDATORY_TOTAL_BOXES,
    0
  );
  ctx.fillStyle = '#5E544B';
  ctx.font = `500 ${Math.round(widthPx * 0.0135)}px "Plus Jakarta Sans", sans-serif`;
  ctx.fillText(
    `${entries.length} Entries  ·  ${totalBoxes} Practice Boxes (10/char)  ·  ${totalStrokes} Strokes`,
    widthPx / 2,
    heightPx * 0.78
  );

  ctx.fillStyle = '#8C8176';
  ctx.font = `400 ${Math.round(widthPx * 0.012)}px "Plus Jakarta Sans", sans-serif`;
  ctx.fillText(
    `${settings.authorName || 'Hanzi Book Studio Press'}  ·  ${new Date().toLocaleDateString()}`,
    widthPx / 2,
    heightPx * 0.83
  );

  return canvas;
}

function renderPaginatedSheetCanvas(
  workbook: WorkbookProject,
  sheet: PaginatedSheet,
  totalPages: number,
  widthPx: number,
  heightPx: number,
  pxPerMm: number
): HTMLCanvasElement {
  const { settings } = workbook;
  const dims = getPageDimensionsMm(settings);
  const canvas = document.createElement('canvas');
  canvas.width = widthPx;
  canvas.height = heightPx;
  const ctx = canvas.getContext('2d')!;

  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, widthPx, heightPx);

  const marginPx = Math.round(settings.pageMarginMm * pxPerMm);
  const contentX = marginPx;
  const contentW = widthPx - marginPx * 2;
  const isCompact = sheet.density === 'compact';
  const isPortrait = settings.orientation === 'portrait';

  // Header
  const headerTopY = Math.round(marginPx * 0.55);
  ctx.fillStyle = '#5E544B';
  ctx.font = `600 17px "Plus Jakarta Sans", sans-serif`;
  ctx.textAlign = 'left';
  ctx.textBaseline = 'top';
  ctx.fillText(workbook.title, contentX, headerTopY);

  ctx.fillStyle = '#8C8176';
  ctx.font = `500 15px "JetBrains Mono", monospace`;
  ctx.textAlign = 'right';
  ctx.fillText(
    `${settings.pageSize} ${settings.orientation.toUpperCase()} · ${sheet.entries.length} Entries · 10 Boxes/Char`,
    contentX + contentW,
    headerTopY
  );

  ctx.fillStyle = settings.accentColor;
  ctx.fillRect(contentX, headerTopY + 23, contentW, 2.5);

  // Footer
  const footerLineY = heightPx - marginPx - 18;
  ctx.fillStyle = '#E5DDD1';
  ctx.fillRect(contentX, footerLineY, contentW, 1.5);
  ctx.fillStyle = '#8C8176';
  ctx.font = `500 14.5px "Plus Jakarta Sans", sans-serif`;
  ctx.textAlign = 'left';
  ctx.textBaseline = 'top';
  ctx.fillText(
    `${settings.authorName || 'Hanzi Book Studio'} · Practice: Trace boxes 1–2, then write from memory in boxes 3–10`,
    contentX,
    footerLineY + 7
  );
  ctx.textAlign = 'right';
  ctx.fillText(`Page ${sheet.pageNumber} of ${totalPages}`, contentX + contentW, footerLineY + 7);

  // Exact vertical geometry inside [bodyTopY, bodyBottomY]
  const bodyTopY = headerTopY + 34;
  const bodyBottomY = footerLineY - 8;
  const maxBodyH = bodyBottomY - bodyTopY;

  const boxSizePx = Math.round(sheet.boxSizeMm * pxPerMm);
  const boxGapPx = Math.max(6, Math.round(sheet.boxGapMm * pxPerMm));
  const stageSizePx = Math.round(sheet.stageSizeMm * pxPerMm);
  const stageGapPx = 5;
  const upperToPracticeGapPx = Math.round(1.8 * pxPerMm);
  const multiRowGapPx = Math.round(1.4 * pxPerMm);
  const leftW = Math.round(contentW * (isPortrait ? 0.34 : 0.27));
  const rightX = contentX + leftW + 18;
  const maxCols = getStrokeColumnsForWidth(dims.usableWidthMm, sheet.stageSizeMm, isPortrait);

  // Pre-measure every entry height so we can distribute remaining vertical breathing room evenly
  const entryHeights = sheet.entries.map((entry) => {
    const inlineStrokes =
      sheet.inlineMultiCharStrokes &&
      canInlineMultiCharStrokes(entry, dims.usableWidthMm, sheet.stageSizeMm, isPortrait);
    let strokeRows = 0;
    if (inlineStrokes) {
      strokeRows = 1;
    } else {
      for (const ch of entry.charResults) {
        strokeRows += Math.ceil(Math.max(1, ch.strokeCount || 1) / maxCols);
      }
    }
    const minInfoH = isCompact ? Math.max(46, stageSizePx + 4) : Math.max(58, stageSizePx + 6);
    const strokeH = strokeRows * stageSizePx + Math.max(0, strokeRows - 1) * stageGapPx;
    const upperH = Math.max(minInfoH, strokeH);
    const practiceH =
      entry.charResults.length * boxSizePx +
      Math.max(0, entry.charResults.length - 1) * multiRowGapPx;
    return { upperH, practiceH, totalH: upperH + upperToPracticeGapPx + practiceH, inlineStrokes };
  });

  const sumEntryH = entryHeights.reduce((sum, m) => sum + m.totalH, 0);
  const gapSlots = Math.max(1, sheet.entries.length - 1);
  const rawGap = sheet.entries.length > 1 ? Math.floor((maxBodyH - sumEntryH) / gapSlots) : 0;
  const entryGapPx = Math.max(10, Math.min(Math.round(sheet.entryGapMm * pxPerMm * 1.6), rawGap));

  // Center the vertical stack nicely within [bodyTopY, bodyBottomY]
  const finalStackH = sumEntryH + Math.max(0, sheet.entries.length - 1) * entryGapPx;
  let cursorY = bodyTopY + Math.max(0, Math.floor((maxBodyH - finalStackH) / 2));

  sheet.entries.forEach((entry, idx) => {
    const isMulti = entry.charResults.length > 1;
    const metrics = entryHeights[idx];

    if (idx > 0) {
      ctx.fillStyle = '#E5DDD1';
      ctx.fillRect(contentX, cursorY - Math.round(entryGapPx / 2), contentW, 1.2);
    }

    const bandTopY = cursorY;

    // Left Vocabulary Info (vertically centered inside upperH)
    const heroFontPx = isCompact ? (isMulti ? 32 : 42) : isMulti ? 42 : 54;
    const infoCenterY = bandTopY + metrics.upperH / 2;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.fillStyle = '#1E1916';
    ctx.font = `400 ${heroFontPx}px ${settings.fontFamily}`;
    ctx.fillText(entry.text, contentX, infoCenterY);

    const measuredHeroW = ctx.measureText(entry.text).width;
    const textColX = contentX + Math.min(leftW * 0.5, measuredHeroW + 14);
    ctx.textBaseline = 'top';
    ctx.fillStyle = settings.accentColor;
    ctx.font = `600 ${isCompact ? 19 : 22}px "Plus Jakarta Sans", sans-serif`;
    ctx.fillText(entry.pinyin || '—', textColX, bandTopY + Math.max(2, (metrics.upperH - 42) / 2));

    const strokeLabel = isMulti
      ? `${entry.totalStrokeCount} str (${entry.charResults.map((c) => c.strokeCount).join('+')})`
      : `${entry.totalStrokeCount} strokes`;
    ctx.fillStyle = '#5E544B';
    ctx.font = `400 ${isCompact ? 14 : 16}px "Plus Jakarta Sans", sans-serif`;
    const maxMeaningLen = isPortrait ? 20 : 32;
    const clippedMeaning = (entry.meaning || 'No English meaning').slice(0, maxMeaningLen);
    ctx.fillText(
      `${clippedMeaning} · ${strokeLabel}`,
      textColX,
      bandTopY + Math.max(2, (metrics.upperH - 42) / 2) + (isCompact ? 23 : 27)
    );

    // Subtle vertical divider between Vocab Info and Stroke Order
    ctx.fillStyle = '#E5DDD1';
    ctx.fillRect(contentX + leftW + 6, bandTopY + 2, 1.2, metrics.upperH - 4);

    // Right Stroke-Order Diagrams (centered vertically inside upperH)
    if (metrics.inlineStrokes) {
      let curX = rightX;
      const tileY = bandTopY + Math.round((metrics.upperH - stageSizePx) / 2);
      for (let cIdx = 0; cIdx < entry.charResults.length; cIdx++) {
        const charRes = entry.charResults[cIdx];
        if (cIdx > 0) curX += 14;
        ctx.fillStyle = '#1E1916';
        ctx.font = `600 18px ${settings.fontFamily}`;
        ctx.textAlign = 'left';
        ctx.textBaseline = 'middle';
        ctx.fillText(charRes.char, curX, tileY + stageSizePx / 2);
        curX += 24;

        if (charRes.status === 'ready' && charRes.strokeData) {
          for (let s = 0; s < charRes.strokeData.strokeCount; s++) {
            const stageCanvas = renderStrokeStageCanvas(charRes.strokeData, s, {
              pixelSize: Math.max(120, stageSizePx * 2),
              newStrokeColor: settings.newStrokeColor,
              completedStrokeColor: settings.completedStrokeColor,
            });
            ctx.drawImage(stageCanvas, curX, tileY, stageSizePx, stageSizePx);
            curX += stageSizePx + stageGapPx;
          }
        }
      }
    } else {
      let strokeY = bandTopY;
      for (const charRes of entry.charResults) {
        let rowOffsetX = rightX;
        if (isMulti) {
          ctx.fillStyle = '#1E1916';
          ctx.font = `600 19px ${settings.fontFamily}`;
          ctx.textAlign = 'left';
          ctx.textBaseline = 'middle';
          ctx.fillText(charRes.char, rightX, strokeY + stageSizePx / 2);
          rowOffsetX += 26;
        }

        if (charRes.status === 'ready' && charRes.strokeData) {
          const count = charRes.strokeData.strokeCount;
          for (let s = 0; s < count; s++) {
            const col = s % maxCols;
            const row = Math.floor(s / maxCols);
            const stageCanvas = renderStrokeStageCanvas(charRes.strokeData, s, {
              pixelSize: Math.max(120, stageSizePx * 2),
              newStrokeColor: settings.newStrokeColor,
              completedStrokeColor: settings.completedStrokeColor,
            });
            ctx.drawImage(
              stageCanvas,
              rowOffsetX + col * (stageSizePx + stageGapPx),
              strokeY + row * (stageSizePx + stageGapPx),
              stageSizePx,
              stageSizePx
            );
          }
          const rowsUsed = Math.ceil(count / maxCols);
          strokeY += rowsUsed * (stageSizePx + stageGapPx);
        } else {
          strokeY += stageSizePx + stageGapPx;
        }
      }
    }

    // Lower Independent 10-Box Practice Rows (1 row of 10 boxes per character in the entry)
    let practiceY = bandTopY + metrics.upperH + upperToPracticeGapPx;
    const labelW = 34;

    for (const charRes of entry.charResults) {
      ctx.fillStyle = '#5E544B';
      ctx.font = `400 22px ${settings.fontFamily}`;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(charRes.char, contentX + labelW / 2 - 4, practiceY + boxSizePx / 2);

      for (let b = 0; b < MANDATORY_TOTAL_BOXES; b++) {
        drawWritingBox(
          ctx,
          contentX + labelW + b * (boxSizePx + boxGapPx),
          practiceY,
          boxSizePx,
          charRes.char,
          b < MANDATORY_TRACE_BOXES,
          settings
        );
      }

      if (!isPortrait) {
        ctx.fillStyle = '#8C8176';
        ctx.font = `500 13px "JetBrains Mono", monospace`;
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(
          `1–${MANDATORY_TRACE_BOXES} Trace · ${MANDATORY_TRACE_BOXES + 1}–${MANDATORY_TOTAL_BOXES} Write`,
          contentX + contentW,
          practiceY + boxSizePx / 2
        );
      }

      practiceY += boxSizePx + multiRowGapPx;
    }

    cursorY += metrics.totalH + entryGapPx;
  });

  return canvas;
}

export async function generateWorkbookPdf(
  workbook: WorkbookProject,
  selectedEntries?: WorkbookEntry[],
  onProgress?: (progress: PdfProgress) => void
): Promise<ExportRecord> {
  const entries = (
    selectedEntries && selectedEntries.length > 0 ? selectedEntries : workbook.entries
  ).filter((e) => e.charResults.length > 0);

  if (entries.length === 0) {
    throw new Error('No entries selected for PDF export.');
  }

  const { settings } = workbook;
  const dims = getPageDimensionsMm(settings);
  const pxPerMm = 5.6;
  const canvasW = Math.round(dims.widthMm * pxPerMm);
  const canvasH = Math.round(dims.heightMm * pxPerMm);

  const sheets = paginateWorkbookEntries(entries, settings);

  const pdf = new jsPDF({
    orientation: settings.orientation,
    unit: 'mm',
    format: settings.pageSize.toLowerCase() as 'a4' | 'a3',
    compress: true,
  });

  let hasAddedFirstPage = false;
  const pushCanvasPage = (canvas: HTMLCanvasElement) => {
    if (hasAddedFirstPage) {
      pdf.addPage(settings.pageSize.toLowerCase(), settings.orientation);
    }
    hasAddedFirstPage = true;
    const dataUrl = canvas.toDataURL('image/jpeg', 0.93);
    pdf.addImage(dataUrl, 'JPEG', 0, 0, dims.widthMm, dims.heightMm);
  };

  if (settings.includeCoverPage) {
    onProgress?.({ current: 0, total: entries.length, message: 'Rendering PDF cover page…' });
    pushCanvasPage(renderCoverCanvas(workbook, entries, canvasW, canvasH));
  }

  let processedEntries = 0;
  for (let p = 0; p < sheets.length; p++) {
    const sheet = sheets[p];
    processedEntries += sheet.entries.length;
    onProgress?.({
      current: Math.min(entries.length, processedEntries),
      total: entries.length,
      message: `Rendering PDF sheet ${sheet.pageNumber} of ${sheets.length}…`,
    });

    const pageCanvas = renderPaginatedSheetCanvas(
      workbook,
      sheet,
      sheets.length,
      canvasW,
      canvasH,
      pxPerMm
    );
    pushCanvasPage(pageCanvas);

    if (p % 3 === 2) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  const safeSlug =
    workbook.title
      .toLowerCase()
      .replace(/[^a-z0-9\u4e00-\u9fff]+/gi, '-')
      .replace(/^-|-$/g, '') || 'hanzi-workbook';
  const fileName = `${safeSlug}-${settings.pageSize.toLowerCase()}-${settings.orientation}.pdf`;
  pdf.save(fileName);

  const record: ExportRecord = {
    id: `exp-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
    workbookId: workbook.id,
    workbookTitle: workbook.title,
    format: 'pdf',
    fileName,
    entryCount: entries.length,
    pageSize: settings.pageSize,
    orientation: settings.orientation,
    exportedAt: Date.now(),
  };

  await repository.saveExport(record);
  return record;
}
