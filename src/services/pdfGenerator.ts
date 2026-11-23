import { jsPDF } from 'jspdf';
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
  const palette = resolveWorkbookPalette(settings);
  ctx.fillStyle = palette.gridBgColor;
  ctx.fillRect(x, y, size, size);

  ctx.save();
  ctx.strokeStyle = palette.gridLineColor;
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
    ctx.strokeStyle = palette.gridDiagonalColor;
    ctx.lineWidth = 0.85;
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x + size, y + size);
    ctx.moveTo(x + size, y);
    ctx.lineTo(x, y + size);
    ctx.stroke();
  }
  ctx.restore();

  ctx.strokeStyle = palette.gridBorderColor;
  ctx.lineWidth = Number((palette.gridBorderWidthPx * 1.18).toFixed(2));
  ctx.strokeRect(x, y, size, size);

  if (isTrace && char) {
    ctx.save();
    ctx.fillStyle = palette.traceGlyphColor;
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
  sheets: PaginatedSheet[],
  widthPx: number,
  heightPx: number
): HTMLCanvasElement {
  const { settings } = workbook;
  const palette = resolveWorkbookPalette(settings);
  const canvas = document.createElement('canvas');
  canvas.width = widthPx;
  canvas.height = heightPx;
  const ctx = canvas.getContext('2d')!;

  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, widthPx, heightPx);

  const inset = Math.round(widthPx * 0.055);
  ctx.strokeStyle = palette.dividerColor;
  ctx.lineWidth = 2;
  ctx.strokeRect(inset, inset, widthPx - inset * 2, heightPx - inset * 2);

  ctx.textAlign = 'center';
  ctx.fillStyle = palette.accentColor;
  ctx.font = `600 ${Math.round(widthPx * 0.014)}px "Plus Jakarta Sans", sans-serif`;
  ctx.fillText('CHINESE CHARACTER WRITING WORKBOOK', widthPx / 2, heightPx * 0.25);

  ctx.fillStyle = palette.primaryInk;
  ctx.font = `700 ${Math.round(widthPx * 0.045)}px "Cormorant Garamond", Georgia, serif`;
  ctx.fillText(workbook.title, widthPx / 2, heightPx * 0.35);

  ctx.fillStyle = palette.accentColor;
  ctx.fillRect(widthPx * 0.38, heightPx * 0.39, widthPx * 0.24, 4);

  ctx.fillStyle = palette.secondaryInk;
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
    ctx.fillStyle = palette.accentColor;
    ctx.font = `400 ${Math.round(widthPx * 0.052)}px ${settings.fontFamily}`;
    ctx.fillText(specimen, widthPx / 2, heightPx * 0.62);
  }

  const totalStrokes = entries.reduce((sum, e) => sum + e.totalStrokeCount, 0);
  const totalBoxes = sheets.reduce((sum, s) => sum + s.totalBoxesOnPage, 0);
  const uniqueBoxesPerChar = Array.from(new Set(sheets.map((s) => s.boxesPerRow)));
  const boxesPerCharLabel =
    uniqueBoxesPerChar.length === 1
      ? `${uniqueBoxesPerChar[0]}/char`
      : `${Math.min(...uniqueBoxesPerChar)}–${Math.max(...uniqueBoxesPerChar)}/char`;
  ctx.fillStyle = palette.secondaryInk;
  ctx.font = `500 ${Math.round(widthPx * 0.0135)}px "Plus Jakarta Sans", sans-serif`;
  ctx.fillText(
    `${entries.length} Entries  ·  ${totalBoxes} Practice Boxes (${boxesPerCharLabel})  ·  ${totalStrokes} Strokes`,
    widthPx / 2,
    heightPx * 0.78
  );

  ctx.fillStyle = palette.subtleInk;
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
  const palette = resolveWorkbookPalette(settings);
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

  const headerTitle = resolveSheetHeaderTitle(workbook.title, settings);
  const footerInstruction = resolveSheetFooterText(settings, sheet.traceBoxes, sheet.boxesPerRow);

  // Header
  const headerTopY = Math.round(marginPx * 0.55);
  const headerRightText = `${settings.pageSize} ${settings.orientation.toUpperCase()} · ${sheet.entries.length} Entries · ${sheet.boxesPerRow} Boxes/Char (${sheet.totalBoxesOnPage} Boxes)`;
  ctx.font = `500 14.5px "JetBrains Mono", monospace`;
  const rightMetaW = ctx.measureText(headerRightText).width;
  ctx.fillStyle = palette.subtleInk;
  ctx.textAlign = 'right';
  ctx.textBaseline = 'top';
  ctx.fillText(headerRightText, contentX + contentW, headerTopY);

  ctx.fillStyle = palette.primaryInk;
  ctx.font = `600 17px "Plus Jakarta Sans", sans-serif`;
  ctx.textAlign = 'left';
  const maxHeaderTitleW = Math.max(120, contentW - rightMetaW - 24);
  ctx.fillText(headerTitle, contentX, headerTopY, maxHeaderTitleW);

  ctx.fillStyle = palette.accentColor;
  ctx.fillRect(contentX, headerTopY + 23, contentW, 2.5);

  let bodyTopY = headerTopY + 34;

  // Optional Printable Student Info Bar
  if (settings.showStudentInfoBar) {
    const barY = headerTopY + 34;
    ctx.fillStyle = palette.primaryInk;
    ctx.font = `600 16px "Plus Jakarta Sans", sans-serif`;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'top';
    ctx.fillText('Name:', contentX, barY);

    const nameLineStartX = contentX + 58;
    const dateLabelX = contentX + contentW - 280;
    const lineColor = palette.isBw ? '#000000' : palette.secondaryInk;

    ctx.fillStyle = lineColor;
    ctx.fillRect(nameLineStartX, barY + 18, Math.max(100, dateLabelX - nameLineStartX - 32), 1.3);

    ctx.fillStyle = palette.primaryInk;
    ctx.fillText('Date:', dateLabelX, barY);
    ctx.fillStyle = lineColor;
    ctx.fillRect(dateLabelX + 50, barY + 18, Math.max(80, contentX + contentW - (dateLabelX + 50)), 1.3);

    ctx.fillStyle = palette.dividerColor;
    ctx.fillRect(contentX, barY + 28, contentW, 1);
    bodyTopY = barY + 36;
  }

  // Footer
  const footerLineY = heightPx - marginPx - 18;
  ctx.fillStyle = palette.dividerColor;
  ctx.fillRect(contentX, footerLineY, contentW, 1.5);
  ctx.fillStyle = palette.subtleInk;
  ctx.font = `500 14px "Plus Jakarta Sans", sans-serif`;
  const pageText = `Page ${sheet.pageNumber} of ${totalPages}`;
  const pageTextW = ctx.measureText(pageText).width;
  ctx.textAlign = 'right';
  ctx.textBaseline = 'top';
  ctx.fillText(pageText, contentX + contentW, footerLineY + 7);

  ctx.textAlign = 'left';
  ctx.fillText(footerInstruction, contentX, footerLineY + 7, Math.max(160, contentW - pageTextW - 24));

  // Exact vertical & horizontal geometry inside [bodyTopY, bodyBottomY]
  const bodyBottomY = footerLineY - 8;
  const maxBodyH = bodyBottomY - bodyTopY;

  const labelW = 34;
  const boxGapPx = Math.max(3, Math.floor(sheet.boxGapMm * pxPerMm));
  const maxFitBoxPx = Math.floor(
    (contentW - labelW - boxGapPx * Math.max(0, sheet.boxesPerRow - 1)) /
      Math.max(1, sheet.boxesPerRow)
  );
  const boxSizePx = Math.max(32, Math.min(Math.floor(sheet.boxSizeMm * pxPerMm), maxFitBoxPx));
  const stageSizePx = Math.round(sheet.stageSizeMm * pxPerMm);
  const stageGapPx = 5;
  const upperToPracticeGapPx = Math.round(1.6 * pxPerMm);
  const multiRowGapPx = Math.round(1.3 * pxPerMm);

  const mainHanziBasePx =
    settings.mainHanziSizePx ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.mainHanziSizePx;
  const metaBasePx = settings.metaFontSizePx ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.metaFontSizePx;
  const pdfPinyinPx = Math.round(metaBasePx * 1.52);
  const pdfEnglishPx = Math.max(12, Math.round(metaBasePx * 1.18));
  const pdfArabicPx = Math.max(13, Math.round(metaBasePx * 1.28));

  const leftRatio = isPortrait ? 0.35 : 0.28;
  const leftW = Math.round(contentW * leftRatio);
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
    const hasArabic =
      settings.showArabicTranslation !== false && Boolean(entry.arabicMeaning?.trim());
    const scaledHanziH = Math.round(mainHanziBasePx * 1.48);
    const metaStackH = pdfPinyinPx + pdfEnglishPx + (hasArabic ? pdfArabicPx + 14 : 10);
    const minInfoH = Math.max(
      isCompact ? (hasArabic ? 54 : 46) : hasArabic ? 66 : 58,
      stageSizePx + 4,
      scaledHanziH + 4,
      metaStackH
    );
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
  const entryGapPx = Math.max(8, Math.min(Math.round(sheet.entryGapMm * pxPerMm * 1.6), rawGap));

  // Center the vertical stack nicely within [bodyTopY, bodyBottomY]
  const finalStackH = sumEntryH + Math.max(0, sheet.entries.length - 1) * entryGapPx;
  let cursorY = bodyTopY + Math.max(0, Math.floor((maxBodyH - finalStackH) / 2));

  sheet.entries.forEach((entry, idx) => {
    const isMulti = entry.charResults.length > 1;
    const metrics = entryHeights[idx];

    if (idx > 0) {
      ctx.fillStyle = palette.dividerColor;
      ctx.fillRect(contentX, cursorY - Math.round(entryGapPx / 2), contentW, 1.2);
    }

    const bandTopY = cursorY;

    // Left Vocabulary Info (vertically centered inside upperH)
    const multiScale =
      entry.charResults.length >= 4
        ? 0.56
        : entry.charResults.length === 3
        ? 0.68
        : isMulti
        ? 0.8
        : 1;
    const heroFontPx = Math.max(26, Math.round(mainHanziBasePx * 1.48 * multiScale));
    const infoCenterY = bandTopY + metrics.upperH / 2;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.fillStyle = palette.primaryInk;
    ctx.font = `400 ${heroFontPx}px ${settings.fontFamily}`;
    ctx.fillText(entry.text, contentX, infoCenterY, leftW * 0.48);

    const measuredHeroW = Math.min(leftW * 0.48, ctx.measureText(entry.text).width);
    const textColX = contentX + measuredHeroW + 14;
    const textColMaxW = Math.max(70, contentX + leftW - textColX - 4);
    const hasArabic =
      settings.showArabicTranslation !== false && Boolean(entry.arabicMeaning?.trim());
    const stackH = pdfPinyinPx + pdfEnglishPx + (hasArabic ? pdfArabicPx + 10 : 6);
    const infoStartY = bandTopY + Math.max(1, (metrics.upperH - stackH) / 2);

    ctx.textBaseline = 'top';
    ctx.fillStyle = palette.accentColor;
    ctx.font = `600 ${pdfPinyinPx}px "Plus Jakarta Sans", sans-serif`;
    ctx.fillText(entry.pinyin || '—', textColX, infoStartY, textColMaxW);

    const strokeLabel = isMulti
      ? `${entry.totalStrokeCount} str (${entry.charResults.map((c) => c.strokeCount).join('+')})`
      : `${entry.totalStrokeCount} strokes`;
    ctx.fillStyle = palette.secondaryInk;
    ctx.font = `400 ${pdfEnglishPx}px "Plus Jakarta Sans", sans-serif`;
    const maxMeaningLen = isPortrait ? 18 : 28;
    const clippedMeaning = (entry.meaning || 'No English meaning').slice(0, maxMeaningLen);
    ctx.fillText(
      `${clippedMeaning} · ${strokeLabel}`,
      textColX,
      infoStartY + pdfPinyinPx + 3,
      textColMaxW
    );

    if (hasArabic) {
      ctx.save();
      ctx.direction = 'rtl';
      ctx.textAlign = 'right';
      ctx.textBaseline = 'top';
      ctx.fillStyle = palette.primaryInk;
      ctx.font = `500 ${pdfArabicPx}px "Noto Sans Arabic", Tahoma, sans-serif`;
      const arabicRightX = contentX + leftW - 4;
      const clippedArabic = entry.arabicMeaning!.trim().slice(0, isPortrait ? 20 : 28);
      ctx.fillText(
        clippedArabic,
        arabicRightX,
        infoStartY + pdfPinyinPx + pdfEnglishPx + 7,
        textColMaxW
      );
      ctx.restore();
    }

    // Subtle vertical divider between Vocab Info and Stroke Order
    ctx.fillStyle = palette.dividerColor;
    ctx.fillRect(contentX + leftW + 6, bandTopY + 2, 1.2, metrics.upperH - 4);

    // Right Stroke-Order Diagrams (centered vertically inside upperH)
    if (metrics.inlineStrokes) {
      let curX = rightX;
      const tileY = bandTopY + Math.round((metrics.upperH - stageSizePx) / 2);
      for (let cIdx = 0; cIdx < entry.charResults.length; cIdx++) {
        const charRes = entry.charResults[cIdx];
        if (cIdx > 0) curX += 14;
        ctx.fillStyle = palette.primaryInk;
        ctx.font = `600 18px ${settings.fontFamily}`;
        ctx.textAlign = 'left';
        ctx.textBaseline = 'middle';
        ctx.fillText(charRes.char, curX, tileY + stageSizePx / 2);
        curX += 24;

        if (charRes.status === 'ready' && charRes.strokeData) {
          for (let s = 0; s < charRes.strokeData.strokeCount; s++) {
            const stageCanvas = renderStrokeStageCanvas(charRes.strokeData, s, {
              pixelSize: Math.max(180, stageSizePx * 3),
              newStrokeColor: palette.newStrokeColor,
              completedStrokeColor: palette.completedStrokeColor,
              borderColor: palette.strokeTileBorderColor,
              guideLineColor: palette.strokeGuideLineColor,
              numberColor: palette.strokeNumberColor,
              backgroundColor: '#FFFFFF',
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
          ctx.fillStyle = palette.primaryInk;
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
              pixelSize: Math.max(180, stageSizePx * 3),
              newStrokeColor: palette.newStrokeColor,
              completedStrokeColor: palette.completedStrokeColor,
              borderColor: palette.strokeTileBorderColor,
              guideLineColor: palette.strokeGuideLineColor,
              numberColor: palette.strokeNumberColor,
              backgroundColor: '#FFFFFF',
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

    // Lower Independent Full-Width Practice Rows (1 row of sheet.boxesPerRow boxes per character)
    let practiceY = bandTopY + metrics.upperH + upperToPracticeGapPx;

    for (const charRes of entry.charResults) {
      ctx.fillStyle = palette.isBw ? '#000000' : '#5E544B';
      ctx.font = `400 22px ${settings.fontFamily}`;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(charRes.char, contentX + labelW / 2 - 4, practiceY + boxSizePx / 2);

      for (let b = 0; b < sheet.boxesPerRow; b++) {
        drawWritingBox(
          ctx,
          contentX + labelW + b * (boxSizePx + boxGapPx),
          practiceY,
          boxSizePx,
          charRes.char,
          b < sheet.traceBoxes,
          settings
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

  if (typeof document !== 'undefined' && 'fonts' in document) {
    try {
      await document.fonts.ready;
    } catch {
      // Proceed even if fonts API rejects
    }
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
    pushCanvasPage(renderCoverCanvas(workbook, entries, sheets, canvasW, canvasH));
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
  const fileName = `${safeSlug}-${settings.pageSize.toLowerCase()}-${settings.orientation}${
    settings.bwPrintMode ? '-bw' : ''
  }.pdf`;
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
