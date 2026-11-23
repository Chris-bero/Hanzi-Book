import React from 'react';
import {
  canInlineMultiCharStrokes,
  getPageDimensionsMm,
  PaginatedSheet,
  paginateWorkbookEntries,
} from '../../services/layoutPagination';
import {
  PRO_TYPOGRAPHY_CONTRAST_DEFAULTS,
  resolveSheetFooterText,
  resolveSheetHeaderTitle,
  resolveWorkbookPalette,
  WorkbookEntry,
  WorkbookProject,
  WorkbookSettings,
} from '../../types/workbook';
import { StrokeStageGrid } from '../Stroke/StrokeStageGrid';
import { WritingGridBox } from './WritingGridBox';

export function getSheetPreviewScale(settings: WorkbookSettings) {
  const dims = getPageDimensionsMm(settings);
  const targetWidthPx =
    settings.orientation === 'landscape'
      ? settings.pageSize === 'A3'
        ? 1180
        : 1040
      : settings.pageSize === 'A3'
      ? 860
      : 760;

  const pxPerMm = targetWidthPx / dims.widthMm;
  const heightPx = Math.round(dims.heightMm * pxPerMm);
  const paddingPx = Math.round(settings.pageMarginMm * pxPerMm);
  return { widthPx: targetWidthPx, heightPx, paddingPx, pxPerMm, dims };
}

interface WorkbookPageProps {
  workbookTitle: string;
  sheet: PaginatedSheet;
  settings: WorkbookSettings;
  totalPages: number;
}

export const WorkbookPage: React.FC<WorkbookPageProps> = ({
  workbookTitle,
  sheet,
  settings,
  totalPages,
}) => {
  const { widthPx, heightPx, paddingPx, pxPerMm, dims } = getSheetPreviewScale(settings);
  const palette = resolveWorkbookPalette(settings);
  const isCompact = sheet.density === 'compact';
  const isPortrait = settings.orientation === 'portrait';

  // Exact inner printable pixel width (subtracting article 2px border + horizontal padding)
  const innerWidthPx = widthPx - 2 - paddingPx * 2;
  const rowLabelWidthPx = 22;
  const rowLabelGapPx = 6;
  const availRowBoxesPx = Math.max(120, innerWidthPx - rowLabelWidthPx - rowLabelGapPx);

  // Strictly clamp boxGapPx and boxSizePx via Math.floor so rightmost boxes NEVER clip
  const boxGapPx = Math.max(2, Math.floor(sheet.boxGapMm * pxPerMm));
  const maxFitBoxPx = Math.floor(
    (availRowBoxesPx - boxGapPx * Math.max(0, sheet.boxesPerRow - 1)) /
      Math.max(1, sheet.boxesPerRow)
  );
  const boxSizePx = Math.max(24, Math.min(Math.floor(sheet.boxSizeMm * pxPerMm), maxFitBoxPx));
  const stageSizePx = Math.round(sheet.stageSizeMm * pxPerMm);
  const entryGapPx = Math.max(6, Math.round(sheet.entryGapMm * pxPerMm));

  const mainHanziBasePx =
    settings.mainHanziSizePx ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.mainHanziSizePx;
  const metaBasePx = settings.metaFontSizePx ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.metaFontSizePx;
  const pinyinFontPx = Math.round(metaBasePx * 1.08);
  const englishFontPx = Math.max(9, Math.round(metaBasePx * 0.86));
  const arabicFontPx = Math.max(10, Math.round(metaBasePx * 0.94));
  const strokeMetaFontPx = Math.max(8.5, Math.round(metaBasePx * 0.76));

  const headerTitle = resolveSheetHeaderTitle(workbookTitle, settings);
  const footerInstruction = resolveSheetFooterText(settings, sheet.traceBoxes, sheet.boxesPerRow);

  // Slightly widen left info column when large typography is chosen, without squeezing stroke grid
  const leftInfoWidthPx = isPortrait
    ? Math.min(245, Math.max(205, Math.round(180 + mainHanziBasePx * 1.1)))
    : Math.min(310, Math.max(255, Math.round(215 + mainHanziBasePx * 1.5)));

  return (
    <article
      className={`print-sheet mx-auto bg-white border flex flex-col justify-between box-border overflow-hidden ${
        palette.isBw
          ? 'bw-print-sheet border-black shadow-none'
          : 'border-hairline shadow-[0_14px_38px_rgba(30,25,22,0.09)]'
      }`}
      style={{
        width: `${widthPx}px`,
        minHeight: `${heightPx}px`,
        padding: `${paddingPx}px`,
        backgroundColor: palette.pageBg,
      }}
      aria-label={`Workbook page ${sheet.pageNumber}`}
    >
      {/* Top Running Header + Optional Printable Student Info Bar */}
      <div>
        <header
          className="flex items-center justify-between gap-3 pb-1.5 border-b-2 text-[10px] font-medium tracking-[0.04em] min-w-0"
          style={{
            borderBottomColor: palette.accentColor,
            color: palette.subtleInk,
          }}
        >
          <span
            className="truncate min-w-0 font-semibold"
            style={{ color: palette.primaryInk }}
            title={headerTitle}
          >
            {headerTitle}
          </span>
          <span
            className="font-mono tabular-nums whitespace-nowrap shrink-0"
            style={{ color: palette.secondaryInk }}
          >
            {settings.pageSize} {settings.orientation === 'landscape' ? 'Landscape' : 'Portrait'} ·{' '}
            {sheet.entries.length} {sheet.entries.length === 1 ? 'Entry' : 'Entries'} ·{' '}
            {sheet.boxesPerRow} Boxes/Char ({sheet.totalBoxesOnPage} Boxes)
          </span>
        </header>

        {settings.showStudentInfoBar && (
          <div
            className="flex items-center justify-between gap-6 pt-2 pb-1 text-[11px] font-medium border-b"
            style={{
              borderBottomColor: palette.dividerColor,
              color: palette.primaryInk,
            }}
          >
            <div className="flex items-baseline gap-2 flex-1 min-w-0">
              <span className="shrink-0 font-semibold">Name:</span>
              <span
                className="flex-1 border-b"
                style={{
                  borderBottomColor: palette.isBw ? '#000000' : palette.secondaryInk,
                  height: '12px',
                }}
              />
            </div>
            <div className="flex items-baseline gap-2 w-52 shrink-0">
              <span className="shrink-0 font-semibold">Date:</span>
              <span
                className="flex-1 border-b"
                style={{
                  borderBottomColor: palette.isBw ? '#000000' : palette.secondaryInk,
                  height: '12px',
                }}
              />
            </div>
          </div>
        )}
      </div>

      {/* Independent Vocabulary Entry Bands */}
      <div
        className="flex-1 flex flex-col justify-evenly py-2"
        style={{ gap: `${entryGapPx}px` }}
      >
        {sheet.entries.map((entry, idx) => {
          const isMultiChar = entry.charResults.length > 1;
          const inlineStrokes =
            sheet.inlineMultiCharStrokes &&
            canInlineMultiCharStrokes(entry, dims.usableWidthMm, sheet.stageSizeMm, isPortrait);

          const multiScale =
            entry.charResults.length >= 4
              ? 0.56
              : entry.charResults.length === 3
              ? 0.68
              : isMultiChar
              ? 0.8
              : 1;
          const heroSizePx = Math.max(18, Math.round(mainHanziBasePx * multiScale));

          const strokeLabel = isMultiChar
            ? `${entry.totalStrokeCount} strokes (${entry.charResults
                .map((c) => c.strokeCount)
                .join('+')})`
            : `${entry.totalStrokeCount} ${entry.totalStrokeCount === 1 ? 'stroke' : 'strokes'}`;

          return (
            <section
              key={entry.id}
              aria-label={`Vocabulary entry ${entry.text}`}
              className={`flex flex-col min-w-0 ${idx > 0 ? 'pt-2 border-t' : ''}`}
              style={{
                gap: isCompact ? '5px' : '9px',
                borderTopColor: idx > 0 ? palette.dividerColor : undefined,
              }}
            >
              {/* Upper Band: Vocabulary Info (Left) + Stroke-Order Diagrams (Right) */}
              <div
                className="grid items-center gap-3.5 min-w-0"
                style={{
                  gridTemplateColumns: `${leftInfoWidthPx}px minmax(0, 1fr)`,
                }}
              >
                {/* Left Vocabulary Info */}
                <div
                  className="pr-3 border-r flex items-center gap-2.5 min-w-0"
                  style={{ borderRightColor: palette.dividerColor }}
                >
                  <div
                    className="leading-none font-normal shrink-0"
                    style={{
                      fontFamily: settings.fontFamily,
                      fontSize: `${heroSizePx}px`,
                      color: palette.primaryInk,
                    }}
                  >
                    {entry.text}
                  </div>
                  <div className="min-w-0 flex-1">
                    <div className="flex items-baseline gap-1.5 flex-wrap">
                      <span
                        className="font-semibold leading-tight truncate"
                        style={{
                          color: palette.accentColor,
                          fontSize: `${pinyinFontPx}px`,
                        }}
                      >
                        {entry.pinyin || '—'}
                      </span>
                      <span
                        className="font-mono tabular-nums whitespace-nowrap"
                        style={{
                          color: palette.subtleInk,
                          fontSize: `${strokeMetaFontPx}px`,
                        }}
                      >
                        {strokeLabel}
                      </span>
                    </div>
                    <p
                      className="m-0 truncate leading-tight"
                      title={entry.meaning}
                      style={{
                        color: palette.secondaryInk,
                        fontSize: `${englishFontPx}px`,
                        marginTop: '1px',
                      }}
                    >
                      {entry.meaning || 'No English meaning provided.'}
                    </p>
                    {settings.showArabicTranslation !== false && entry.arabicMeaning?.trim() && (
                      <p
                        dir="rtl"
                        className="font-medium m-0 truncate leading-tight text-left"
                        title={entry.arabicMeaning}
                        style={{
                          fontFamily: 'var(--font-arabic)',
                          color: palette.primaryInk,
                          fontSize: `${arabicFontPx}px`,
                          marginTop: '1px',
                        }}
                      >
                        {entry.arabicMeaning}
                      </p>
                    )}
                  </div>
                </div>

                {/* Right Stroke Order Diagrams */}
                <div
                  className={`min-w-0 flex ${
                    inlineStrokes ? 'flex-row flex-wrap items-center gap-3' : 'flex-col gap-1'
                  }`}
                >
                  {entry.charResults.map((charRes, cIdx) => (
                    <div
                      key={`${entry.id}-strokes-${cIdx}`}
                      className="flex items-center gap-1.5 min-w-0"
                    >
                      {isMultiChar && (
                        <span
                          className="text-xs font-semibold shrink-0"
                          style={{
                            fontFamily: settings.fontFamily,
                            color: palette.primaryInk,
                          }}
                        >
                          {charRes.char}
                        </span>
                      )}
                      {charRes.status === 'ready' && charRes.strokeData ? (
                        <StrokeStageGrid
                          strokeData={charRes.strokeData}
                          stageSize={stageSizePx}
                          newStrokeColor={palette.newStrokeColor}
                          completedStrokeColor={palette.completedStrokeColor}
                          tileBorderColor={palette.strokeTileBorderColor}
                          guideLineColor={palette.strokeGuideLineColor}
                          numberColor={palette.strokeNumberColor}
                          compact
                        />
                      ) : (
                        <span
                          className="text-[11px]"
                          style={{ color: palette.isBw ? '#000000' : '#A51D24' }}
                        >
                          Stroke-order vector unavailable for “{charRes.char}”
                        </span>
                      )}
                    </div>
                  ))}
                </div>
              </div>

              {/* Lower Band: Independent Full-Width Practice Row for EACH Character */}
              <div className="flex flex-col gap-1.5 min-w-0">
                {entry.charResults.map((charRes, rowIdx) => (
                  <div
                    key={`${entry.id}-practice-${rowIdx}`}
                    className="flex items-center min-w-0"
                    style={{ gap: `${rowLabelGapPx}px` }}
                  >
                    {/* Character marker for each independent practice row */}
                    <div
                      className="shrink-0 flex items-center justify-center"
                      style={{ width: `${rowLabelWidthPx}px` }}
                    >
                      <span
                        className="text-sm leading-none"
                        style={{
                          fontFamily: settings.fontFamily,
                          color: palette.primaryInk,
                        }}
                      >
                        {charRes.char}
                      </span>
                    </div>

                    {/* Full-width equal square writing practice boxes strictly clamped inside innerWidthPx */}
                    <div className="flex items-center min-w-0" style={{ gap: `${boxGapPx}px` }}>
                      {Array.from({ length: sheet.boxesPerRow }).map((_, boxIdx) => (
                        <WritingGridBox
                          key={`box-${boxIdx}`}
                          char={charRes.char}
                          boxIndex={boxIdx + 1}
                          isTrace={boxIdx < sheet.traceBoxes}
                          boxSizePx={boxSizePx}
                          gridStyle={settings.gridStyle}
                          gridBgColor={palette.gridBgColor}
                          gridBorderColor={palette.gridBorderColor}
                          gridBorderWidthPx={palette.gridBorderWidthPx}
                          gridLineColor={palette.gridLineColor}
                          gridDiagonalColor={palette.gridDiagonalColor}
                          traceGlyphColor={palette.traceGlyphColor}
                          fontFamily={settings.fontFamily}
                        />
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </section>
          );
        })}
      </div>

      {/* Footer */}
      <footer
        className="flex items-center justify-between gap-4 pt-1.5 border-t text-[10px] min-w-0"
        style={{
          borderTopColor: palette.dividerColor,
          color: palette.subtleInk,
        }}
      >
        <span className="truncate min-w-0" title={footerInstruction}>
          {footerInstruction}
        </span>
        <span className="font-mono tabular-nums whitespace-nowrap shrink-0">
          Page {sheet.pageNumber} of {totalPages}
        </span>
      </footer>
    </article>
  );
};

export const WorkbookCoverPreview: React.FC<{
  workbook: WorkbookProject;
  entries: WorkbookEntry[];
}> = ({ workbook, entries }) => {
  const { settings } = workbook;
  const palette = resolveWorkbookPalette(settings);
  const { widthPx, heightPx, paddingPx } = getSheetPreviewScale(settings);
  const specimen = entries.slice(0, 6).map((e) => e.text);
  const totalStrokes = entries.reduce((sum, e) => sum + e.totalStrokeCount, 0);
  const sheets = paginateWorkbookEntries(entries, settings);
  const totalPracticeBoxes = sheets.reduce((sum, s) => sum + s.totalBoxesOnPage, 0);

  return (
    <article
      className={`print-sheet mx-auto bg-white border flex flex-col justify-between text-center ${
        palette.isBw
          ? 'bw-print-sheet border-black shadow-none'
          : 'border-hairline shadow-[0_14px_38px_rgba(30,25,22,0.09)]'
      }`}
      style={{
        width: `${widthPx}px`,
        minHeight: `${heightPx}px`,
        padding: `${paddingPx * 1.4}px`,
        backgroundColor: palette.pageBg,
      }}
    >
      <div
        className="border flex-1 flex flex-col items-center justify-center px-12 py-12"
        style={{ borderColor: palette.dividerColor }}
      >
        <div
          className="text-xs font-semibold tracking-[0.14em] mb-5"
          style={{ color: palette.accentColor }}
        >
          CHINESE CHARACTER WRITING WORKBOOK
        </div>
        <h2
          className="font-display text-5xl font-bold max-w-2xl leading-tight m-0"
          style={{ color: palette.primaryInk }}
        >
          {workbook.title}
        </h2>
        <div className="w-24 h-0.5 my-6" style={{ backgroundColor: palette.accentColor }} />
        <p className="text-base max-w-xl m-0" style={{ color: palette.secondaryInk }}>
          {workbook.subtitle || 'Guided Stroke-Order & Calligraphy Practice Sheets'}
        </p>

        {specimen.length > 0 && (
          <div
            className="flex flex-wrap items-center justify-center gap-7 my-10"
            style={{
              fontFamily: settings.fontFamily,
              color: palette.accentColor,
              fontSize: '54px',
            }}
          >
            {specimen.map((ch, i) => (
              <span key={`${ch}-${i}`}>{ch}</span>
            ))}
          </div>
        )}

        <div
          className="mt-4 text-sm flex flex-wrap items-center justify-center gap-3 tabular-nums"
          style={{ color: palette.secondaryInk }}
        >
          <span>
            {entries.length} Vocabulary {entries.length === 1 ? 'Entry' : 'Entries'}
          </span>
          <span aria-hidden="true">·</span>
          <span>{totalPracticeBoxes} Practice Boxes</span>
          <span aria-hidden="true">·</span>
          <span>{totalStrokes} Total Strokes</span>
          <span aria-hidden="true">·</span>
          <span>
            {settings.pageSize} {settings.orientation === 'landscape' ? 'Landscape' : 'Portrait'}
          </span>
        </div>

        <div className="mt-2 text-xs" style={{ color: palette.subtleInk }}>
          {settings.authorName || 'Hanzi Book Studio Press'} · {new Date().toLocaleDateString()}
        </div>
      </div>
    </article>
  );
};

export const WorkbookIndexPreview: React.FC<{
  workbook: WorkbookProject;
  entries: WorkbookEntry[];
}> = ({ workbook, entries }) => {
  const { settings } = workbook;
  const palette = resolveWorkbookPalette(settings);
  const { widthPx, heightPx, paddingPx } = getSheetPreviewScale(settings);

  const sheets = paginateWorkbookEntries(entries, settings);
  const entryPageMap = new Map<string, number>();
  const entryBoxesPerRowMap = new Map<string, number>();
  for (const sheet of sheets) {
    for (const e of sheet.entries) {
      entryPageMap.set(e.id, sheet.pageNumber);
      entryBoxesPerRowMap.set(e.id, sheet.boxesPerRow);
    }
  }

  return (
    <article
      className={`print-sheet mx-auto bg-white border flex flex-col justify-between ${
        palette.isBw
          ? 'bw-print-sheet border-black shadow-none'
          : 'border-hairline shadow-[0_14px_38px_rgba(30,25,22,0.09)]'
      }`}
      style={{
        width: `${widthPx}px`,
        minHeight: `${heightPx}px`,
        padding: `${paddingPx}px`,
        backgroundColor: palette.pageBg,
      }}
    >
      <div>
        <div
          className="pb-3 mb-5 border-b-2 flex items-baseline justify-between"
          style={{ borderBottomColor: palette.accentColor }}
        >
          <h2 className="font-display text-3xl font-bold m-0" style={{ color: palette.primaryInk }}>
            Character Index
          </h2>
          <span className="text-xs font-mono tabular-nums" style={{ color: palette.subtleInk }}>
            {entries.length} entries · {sheets.length} practice{' '}
            {sheets.length === 1 ? 'sheet' : 'sheets'}
          </span>
        </div>

        <table className="w-full border-collapse text-sm">
          <thead>
            <tr
              className="text-left text-xs font-semibold border-b"
              style={{
                backgroundColor: palette.indexHeaderBg,
                color: palette.accentColor,
                borderBottomColor: palette.isBw ? '#000000' : '#CFC4B3',
              }}
            >
              <th className="py-2 px-3 w-12 text-center">#</th>
              <th className="py-2 px-3 w-24">Vocabulary</th>
              <th className="py-2 px-3 w-32">Pinyin</th>
              <th className="py-2 px-3">English Meaning</th>
              {settings.showArabicTranslation !== false && (
                <th className="py-2 px-3 w-40 text-right">Arabic (العربية)</th>
              )}
              <th className="py-2 px-3 w-20 text-center">Boxes</th>
              <th className="py-2 px-3 w-20 text-center">Strokes</th>
              <th className="py-2 px-3 w-16 text-center">Sheet</th>
            </tr>
          </thead>
          <tbody>
            {entries.slice(0, 22).map((entry, idx) => (
              <tr
                key={entry.id}
                className="border-b"
                style={{ borderBottomColor: palette.isBw ? '#000000' : '#E5DDD1' }}
              >
                <td
                  className="py-2 px-3 text-center font-mono text-xs tabular-nums"
                  style={{ color: palette.subtleInk }}
                >
                  {idx + 1}
                </td>
                <td
                  className="py-2 px-3 text-xl"
                  style={{ fontFamily: settings.fontFamily, color: palette.primaryInk }}
                >
                  {entry.text}
                </td>
                <td className="py-2 px-3 font-medium" style={{ color: palette.accentColor }}>
                  {entry.pinyin || '—'}
                </td>
                <td className="py-2 px-3" style={{ color: palette.secondaryInk }}>
                  {entry.meaning || '—'}
                </td>
                {settings.showArabicTranslation !== false && (
                  <td
                    dir="rtl"
                    className="py-2 px-3 text-right font-medium"
                    style={{ fontFamily: 'var(--font-arabic)', color: palette.primaryInk }}
                  >
                    {entry.arabicMeaning || '—'}
                  </td>
                )}
                <td
                  className="py-2 px-3 text-center font-mono text-xs tabular-nums"
                  style={{ color: palette.primaryInk }}
                >
                  {Math.max(1, entry.charResults.length) *
                    (entryBoxesPerRowMap.get(entry.id) ?? 10)}
                </td>
                <td
                  className="py-2 px-3 text-center font-mono text-xs tabular-nums"
                  style={{ color: palette.primaryInk }}
                >
                  {entry.totalStrokeCount}
                </td>
                <td
                  className="py-2 px-3 text-center font-mono text-xs tabular-nums"
                  style={{ color: palette.primaryInk }}
                >
                  {entryPageMap.get(entry.id) ?? 1}
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        {entries.length > 22 && (
          <p className="text-xs mt-3" style={{ color: palette.subtleInk }}>
            Showing first 22 entries in preview. All {entries.length} entries are included in the
            exported Word index.
          </p>
        )}
      </div>

      <footer
        className="flex items-center justify-between pt-2.5 border-t text-[11px]"
        style={{ borderTopColor: palette.dividerColor, color: palette.subtleInk }}
      >
        <span>{workbook.title}</span>
        <span>Character Index</span>
      </footer>
    </article>
  );
};
