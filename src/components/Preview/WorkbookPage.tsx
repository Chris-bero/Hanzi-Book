import React from 'react';
import {
  canInlineMultiCharStrokes,
  getPageDimensionsMm,
  PaginatedSheet,
  paginateWorkbookEntries,
} from '../../services/layoutPagination';
import { WorkbookEntry, WorkbookProject, WorkbookSettings } from '../../types/workbook';
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
  const isCompact = sheet.density === 'compact';
  const isBalanced = sheet.density === 'balanced';
  const isPortrait = settings.orientation === 'portrait';

  const boxSizePx = Math.round(sheet.boxSizeMm * pxPerMm);
  const boxGapPx = Math.max(4, Math.round(sheet.boxGapMm * pxPerMm));
  const stageSizePx = Math.round(sheet.stageSizeMm * pxPerMm);
  const entryGapPx = Math.max(6, Math.round(sheet.entryGapMm * pxPerMm));

  return (
    <article
      className="print-sheet mx-auto bg-white border border-hairline shadow-[0_14px_38px_rgba(30,25,22,0.09)] flex flex-col justify-between"
      style={{
        width: `${widthPx}px`,
        minHeight: `${heightPx}px`,
        padding: `${paddingPx}px`,
      }}
      aria-label={`Workbook page ${sheet.pageNumber}`}
    >
      {/* Top Running Header */}
      <header
        className="flex items-center justify-between pb-1.5 border-b-2 text-[10.5px] font-medium tracking-[0.05em] text-ink-subtle"
        style={{ borderBottomColor: settings.accentColor }}
      >
        <span className="truncate max-w-[52%] text-ink-muted font-semibold">{workbookTitle}</span>
        <span className="font-mono tabular-nums whitespace-nowrap">
          {settings.pageSize} {settings.orientation === 'landscape' ? 'Landscape' : 'Portrait'} ·{' '}
          {sheet.entries.length} {sheet.entries.length === 1 ? 'Entry' : 'Entries'} ·{' '}
          {sheet.boxesPerRow} Boxes/Char ({sheet.totalBoxesOnPage} Boxes)
        </span>
      </header>

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

          const heroSizePx = isCompact
            ? isMultiChar
              ? 23
              : 30
            : isBalanced
            ? isMultiChar
              ? 28
              : 38
            : Math.max(
                40,
                Math.min(
                  72,
                  isMultiChar ? settings.characterSizePx * 0.8 : settings.characterSizePx * 1.25
                )
              );

          const strokeLabel = isMultiChar
            ? `${entry.totalStrokeCount} strokes (${entry.charResults
                .map((c) => c.strokeCount)
                .join('+')})`
            : `${entry.totalStrokeCount} ${entry.totalStrokeCount === 1 ? 'stroke' : 'strokes'}`;

          return (
            <section
              key={entry.id}
              aria-label={`Vocabulary entry ${entry.text}`}
              className={`flex flex-col ${idx > 0 ? 'pt-2 border-t border-hairline' : ''}`}
              style={{ gap: isCompact ? '5px' : '9px' }}
            >
              {/* Upper Band: Vocabulary Info (Left) + Stroke-Order Diagrams (Right) */}
              <div
                className="grid items-center gap-3.5"
                style={{
                  gridTemplateColumns: isPortrait ? '205px 1fr' : '255px 1fr',
                }}
              >
                {/* Left Vocabulary Info */}
                <div className="pr-3 border-r border-hairline flex items-center gap-3 min-w-0">
                  <div
                    className="leading-none text-ink font-normal shrink-0"
                    style={{
                      fontFamily: settings.fontFamily,
                      fontSize: `${heroSizePx}px`,
                    }}
                  >
                    {entry.text}
                  </div>
                  <div className="min-w-0 flex-1">
                    <div className="flex items-baseline gap-2 flex-wrap">
                      <span
                        className="font-semibold leading-tight truncate"
                        style={{
                          color: settings.accentColor,
                          fontSize: isCompact ? '13px' : '16px',
                        }}
                      >
                        {entry.pinyin || '—'}
                      </span>
                      <span className="font-mono text-[9.5px] text-ink-subtle tabular-nums whitespace-nowrap">
                        {strokeLabel}
                      </span>
                    </div>
                    <p
                      className="text-ink-muted m-0 truncate"
                      title={entry.meaning}
                      style={{
                        fontSize: isCompact ? '10.5px' : '12.5px',
                        marginTop: '1px',
                      }}
                    >
                      {entry.meaning || 'No English meaning provided.'}
                    </p>
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
                          className="text-xs text-ink font-semibold shrink-0"
                          style={{ fontFamily: settings.fontFamily }}
                        >
                          {charRes.char}
                        </span>
                      )}
                      {charRes.status === 'ready' && charRes.strokeData ? (
                        <StrokeStageGrid
                          strokeData={charRes.strokeData}
                          stageSize={stageSizePx}
                          newStrokeColor={settings.newStrokeColor}
                          completedStrokeColor={settings.completedStrokeColor}
                          compact
                        />
                      ) : (
                        <span className="text-[11px] text-[#A51D24]">
                          Stroke-order vector unavailable for “{charRes.char}”
                        </span>
                      )}
                    </div>
                  ))}
                </div>
              </div>

              {/* Lower Band: Independent Full-Width Practice Row for EACH Character */}
              <div className="flex flex-col gap-1.5">
                {entry.charResults.map((charRes, rowIdx) => (
                  <div
                    key={`${entry.id}-practice-${rowIdx}`}
                    className="flex items-center gap-1.5"
                  >
                    {/* Character marker for each independent practice row */}
                    <div
                      className="shrink-0 flex items-center justify-center text-ink-subtle"
                      style={{ width: '20px' }}
                    >
                      <span
                        className="text-sm text-ink leading-none"
                        style={{ fontFamily: settings.fontFamily }}
                      >
                        {charRes.char}
                      </span>
                    </div>

                    {/* Full-width equal square writing practice boxes */}
                    <div className="flex items-center" style={{ gap: `${boxGapPx}px` }}>
                      {Array.from({ length: sheet.boxesPerRow }).map((_, boxIdx) => (
                        <WritingGridBox
                          key={`box-${boxIdx}`}
                          char={charRes.char}
                          boxIndex={boxIdx + 1}
                          isTrace={boxIdx < sheet.traceBoxes}
                          boxSizePx={boxSizePx}
                          gridStyle={settings.gridStyle}
                          gridBgColor={settings.gridBgColor}
                          gridLineColor={settings.gridLineColor}
                          traceGlyphColor={settings.traceGlyphColor}
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
      <footer className="flex items-center justify-between pt-1.5 border-t border-hairline text-[10.5px] text-ink-subtle">
        <span>
          {settings.authorName || 'Hanzi Book Studio'} · Practice: Trace boxes 1–{sheet.traceBoxes},
          then write from memory in boxes {sheet.traceBoxes + 1}–{sheet.boxesPerRow}
        </span>
        <span className="font-mono tabular-nums whitespace-nowrap">
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
  const { widthPx, heightPx, paddingPx } = getSheetPreviewScale(settings);
  const specimen = entries.slice(0, 6).map((e) => e.text);
  const totalStrokes = entries.reduce((sum, e) => sum + e.totalStrokeCount, 0);
  const sheets = paginateWorkbookEntries(entries, settings);
  const totalPracticeBoxes = sheets.reduce((sum, s) => sum + s.totalBoxesOnPage, 0);

  return (
    <article
      className="print-sheet mx-auto bg-white border border-hairline shadow-[0_14px_38px_rgba(30,25,22,0.09)] flex flex-col justify-between text-center"
      style={{
        width: `${widthPx}px`,
        minHeight: `${heightPx}px`,
        padding: `${paddingPx * 1.4}px`,
      }}
    >
      <div className="border border-hairline flex-1 flex flex-col items-center justify-center px-12 py-12">
        <div
          className="text-xs font-semibold tracking-[0.14em] mb-5"
          style={{ color: settings.accentColor }}
        >
          CHINESE CHARACTER WRITING WORKBOOK
        </div>
        <h2 className="font-display text-5xl font-bold text-ink max-w-2xl leading-tight m-0">
          {workbook.title}
        </h2>
        <div className="w-24 h-0.5 my-6" style={{ backgroundColor: settings.accentColor }} />
        <p className="text-base text-ink-muted max-w-xl m-0">
          {workbook.subtitle || 'Guided Stroke-Order & Calligraphy Practice Sheets'}
        </p>

        {specimen.length > 0 && (
          <div
            className="flex flex-wrap items-center justify-center gap-7 my-10"
            style={{
              fontFamily: settings.fontFamily,
              color: settings.accentColor,
              fontSize: '54px',
            }}
          >
            {specimen.map((ch, i) => (
              <span key={`${ch}-${i}`}>{ch}</span>
            ))}
          </div>
        )}

        <div className="mt-4 text-sm text-ink-muted flex flex-wrap items-center justify-center gap-3 tabular-nums">
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

        <div className="mt-2 text-xs text-ink-subtle">
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
      className="print-sheet mx-auto bg-white border border-hairline shadow-[0_14px_38px_rgba(30,25,22,0.09)] flex flex-col justify-between"
      style={{
        width: `${widthPx}px`,
        minHeight: `${heightPx}px`,
        padding: `${paddingPx}px`,
      }}
    >
      <div>
        <div
          className="pb-3 mb-5 border-b-2 flex items-baseline justify-between"
          style={{ borderBottomColor: settings.accentColor }}
        >
          <h2 className="font-display text-3xl font-bold text-ink m-0">Character Index</h2>
          <span className="text-xs text-ink-subtle font-mono tabular-nums">
            {entries.length} entries · {sheets.length} practice{' '}
            {sheets.length === 1 ? 'sheet' : 'sheets'}
          </span>
        </div>

        <table className="w-full border-collapse text-sm">
          <thead>
            <tr
              className="text-left text-xs font-semibold border-b border-hairline-strong"
              style={{ backgroundColor: settings.gridBgColor, color: settings.accentColor }}
            >
              <th className="py-2 px-3 w-12 text-center">#</th>
              <th className="py-2 px-3 w-28">Vocabulary</th>
              <th className="py-2 px-3 w-36">Pinyin</th>
              <th className="py-2 px-3">English Meaning</th>
              <th className="py-2 px-3 w-24 text-center">Boxes</th>
              <th className="py-2 px-3 w-20 text-center">Strokes</th>
              <th className="py-2 px-3 w-20 text-center">Sheet</th>
            </tr>
          </thead>
          <tbody>
            {entries.slice(0, 22).map((entry, idx) => (
              <tr key={entry.id} className="border-b border-hairline">
                <td className="py-2 px-3 text-center font-mono text-xs text-ink-subtle tabular-nums">
                  {idx + 1}
                </td>
                <td
                  className="py-2 px-3 text-xl text-ink"
                  style={{ fontFamily: settings.fontFamily }}
                >
                  {entry.text}
                </td>
                <td className="py-2 px-3 font-medium" style={{ color: settings.accentColor }}>
                  {entry.pinyin || '—'}
                </td>
                <td className="py-2 px-3 text-ink-muted">{entry.meaning || '—'}</td>
                <td className="py-2 px-3 text-center font-mono text-xs tabular-nums">
                  {Math.max(1, entry.charResults.length) *
                    (entryBoxesPerRowMap.get(entry.id) ?? 10)}
                </td>
                <td className="py-2 px-3 text-center font-mono text-xs tabular-nums">
                  {entry.totalStrokeCount}
                </td>
                <td className="py-2 px-3 text-center font-mono text-xs tabular-nums">
                  {entryPageMap.get(entry.id) ?? 1}
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        {entries.length > 22 && (
          <p className="text-xs text-ink-subtle mt-3">
            Showing first 22 entries in preview. All {entries.length} entries are included in the
            exported Word index.
          </p>
        )}
      </div>

      <footer className="flex items-center justify-between pt-2.5 border-t border-hairline text-[11px] text-ink-subtle">
        <span>{workbook.title}</span>
        <span>Character Index</span>
      </footer>
    </article>
  );
};
