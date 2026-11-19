import React, { useMemo, useState } from 'react';
import {
  ChevronLeft,
  ChevronRight,
  Download,
  FileDown,
  FileText,
  Printer,
  SlidersHorizontal,
} from 'lucide-react';
import { DocxProgress, generateWorkbookDocx } from '../../services/docxGenerator';
import {
  MANDATORY_EMPTY_BOXES,
  MANDATORY_TOTAL_BOXES,
  MANDATORY_TRACE_BOXES,
  paginateWorkbookEntries,
} from '../../services/layoutPagination';
import { generateWorkbookPdf, PdfProgress } from '../../services/pdfGenerator';
import { ExportRecord, WorkbookProject } from '../../types/workbook';
import {
  WorkbookCoverPreview,
  WorkbookIndexPreview,
  WorkbookPage,
} from './WorkbookPage';

interface WorkbookPreviewViewProps {
  workbook: WorkbookProject | null;
  selectedEntryIds: Set<string>;
  onExportComplete: (record: ExportRecord) => void;
  onGoToCharacters: () => void;
  onGoToSettings: () => void;
}

type PreviewSection = 'single' | 'all' | 'cover' | 'index';

export const WorkbookPreviewView: React.FC<WorkbookPreviewViewProps> = ({
  workbook,
  selectedEntryIds,
  onExportComplete,
  onGoToCharacters,
  onGoToSettings,
}) => {
  const [sectionMode, setSectionMode] = useState<PreviewSection>('single');
  const [pageIndex, setPageIndex] = useState(0);
  const [exportScope, setExportScope] = useState<'all' | 'selected'>('all');
  const [exportProgress, setExportProgress] = useState<DocxProgress | PdfProgress | null>(null);
  const [exportNotice, setExportNotice] = useState<{ type: 'ok' | 'error'; text: string } | null>(
    null
  );

  const activeEntries = useMemo(() => {
    if (!workbook) return [];
    if (exportScope === 'selected' && selectedEntryIds.size > 0) {
      return workbook.entries.filter((e) => selectedEntryIds.has(e.id));
    }
    return workbook.entries;
  }, [workbook, exportScope, selectedEntryIds]);

  const paginatedSheets = useMemo(() => {
    if (!workbook) return [];
    return paginateWorkbookEntries(activeEntries, workbook.settings);
  }, [workbook, activeEntries]);

  if (!workbook || workbook.entries.length === 0) {
    return (
      <div className="bg-white border border-hairline p-12 text-center">
        <FileText className="w-9 h-9 text-ink-subtle mx-auto mb-3" />
        <h2 className="font-display text-3xl font-bold text-ink m-0">
          No Characters in Active Workbook
        </h2>
        <p className="mt-2 text-sm text-ink-muted max-w-md mx-auto">
          Add Chinese characters or words such as 不, 我, 你, 好, and 你好 to inspect the formatted
          practice sheets and export to Microsoft Word.
        </p>
        <button
          type="button"
          onClick={onGoToCharacters}
          className="mt-5 px-5 py-2.5 text-sm font-semibold bg-burgundy text-white hover:bg-burgundy-dark cursor-pointer"
        >
          Open Character Management
        </button>
      </div>
    );
  }

  const safePageIndex = Math.min(pageIndex, Math.max(0, paginatedSheets.length - 1));
  const currentSheet = paginatedSheets[safePageIndex];

  const handleExportDocx = async () => {
    if (exportProgress) return;
    setExportNotice(null);
    setExportProgress({
      current: 0,
      total: activeEntries.length,
      message: 'Preparing Word document…',
    });
    try {
      const record = await generateWorkbookDocx(workbook, activeEntries, (p) =>
        setExportProgress(p)
      );
      onExportComplete(record);
      setExportNotice({
        type: 'ok',
        text: `Downloaded editable Word workbook “${record.fileName}” (${record.entryCount} entries · 10 editable boxes per character).`,
      });
    } catch (err) {
      setExportNotice({
        type: 'error',
        text: err instanceof Error ? err.message : 'Word export failed.',
      });
    } finally {
      setExportProgress(null);
    }
  };

  const handleExportPdf = async () => {
    if (exportProgress) return;
    setExportNotice(null);
    setExportProgress({
      current: 0,
      total: activeEntries.length,
      message: 'Preparing PDF workbook…',
    });
    try {
      const record = await generateWorkbookPdf(workbook, activeEntries, (p) =>
        setExportProgress(p)
      );
      onExportComplete(record);
      setExportNotice({
        type: 'ok',
        text: `Downloaded printable PDF workbook “${record.fileName}” (${record.entryCount} entries).`,
      });
    } catch (err) {
      setExportNotice({
        type: 'error',
        text: err instanceof Error ? err.message : 'PDF export failed.',
      });
    } finally {
      setExportProgress(null);
    }
  };

  const pct =
    exportProgress && exportProgress.total > 0
      ? Math.round((exportProgress.current / exportProgress.total) * 100)
      : 0;

  return (
    <div className="space-y-6">
      {/* Header & Export Actions */}
      <div className="no-print flex flex-wrap items-end justify-between gap-4 pb-5 border-b border-hairline">
        <div>
          <h1 className="font-display text-4xl font-bold text-ink leading-tight m-0">
            Workbook Page Preview
          </h1>
          <div className="mt-1.5 flex flex-wrap items-center gap-2 text-xs text-ink-muted tabular-nums">
            <span className="font-semibold text-ink">{workbook.title}</span>
            <span aria-hidden="true">·</span>
            <span>
              {activeEntries.length} {activeEntries.length === 1 ? 'entry' : 'entries'} across{' '}
              {paginatedSheets.length} {paginatedSheets.length === 1 ? 'sheet' : 'sheets'}
            </span>
            <span aria-hidden="true">·</span>
            <span>
              {currentSheet
                ? `${currentSheet.boxesPerRow} boxes/char (${currentSheet.traceBoxes} trace · ${currentSheet.emptyBoxes} empty) · ${currentSheet.totalBoxesOnPage} boxes on page`
                : `${MANDATORY_TOTAL_BOXES}+ boxes/char (${MANDATORY_TRACE_BOXES} trace · ${MANDATORY_EMPTY_BOXES}+ empty)`}
            </span>
            <span aria-hidden="true">·</span>
            <span>
              {workbook.settings.pageSize} {workbook.settings.orientation}
            </span>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          <button
            type="button"
            onClick={onGoToSettings}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium border border-hairline-strong bg-white text-ink hover:bg-ivory-deep transition-colors cursor-pointer whitespace-nowrap"
          >
            <SlidersHorizontal className="w-3.5 h-3.5" />
            Customize Layout
          </button>

          <button
            type="button"
            onClick={() => window.print()}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium border border-hairline-strong bg-white text-ink hover:bg-ivory-deep transition-colors cursor-pointer whitespace-nowrap"
          >
            <Printer className="w-3.5 h-3.5" />
            Print
          </button>

          <button
            type="button"
            onClick={handleExportPdf}
            disabled={Boolean(exportProgress)}
            className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold border border-burgundy text-burgundy bg-white hover:bg-burgundy-tint disabled:opacity-50 transition-colors cursor-pointer whitespace-nowrap"
          >
            <FileDown className="w-3.5 h-3.5" />
            Export .PDF
          </button>

          <button
            type="button"
            onClick={handleExportDocx}
            disabled={Boolean(exportProgress)}
            className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold bg-burgundy text-white hover:bg-burgundy-dark disabled:opacity-50 transition-colors cursor-pointer whitespace-nowrap"
          >
            <Download className="w-3.5 h-3.5" />
            Export Editable .DOCX
          </button>
        </div>
      </div>

      {/* Live Generation Progress */}
      {exportProgress && (
        <div className="no-print bg-white border border-hairline p-4 space-y-2">
          <div className="flex items-center justify-between text-xs">
            <span className="font-medium text-ink">{exportProgress.message}</span>
            <span className="font-mono text-ink-muted tabular-nums">{pct}%</span>
          </div>
          <div className="w-full h-1.5 bg-ivory-deep overflow-hidden">
            <div
              className="h-full bg-burgundy transition-all duration-150"
              style={{ width: `${Math.max(4, pct)}%` }}
            />
          </div>
        </div>
      )}

      {/* Export Result Notice */}
      {exportNotice && (
        <div
          className={`no-print px-4 py-3 text-xs border ${
            exportNotice.type === 'ok'
              ? 'bg-[#EEF6F0] border-[#BBD9C4] text-[#1F6136]'
              : 'bg-[#FDF4F4] border-[#EBC6C8] text-[#A51D24]'
          }`}
        >
          {exportNotice.text}
        </div>
      )}

      {/* Preview Navigation Toolbar */}
      <div className="no-print bg-white border border-hairline px-5 py-3.5 flex flex-wrap items-center justify-between gap-4">
        <div className="inline-flex bg-ivory p-0.5 border border-hairline">
          {(
            [
              ['cover', 'Cover Page'],
              ['index', 'Character Index'],
              ['single', 'Single Sheet'],
              ['all', `All Sheets (${paginatedSheets.length})`],
            ] as Array<[PreviewSection, string]>
          ).map(([mode, label]) => (
            <button
              key={mode}
              type="button"
              onClick={() => setSectionMode(mode)}
              className={`px-3 py-1.5 text-xs font-medium transition-colors cursor-pointer whitespace-nowrap ${
                sectionMode === mode
                  ? 'bg-white text-burgundy font-semibold'
                  : 'text-ink-muted hover:text-ink'
              }`}
            >
              {label}
            </button>
          ))}
        </div>

        <div className="flex flex-wrap items-center gap-4">
          {selectedEntryIds.size > 0 && (
            <div className="inline-flex bg-ivory p-0.5 border border-hairline">
              <button
                type="button"
                onClick={() => {
                  setExportScope('all');
                  setPageIndex(0);
                }}
                className={`px-2.5 py-1 text-xs font-medium cursor-pointer whitespace-nowrap ${
                  exportScope === 'all' ? 'bg-white text-ink font-semibold' : 'text-ink-muted'
                }`}
              >
                Entire Book ({workbook.entries.length})
              </button>
              <button
                type="button"
                onClick={() => {
                  setExportScope('selected');
                  setPageIndex(0);
                }}
                className={`px-2.5 py-1 text-xs font-medium cursor-pointer whitespace-nowrap ${
                  exportScope === 'selected'
                    ? 'bg-white text-burgundy font-semibold'
                    : 'text-ink-muted'
                }`}
              >
                Selected Only ({selectedEntryIds.size})
              </button>
            </div>
          )}

          {sectionMode === 'single' && paginatedSheets.length > 0 && (
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => setPageIndex((p) => Math.max(0, p - 1))}
                disabled={safePageIndex === 0}
                className="p-1.5 border border-hairline-strong bg-white text-ink hover:bg-ivory-deep disabled:opacity-40 cursor-pointer"
                aria-label="Previous sheet"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              <span className="font-mono text-xs text-ink tabular-nums px-1.5">
                {safePageIndex + 1} / {paginatedSheets.length}
              </span>
              <button
                type="button"
                onClick={() => setPageIndex((p) => Math.min(paginatedSheets.length - 1, p + 1))}
                disabled={safePageIndex >= paginatedSheets.length - 1}
                className="p-1.5 border border-hairline-strong bg-white text-ink hover:bg-ivory-deep disabled:opacity-40 cursor-pointer"
                aria-label="Next sheet"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Quick Sheet Jump Strip */}
      {sectionMode === 'single' && paginatedSheets.length > 1 && (
        <div className="no-print flex items-center gap-1.5 flex-wrap">
          {paginatedSheets.map((sheet, idx) => (
            <button
              key={`jump-${idx}`}
              type="button"
              onClick={() => setPageIndex(idx)}
              style={{ fontFamily: workbook.settings.fontFamily }}
              className={`px-3 py-1.5 text-sm border transition-colors cursor-pointer ${
                idx === safePageIndex
                  ? 'border-burgundy bg-burgundy text-white'
                  : 'border-hairline bg-white text-ink hover:border-hairline-strong'
              }`}
            >
              <span className="font-mono text-[10px] mr-1.5 opacity-75">P{sheet.pageNumber}</span>
              {sheet.entries.map((e) => e.text).join(' · ')}
            </button>
          ))}
        </div>
      )}

      {/* Rendered Sheet Canvas */}
      <div className="w-full overflow-x-auto pb-8 pt-2">
        {sectionMode === 'cover' && (
          <WorkbookCoverPreview workbook={workbook} entries={activeEntries} />
        )}

        {sectionMode === 'index' && (
          <WorkbookIndexPreview workbook={workbook} entries={activeEntries} />
        )}

        {sectionMode === 'single' && currentSheet && (
          <WorkbookPage
            workbookTitle={workbook.title}
            sheet={currentSheet}
            settings={workbook.settings}
            totalPages={paginatedSheets.length}
          />
        )}

        {sectionMode === 'all' && (
          <div className="flex flex-col gap-8">
            {workbook.settings.includeCoverPage && (
              <WorkbookCoverPreview workbook={workbook} entries={activeEntries} />
            )}
            {workbook.settings.includeTableOfContents && (
              <WorkbookIndexPreview workbook={workbook} entries={activeEntries} />
            )}
            {paginatedSheets.map((sheet) => (
              <WorkbookPage
                key={`sheet-${sheet.pageNumber}`}
                workbookTitle={workbook.title}
                sheet={sheet}
                settings={workbook.settings}
                totalPages={paginatedSheets.length}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  );
};
