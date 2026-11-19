import React, { useEffect, useRef, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  Pause,
  Play,
  RotateCcw,
  Upload,
} from 'lucide-react';
import {
  createWorkbookEntryFromItem,
  downloadImportTemplate,
  parseSpreadsheetFile,
} from '../../services/excelImportService';
import { extractChineseCharacters } from '../../services/strokeEngine';
import { repository } from '../../storage/db';
import {
  BatchQueueItem,
  PersistedBatchJob,
  WorkbookEntry,
  WorkbookProject,
} from '../../types/workbook';

interface BatchImportViewProps {
  workbook: WorkbookProject | null;
  onAppendEntries: (entries: WorkbookEntry[]) => Promise<void>;
  onGoToCharacters: () => void;
  onGoToPreview: () => void;
}

const BATCH_META_KEY = 'active_batch_job_v1';

export const BatchImportView: React.FC<BatchImportViewProps> = ({
  workbook,
  onAppendEntries,
  onGoToCharacters,
  onGoToPreview,
}) => {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const pauseRequestedRef = useRef(false);

  const [queuedItems, setQueuedItems] = useState<BatchQueueItem[]>([]);
  const [sourceSummary, setSourceSummary] = useState<string | null>(null);
  const [pasteText, setPasteText] = useState('');
  const [parseError, setParseError] = useState<string | null>(null);

  const [isRunning, setIsRunning] = useState(false);
  const [isPaused, setIsPaused] = useState(false);
  const [processedCount, setProcessedCount] = useState(0);
  const [cacheHits, setCacheHits] = useState(0);
  const [networkFetches, setNetworkFetches] = useState(0);
  const [unsupportedList, setUnsupportedList] = useState<Array<{ text: string; reason: string }>>(
    []
  );
  const [resumableJob, setResumableJob] = useState<PersistedBatchJob | null>(null);
  const [completedBanner, setCompletedBanner] = useState<string | null>(null);

  // Check for interrupted batch job saved in IndexedDB
  useEffect(() => {
    const loadSavedJob = async () => {
      const raw = await repository.getMeta(BATCH_META_KEY);
      if (!raw) return;
      try {
        const parsed = JSON.parse(raw) as PersistedBatchJob;
        if (parsed.items.length > parsed.processedIndex) {
          setResumableJob(parsed);
        }
      } catch {
        // Ignore malformed meta
      }
    };
    loadSavedJob();
  }, []);

  if (!workbook) {
    return (
      <div className="bg-white border border-hairline p-12 text-center">
        <h2 className="font-display text-3xl font-bold text-ink m-0">No Active Workbook</h2>
        <p className="mt-2 text-sm text-ink-muted">
          Create or select a workbook before importing an Excel or CSV character list.
        </p>
      </div>
    );
  }

  const handleFileSelected = async (file: File) => {
    setParseError(null);
    setCompletedBanner(null);
    try {
      const parsed = await parseSpreadsheetFile(file);
      if (parsed.items.length === 0) {
        setParseError('No Chinese characters were found in this file.');
        return;
      }
      setQueuedItems(parsed.items);
      setProcessedCount(0);
      setCacheHits(0);
      setNetworkFetches(0);
      setUnsupportedList([]);
      setIsPaused(false);
      setSourceSummary(
        `${file.name} · ${parsed.items.length} valid entries (${parsed.skippedRows} rows skipped) · Columns: ${parsed.detectedColumns.characterCol}${
          parsed.detectedColumns.pinyinCol ? `, ${parsed.detectedColumns.pinyinCol}` : ''
        }${parsed.detectedColumns.meaningCol ? `, ${parsed.detectedColumns.meaningCol}` : ''}`
      );
    } catch (err) {
      setParseError(err instanceof Error ? err.message : 'Unable to parse file.');
    }
  };

  const handleStagePastedList = () => {
    setParseError(null);
    setCompletedBanner(null);
    const lines = pasteText
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);

    const parsed: BatchQueueItem[] = [];
    for (const line of lines) {
      // Support tab/comma separated "Character, Pinyin, Meaning" or space-separated tokens
      if (line.includes('\t') || line.includes(',')) {
        const parts = line.split(/[\t,]/).map((p) => p.trim());
        const chars = extractChineseCharacters(parts[0] ?? '').join('');
        if (chars) {
          parsed.push({
            text: chars,
            pinyin: parts[1] || undefined,
            meaning: parts.slice(2).join(', ') || undefined,
          });
        }
      } else {
        const tokens = line
          .split(/[\s，;；]+/)
          .map((t) => extractChineseCharacters(t).join(''))
          .filter(Boolean);
        for (const token of tokens) {
          parsed.push({ text: token });
        }
      }
    }

    if (parsed.length === 0) {
      setParseError('No Chinese characters detected in the pasted list.');
      return;
    }

    setQueuedItems(parsed);
    setProcessedCount(0);
    setCacheHits(0);
    setNetworkFetches(0);
    setUnsupportedList([]);
    setIsPaused(false);
    setSourceSummary(`Pasted batch list · ${parsed.length} entries ready for processing`);
  };

  const runBatchFromIndex = async (
    items: BatchQueueItem[],
    startIndex: number,
    initialCacheHits: number,
    initialFetches: number,
    initialUnsupported: Array<{ text: string; reason: string }>
  ) => {
    setIsRunning(true);
    setIsPaused(false);
    setCompletedBanner(null);
    pauseRequestedRef.current = false;

    let idx = startIndex;
    let totalHits = initialCacheHits;
    let totalFetches = initialFetches;
    const unsupported = [...initialUnsupported];
    const chunkBuffer: WorkbookEntry[] = [];
    const baseOrder = workbook.entries.length;

    const persistJobState = async (currentIndex: number) => {
      const job: PersistedBatchJob = {
        workbookId: workbook.id,
        items,
        processedIndex: currentIndex,
        unsupported,
        cacheHits: totalHits,
        networkFetches: totalFetches,
        updatedAt: Date.now(),
      };
      await repository.setMeta(BATCH_META_KEY, JSON.stringify(job));
      setResumableJob(job);
    };

    try {
      while (idx < items.length) {
        if (pauseRequestedRef.current) {
          if (chunkBuffer.length > 0) {
            await onAppendEntries(chunkBuffer.splice(0, chunkBuffer.length));
          }
          await persistJobState(idx);
          setIsPaused(true);
          return;
        }

        const item = items[idx];
        const res = await createWorkbookEntryFromItem(item, baseOrder + idx);
        totalHits += res.cacheHits;
        totalFetches += res.networkFetches;

        if (res.entry) {
          chunkBuffer.push(res.entry);
          if (res.entry.status !== 'ready') {
            const missingChars = res.entry.charResults
              .filter((c) => c.status !== 'ready')
              .map((c) => c.char)
              .join(', ');
            unsupported.push({
              text: res.entry.text,
              reason: `Missing stroke vector data for: ${missingChars}`,
            });
          }
        }

        idx++;
        setProcessedCount(idx);
        setCacheHits(totalHits);
        setNetworkFetches(totalFetches);
        setUnsupportedList([...unsupported]);

        // Commit to workbook & IndexedDB every 10 items to avoid memory spikes on 500+ books
        if (chunkBuffer.length >= 10 || idx === items.length) {
          await onAppendEntries(chunkBuffer.splice(0, chunkBuffer.length));
          if (idx < items.length) {
            await persistJobState(idx);
          }
          await new Promise((resolve) => setTimeout(resolve, 0));
        }
      }

      // Finished cleanly
      await repository.deleteMeta(BATCH_META_KEY);
      setResumableJob(null);
      setCompletedBanner(
        `Processed all ${items.length} entries (${totalHits} cache hits, ${totalFetches} fetched from Hanzi Writer Data, ${unsupported.length} warnings).`
      );
    } finally {
      setIsRunning(false);
    }
  };

  const handleStartOrResume = () => {
    if (queuedItems.length === 0) return;
    runBatchFromIndex(
      queuedItems,
      processedCount,
      cacheHits,
      networkFetches,
      unsupportedList
    );
  };

  const handleResumeSavedJob = () => {
    if (!resumableJob) return;
    setQueuedItems(resumableJob.items);
    setProcessedCount(resumableJob.processedIndex);
    setCacheHits(resumableJob.cacheHits);
    setNetworkFetches(resumableJob.networkFetches);
    setUnsupportedList(resumableJob.unsupported);
    setSourceSummary(
      `Resumed interrupted batch · ${resumableJob.processedIndex} / ${resumableJob.items.length} completed`
    );
    runBatchFromIndex(
      resumableJob.items,
      resumableJob.processedIndex,
      resumableJob.cacheHits,
      resumableJob.networkFetches,
      resumableJob.unsupported
    );
  };

  const handleDiscardJob = async () => {
    await repository.deleteMeta(BATCH_META_KEY);
    setResumableJob(null);
    setQueuedItems([]);
    setProcessedCount(0);
    setUnsupportedList([]);
    setSourceSummary(null);
    setIsPaused(false);
  };

  const progressPct =
    queuedItems.length > 0 ? Math.round((processedCount / queuedItems.length) * 100) : 0;

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-wrap items-end justify-between gap-4 pb-5 border-b border-hairline">
        <div>
          <h1 className="font-display text-4xl font-bold text-ink leading-tight m-0">
            Import Excel / CSV &amp; Batch Processor
          </h1>
          <p className="mt-1.5 text-sm text-ink-muted m-0">
            Import 100–500+ characters into <strong>{workbook.title}</strong> with automatic
            IndexedDB stroke caching, pause/resume support, and missing-character diagnostics.
          </p>
        </div>

        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={() => downloadImportTemplate('xlsx')}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium border border-hairline-strong bg-white text-ink hover:bg-ivory-deep cursor-pointer whitespace-nowrap"
          >
            <Download className="w-3.5 h-3.5" />
            Sample .XLSX
          </button>
          <button
            type="button"
            onClick={() => downloadImportTemplate('csv')}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium border border-hairline-strong bg-white text-ink hover:bg-ivory-deep cursor-pointer whitespace-nowrap"
          >
            <Download className="w-3.5 h-3.5" />
            Sample .CSV
          </button>
        </div>
      </div>

      {/* Resumable Interrupted Batch Notice */}
      {resumableJob && !isRunning && queuedItems.length === 0 && (
        <div className="bg-[#FFF8EB] border border-[#E5C185] p-5 flex flex-wrap items-center justify-between gap-4">
          <div className="text-xs text-[#7D4300]">
            <div className="font-semibold text-sm">Interrupted Batch Ready to Resume</div>
            <p className="mt-1 m-0 tabular-nums">
              {resumableJob.processedIndex} of {resumableJob.items.length} entries completed before
              pausing. Resume from where you left off without re-fetching cached characters.
            </p>
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={handleResumeSavedJob}
              className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold bg-burgundy text-white hover:bg-burgundy-dark cursor-pointer"
            >
              <Play className="w-3.5 h-3.5" />
              Resume Batch
            </button>
            <button
              type="button"
              onClick={handleDiscardJob}
              className="px-3 py-2 text-xs font-medium border border-hairline-strong bg-white text-ink hover:bg-ivory-deep cursor-pointer"
            >
              Discard
            </button>
          </div>
        </div>
      )}

      {/* Two-Column Source Selector */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Excel / CSV Dropzone */}
        <section className="bg-white border border-hairline p-6 flex flex-col justify-between">
          <div>
            <h2 className="font-display text-2xl font-bold text-ink m-0">
              Upload Spreadsheet (.xlsx, .xls, .csv)
            </h2>
            <p className="mt-1 text-xs text-ink-muted">
              Columns are auto-detected for <strong>Character / Hanzi</strong>,{' '}
              <strong>Pinyin</strong> (optional), and <strong>Meaning / English</strong> (optional).
            </p>

            <div
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                const file = e.dataTransfer.files?.[0];
                if (file) handleFileSelected(file);
              }}
              onClick={() => fileInputRef.current?.click()}
              className="mt-4 border border-dashed border-hairline-strong bg-ivory/70 hover:bg-ivory p-8 text-center cursor-pointer transition-colors"
            >
              <Upload className="w-6 h-6 text-burgundy mx-auto mb-2.5" />
              <div className="text-sm font-semibold text-ink">
                Click to choose an Excel or CSV file
              </div>
              <p className="mt-1 text-xs text-ink-subtle m-0">
                Or drag and drop .xlsx, .xls, or .csv here
              </p>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                className="hidden"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) handleFileSelected(file);
                  e.target.value = '';
                }}
              />
            </div>
          </div>

          {parseError && (
            <div className="mt-4 text-xs text-[#A51D24] bg-[#FDF4F4] border border-[#EBC6C8] px-3.5 py-2.5">
              {parseError}
            </div>
          )}
        </section>

        {/* High-Volume Paste Staging */}
        <section className="bg-white border border-hairline p-6 flex flex-col justify-between">
          <div>
            <h2 className="font-display text-2xl font-bold text-ink m-0">
              High-Volume Paste Queue
            </h2>
            <p className="mt-1 text-xs text-ink-muted">
              Paste 100–500+ characters or tab/comma-separated rows (Character, Pinyin, Meaning).
            </p>
            <textarea
              rows={6}
              value={pasteText}
              onChange={(e) => setPasteText(e.target.value)}
              placeholder={
                '不, bù, no / not\n我, wǒ, I / me\n你, nǐ, you\n好, hǎo, good\n你好, nǐ hǎo, hello'
              }
              className="mt-4 w-full border border-hairline-strong bg-white px-3.5 py-2.5 text-sm text-ink focus-visible:outline-2 focus-visible:outline-burgundy"
            />
          </div>

          <div className="mt-4 flex items-center justify-between">
            <span className="text-xs text-ink-subtle">
              Pinyin &amp; meanings auto-fill when omitted
            </span>
            <button
              type="button"
              onClick={handleStagePastedList}
              className="px-4 py-2 text-xs font-semibold border border-hairline-strong bg-white text-ink hover:bg-ivory-deep cursor-pointer"
            >
              Stage List for Batch
            </button>
          </div>
        </section>
      </div>

      {/* Staged Queue & Live Batch Runner */}
      {queuedItems.length > 0 && (
        <section className="bg-white border border-hairline p-6 space-y-5">
          <div className="flex flex-wrap items-center justify-between gap-4 pb-4 border-b border-hairline">
            <div>
              <div className="flex items-center gap-2 text-ink font-semibold text-base">
                <FileSpreadsheet className="w-4 h-4 text-burgundy" />
                <span>Batch Queue Ready</span>
              </div>
              {sourceSummary && (
                <div className="mt-1 text-xs text-ink-muted">{sourceSummary}</div>
              )}
            </div>

            <div className="flex items-center gap-2.5">
              {!isRunning ? (
                <button
                  type="button"
                  onClick={handleStartOrResume}
                  disabled={processedCount >= queuedItems.length}
                  className="inline-flex items-center gap-2 px-5 py-2.5 text-xs font-semibold bg-burgundy text-white hover:bg-burgundy-dark disabled:opacity-50 cursor-pointer whitespace-nowrap"
                >
                  <Play className="w-3.5 h-3.5" />
                  {isPaused
                    ? `Resume (${processedCount}/${queuedItems.length})`
                    : `Process ${queuedItems.length} Entries`}
                </button>
              ) : (
                <button
                  type="button"
                  onClick={() => {
                    pauseRequestedRef.current = true;
                  }}
                  className="inline-flex items-center gap-2 px-4 py-2.5 text-xs font-semibold border border-burgundy text-burgundy bg-white hover:bg-burgundy-tint cursor-pointer whitespace-nowrap"
                >
                  <Pause className="w-3.5 h-3.5" />
                  Pause Safely
                </button>
              )}

              <button
                type="button"
                onClick={handleDiscardJob}
                disabled={isRunning}
                className="inline-flex items-center gap-1.5 px-3 py-2.5 text-xs font-medium border border-hairline-strong text-ink-muted hover:text-ink disabled:opacity-40 cursor-pointer"
              >
                <RotateCcw className="w-3.5 h-3.5" />
                Clear
              </button>
            </div>
          </div>

          {/* Progress Bar & Telemetry */}
          <div className="space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2 text-xs tabular-nums">
              <span className="font-medium text-ink">
                Processed {processedCount} of {queuedItems.length} entries ({progressPct}%)
              </span>
              <div className="flex items-center gap-3 text-ink-muted font-mono">
                <span>Cache hits: {cacheHits}</span>
                <span aria-hidden="true">·</span>
                <span>Fetched: {networkFetches}</span>
                <span aria-hidden="true">·</span>
                <span>Warnings: {unsupportedList.length}</span>
              </div>
            </div>
            <div className="w-full h-2 bg-ivory-deep overflow-hidden">
              <div
                className="h-full bg-burgundy transition-all duration-150"
                style={{ width: `${progressPct}%` }}
              />
            </div>
          </div>

          {completedBanner && (
            <div className="bg-[#EEF6F0] border border-[#BBD9C4] px-4 py-3 flex flex-wrap items-center justify-between gap-3 text-xs text-[#1F6136]">
              <div className="flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 shrink-0" />
                <span>{completedBanner}</span>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={onGoToCharacters}
                  className="underline font-semibold cursor-pointer"
                >
                  Review Characters
                </button>
                <span aria-hidden="true">·</span>
                <button
                  type="button"
                  onClick={onGoToPreview}
                  className="underline font-semibold cursor-pointer"
                >
                  Preview &amp; Export Book
                </button>
              </div>
            </div>
          )}

          {/* Unsupported Characters Diagnostic Report */}
          {unsupportedList.length > 0 && (
            <div className="border border-[#E5C185] bg-[#FFF8EB] p-4 space-y-2">
              <div className="flex items-center gap-2 text-xs font-semibold text-[#7D4300]">
                <AlertTriangle className="w-4 h-4" />
                <span>Unsupported Character Report ({unsupportedList.length})</span>
              </div>
              <div className="max-h-40 overflow-y-auto divide-y divide-[#F0D9B0] text-xs">
                {unsupportedList.map((u, idx) => (
                  <div key={`${u.text}-${idx}`} className="py-1.5 flex items-center justify-between">
                    <span className="font-semibold text-ink">{u.text}</span>
                    <span className="text-[#7D4300]">{u.reason}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Preview of first 12 queued rows */}
          <div className="overflow-x-auto border border-hairline">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-ivory border-b border-hairline text-left text-ink-subtle">
                  <th className="py-2 px-3 w-12 text-center">#</th>
                  <th className="py-2 px-3 w-28">Entry</th>
                  <th className="py-2 px-3 w-36">Pinyin</th>
                  <th className="py-2 px-3">Meaning</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-hairline">
                {queuedItems.slice(0, 12).map((item, i) => (
                  <tr key={`q-${i}`}>
                    <td className="py-2 px-3 text-center font-mono text-ink-subtle tabular-nums">
                      {i + 1}
                    </td>
                    <td className="py-2 px-3 text-base text-ink">{item.text}</td>
                    <td className="py-2 px-3 text-burgundy">{item.pinyin || 'Auto-lookup'}</td>
                    <td className="py-2 px-3 text-ink-muted">{item.meaning || 'Auto-lookup'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {queuedItems.length > 12 && (
              <div className="px-3 py-2 text-[11px] text-ink-subtle bg-ivory/60 border-t border-hairline">
                Showing 12 of {queuedItems.length} staged rows.
              </div>
            )}
          </div>
        </section>
      )}
    </div>
  );
};
