import React, { useState } from 'react';
import {
  ArrowRight,
  BookPlus,
  Download,
  FileSpreadsheet,
  FolderOpen,
  Languages,
  SlidersHorizontal,
  Trash2,
} from 'lucide-react';
import { ExportRecord, NavigationTab, WorkbookProject } from '../../types/workbook';

interface DashboardViewProps {
  workbooks: WorkbookProject[];
  activeWorkbookId: string | null;
  cachedStrokeCount: number;
  exportsHistory: ExportRecord[];
  onCreateWorkbook: (title: string, subtitle: string, includeSampleSet: boolean) => Promise<void>;
  onSelectWorkbook: (id: string, navigateTo?: NavigationTab) => void;
  onDeleteWorkbook: (id: string) => Promise<void>;
  onNavigate: (tab: NavigationTab) => void;
}

export const DashboardView: React.FC<DashboardViewProps> = ({
  workbooks,
  activeWorkbookId,
  cachedStrokeCount,
  exportsHistory,
  onCreateWorkbook,
  onSelectWorkbook,
  onDeleteWorkbook,
  onNavigate,
}) => {
  const [showCreator, setShowCreator] = useState(false);
  const [newTitle, setNewTitle] = useState('');
  const [newSubtitle, setNewSubtitle] = useState('A4 Landscape Chinese Writing Practice Book');
  const [seedSamples, setSeedSamples] = useState(true);
  const [creating, setCreating] = useState(false);

  const totalCharacters = workbooks.reduce((sum, wb) => sum + wb.entries.length, 0);
  const totalUniqueStrokes = workbooks.reduce(
    (sum, wb) => sum + wb.entries.reduce((s, e) => s + e.totalStrokeCount, 0),
    0
  );

  const handleCreate = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newTitle.trim() || creating) return;
    setCreating(true);
    try {
      await onCreateWorkbook(newTitle.trim(), newSubtitle.trim(), seedSamples);
      setNewTitle('');
      setShowCreator(false);
    } finally {
      setCreating(false);
    }
  };

  return (
    <div className="space-y-8">
      {/* Header Row */}
      <div className="flex flex-wrap items-end justify-between gap-4 pb-6 border-b border-hairline">
        <div>
          <h1 className="font-display text-4xl font-bold text-ink leading-tight m-0">
            Publishing Dashboard
          </h1>
          <p className="mt-1.5 text-sm text-ink-muted m-0">
            Design, preview, and export editable Microsoft Word Chinese handwriting practice books.
          </p>
        </div>

        <button
          type="button"
          onClick={() => setShowCreator((v) => !v)}
          className="inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold bg-burgundy text-white hover:bg-burgundy-dark transition-colors cursor-pointer whitespace-nowrap shrink-0"
        >
          <BookPlus className="w-4 h-4" />
          Create New Workbook
        </button>
      </div>

      {/* Inline Workbook Creator */}
      {showCreator && (
        <form
          onSubmit={handleCreate}
          className="bg-white border border-hairline p-6 space-y-5"
        >
          <div className="flex items-center justify-between">
            <h2 className="font-display text-2xl font-bold text-ink m-0">
              New Workbook Project
            </h2>
            <button
              type="button"
              onClick={() => setShowCreator(false)}
              className="text-xs font-medium text-ink-muted hover:text-ink cursor-pointer"
            >
              Cancel
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label
                htmlFor="wb-title"
                className="block text-xs font-semibold text-ink-muted mb-1.5"
              >
                Workbook Title
              </label>
              <input
                id="wb-title"
                type="text"
                required
                value={newTitle}
                onChange={(e) => setNewTitle(e.target.value)}
                placeholder="e.g. HSK 1 Essential Writing Book"
                className="w-full border border-hairline-strong bg-white px-3.5 py-2.5 text-sm text-ink focus-visible:outline-2 focus-visible:outline-burgundy"
              />
            </div>

            <div>
              <label
                htmlFor="wb-subtitle"
                className="block text-xs font-semibold text-ink-muted mb-1.5"
              >
                Subtitle / Edition Note
              </label>
              <input
                id="wb-subtitle"
                type="text"
                value={newSubtitle}
                onChange={(e) => setNewSubtitle(e.target.value)}
                className="w-full border border-hairline-strong bg-white px-3.5 py-2.5 text-sm text-ink focus-visible:outline-2 focus-visible:outline-burgundy"
              />
            </div>
          </div>

          <label className="flex items-center gap-2.5 text-sm text-ink-muted cursor-pointer select-none">
            <input
              type="checkbox"
              checked={seedSamples}
              onChange={(e) => setSeedSamples(e.target.checked)}
              className="accent-burgundy w-4 h-4"
            />
            <span>Pre-load starter entries (不, 我, 你, 好, 你好) with real Hanzi Writer stroke data</span>
          </label>

          <div className="flex items-center gap-3">
            <button
              type="submit"
              disabled={creating}
              className="inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold bg-burgundy text-white hover:bg-burgundy-dark disabled:opacity-55 transition-colors cursor-pointer whitespace-nowrap"
            >
              {creating ? 'Retrieving Stroke Vectors…' : 'Create & Open Workbook'}
            </button>
          </div>
        </form>
      )}

      {/* Full-Width 4-Column KPI Metric Strip */}
      <section
        aria-label="Workspace summary metrics"
        className="grid grid-cols-2 lg:grid-cols-4 bg-white border border-hairline divide-y lg:divide-y-0 lg:divide-x divide-hairline"
      >
        <div className="p-6">
          <div className="text-xs font-medium text-ink-subtle">Total Created Books</div>
          <div className="mt-2 font-display text-4xl font-bold text-ink tabular-nums">
            {workbooks.length}
          </div>
          <div className="mt-1 text-xs text-ink-subtle">Saved locally in IndexedDB</div>
        </div>

        <div className="p-6">
          <div className="text-xs font-medium text-ink-subtle">Total Added Entries</div>
          <div className="mt-2 font-display text-4xl font-bold text-ink tabular-nums">
            {totalCharacters}
          </div>
          <div className="mt-1 text-xs text-ink-subtle tabular-nums">
            {totalUniqueStrokes} verified strokes total
          </div>
        </div>

        <div className="p-6">
          <div className="text-xs font-medium text-ink-subtle">Cached Character Vectors</div>
          <div className="mt-2 font-display text-4xl font-bold text-ink tabular-nums">
            {cachedStrokeCount}
          </div>
          <div className="mt-1 text-xs text-ink-subtle">Zero duplicate API requests</div>
        </div>

        <div className="p-6">
          <div className="text-xs font-medium text-ink-subtle">Exported Editions</div>
          <div className="mt-2 font-display text-4xl font-bold text-ink tabular-nums">
            {exportsHistory.length}
          </div>
          <div className="mt-1 text-xs text-ink-subtle">.docx &amp; printable PDF</div>
        </div>
      </section>

      {/* Primary Workflow Links */}
      <section aria-label="Publishing workflow tools" className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
        <button
          type="button"
          onClick={() => onNavigate('characters')}
          className="group bg-white border border-hairline p-5 text-left hover:border-burgundy transition-colors cursor-pointer"
        >
          <div className="flex items-center justify-between text-burgundy">
            <Languages className="w-5 h-5" />
            <ArrowRight className="w-4 h-4 opacity-0 group-hover:opacity-100 transition-opacity" />
          </div>
          <div className="mt-3 text-sm font-semibold text-ink">Character Management</div>
          <p className="mt-1 text-xs text-ink-muted leading-relaxed m-0">
            Add single characters, multi-character words, or bulk lists with automatic stroke lookup.
          </p>
        </button>

        <button
          type="button"
          onClick={() => onNavigate('import')}
          className="group bg-white border border-hairline p-5 text-left hover:border-burgundy transition-colors cursor-pointer"
        >
          <div className="flex items-center justify-between text-burgundy">
            <FileSpreadsheet className="w-5 h-5" />
            <ArrowRight className="w-4 h-4 opacity-0 group-hover:opacity-100 transition-opacity" />
          </div>
          <div className="mt-3 text-sm font-semibold text-ink">Import Excel / CSV</div>
          <p className="mt-1 text-xs text-ink-muted leading-relaxed m-0">
            Upload vocabulary spreadsheets of 100–500+ items with resumable batch processing.
          </p>
        </button>

        <button
          type="button"
          onClick={() => onNavigate('settings')}
          className="group bg-white border border-hairline p-5 text-left hover:border-burgundy transition-colors cursor-pointer"
        >
          <div className="flex items-center justify-between text-burgundy">
            <SlidersHorizontal className="w-5 h-5" />
            <ArrowRight className="w-4 h-4 opacity-0 group-hover:opacity-100 transition-opacity" />
          </div>
          <div className="mt-3 text-sm font-semibold text-ink">Workbook Settings</div>
          <p className="mt-1 text-xs text-ink-muted leading-relaxed m-0">
            Customize A4/A3 size, orientation, tracing grid boxes, Chinese fonts, and margins.
          </p>
        </button>

        <button
          type="button"
          onClick={() => onNavigate('history')}
          className="group bg-white border border-hairline p-5 text-left hover:border-burgundy transition-colors cursor-pointer"
        >
          <div className="flex items-center justify-between text-burgundy">
            <Download className="w-5 h-5" />
            <ArrowRight className="w-4 h-4 opacity-0 group-hover:opacity-100 transition-opacity" />
          </div>
          <div className="mt-3 text-sm font-semibold text-ink">Export History</div>
          <p className="mt-1 text-xs text-ink-muted leading-relaxed m-0">
            Review generated Word documents and printable PDF editions across projects.
          </p>
        </button>
      </section>

      {/* Recent Projects Table */}
      <section className="bg-white border border-hairline">
        <div className="px-6 py-4 border-b border-hairline flex items-center justify-between">
          <h2 className="font-display text-2xl font-bold text-ink m-0">Recent Projects</h2>
          <span className="text-xs text-ink-subtle font-mono tabular-nums">
            {workbooks.length} {workbooks.length === 1 ? 'workbook' : 'workbooks'}
          </span>
        </div>

        {workbooks.length === 0 ? (
          <div className="px-6 py-12 text-center">
            <p className="text-sm text-ink-muted m-0">
              No workbooks created yet. Start by creating your first project.
            </p>
            <button
              type="button"
              onClick={() => setShowCreator(true)}
              className="mt-4 inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold bg-burgundy text-white hover:bg-burgundy-dark cursor-pointer"
            >
              <BookPlus className="w-4 h-4" />
              Create First Workbook
            </button>
          </div>
        ) : (
          <div className="divide-y divide-hairline">
            {workbooks.map((wb) => {
              const isActive = wb.id === activeWorkbookId;
              const sampleGlyphs = wb.entries
                .slice(0, 5)
                .map((e) => e.text)
                .join(' ');
              return (
                <div
                  key={wb.id}
                  className={`px-6 py-4 flex flex-wrap items-center justify-between gap-4 transition-colors ${
                    isActive ? 'bg-burgundy-tint/45' : 'hover:bg-ivory'
                  }`}
                >
                  <div className="min-w-0">
                    <div className="flex items-center gap-2.5">
                      <span className="text-base font-semibold text-ink truncate">
                        {wb.title}
                      </span>
                      {isActive && (
                        <span className="text-xs font-semibold text-burgundy">
                          Active
                        </span>
                      )}
                    </div>
                    <div className="mt-1 flex flex-wrap items-center gap-2 text-xs text-ink-subtle tabular-nums">
                      <span>{wb.entries.length} entries</span>
                      <span aria-hidden="true">·</span>
                      <span>
                        {wb.settings.pageSize} {wb.settings.orientation}
                      </span>
                      <span aria-hidden="true">·</span>
                      <span>Updated {new Date(wb.updatedAt).toLocaleDateString()}</span>
                      {sampleGlyphs && (
                        <>
                          <span aria-hidden="true">·</span>
                          <span
                            className="text-ink-muted"
                            style={{ fontFamily: wb.settings.fontFamily }}
                          >
                            {sampleGlyphs}
                          </span>
                        </>
                      )}
                    </div>
                  </div>

                  <div className="flex items-center gap-2 shrink-0">
                    <button
                      type="button"
                      onClick={() => onSelectWorkbook(wb.id, 'characters')}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border border-hairline-strong bg-white text-ink hover:bg-ivory-deep transition-colors cursor-pointer whitespace-nowrap"
                    >
                      <FolderOpen className="w-3.5 h-3.5" />
                      Characters
                    </button>
                    <button
                      type="button"
                      onClick={() => onSelectWorkbook(wb.id, 'preview')}
                      className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-semibold bg-burgundy text-white hover:bg-burgundy-dark transition-colors cursor-pointer whitespace-nowrap"
                    >
                      Preview &amp; Export
                    </button>
                    <button
                      type="button"
                      onClick={() => onDeleteWorkbook(wb.id)}
                      aria-label={`Delete ${wb.title}`}
                      className="p-2 text-ink-subtle hover:text-[#A51D24] hover:bg-[#FDF4F4] transition-colors cursor-pointer"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </section>
    </div>
  );
};
