import React, { useMemo } from 'react';
import { Eye, LayoutGrid, RotateCcw } from 'lucide-react';
import { paginateWorkbookEntries } from '../../services/layoutPagination';
import {
  BoxDensity,
  BoxFillMode,
  DEFAULT_WORKBOOK_SETTINGS,
  GridStyle,
  HANZI_FONT_MAP,
  HanziFontOption,
  PageOrientation,
  PageSize,
  WorkbookProject,
  WorkbookSettings,
} from '../../types/workbook';
import { WorkbookPage } from '../Preview/WorkbookPage';

interface WorkbookSettingsViewProps {
  workbook: WorkbookProject | null;
  onUpdateWorkbookMeta: (title: string, subtitle: string) => Promise<void>;
  onUpdateSettings: (settings: WorkbookSettings) => Promise<void>;
  onGoToPreview: () => void;
}

const ACCENT_PRESETS = [
  { name: 'Publishing Burgundy', value: '#6B1D2F' },
  { name: 'Scholar Ink', value: '#23201D' },
  { name: 'Imperial Jade', value: '#205245' },
  { name: 'Indigo Slate', value: '#273B5B' },
];

export const WorkbookSettingsView: React.FC<WorkbookSettingsViewProps> = ({
  workbook,
  onUpdateWorkbookMeta,
  onUpdateSettings,
  onGoToPreview,
}) => {
  const previewSheets = useMemo(() => {
    if (!workbook) return [];
    return paginateWorkbookEntries(workbook.entries, workbook.settings);
  }, [workbook]);

  if (!workbook) {
    return (
      <div className="bg-white border border-hairline p-12 text-center">
        <h2 className="font-display text-3xl font-bold text-ink m-0">No Active Workbook</h2>
        <p className="mt-2 text-sm text-ink-muted">
          Create or select a workbook from the Dashboard to customize page layout and typography.
        </p>
      </div>
    );
  }

  const { settings } = workbook;
  const firstSheet = previewSheets[0] ?? null;

  const patch = (partial: Partial<WorkbookSettings>) => {
    onUpdateSettings({
      ...settings,
      ...partial,
      layoutVersion: 3,
    });
  };

  const handleFontChange = (fontOption: HanziFontOption) => {
    const def = HANZI_FONT_MAP[fontOption];
    patch({
      fontOption,
      fontFamily: def.cssFamily,
      wordFontName: def.wordFontName,
    });
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-wrap items-end justify-between gap-4 pb-5 border-b border-hairline">
        <div>
          <h1 className="font-display text-4xl font-bold text-ink leading-tight m-0">
            Workbook Settings
          </h1>
          <p className="mt-1.5 text-sm text-ink-muted m-0">
            Configure paper format, page density, grid guidelines, typography, and colors. Changes
            apply identically to Live Preview, editable Word (.docx), and PDF exports.
          </p>
        </div>

        <div className="flex items-center gap-2.5">
          <button
            type="button"
            onClick={() => onUpdateSettings({ ...DEFAULT_WORKBOOK_SETTINGS })}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium border border-hairline-strong bg-white text-ink hover:bg-ivory-deep transition-colors cursor-pointer whitespace-nowrap"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            Reset Defaults
          </button>
          <button
            type="button"
            onClick={onGoToPreview}
            className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold bg-burgundy text-white hover:bg-burgundy-dark transition-colors cursor-pointer whitespace-nowrap"
          >
            <Eye className="w-3.5 h-3.5" />
            Full Preview &amp; Export
          </button>
        </div>
      </div>

      {/* Controls Grid */}
      <div className="grid grid-cols-1 xl:grid-cols-12 gap-6 items-start">
        {/* Left: Configuration Panels */}
        <div className="xl:col-span-5 space-y-5">
          {/* Edition Metadata */}
          <section className="bg-white border border-hairline p-5 space-y-4">
            <h2 className="font-display text-xl font-bold text-ink m-0 pb-2 border-b border-hairline">
              Edition &amp; Structure
            </h2>

            <div>
              <label className="block text-xs font-semibold text-ink-muted mb-1">
                Book Title
              </label>
              <input
                type="text"
                value={workbook.title}
                onChange={(e) => onUpdateWorkbookMeta(e.target.value, workbook.subtitle)}
                className="w-full border border-hairline-strong px-3 py-2 text-sm text-ink"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-ink-muted mb-1">
                Subtitle
              </label>
              <input
                type="text"
                value={workbook.subtitle}
                onChange={(e) => onUpdateWorkbookMeta(workbook.title, e.target.value)}
                className="w-full border border-hairline-strong px-3 py-2 text-sm text-ink"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-ink-muted mb-1">
                Publisher / Author Imprint
              </label>
              <input
                type="text"
                value={settings.authorName}
                onChange={(e) => patch({ authorName: e.target.value })}
                className="w-full border border-hairline-strong px-3 py-2 text-sm text-ink"
              />
            </div>

            <div className="flex flex-col gap-2 pt-1">
              <label className="flex items-center gap-2.5 text-xs text-ink cursor-pointer">
                <input
                  type="checkbox"
                  checked={settings.includeCoverPage}
                  onChange={(e) => patch({ includeCoverPage: e.target.checked })}
                  className="accent-burgundy w-4 h-4"
                />
                <span>Include Cover Page in exported workbook</span>
              </label>
              <label className="flex items-center gap-2.5 text-xs text-ink cursor-pointer">
                <input
                  type="checkbox"
                  checked={settings.includeTableOfContents}
                  onChange={(e) => patch({ includeTableOfContents: e.target.checked })}
                  className="accent-burgundy w-4 h-4"
                />
                <span>Include Character Index / Table of Contents</span>
              </label>
              <label className="flex items-center gap-2.5 text-xs text-ink cursor-pointer">
                <input
                  type="checkbox"
                  checked={settings.showArabicTranslation !== false}
                  onChange={(e) => patch({ showArabicTranslation: e.target.checked })}
                  className="accent-burgundy w-4 h-4"
                />
                <span>
                  Show Arabic Translation (<span style={{ fontFamily: 'var(--font-arabic)' }}>العربية</span>) on Practice Sheets &amp; Index
                </span>
              </label>
            </div>
          </section>

          {/* Paper & Auto-Pagination Target */}
          <section className="bg-white border border-hairline p-5 space-y-4">
            <h2 className="font-display text-xl font-bold text-ink m-0 pb-2 border-b border-hairline">
              Page Format &amp; Pagination
            </h2>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <span className="block text-xs font-semibold text-ink-muted mb-1.5">
                  Paper Size
                </span>
                <div className="grid grid-cols-2 bg-ivory p-0.5 border border-hairline">
                  {(['A4', 'A3'] as PageSize[]).map((size) => (
                    <button
                      key={size}
                      type="button"
                      onClick={() => patch({ pageSize: size })}
                      className={`py-1.5 text-xs font-semibold cursor-pointer ${
                        settings.pageSize === size
                          ? 'bg-white text-burgundy'
                          : 'text-ink-muted hover:text-ink'
                      }`}
                    >
                      {size}
                    </button>
                  ))}
                </div>
              </div>

              <div>
                <span className="block text-xs font-semibold text-ink-muted mb-1.5">
                  Orientation
                </span>
                <div className="grid grid-cols-2 bg-ivory p-0.5 border border-hairline">
                  {(['landscape', 'portrait'] as PageOrientation[]).map((ori) => (
                    <button
                      key={ori}
                      type="button"
                      onClick={() => patch({ orientation: ori })}
                      className={`py-1.5 text-xs font-semibold capitalize cursor-pointer ${
                        settings.orientation === ori
                          ? 'bg-white text-burgundy'
                          : 'text-ink-muted hover:text-ink'
                      }`}
                    >
                      {ori}
                    </button>
                  ))}
                </div>
              </div>
            </div>

            <div>
              <div className="flex justify-between text-xs mb-1.5">
                <span className="font-semibold text-ink-muted">
                  Target Vocabulary Entries per Page
                </span>
                <span className="font-mono text-burgundy font-semibold tabular-nums">
                  {settings.charactersPerPage} / page
                </span>
              </div>
              <div className="grid grid-cols-7 bg-ivory p-0.5 border border-hairline">
                {([1, 2, 3, 4, 5, 6, 7] as Array<1 | 2 | 3 | 4 | 5 | 6 | 7>).map((num) => (
                  <button
                    key={num}
                    type="button"
                    onClick={() => patch({ charactersPerPage: num })}
                    className={`py-1.5 text-xs font-semibold font-mono cursor-pointer ${
                      settings.charactersPerPage === num
                        ? 'bg-white text-burgundy'
                        : 'text-ink-muted hover:text-ink'
                    }`}
                  >
                    {num}
                  </button>
                ))}
              </div>
              <p className="mt-1.5 text-[11px] text-ink-subtle m-0">
                Default is 5 entries per A4 Landscape sheet (up to 7 in Compact density).
                Multi-character words auto-paginate to prevent overflow.
              </p>
            </div>

            <div>
              <div className="flex justify-between text-xs mb-1.5">
                <span className="font-semibold text-ink-muted">Page Margin</span>
                <span className="font-mono text-ink tabular-nums">{settings.pageMarginMm} mm</span>
              </div>
              <input
                type="range"
                min={8}
                max={22}
                step={1}
                value={settings.pageMarginMm}
                onChange={(e) => patch({ pageMarginMm: Number(e.target.value) })}
                className="w-full accent-burgundy"
              />
            </div>
          </section>

          {/* Practice Grid Maximization & Density */}
          <section className="bg-white border border-hairline p-5 space-y-4">
            <div className="flex items-center justify-between pb-2 border-b border-hairline">
              <h2 className="font-display text-xl font-bold text-ink m-0">
                Practice Grid &amp; Density
              </h2>
              <span className="inline-flex items-center gap-1.5 font-mono text-xs text-burgundy font-semibold tabular-nums">
                <LayoutGrid className="w-3.5 h-3.5" />
                {firstSheet ? `${firstSheet.boxesPerRow} boxes / char` : 'Full Width'}
              </span>
            </div>

            <div>
              <span className="block text-xs font-semibold text-ink-muted mb-1.5">
                Row Fill Mode
              </span>
              <div className="grid grid-cols-2 bg-ivory p-0.5 border border-hairline">
                {(
                  [
                    ['maximize', 'Maximize Full Width'],
                    ['fixed10', 'Fixed 10 Boxes'],
                  ] as Array<[BoxFillMode, string]>
                ).map(([mode, label]) => (
                  <button
                    key={mode}
                    type="button"
                    onClick={() => patch({ boxFillMode: mode })}
                    className={`py-1.5 text-xs font-semibold cursor-pointer ${
                      settings.boxFillMode === mode
                        ? 'bg-white text-burgundy'
                        : 'text-ink-muted hover:text-ink'
                    }`}
                  >
                    {label}
                  </button>
                ))}
              </div>
            </div>

            <div>
              <span className="block text-xs font-semibold text-ink-muted mb-1.5">
                Box Size Density
              </span>
              <div className="grid grid-cols-3 bg-ivory p-0.5 border border-hairline">
                {(
                  [
                    ['compact', 'Compact (~12mm)'],
                    ['standard', 'Standard (~14mm)'],
                    ['large', 'Large (~16mm)'],
                  ] as Array<[BoxDensity, string]>
                ).map(([density, label]) => (
                  <button
                    key={density}
                    type="button"
                    onClick={() => patch({ boxDensity: density })}
                    className={`py-1.5 text-xs font-semibold cursor-pointer ${
                      settings.boxDensity === density
                        ? 'bg-white text-burgundy'
                        : 'text-ink-muted hover:text-ink'
                    }`}
                  >
                    {label}
                  </button>
                ))}
              </div>
            </div>

            <div>
              <div className="flex justify-between text-xs mb-1.5">
                <span className="font-semibold text-ink-muted">Tracing Boxes per Row</span>
                <span className="font-mono text-burgundy font-semibold tabular-nums">
                  Boxes 1–{settings.traceBoxCount}
                </span>
              </div>
              <div className="grid grid-cols-3 bg-ivory p-0.5 border border-hairline">
                {[2, 3, 4].map((count) => (
                  <button
                    key={count}
                    type="button"
                    onClick={() => patch({ traceBoxCount: count })}
                    className={`py-1.5 text-xs font-semibold font-mono cursor-pointer ${
                      settings.traceBoxCount === count
                        ? 'bg-white text-burgundy'
                        : 'text-ink-muted hover:text-ink'
                    }`}
                  >
                    {count} Trace
                  </button>
                ))}
              </div>
            </div>

            {firstSheet && (
              <div className="bg-ivory border border-hairline px-3.5 py-2.5 text-xs text-ink-muted space-y-0.5">
                <div className="font-semibold text-ink tabular-nums">
                  Sheet 1 Capacity: {firstSheet.totalBoxesOnPage} Total Practice Boxes
                </div>
                <div className="font-mono text-[11px] tabular-nums">
                  {firstSheet.boxesPerRow} boxes/char · Boxes 1–{firstSheet.traceBoxes} tracing ·
                  Boxes {firstSheet.traceBoxes + 1}–{firstSheet.boxesPerRow} empty ({firstSheet.emptyBoxes})
                </div>
                <div className="text-[11px] text-ink-subtle">
                  Each character of multi-character words (e.g. 你好) gets its own full independent
                  row.
                </div>
              </div>
            )}

            <div>
              <span className="block text-xs font-semibold text-ink-muted mb-1.5">
                Traditional Writing Guidelines
              </span>
              <div className="grid grid-cols-3 bg-ivory p-0.5 border border-hairline">
                {(
                  [
                    ['mi', 'Mi-Zi-Ge 米'],
                    ['tian', 'Tian-Zi-Ge 田'],
                    ['cross', 'Cross 十字'],
                  ] as Array<[GridStyle, string]>
                ).map(([style, label]) => (
                  <button
                    key={style}
                    type="button"
                    onClick={() => patch({ gridStyle: style })}
                    className={`py-1.5 text-xs font-semibold cursor-pointer ${
                      settings.gridStyle === style
                        ? 'bg-white text-burgundy'
                        : 'text-ink-muted hover:text-ink'
                    }`}
                  >
                    {label}
                  </button>
                ))}
              </div>
            </div>
          </section>

          {/* Typography & Colors */}
          <section className="bg-white border border-hairline p-5 space-y-4">
            <h2 className="font-display text-xl font-bold text-ink m-0 pb-2 border-b border-hairline">
              Chinese Typography &amp; Palette
            </h2>

            <div className="space-y-2">
              <span className="block text-xs font-semibold text-ink-muted">
                Character Font Family
              </span>
              {(Object.keys(HANZI_FONT_MAP) as HanziFontOption[]).map((key) => {
                const font = HANZI_FONT_MAP[key];
                const active = settings.fontOption === key;
                return (
                  <button
                    key={key}
                    type="button"
                    onClick={() => handleFontChange(key)}
                    className={`w-full px-3.5 py-2.5 border text-left flex items-center justify-between transition-colors cursor-pointer ${
                      active
                        ? 'border-burgundy bg-burgundy-tint/40'
                        : 'border-hairline bg-white hover:bg-ivory'
                    }`}
                  >
                    <div>
                      <div className="text-xs font-semibold text-ink">{font.label}</div>
                      <div className="text-[11px] text-ink-subtle">{font.description}</div>
                    </div>
                    <span className="text-xl text-ink" style={{ fontFamily: font.cssFamily }}>
                      永和
                    </span>
                  </button>
                );
              })}
            </div>

            <div className="pt-2 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-ink-muted">Accent Presets</span>
                <div className="flex items-center gap-2">
                  {ACCENT_PRESETS.map((preset) => (
                    <button
                      key={preset.value}
                      type="button"
                      title={preset.name}
                      onClick={() => patch({ accentColor: preset.value })}
                      className={`w-5 h-5 border cursor-pointer ${
                        settings.accentColor.toLowerCase() === preset.value.toLowerCase()
                          ? 'ring-2 ring-offset-1 ring-ink border-white'
                          : 'border-hairline-strong'
                      }`}
                      style={{ backgroundColor: preset.value }}
                    />
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3 text-xs">
                <label className="flex items-center justify-between border border-hairline px-3 py-2">
                  <span className="text-ink-muted">Accent</span>
                  <input
                    type="color"
                    value={settings.accentColor}
                    onChange={(e) => patch({ accentColor: e.target.value })}
                    className="w-6 h-5 border-0 bg-transparent cursor-pointer"
                  />
                </label>

                <label className="flex items-center justify-between border border-hairline px-3 py-2">
                  <span className="text-ink-muted">New Stroke</span>
                  <input
                    type="color"
                    value={settings.newStrokeColor}
                    onChange={(e) => patch({ newStrokeColor: e.target.value })}
                    className="w-6 h-5 border-0 bg-transparent cursor-pointer"
                  />
                </label>

                <label className="flex items-center justify-between border border-hairline px-3 py-2">
                  <span className="text-ink-muted">Grid Paper</span>
                  <input
                    type="color"
                    value={settings.gridBgColor}
                    onChange={(e) => patch({ gridBgColor: e.target.value })}
                    className="w-6 h-5 border-0 bg-transparent cursor-pointer"
                  />
                </label>

                <label className="flex items-center justify-between border border-hairline px-3 py-2">
                  <span className="text-ink-muted">Trace Ink</span>
                  <input
                    type="color"
                    value={settings.traceGlyphColor}
                    onChange={(e) => patch({ traceGlyphColor: e.target.value })}
                    className="w-6 h-5 border-0 bg-transparent cursor-pointer"
                  />
                </label>
              </div>
            </div>
          </section>
        </div>

        {/* Right: Sticky Live Sheet Preview */}
        <div className="xl:col-span-7 xl:sticky xl:top-6 bg-ivory-deep/70 border border-hairline p-4 overflow-hidden">
          <div className="flex items-center justify-between mb-3 px-1">
            <span className="text-xs font-semibold text-ink-muted">Live Sheet 1 Preview</span>
            <span className="font-mono text-[11px] text-ink-subtle tabular-nums">
              {settings.pageSize} {settings.orientation} ·{' '}
              {firstSheet
                ? `${firstSheet.entries.length} entries · ${firstSheet.boxesPerRow} boxes/char (${firstSheet.totalBoxesOnPage} boxes)`
                : 'Empty'}
            </span>
          </div>

          {firstSheet ? (
            <div className="overflow-x-auto pb-2">
              <div
                className="origin-top-left"
                style={{
                  transform: 'scale(0.68)',
                  width: '147%',
                  marginBottom: '-220px',
                }}
              >
                <WorkbookPage
                  workbookTitle={workbook.title}
                  sheet={firstSheet}
                  settings={settings}
                  totalPages={previewSheets.length}
                />
              </div>
            </div>
          ) : (
            <div className="bg-white border border-hairline p-10 text-center text-xs text-ink-muted">
              Add at least one character to see a live preview of your customized sheet.
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
