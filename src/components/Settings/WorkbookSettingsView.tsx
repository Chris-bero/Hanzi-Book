import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Eye, LayoutGrid, RotateCcw } from 'lucide-react';
import { paginateWorkbookEntries } from '../../services/layoutPagination';
import {
  BoxDensity,
  BoxFillMode,
  DEFAULT_WORKBOOK_SETTINGS,
  GridBorderThickness,
  GridStyle,
  HANZI_FONT_MAP,
  HanziFontOption,
  PageOrientation,
  PageSize,
  PRO_TYPOGRAPHY_CONTRAST_DEFAULTS,
  resolveWorkbookPalette,
  WorkbookProject,
  WorkbookSettings,
} from '../../types/workbook';
import { getSheetPreviewScale, WorkbookPage } from '../Preview/WorkbookPage';

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

const PRINT_CONTRAST_PRESETS: Array<{
  label: string;
  traceOpacityPct: number;
  gridBorderDarknessPct: number;
  gridBorderThickness: GridBorderThickness;
  gridCrosshairContrastPct: number;
}> = [
  {
    label: 'Light Ink',
    traceOpacityPct: 22,
    gridBorderDarknessPct: 65,
    gridBorderThickness: 'thin',
    gridCrosshairContrastPct: 22,
  },
  {
    label: 'Balanced',
    traceOpacityPct: PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.traceOpacityPct,
    gridBorderDarknessPct: PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridBorderDarknessPct,
    gridBorderThickness: PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridBorderThickness,
    gridCrosshairContrastPct: PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridCrosshairContrastPct,
  },
  {
    label: 'Crisp Print',
    traceOpacityPct: 50,
    gridBorderDarknessPct: 100,
    gridBorderThickness: 'bold',
    gridCrosshairContrastPct: 52,
  },
];

export const WorkbookSettingsView: React.FC<WorkbookSettingsViewProps> = ({
  workbook,
  onUpdateWorkbookMeta,
  onUpdateSettings,
  onGoToPreview,
}) => {
  const previewContainerRef = useRef<HTMLDivElement | null>(null);
  const sheetContentRef = useRef<HTMLDivElement | null>(null);
  const [previewContainerW, setPreviewContainerW] = useState<number>(680);
  const [measuredSheetH, setMeasuredSheetH] = useState<number>(0);

  useEffect(() => {
    const containerEl = previewContainerRef.current;
    const sheetEl = sheetContentRef.current;
    if (!containerEl) return;

    const updateMetrics = () => {
      const w = containerEl.clientWidth;
      if (w > 0) setPreviewContainerW(w);
      if (sheetEl && sheetEl.offsetHeight > 0) {
        setMeasuredSheetH(sheetEl.offsetHeight);
      }
    };
    updateMetrics();

    if (typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(updateMetrics);
    observer.observe(containerEl);
    if (sheetEl) observer.observe(sheetEl);
    return () => observer.disconnect();
  }, [workbook?.id, workbook?.settings]);

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
  const palette = resolveWorkbookPalette(settings);
  const sheetScaleMetrics = getSheetPreviewScale(settings);
  const autoFitScale = Math.min(1, Math.max(0.32, (previewContainerW - 4) / sheetScaleMetrics.widthPx));
  const effectiveSheetH = Math.max(sheetScaleMetrics.heightPx, measuredSheetH);
  const scaledHeightPx = Math.ceil(effectiveSheetH * autoFitScale) + 2;

  const mainHanziSize =
    settings.mainHanziSizePx ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.mainHanziSizePx;
  const metaFontSize = settings.metaFontSizePx ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.metaFontSizePx;
  const traceOpacityPct =
    settings.traceOpacityPct ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.traceOpacityPct;
  const gridBorderDarknessPct =
    settings.gridBorderDarknessPct ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridBorderDarknessPct;
  const gridBorderThickness =
    settings.gridBorderThickness ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridBorderThickness;
  const gridCrosshairContrastPct =
    settings.gridCrosshairContrastPct ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridCrosshairContrastPct;
  const traceCount = typeof settings.traceBoxCount === 'number' ? settings.traceBoxCount : 2;

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

  const handleResetTypographyContrast = () => {
    patch({
      ...PRO_TYPOGRAPHY_CONTRAST_DEFAULTS,
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
            Configure paper format, typography scale, print contrast, headers, and practice grid
            layout. Changes apply identically to Live Preview, editable Word (.docx), and PDF
            exports.
          </p>
        </div>

        <div className="flex items-center gap-2.5">
          <button
            type="button"
            onClick={() => onUpdateSettings({ ...DEFAULT_WORKBOOK_SETTINGS })}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium border border-hairline-strong bg-white text-ink hover:bg-ivory-deep transition-colors cursor-pointer whitespace-nowrap shrink-0"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            Reset All Defaults
          </button>
          <button
            type="button"
            onClick={onGoToPreview}
            className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold bg-burgundy text-white hover:bg-burgundy-dark transition-colors cursor-pointer whitespace-nowrap shrink-0"
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
          {/* Edition, Header & Student Info Bar */}
          <section className="bg-white border border-hairline p-5 space-y-4">
            <h2 className="font-display text-xl font-bold text-ink m-0 pb-2 border-b border-hairline">
              Edition, Header &amp; Student Bar
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
                Worksheet Running Header Title
              </label>
              <input
                type="text"
                value={settings.customHeaderTitle ?? ''}
                placeholder={workbook.title || 'Defaults to Book Title'}
                onChange={(e) => patch({ customHeaderTitle: e.target.value })}
                className="w-full border border-hairline-strong px-3 py-2 text-sm text-ink placeholder:text-ink-subtle"
              />
              <p className="mt-1 text-[11px] text-ink-subtle m-0">
                Overrides the top-left title printed on every practice sheet. Leave blank to use
                the Book Title.
              </p>
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

            <div>
              <label className="block text-xs font-semibold text-ink-muted mb-1">
                Custom Footer Instruction
              </label>
              <input
                type="text"
                value={settings.customFooterText ?? ''}
                placeholder="Auto-generated practice instruction (leave blank for default)"
                onChange={(e) => patch({ customFooterText: e.target.value })}
                className="w-full border border-hairline-strong px-3 py-2 text-sm text-ink placeholder:text-ink-subtle"
              />
            </div>

            {/* Printable Student Info Bar Toggle */}
            <div className="flex items-center justify-between gap-3 border border-hairline-strong bg-ivory/60 px-3.5 py-2.5">
              <div>
                <div className="text-xs font-semibold text-ink">Student Info Bar</div>
                <div className="text-[11px] text-ink-subtle">
                  Adds a printable “Name: ______  Date: ______” line beneath the header
                </div>
              </div>
              <button
                type="button"
                role="switch"
                aria-label="Student Info Bar"
                aria-checked={Boolean(settings.showStudentInfoBar)}
                onClick={() => patch({ showStudentInfoBar: !settings.showStudentInfoBar })}
                className={`relative inline-flex h-5 w-9 shrink-0 items-center transition-colors cursor-pointer ${
                  settings.showStudentInfoBar ? 'bg-burgundy' : 'bg-hairline-strong'
                }`}
              >
                <span
                  className={`inline-block h-3.5 w-3.5 bg-white transition-transform ${
                    settings.showStudentInfoBar ? 'translate-x-4.5' : 'translate-x-0.5'
                  }`}
                />
              </button>
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
                  Show Arabic Translation (
                  <span style={{ fontFamily: 'var(--font-arabic)' }}>العربية</span>) on Practice
                  Sheets &amp; Index
                </span>
              </label>
            </div>
          </section>

          {/* Pro Typography & Print Contrast Controls */}
          <section className="bg-white border border-hairline p-5 space-y-4">
            <div className="flex items-center justify-between gap-2 pb-2 border-b border-hairline">
              <h2 className="font-display text-xl font-bold text-ink m-0">
                Typography &amp; Print Contrast
              </h2>
              <button
                type="button"
                onClick={handleResetTypographyContrast}
                className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-medium border border-hairline-strong bg-white text-ink-muted hover:text-ink hover:bg-ivory transition-colors cursor-pointer whitespace-nowrap shrink-0"
                title="Reset typography and print contrast sliders to balanced defaults"
              >
                <RotateCcw className="w-3 h-3" />
                Reset Sliders
              </button>
            </div>

            {/* Typography Sliders */}
            <div className="space-y-3.5">
              <div>
                <div className="flex items-center justify-between text-xs mb-1">
                  <label htmlFor="slider-main-hanzi" className="font-semibold text-ink-muted">
                    Main Hanzi Size
                  </label>
                  <span className="font-mono text-burgundy font-semibold tabular-nums">
                    {mainHanziSize}px
                  </span>
                </div>
                <input
                  id="slider-main-hanzi"
                  type="range"
                  min={24}
                  max={48}
                  step={1}
                  value={mainHanziSize}
                  onChange={(e) => patch({ mainHanziSizePx: Number(e.target.value) })}
                  className="w-full accent-burgundy"
                />
                <div className="flex justify-between text-[10px] font-mono text-ink-subtle tabular-nums">
                  <span>24px Compact</span>
                  <span>34px Default</span>
                  <span>48px Large</span>
                </div>
              </div>

              <div>
                <div className="flex items-center justify-between text-xs mb-1">
                  <label htmlFor="slider-meta-font" className="font-semibold text-ink-muted">
                    Pinyin &amp; Meaning Font Size
                  </label>
                  <span className="font-mono text-burgundy font-semibold tabular-nums">
                    {metaFontSize}px
                  </span>
                </div>
                <input
                  id="slider-meta-font"
                  type="range"
                  min={10}
                  max={20}
                  step={1}
                  value={metaFontSize}
                  onChange={(e) => patch({ metaFontSizePx: Number(e.target.value) })}
                  className="w-full accent-burgundy"
                />
                <div className="flex justify-between text-[10px] font-mono text-ink-subtle tabular-nums">
                  <span>10px Fine</span>
                  <span>13px Default</span>
                  <span>20px Large Print</span>
                </div>
              </div>
            </div>

            {/* Print Contrast & Darkness Controls */}
            <div className="pt-3 border-t border-hairline space-y-3.5">
              <div className="flex items-center justify-between gap-2">
                <span className="text-xs font-semibold text-ink-muted">
                  Printer Calibration Presets
                </span>
                <div className="inline-flex bg-ivory p-0.5 border border-hairline">
                  {PRINT_CONTRAST_PRESETS.map((preset) => {
                    const isActive =
                      traceOpacityPct === preset.traceOpacityPct &&
                      gridBorderDarknessPct === preset.gridBorderDarknessPct &&
                      gridBorderThickness === preset.gridBorderThickness &&
                      gridCrosshairContrastPct === preset.gridCrosshairContrastPct;
                    return (
                      <button
                        key={preset.label}
                        type="button"
                        onClick={() =>
                          patch({
                            traceOpacityPct: preset.traceOpacityPct,
                            gridBorderDarknessPct: preset.gridBorderDarknessPct,
                            gridBorderThickness: preset.gridBorderThickness,
                            gridCrosshairContrastPct: preset.gridCrosshairContrastPct,
                          })
                        }
                        className={`px-2 py-1 text-[11px] font-medium transition-colors cursor-pointer whitespace-nowrap shrink-0 ${
                          isActive
                            ? 'bg-white text-burgundy font-semibold'
                            : 'text-ink-muted hover:text-ink'
                        }`}
                      >
                        {preset.label}
                      </button>
                    );
                  })}
                </div>
              </div>

              <div>
                <div className="flex items-center justify-between text-xs mb-1">
                  <label htmlFor="slider-trace-opacity" className="font-semibold text-ink-muted">
                    Tracing Character Darkness
                  </label>
                  <span className="inline-flex items-center gap-1.5 font-mono text-ink font-semibold tabular-nums">
                    <span
                      className="inline-block w-3 h-3 border border-hairline-strong"
                      style={{ backgroundColor: palette.traceGlyphColor }}
                      aria-hidden="true"
                    />
                    {traceOpacityPct}%
                  </span>
                </div>
                <input
                  id="slider-trace-opacity"
                  type="range"
                  min={10}
                  max={70}
                  step={1}
                  value={traceOpacityPct}
                  onChange={(e) => patch({ traceOpacityPct: Number(e.target.value) })}
                  className="w-full accent-burgundy"
                />
                <div className="flex justify-between text-[10px] font-mono text-ink-subtle tabular-nums">
                  <span>10% Faint Ghost</span>
                  <span>35% Balanced</span>
                  <span>70% Dark Ink</span>
                </div>
              </div>

              <div>
                <div className="flex items-center justify-between text-xs mb-1">
                  <label htmlFor="slider-border-darkness" className="font-semibold text-ink-muted">
                    Outer Box Border Darkness
                  </label>
                  <span className="inline-flex items-center gap-1.5 font-mono text-ink font-semibold tabular-nums">
                    <span
                      className="inline-block w-3 h-3 border border-hairline-strong"
                      style={{ backgroundColor: palette.gridBorderColor }}
                      aria-hidden="true"
                    />
                    {gridBorderDarknessPct}%
                  </span>
                </div>
                <input
                  id="slider-border-darkness"
                  type="range"
                  min={20}
                  max={100}
                  step={5}
                  value={gridBorderDarknessPct}
                  onChange={(e) => patch({ gridBorderDarknessPct: Number(e.target.value) })}
                  className="w-full accent-burgundy"
                />
              </div>

              <div>
                <span className="block text-xs font-semibold text-ink-muted mb-1.5">
                  Outer Box Border Thickness
                </span>
                <div className="grid grid-cols-3 bg-ivory p-0.5 border border-hairline">
                  {(
                    [
                      ['thin', 'Thin (1px)'],
                      ['standard', 'Standard'],
                      ['bold', 'Bold (2px)'],
                    ] as Array<[GridBorderThickness, string]>
                  ).map(([val, label]) => (
                    <button
                      key={val}
                      type="button"
                      onClick={() => patch({ gridBorderThickness: val })}
                      className={`py-1.5 text-xs font-semibold cursor-pointer whitespace-nowrap ${
                        gridBorderThickness === val
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
                <div className="flex items-center justify-between text-xs mb-1">
                  <label
                    htmlFor="slider-crosshair-contrast"
                    className="font-semibold text-ink-muted"
                  >
                    Inner Dashed Guidelines (Mi / Tian Zi Ge)
                  </label>
                  <span className="inline-flex items-center gap-1.5 font-mono text-ink font-semibold tabular-nums">
                    <span
                      className="inline-block w-3 h-3 border border-hairline-strong"
                      style={{ backgroundColor: palette.gridLineColor }}
                      aria-hidden="true"
                    />
                    {gridCrosshairContrastPct}%
                  </span>
                </div>
                <input
                  id="slider-crosshair-contrast"
                  type="range"
                  min={10}
                  max={80}
                  step={1}
                  value={gridCrosshairContrastPct}
                  onChange={(e) => patch({ gridCrosshairContrastPct: Number(e.target.value) })}
                  className="w-full accent-burgundy"
                />
                <div className="flex justify-between text-[10px] font-mono text-ink-subtle tabular-nums">
                  <span>10% Subtle</span>
                  <span>35% Print Ready</span>
                  <span>80% High Contrast</span>
                </div>
              </div>
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
                  {traceCount === 0
                    ? '0 (Blank Dictation)'
                    : traceCount === 1
                    ? 'Box 1 Trace'
                    : `Boxes 1–${traceCount} Trace`}
                </span>
              </div>
              <div className="grid grid-cols-5 bg-ivory p-0.5 border border-hairline">
                {[0, 1, 2, 3, 4].map((count) => (
                  <button
                    key={count}
                    type="button"
                    onClick={() => patch({ traceBoxCount: count })}
                    className={`py-1.5 text-xs font-semibold font-mono cursor-pointer whitespace-nowrap ${
                      traceCount === count
                        ? 'bg-white text-burgundy'
                        : 'text-ink-muted hover:text-ink'
                    }`}
                  >
                    {count === 0 ? '0 Blank' : `${count}`}
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
                  {firstSheet.traceBoxes === 0
                    ? `${firstSheet.boxesPerRow} boxes/char · All ${firstSheet.boxesPerRow} blank (Dictation mode)`
                    : `${firstSheet.boxesPerRow} boxes/char · ${
                        firstSheet.traceBoxes === 1
                          ? 'Box 1 tracing'
                          : `Boxes 1–${firstSheet.traceBoxes} tracing`
                      } · Boxes ${firstSheet.traceBoxes + 1}–${firstSheet.boxesPerRow} empty (${
                        firstSheet.emptyBoxes
                      })`}
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

          {/* Font Family & Palette */}
          <section className="bg-white border border-hairline p-5 space-y-4">
            <h2 className="font-display text-xl font-bold text-ink m-0 pb-2 border-b border-hairline">
              Chinese Font &amp; Palette
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
              {/* Black & White / Ink-Saver Print Mode Switch */}
              <div className="flex items-center justify-between gap-3 border border-hairline-strong bg-ivory/60 px-3.5 py-2.5">
                <div>
                  <div className="text-xs font-semibold text-ink">B&amp;W Print Mode</div>
                  <div className="text-[11px] text-ink-subtle">
                    Ink-saver pure white &amp; black layout with calibrated neutral gray guidelines
                    for Preview, Print, Word &amp; PDF
                  </div>
                </div>
                <button
                  type="button"
                  role="switch"
                  aria-label="B&W Print Mode"
                  aria-checked={Boolean(settings.bwPrintMode)}
                  onClick={() => patch({ bwPrintMode: !settings.bwPrintMode })}
                  className={`relative inline-flex h-5 w-9 shrink-0 items-center transition-colors cursor-pointer ${
                    settings.bwPrintMode ? 'bg-black' : 'bg-hairline-strong'
                  }`}
                >
                  <span
                    className={`inline-block h-3.5 w-3.5 bg-white transition-transform ${
                      settings.bwPrintMode ? 'translate-x-4.5' : 'translate-x-0.5'
                    }`}
                  />
                </button>
              </div>

              <div className={`space-y-3 ${settings.bwPrintMode ? 'opacity-50' : ''}`}>
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
                    <span className="text-ink-muted">Prev Strokes</span>
                    <input
                      type="color"
                      value={settings.completedStrokeColor}
                      onChange={(e) => patch({ completedStrokeColor: e.target.value })}
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
                    <span className="text-ink-muted">Trace Tint</span>
                    <input
                      type="color"
                      value={settings.traceGlyphColor}
                      onChange={(e) => patch({ traceGlyphColor: e.target.value })}
                      className="w-6 h-5 border-0 bg-transparent cursor-pointer"
                    />
                  </label>
                </div>
              </div>
            </div>
          </section>
        </div>

        {/* Right: Sticky Auto-Fit Live Sheet Preview */}
        <div className="xl:col-span-7 xl:sticky xl:top-6 bg-ivory-deep/70 border border-hairline p-4">
          <div className="flex items-center justify-between gap-3 mb-3 px-1 min-w-0">
            <span className="text-xs font-semibold text-ink-muted whitespace-nowrap">
              Live Sheet 1 Preview
            </span>
            <span className="font-mono text-[11px] text-ink-subtle tabular-nums truncate">
              {settings.pageSize} {settings.orientation} ·{' '}
              {firstSheet
                ? `${firstSheet.entries.length} entries · ${firstSheet.boxesPerRow} boxes/char (${firstSheet.totalBoxesOnPage} boxes) · ${Math.round(
                    autoFitScale * 100
                  )}% fit`
                : 'Empty'}
            </span>
          </div>

          <div ref={previewContainerRef} className="w-full overflow-hidden">
            {firstSheet ? (
              <div
                className="relative mx-auto"
                style={{
                  width: `${Math.ceil(sheetScaleMetrics.widthPx * autoFitScale)}px`,
                  height: `${scaledHeightPx}px`,
                }}
              >
                <div
                  ref={sheetContentRef}
                  className="origin-top-left"
                  style={{
                    width: `${sheetScaleMetrics.widthPx}px`,
                    transform: `scale(${autoFitScale})`,
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
    </div>
  );
};
