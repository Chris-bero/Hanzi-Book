export type NavigationTab =
  | 'dashboard'
  | 'characters'
  | 'preview'
  | 'settings'
  | 'import'
  | 'history';

export interface HanziCharStrokeData {
  char: string;
  strokes: string[];
  medians: number[][][];
  radStrokes?: number[];
  strokeCount: number;
  source: 'hanzi-writer-data';
  fetchedAt: number;
}

export interface DictionaryCharRecord {
  character: string;
  pinyin: string[];
  definition: string;
  decomposition?: string;
  radical?: string;
}

export interface DictionaryLookupResult {
  text: string;
  pinyin: string;
  meaning: string;
  arabicMeaning: string;
  charBreakdowns: Array<{
    char: string;
    pinyin: string;
    meaning: string;
    radical?: string;
  }>;
  source: 'makemeahanzi' | 'builtin-word' | 'partial' | 'none';
}

export interface CharacterStrokeResult {
  char: string;
  status: 'ready' | 'unsupported';
  strokeCount: number;
  strokeData?: HanziCharStrokeData;
  error?: string;
}

export type EntryStatus = 'ready' | 'partial' | 'unsupported';

export interface WorkbookEntry {
  id: string;
  text: string;
  pinyin: string;
  meaning: string;
  arabicMeaning?: string;
  isManualPinyin: boolean;
  isManualMeaning: boolean;
  isManualArabicMeaning?: boolean;
  order: number;
  charResults: CharacterStrokeResult[];
  totalStrokeCount: number;
  status: EntryStatus;
  createdAt: number;
  updatedAt: number;
}

export type PageSize = 'A4' | 'A3';
export type PageOrientation = 'landscape' | 'portrait';
export type GridStyle = 'mi' | 'tian' | 'cross';
export type HanziFontOption = 'kaiti' | 'serif' | 'sans';
export type BoxFillMode = 'maximize' | 'fixed10';
export type BoxDensity = 'compact' | 'standard' | 'large';
export type GridBorderThickness = 'thin' | 'standard' | 'bold';

export interface WorkbookSettings {
  layoutVersion?: number;
  pageSize: PageSize;
  orientation: PageOrientation;
  boxFillMode: BoxFillMode;
  boxDensity: BoxDensity;
  traceBoxCount: number;
  emptyBoxCount: number;
  gridStyle: GridStyle;
  characterSizePx: number;
  mainHanziSizePx?: number;
  metaFontSizePx?: number;
  strokeStageSizePx: number;
  gridSizePx: number;
  pageMarginMm: number;
  charactersPerPage: 1 | 2 | 3 | 4 | 5 | 6 | 7;
  fontOption: HanziFontOption;
  fontFamily: string;
  wordFontName: string;
  accentColor: string;
  newStrokeColor: string;
  completedStrokeColor: string;
  gridBgColor: string;
  gridLineColor: string;
  traceGlyphColor: string;
  traceOpacityPct?: number;
  gridBorderDarknessPct?: number;
  gridBorderThickness?: GridBorderThickness;
  gridCrosshairContrastPct?: number;
  customHeaderTitle?: string;
  showStudentInfoBar?: boolean;
  customFooterText?: string;
  includeCoverPage: boolean;
  includeTableOfContents: boolean;
  showArabicTranslation: boolean;
  bwPrintMode?: boolean;
  authorName: string;
}

export const PRO_TYPOGRAPHY_CONTRAST_DEFAULTS = {
  mainHanziSizePx: 34,
  metaFontSizePx: 13,
  traceOpacityPct: 35,
  gridBorderDarknessPct: 100,
  gridBorderThickness: 'standard' as GridBorderThickness,
  gridCrosshairContrastPct: 35,
};

export interface EffectiveWorkbookPalette {
  isBw: boolean;
  pageBg: string;
  primaryInk: string;
  secondaryInk: string;
  subtleInk: string;
  accentColor: string;
  dividerColor: string;
  gridBgColor: string;
  gridBorderColor: string;
  gridBorderWidthPx: number;
  gridLineColor: string;
  gridDiagonalColor: string;
  traceGlyphColor: string;
  newStrokeColor: string;
  completedStrokeColor: string;
  strokeTileBorderColor: string;
  strokeGuideLineColor: string;
  strokeNumberColor: string;
  indexHeaderBg: string;
}

function parseHexRgb(hex: string, fallback: [number, number, number]): [number, number, number] {
  const clean = (hex || '').replace('#', '').trim();
  if (!/^[0-9a-fA-F]{6}$/.test(clean)) return fallback;
  return [
    parseInt(clean.slice(0, 2), 16),
    parseInt(clean.slice(2, 4), 16),
    parseInt(clean.slice(4, 6), 16),
  ];
}

function toHexByte(n: number): string {
  return Math.max(0, Math.min(255, Math.round(n)))
    .toString(16)
    .padStart(2, '0')
    .toUpperCase();
}

export function blendHex(bgHex: string, fgHex: string, ratio: number): string {
  const t = Math.max(0, Math.min(1, ratio));
  const [br, bg, bb] = parseHexRgb(bgHex, [255, 255, 255]);
  const [fr, fg, fb] = parseHexRgb(fgHex, [0, 0, 0]);
  return `#${toHexByte(br + (fr - br) * t)}${toHexByte(bg + (fg - bg) * t)}${toHexByte(
    bb + (fb - bb) * t
  )}`;
}

export function resolveGridBorderWidthPx(thickness?: GridBorderThickness): number {
  if (thickness === 'thin') return 1;
  if (thickness === 'bold') return 2;
  return 1.25;
}

export function resolveWorkbookPalette(settings: WorkbookSettings): EffectiveWorkbookPalette {
  const tracePct = Math.max(
    10,
    Math.min(70, settings.traceOpacityPct ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.traceOpacityPct)
  );
  const borderPct = Math.max(
    20,
    Math.min(
      100,
      settings.gridBorderDarknessPct ?? PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridBorderDarknessPct
    )
  );
  const crosshairPct = Math.max(
    10,
    Math.min(
      80,
      settings.gridCrosshairContrastPct ??
        PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridCrosshairContrastPct
    )
  );
  const gridBorderWidthPx = resolveGridBorderWidthPx(settings.gridBorderThickness);

  if (settings.bwPrintMode) {
    const traceGlyphColor = blendHex('#FFFFFF', '#000000', tracePct / 100);
    const gridBorderColor = blendHex('#FFFFFF', '#000000', borderPct / 100);
    const gridLineColor = blendHex('#FFFFFF', '#000000', crosshairPct / 100);
    const gridDiagonalColor = blendHex('#FFFFFF', '#000000', (crosshairPct / 100) * 0.65);

    return {
      isBw: true,
      pageBg: '#FFFFFF',
      primaryInk: '#000000',
      secondaryInk: '#000000',
      subtleInk: '#444444',
      accentColor: '#000000',
      dividerColor: '#000000',
      gridBgColor: '#FFFFFF',
      gridBorderColor,
      gridBorderWidthPx,
      gridLineColor,
      gridDiagonalColor,
      traceGlyphColor,
      newStrokeColor: '#666666',
      completedStrokeColor: '#BDBDBD',
      strokeTileBorderColor: '#000000',
      strokeGuideLineColor: '#CCCCCC',
      strokeNumberColor: '#666666',
      indexHeaderBg: '#FFFFFF',
    };
  }

  const gridBg = settings.gridBgColor || '#F5EFE4';
  const defaultTraceHex = '#CFC6B8';
  const hasCustomTraceTint =
    Boolean(settings.traceGlyphColor) &&
    settings.traceGlyphColor.toUpperCase() !== defaultTraceHex;

  const traceGlyphColor = hasCustomTraceTint
    ? tracePct <= 35
      ? blendHex(gridBg, settings.traceGlyphColor, Math.max(0.25, tracePct / 35))
      : blendHex(settings.traceGlyphColor, '#1E1916', ((tracePct - 35) / 35) * 0.55)
    : blendHex(gridBg, '#1E1916', (tracePct / 100) * 0.78);

  const gridBorderColor = blendHex(gridBg, '#665645', (borderPct / 100) * 0.75);
  const gridLineColor = blendHex(gridBg, '#5C4D3E', (crosshairPct / 100) * 0.9);
  const gridDiagonalColor = blendHex(gridBg, '#5C4D3E', (crosshairPct / 100) * 0.55);

  return {
    isBw: false,
    pageBg: '#FFFFFF',
    primaryInk: '#1E1916',
    secondaryInk: '#5E544B',
    subtleInk: '#8C8176',
    accentColor: settings.accentColor,
    dividerColor: '#E5DDD1',
    gridBgColor: gridBg,
    gridBorderColor,
    gridBorderWidthPx,
    gridLineColor,
    gridDiagonalColor,
    traceGlyphColor,
    newStrokeColor: settings.newStrokeColor,
    completedStrokeColor: settings.completedStrokeColor,
    strokeTileBorderColor: '#DDD3C5',
    strokeGuideLineColor: '#EAE2D6',
    strokeNumberColor: '#7A7066',
    indexHeaderBg: gridBg,
  };
}

export interface WorkbookProject {
  id: string;
  title: string;
  subtitle: string;
  entries: WorkbookEntry[];
  settings: WorkbookSettings;
  createdAt: number;
  updatedAt: number;
}

export interface ExportRecord {
  id: string;
  workbookId: string;
  workbookTitle: string;
  format: 'docx' | 'pdf';
  fileName: string;
  entryCount: number;
  pageSize: PageSize;
  orientation: PageOrientation;
  exportedAt: number;
}

export interface BatchQueueItem {
  text: string;
  pinyin?: string;
  meaning?: string;
  arabicMeaning?: string;
}

export interface PersistedBatchJob {
  workbookId: string;
  items: BatchQueueItem[];
  processedIndex: number;
  unsupported: Array<{ text: string; reason: string }>;
  cacheHits: number;
  networkFetches: number;
  updatedAt: number;
}

export const HANZI_FONT_MAP: Record<
  HanziFontOption,
  { label: string; cssFamily: string; wordFontName: string; description: string }
> = {
  kaiti: {
    label: 'KaiTi Calligraphy (楷体)',
    cssFamily: '"KaiTi", "STKaiti", "Kaiti SC", "Noto Serif SC", serif',
    wordFontName: 'KaiTi',
    description: 'Traditional handwritten textbook style for learners',
  },
  serif: {
    label: 'Song / Serif (宋体)',
    cssFamily: '"Noto Serif SC", "Songti SC", "SimSun", serif',
    wordFontName: 'SimSun',
    description: 'Classic book publishing print serif',
  },
  sans: {
    label: 'HeiTi Sans (黑体)',
    cssFamily: '"Noto Sans SC", "PingFang SC", "Microsoft YaHei", sans-serif',
    wordFontName: 'Microsoft YaHei',
    description: 'Clean geometric strokes for modern clarity',
  },
};

export const DEFAULT_WORKBOOK_SETTINGS: WorkbookSettings = {
  layoutVersion: 3,
  pageSize: 'A4',
  orientation: 'landscape',
  boxFillMode: 'maximize',
  boxDensity: 'standard',
  traceBoxCount: 2,
  emptyBoxCount: 8,
  gridStyle: 'mi',
  characterSizePx: 46,
  mainHanziSizePx: PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.mainHanziSizePx,
  metaFontSizePx: PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.metaFontSizePx,
  strokeStageSizePx: 38,
  gridSizePx: 54,
  pageMarginMm: 10,
  charactersPerPage: 5,
  fontOption: 'kaiti',
  fontFamily: HANZI_FONT_MAP.kaiti.cssFamily,
  wordFontName: HANZI_FONT_MAP.kaiti.wordFontName,
  accentColor: '#6B1D2F',
  newStrokeColor: '#C83232',
  completedStrokeColor: '#C8C2B8',
  gridBgColor: '#F5EFE4',
  gridLineColor: '#C9BCA9',
  traceGlyphColor: '#CFC6B8',
  traceOpacityPct: PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.traceOpacityPct,
  gridBorderDarknessPct: PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridBorderDarknessPct,
  gridBorderThickness: PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridBorderThickness,
  gridCrosshairContrastPct: PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridCrosshairContrastPct,
  customHeaderTitle: '',
  showStudentInfoBar: false,
  customFooterText: '',
  includeCoverPage: true,
  includeTableOfContents: true,
  showArabicTranslation: true,
  bwPrintMode: false,
  authorName: 'Hanzi Book Studio Press',
};

export function resolveSheetHeaderTitle(workbookTitle: string, settings: WorkbookSettings): string {
  const custom = settings.customHeaderTitle?.trim();
  return custom || workbookTitle || 'Chinese Character Workbook';
}

export function resolveSheetFooterText(
  settings: WorkbookSettings,
  traceBoxes: number,
  boxesPerRow: number
): string {
  const custom = settings.customFooterText?.trim();
  if (custom) return custom;

  const imprint = settings.authorName?.trim() || 'Hanzi Book Studio';
  if (traceBoxes <= 0) {
    return `${imprint} · Dictation Practice: Write from memory in boxes 1–${boxesPerRow}`;
  }
  if (traceBoxes === 1) {
    return `${imprint} · Practice: Trace box 1, then write from memory in boxes 2–${boxesPerRow}`;
  }
  return `${imprint} · Practice: Trace boxes 1–${traceBoxes}, then write from memory in boxes ${
    traceBoxes + 1
  }–${boxesPerRow}`;
}

/**
 * Normalizes persisted workbooks cleanly while preserving user preferences
 * and upgrading older saved workbooks to v3 maximized full-width boxes.
 */
export function normalizeWorkbookSettings(raw?: Partial<WorkbookSettings>): WorkbookSettings {
  const merged: WorkbookSettings = {
    ...DEFAULT_WORKBOOK_SETTINGS,
    ...(raw ?? {}),
  };

  if (!raw?.layoutVersion || raw.layoutVersion < 3) {
    merged.boxFillMode = 'maximize';
    merged.boxDensity = 'standard';
    if (!raw?.layoutVersion || raw.layoutVersion < 2) {
      merged.charactersPerPage = 5;
      merged.pageMarginMm = 10;
      merged.characterSizePx = 46;
      merged.strokeStageSizePx = 38;
      merged.gridSizePx = 54;
    }
    merged.layoutVersion = 3;
  }

  if (raw?.showArabicTranslation === undefined) {
    merged.showArabicTranslation = true;
  }
  merged.bwPrintMode = Boolean(raw?.bwPrintMode);
  merged.showStudentInfoBar = Boolean(raw?.showStudentInfoBar);
  merged.customHeaderTitle = typeof raw?.customHeaderTitle === 'string' ? raw.customHeaderTitle : '';
  merged.customFooterText = typeof raw?.customFooterText === 'string' ? raw.customFooterText : '';

  merged.mainHanziSizePx = Math.max(
    24,
    Math.min(
      48,
      Math.round(
        typeof raw?.mainHanziSizePx === 'number'
          ? raw.mainHanziSizePx
          : PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.mainHanziSizePx
      )
    )
  );
  merged.metaFontSizePx = Math.max(
    10,
    Math.min(
      20,
      Math.round(
        typeof raw?.metaFontSizePx === 'number'
          ? raw.metaFontSizePx
          : PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.metaFontSizePx
      )
    )
  );
  merged.traceOpacityPct = Math.max(
    10,
    Math.min(
      70,
      Math.round(
        typeof raw?.traceOpacityPct === 'number'
          ? raw.traceOpacityPct
          : PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.traceOpacityPct
      )
    )
  );
  merged.gridBorderDarknessPct = Math.max(
    20,
    Math.min(
      100,
      Math.round(
        typeof raw?.gridBorderDarknessPct === 'number'
          ? raw.gridBorderDarknessPct
          : PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridBorderDarknessPct
      )
    )
  );
  merged.gridBorderThickness =
    raw?.gridBorderThickness === 'thin' ||
    raw?.gridBorderThickness === 'standard' ||
    raw?.gridBorderThickness === 'bold'
      ? raw.gridBorderThickness
      : PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridBorderThickness;
  merged.gridCrosshairContrastPct = Math.max(
    10,
    Math.min(
      80,
      Math.round(
        typeof raw?.gridCrosshairContrastPct === 'number'
          ? raw.gridCrosshairContrastPct
          : PRO_TYPOGRAPHY_CONTRAST_DEFAULTS.gridCrosshairContrastPct
      )
    )
  );

  if (!merged.newStrokeColor || merged.newStrokeColor.toUpperCase() === '#C62828') {
    merged.newStrokeColor = '#C83232';
  }
  if (!merged.completedStrokeColor || merged.completedStrokeColor.toUpperCase() === '#C4BEB4') {
    merged.completedStrokeColor = '#C8C2B8';
  }

  const rawTrace = typeof raw?.traceBoxCount === 'number' ? raw.traceBoxCount : 2;
  merged.traceBoxCount = Math.max(0, Math.min(4, Math.round(rawTrace)));
  merged.emptyBoxCount = Math.max(6, merged.emptyBoxCount || 8);

  return merged;
}

/**
 * Storage repository interface structured so IndexedDB can be swapped for
 * a remote REST / PostgreSQL / Firebase backend without changing UI or engine code.
 */
export interface WorkbookRepository {
  listWorkbooks(): Promise<WorkbookProject[]>;
  getWorkbook(id: string): Promise<WorkbookProject | null>;
  saveWorkbook(workbook: WorkbookProject): Promise<void>;
  deleteWorkbook(id: string): Promise<void>;

  getCachedStrokeData(char: string): Promise<HanziCharStrokeData | null>;
  saveCachedStrokeData(data: HanziCharStrokeData): Promise<void>;
  countCachedStrokeChars(): Promise<number>;

  getDictionaryChar(char: string): Promise<DictionaryCharRecord | null>;
  bulkSaveDictionaryChars(records: DictionaryCharRecord[]): Promise<void>;
  getMeta(key: string): Promise<string | null>;
  setMeta(key: string, value: string): Promise<void>;
  deleteMeta(key: string): Promise<void>;

  listExports(): Promise<ExportRecord[]>;
  saveExport(record: ExportRecord): Promise<void>;
  deleteExport(id: string): Promise<void>;
}
