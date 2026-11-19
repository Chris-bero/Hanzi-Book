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
  isManualPinyin: boolean;
  isManualMeaning: boolean;
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
  includeCoverPage: boolean;
  includeTableOfContents: boolean;
  authorName: string;
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
  strokeStageSizePx: 38,
  gridSizePx: 54,
  pageMarginMm: 10,
  charactersPerPage: 5,
  fontOption: 'kaiti',
  fontFamily: HANZI_FONT_MAP.kaiti.cssFamily,
  wordFontName: HANZI_FONT_MAP.kaiti.wordFontName,
  accentColor: '#6B1D2F',
  newStrokeColor: '#C62828',
  completedStrokeColor: '#C4BEB4',
  gridBgColor: '#F5EFE4',
  gridLineColor: '#C9BCA9',
  traceGlyphColor: '#CFC6B8',
  includeCoverPage: true,
  includeTableOfContents: true,
  authorName: 'Hanzi Book Studio Press',
};

/**
 * Normalizes persisted workbooks to enforce minimum 10-box / 2-trace standard
 * and upgrade older saved workbooks cleanly to v3 maximized full-width boxes.
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

  merged.traceBoxCount = Math.max(2, Math.min(4, merged.traceBoxCount || 2));
  merged.emptyBoxCount = Math.max(8, merged.emptyBoxCount || 8);

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
