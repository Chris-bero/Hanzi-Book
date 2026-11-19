import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { BatchQueueItem, WorkbookEntry } from '../types/workbook';
import { extractChineseCharacters, resolveEntryStrokes } from './strokeEngine';
import { lookupChineseText } from './dictionaryService';

export interface ParsedImportResult {
  items: BatchQueueItem[];
  skippedRows: number;
  detectedColumns: {
    characterCol: string;
    pinyinCol?: string;
    meaningCol?: string;
    arabicCol?: string;
  };
}

const CHAR_HEADER_ALIASES = [
  'character',
  'characters',
  'hanzi',
  'chinese',
  'word',
  'words',
  'term',
  '汉字',
  '词语',
  '字',
];
const PINYIN_HEADER_ALIASES = ['pinyin', 'pronunciation', 'reading', '拼音'];
const MEANING_HEADER_ALIASES = ['meaning', 'english', 'definition', 'translation', 'gloss', '意思', '英文'];
const ARABIC_HEADER_ALIASES = [
  'arabic',
  'ar',
  'arabic meaning',
  'arabic translation',
  'العربية',
  'عربي',
  'المعنى',
  'الترجمة',
  'الترجمة العربية',
  '阿拉伯语',
];

function findMatchingKey(keys: string[], aliases: string[]): string | undefined {
  return keys.find((k) => {
    const normalized = k.trim().toLowerCase();
    return aliases.some((alias) => normalized === alias || normalized.includes(alias));
  });
}

export async function parseSpreadsheetFile(file: File): Promise<ParsedImportResult> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error('The uploaded spreadsheet contains no sheets.');
  }

  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    defval: '',
    raw: false,
  });

  if (rows.length === 0) {
    throw new Error('The spreadsheet is empty.');
  }

  const keys = Object.keys(rows[0]);
  const charKey = findMatchingKey(keys, CHAR_HEADER_ALIASES) ?? keys[0];
  const pinyinKey = findMatchingKey(keys, PINYIN_HEADER_ALIASES);
  const arabicKey = findMatchingKey(keys, ARABIC_HEADER_ALIASES);
  const meaningKey = findMatchingKey(
    keys.filter((k) => k !== arabicKey),
    MEANING_HEADER_ALIASES
  );

  const items: BatchQueueItem[] = [];
  let skippedRows = 0;

  for (const row of rows) {
    const rawChar = String(row[charKey] ?? '').trim();
    const chineseChars = extractChineseCharacters(rawChar);
    if (chineseChars.length === 0) {
      skippedRows++;
      continue;
    }

    items.push({
      text: chineseChars.join(''),
      pinyin: pinyinKey ? String(row[pinyinKey] ?? '').trim() || undefined : undefined,
      meaning: meaningKey ? String(row[meaningKey] ?? '').trim() || undefined : undefined,
      arabicMeaning: arabicKey ? String(row[arabicKey] ?? '').trim() || undefined : undefined,
    });
  }

  return {
    items,
    skippedRows,
    detectedColumns: {
      characterCol: charKey,
      pinyinCol: pinyinKey,
      meaningCol: meaningKey,
      arabicCol: arabicKey,
    },
  };
}

export function downloadImportTemplate(format: 'xlsx' | 'csv') {
  const sampleRows = [
    { Character: '不', Pinyin: 'bù', Meaning: 'no, not; negative prefix', Arabic: 'لا / ليس' },
    { Character: '我', Pinyin: 'wǒ', Meaning: 'I, me, my', Arabic: 'أنا' },
    { Character: '你', Pinyin: 'nǐ', Meaning: 'you', Arabic: 'أنت' },
    { Character: '好', Pinyin: 'hǎo', Meaning: 'good, fine, well', Arabic: 'جيد / بخير' },
    { Character: '你好', Pinyin: 'nǐ hǎo', Meaning: 'hello; hi', Arabic: 'مرحباً' },
    { Character: '学习', Pinyin: 'xué xí', Meaning: 'to study; to learn', Arabic: 'يدرس / يتعلم' },
    { Character: '中国', Pinyin: 'Zhōng guó', Meaning: 'China', Arabic: 'الصين' },
  ];

  const ws = XLSX.utils.json_to_sheet(sampleRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Characters');

  if (format === 'csv') {
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    saveAs(blob, 'hanzi-book-studio-template.csv');
  } else {
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    saveAs(blob, 'hanzi-book-studio-template.xlsx');
  }
}

export async function createWorkbookEntryFromItem(
  item: BatchQueueItem,
  order: number
): Promise<{
  entry: WorkbookEntry | null;
  cacheHits: number;
  networkFetches: number;
}> {
  const chars = extractChineseCharacters(item.text);
  if (chars.length === 0) {
    return { entry: null, cacheHits: 0, networkFetches: 0 };
  }
  const cleanText = chars.join('');

  const [lookup, strokeRes] = await Promise.all([
    lookupChineseText(cleanText),
    resolveEntryStrokes(cleanText),
  ]);

  const { results: charResults, cacheHits, networkFetches } = strokeRes;
  const readyCount = charResults.filter((c) => c.status === 'ready').length;
  const status =
    readyCount === charResults.length
      ? 'ready'
      : readyCount > 0
      ? 'partial'
      : 'unsupported';

  const now = Date.now();
  const entry: WorkbookEntry = {
    id: `${now}-${Math.random().toString(36).slice(2, 8)}`,
    text: cleanText,
    pinyin: item.pinyin?.trim() || lookup.pinyin,
    meaning: item.meaning?.trim() || lookup.meaning,
    arabicMeaning: item.arabicMeaning?.trim() || lookup.arabicMeaning || '',
    isManualPinyin: Boolean(item.pinyin?.trim()),
    isManualMeaning: Boolean(item.meaning?.trim()),
    isManualArabicMeaning: Boolean(item.arabicMeaning?.trim()),
    order,
    charResults,
    totalStrokeCount: charResults.reduce((sum, c) => sum + c.strokeCount, 0),
    status,
    createdAt: now,
    updatedAt: now,
  };

  return { entry, cacheHits, networkFetches };
}
