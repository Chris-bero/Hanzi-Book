import { repository } from '../storage/db';
import { extractChineseCharacters } from './strokeEngine';
import { DictionaryCharRecord, DictionaryLookupResult } from '../types/workbook';

const DICT_HYDRATED_KEY = 'makemeahanzi_dict_v1';

const DICTIONARY_SOURCES = [
  'https://cdn.jsdelivr.net/gh/skishore/makemeahanzi@master/dictionary.txt',
  'https://raw.githubusercontent.com/skishore/makemeahanzi/master/dictionary.txt',
];

/**
 * Curated high-frequency multi-character words and core starter characters
 * so lookups respond immediately even before full 9,500+ character hydration completes.
 */
const BUILTIN_WORD_ENTRIES: Record<string, { pinyin: string; meaning: string }> = {
  你好: { pinyin: 'nǐ hǎo', meaning: 'hello; hi' },
  谢谢: { pinyin: 'xiè xie', meaning: 'thank you; to thank' },
  再见: { pinyin: 'zài jiàn', meaning: 'goodbye; see you again' },
  中国: { pinyin: 'Zhōng guó', meaning: 'China' },
  中文: { pinyin: 'Zhōng wén', meaning: 'Chinese language' },
  汉语: { pinyin: 'Hàn yǔ', meaning: 'Chinese language (Mandarin)' },
  学习: { pinyin: 'xué xí', meaning: 'to study; to learn' },
  朋友: { pinyin: 'péng you', meaning: 'friend' },
  老师: { pinyin: 'lǎo shī', meaning: 'teacher' },
  学生: { pinyin: 'xué sheng', meaning: 'student' },
  学校: { pinyin: 'xué xiào', meaning: 'school' },
  家人: { pinyin: 'jiā rén', meaning: 'family members' },
  今天: { pinyin: 'jīn tiān', meaning: 'today' },
  明天: { pinyin: 'míng tiān', meaning: 'tomorrow' },
  昨天: { pinyin: 'zuó tiān', meaning: 'yesterday' },
  高兴: { pinyin: 'gāo xìng', meaning: 'happy; glad' },
  喜欢: { pinyin: 'xǐ huan', meaning: 'to like; to be fond of' },
  天气: { pinyin: 'tiān qì', meaning: 'weather' },
  时间: { pinyin: 'shí jiān', meaning: 'time; period' },
  工作: { pinyin: 'gōng zuò', meaning: 'work; job' },
  对不起: { pinyin: 'duì bu qǐ', meaning: 'sorry; excuse me' },
  没关系: { pinyin: 'méi guān xi', meaning: 'it does not matter; never mind' },
  不客气: { pinyin: 'bù kè qi', meaning: "you're welcome" },
};

const BUILTIN_CHAR_FALLBACK: Record<string, DictionaryCharRecord> = {
  不: { character: '不', pinyin: ['bù'], definition: 'no, not; un-; negative prefix', radical: '一' },
  我: { character: '我', pinyin: ['wǒ'], definition: 'I, me, my; our, us', radical: '戈' },
  你: { character: '你', pinyin: ['nǐ'], definition: 'you, second person pronoun', radical: '亻' },
  好: { character: '好', pinyin: ['hǎo', 'hào'], definition: 'good, fine, well; to be fond of', radical: '女' },
  人: { character: '人', pinyin: ['rén'], definition: 'person, people, human being', radical: '人' },
  大: { character: '大', pinyin: ['dà'], definition: 'big, great, large, vast', radical: '大' },
  小: { character: '小', pinyin: ['xiǎo'], definition: 'small, tiny, little', radical: '小' },
  中: { character: '中', pinyin: ['zhōng'], definition: 'middle, center; China', radical: '丨' },
  国: { character: '国', pinyin: ['guó'], definition: 'country, nation, state', radical: '囗' },
  学: { character: '学', pinyin: ['xué'], definition: 'to study, to learn; school', radical: '子' },
  水: { character: '水', pinyin: ['shuǐ'], definition: 'water, river, liquid', radical: '水' },
  火: { character: '火', pinyin: ['huǒ'], definition: 'fire, flame', radical: '火' },
  山: { character: '山', pinyin: ['shān'], definition: 'mountain, hill', radical: '山' },
  日: { character: '日', pinyin: ['rì'], definition: 'sun; day, date', radical: '日' },
  月: { character: '月', pinyin: ['yuè'], definition: 'moon; month', radical: '月' },
};

let hydrationPromise: Promise<boolean> | null = null;

export async function ensureDictionaryHydrated(
  onProgress?: (loadedLines: number) => void
): Promise<boolean> {
  const alreadyDone = await repository.getMeta(DICT_HYDRATED_KEY);
  if (alreadyDone === 'true') return true;
  if (hydrationPromise) return hydrationPromise;

  hydrationPromise = (async () => {
    for (const url of DICTIONARY_SOURCES) {
      try {
        const res = await fetch(url);
        if (!res.ok) continue;
        const text = await res.text();
        const lines = text.split('\n');
        const batch: DictionaryCharRecord[] = [];
        let processed = 0;

        for (const rawLine of lines) {
          const line = rawLine.trim();
          if (!line) continue;
          try {
            const parsed = JSON.parse(line) as {
              character?: string;
              pinyin?: string[];
              definition?: string;
              decomposition?: string;
              radical?: string;
            };
            if (!parsed.character) continue;
            batch.push({
              character: parsed.character,
              pinyin: Array.isArray(parsed.pinyin) ? parsed.pinyin : [],
              definition: parsed.definition ?? '',
              decomposition: parsed.decomposition,
              radical: parsed.radical,
            });
            processed++;

            if (batch.length >= 600) {
              await repository.bulkSaveDictionaryChars(batch.splice(0, batch.length));
              onProgress?.(processed);
            }
          } catch {
            // Skip malformed NDJSON lines safely
          }
        }

        if (batch.length > 0) {
          await repository.bulkSaveDictionaryChars(batch);
          onProgress?.(processed);
        }

        await repository.setMeta(DICT_HYDRATED_KEY, 'true');
        return true;
      } catch {
        // Fall through to alternate mirror
      }
    }
    return false;
  })();

  try {
    return await hydrationPromise;
  } finally {
    hydrationPromise = null;
  }
}

export async function lookupChineseText(
  rawText: string
): Promise<DictionaryLookupResult> {
  const cleanText = rawText.trim();
  const chars = extractChineseCharacters(cleanText);

  if (chars.length === 0) {
    return {
      text: cleanText,
      pinyin: '',
      meaning: '',
      charBreakdowns: [],
      source: 'none',
    };
  }

  const breakdowns: DictionaryLookupResult['charBreakdowns'] = [];

  for (const ch of chars) {
    const dbRecord =
      (await repository.getDictionaryChar(ch)) ?? BUILTIN_CHAR_FALLBACK[ch] ?? null;

    breakdowns.push({
      char: ch,
      pinyin: dbRecord?.pinyin?.[0] ?? '',
      meaning: dbRecord?.definition ?? '',
      radical: dbRecord?.radical,
    });
  }

  const joined = chars.join('');
  const builtinWord = BUILTIN_WORD_ENTRIES[joined];
  if (builtinWord) {
    return {
      text: joined,
      pinyin: builtinWord.pinyin,
      meaning: builtinWord.meaning,
      charBreakdowns: breakdowns,
      source: 'builtin-word',
    };
  }

  if (chars.length === 1) {
    const single = breakdowns[0];
    return {
      text: single.char,
      pinyin: single.pinyin,
      meaning: single.meaning,
      charBreakdowns: breakdowns,
      source: single.pinyin || single.meaning ? 'makemeahanzi' : 'none',
    };
  }

  const combinedPinyin = breakdowns
    .map((b) => b.pinyin)
    .filter(Boolean)
    .join(' ');
  const combinedMeaning = breakdowns
    .map((b) => (b.meaning ? `${b.char}: ${b.meaning.split(';')[0]}` : ''))
    .filter(Boolean)
    .join(' · ');

  return {
    text: joined,
    pinyin: combinedPinyin,
    meaning: combinedMeaning,
    charBreakdowns: breakdowns,
    source: combinedPinyin || combinedMeaning ? 'partial' : 'none',
  };
}
