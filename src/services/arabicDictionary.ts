import { repository } from '../storage/db';
import { extractChineseCharacters } from './strokeEngine';

const AR_CACHE_PREFIX = 'ar_trans_v1_';

/**
 * Curated offline Chinese -> Arabic dictionary covering starter vocabulary,
 * HSK 1–2 essential characters, pronouns, numbers, greetings, family, time,
 * nature, and common everyday words.
 */
export const BUILTIN_ARABIC_WORDS: Record<string, string> = {
  // Core greetings & polite phrases
  你好: 'مرحباً',
  您好: 'مرحباً (للاحترام)',
  谢谢: 'شكراً',
  再见: 'إلى اللقاء',
  对不起: 'آسف / عفواً',
  没关系: 'لا بأس',
  不客气: 'على الرحب والسعة',
  请问: 'لو سمحت / أستفسر',
  早上好: 'صباح الخير',
  晚上好: 'مساء الخير',
  晚安: 'تصبح على خير',

  // Countries, languages & education
  中国: 'الصين',
  中文: 'اللغة الصينية',
  汉语: 'اللغة الصينية (الماندرين)',
  汉字: 'الحروف الصينية',
  阿拉伯语: 'اللغة العربية',
  英语: 'اللغة الإنجليزية',
  学习: 'يدرس / يتعلم',
  朋友: 'صديق',
  同学: 'زميل دراسة',
  老师: 'معلم / أستاذ',
  学生: 'طالب',
  学校: 'مدرسة',
  大学: 'جامعة',
  书本: 'كتاب',
  写字: 'يكتب حروفاً',
  读书: 'يقرأ / يدرس',

  // People & family
  家人: 'العائلة',
  爸爸: 'أب',
  妈妈: 'أم',
  哥哥: 'أخ أكبر',
  弟弟: 'أخ أصغر',
  姐姐: 'أخت كبرى',
  妹妹: 'أخت صغرى',
  儿子: 'ابن',
  女儿: 'ابنة',
  先生: 'سيد',
  小姐: 'آنسة',
  医生: 'طبيب',
  我们: 'نحن',
  你们: 'أنتم',
  他们: 'هم',
  大家: 'الجميع',

  // Time & dates
  今天: 'اليوم',
  明天: 'غداً',
  昨天: 'أمس',
  现在: 'الآن',
  时候: 'وقت / حين',
  时间: 'الوقت',
  上午: 'قبل الظهر',
  中午: 'الظهر',
  下午: 'بعد الظهر',
  晚上: 'المساء',
  星期: 'أسبوع',
  今年: 'هذا العام',
  明年: 'العام القادم',
  去年: 'العام الماضي',
  分钟: 'دقيقة',
  小时: 'ساعة',

  // Everyday verbs, adjectives & nouns
  高兴: 'سعيد / مسرور',
  开心: 'سعيد',
  喜欢: 'يحب / يعجب بـ',
  天气: 'الطقس',
  工作: 'عمل / وظيفة',
  名字: 'اسم',
  什么: 'ماذا / ما',
  怎么: 'كيف',
  怎么样: 'كيف الحال',
  多少: 'كم (للعدد/السعر)',
  哪里: 'أين',
  这里: 'هنا',
  那里: 'هناك',
  没有: 'لا يوجد / ليس لديه',
  可以: 'يمكن / يستطيع',
  知道: 'يعرف',
  认识: 'يتعرف على / يعرف',
  看见: 'يرى',
  听见: 'يسمع',
  说话: 'يتحدث / يتكلم',
  吃饭: 'يأكل الطعام',
  喝水: 'يشرب ماء',
  喝茶: 'يشرب شاياً',
  睡觉: 'ينام',
  回家: 'يعود إلى المنزل',
  东西: 'شيء / أغراض',
  电脑: 'حاسوب',
  电视: 'تلفاز',
  电影: 'فيلم',
  电话: 'هاتف',
  手机: 'هاتف محمول',
  飞机: 'طائرة',
  火车: 'قطار',
  出租车: 'سيارة أجرة',
  衣服: 'ملابس',
  水果: 'فواكه',
  苹果: 'تفاحة',
  米饭: 'أرز مطبوخ',
  医院: 'مستشفى',
  饭店: 'مطعم / فندق',
  商店: 'متجر',
  房间: 'غرفة',
  问题: 'سؤال / مشكلة',
  意思: 'معنى',
  准备: 'يستعد / يجهز',
  开始: 'يبدأ',
  希望: 'يأمل / أمل',
  帮助: 'يساعد / مساعدة',
  运动: 'رياضة',
  旅游: 'سفر / سياحة',
  健康: 'صحة / سليم',
  漂亮: 'جميل',
  好吃: 'لذيذ',
  便宜: 'رخيص',
  非常: 'جداً / للغاية',
  一起: 'معاً',
  已经: 'بالفعل / قد',
  因为: 'لأن',
  所以: 'لذلك',
  但是: 'ولكن',
  如果: 'إذا',
};

export const BUILTIN_ARABIC_CHARS: Record<string, string> = {
  // Starter & essential pronouns / particles
  不: 'لا / ليس',
  我: 'أنا',
  你: 'أنت',
  您: 'حضرتك / أنت (للاحترام)',
  他: 'هو',
  她: 'هي',
  它: 'هو/هي (لغير العاقل)',
  们: 'علامة الجمع',
  的: 'أداة ملكية/وصف',
  了: 'أداة إتمام الفعل',
  吗: 'أداة استفهام (هل)',
  呢: 'ماذا عن...؟',
  吧: 'أداة اقتراح',
  和: 'و (حرف عطف) / سلام',
  在: 'في / موجود',
  是: 'يكون / نعم',
  有: 'يملك / يوجد',
  没: 'لا / ليس',
  这: 'هذا',
  那: 'ذلك',
  哪: 'أي / أين',
  谁: 'مَن',
  几: 'كم (لعدد صغير)',
  很: 'جداً',
  太: 'جداً / للغاية',
  都: 'الكل / جميعاً',
  也: 'أيضاً',
  还: 'ما زال / أيضاً',
  就: 'فوراً / إذن',
  再: 'مرة أخرى',
  最: 'الأكثر',
  真: 'حقيقي / حقاً',

  // Numbers
  零: 'صفر',
  一: 'واحد',
  二: 'اثنان',
  两: 'اثنان',
  三: 'ثلاثة',
  四: 'أربعة',
  五: 'خمسة',
  六: 'ستة',
  七: 'سبعة',
  八: 'ثمانية',
  九: 'تسعة',
  十: 'عشرة',
  百: 'مئة',
  千: 'ألف',
  万: 'عشرة آلاف',

  // Core adjectives
  好: 'جيد / بخير',
  大: 'كبير',
  小: 'صغير',
  多: 'كثير',
  少: 'قليل',
  高: 'عالي / طويل',
  矮: 'قصير',
  长: 'طويل',
  短: 'قصير',
  新: 'جديد',
  旧: 'قديم',
  老: 'كبير السن / قديم',
  快: 'سريع',
  慢: 'بطيء',
  早: 'مبكر / صباح',
  晚: 'متأخر / مساء',
  冷: 'بارد',
  热: 'حار',
  红: 'أحمر',
  白: 'أبيض',
  黑: 'أسود',
  蓝: 'أزرق',
  绿: 'أخضر',
  黄: 'أصفر',
  忙: 'مشغول',
  累: 'متعب',
  饿: 'جائع',
  贵: 'غالي الثمن',
  错: 'خطأ',
  对: 'صحيح / نحو',
  远: 'بعيد',
  近: 'قريب',

  // People, body & nature
  人: 'إنسان / شخص',
  男: 'رجل / ذكر',
  女: 'امرأة / أنثى',
  子: 'ابن / طفل',
  儿: 'طفل / ابن',
  父: 'أب',
  母: 'أم',
  家: 'بيت / عائلة',
  友: 'صديق',
  口: 'فم',
  目: 'عين',
  耳: 'أذن',
  手: 'يد',
  足: 'قدم',
  心: 'قلب',
  头: 'رأس',
  水: 'ماء',
  火: 'نار',
  山: 'جبل',
  木: 'خشب / شجرة',
  林: 'غابة',
  森: 'غابة كثيفة',
  土: 'تراب / أرض',
  金: 'ذهب / معدن',
  石: 'حجر',
  田: 'حقل',
  日: 'شمس / يوم',
  月: 'قمر / شهر',
  星: 'نجم',
  天: 'سماء / يوم',
  气: 'هواء / طقس',
  风: 'رياح',
  雨: 'مطر',
  雪: 'ثلج',
  云: 'سحاب',
  花: 'زهرة',
  草: 'عشب',
  树: 'شجرة',
  鸟: 'طائر',
  鱼: 'سمك',
  马: 'حصان',
  牛: 'بقرة',
  羊: 'خروف',
  狗: 'كلب',
  猫: 'قطة',
  龙: 'تنين',

  // Space, time & directions
  中: 'وسط / الصين',
  国: 'دولة / بلد',
  上: 'فوق / أعلى',
  下: 'تحت / أسفل',
  左: 'يسار',
  右: 'يمين',
  前: 'أمام / قبل',
  后: 'خلف / بعد',
  里: 'داخل',
  外: 'خارج',
  东: 'شرق',
  西: 'غرب',
  南: 'جنوب',
  北: 'شمال',
  年: 'سنة / عام',
  今: 'اليوم / الآن',
  明: 'غداً / مشرق',
  时: 'وقت / ساعة',
  分: 'دقيقة / يقسم',
  点: 'ساعة / نقطة',
  秒: 'ثانية',
  春: 'ربيع',
  夏: 'صيف',
  秋: 'خريف',
  冬: 'شتاء',

  // Core verbs & study / life
  学: 'يتعلم / يدرس',
  习: 'يتدرب / يمارس',
  文: 'لغة / نص / كتابة',
  字: 'حرف / كلمة',
  语: 'لغة / كلام',
  言: 'كلام / قول',
  话: 'حديث / كلام',
  书: 'كتاب',
  读: 'يقرأ',
  写: 'يكتب',
  听: 'يستمع',
  说: 'يقول / يتحدث',
  看: 'ينظر / يرى / يقرأ',
  见: 'يرى / يقابل',
  问: 'يسأل',
  答: 'يجيب',
  叫: 'يُدعى / ينادي',
  去: 'يذهب',
  来: 'يأتي',
  回: 'يعود',
  走: 'يمشي',
  跑: 'يركض',
  坐: 'يجلس',
  站: 'يقف / محطة',
  住: 'يسكن / يعيش',
  吃: 'يأكل',
  喝: 'يشرب',
  买: 'يشتري',
  卖: 'يبيع',
  给: 'يعطي',
  开: 'يفتح / يقود',
  关: 'يغلق',
  做: 'يفعل / يصنع',
  作: 'يعمل / يؤلف',
  工: 'عمل',
  用: 'يستخدم',
  想: 'يفكر / يريد / يشتاق',
  要: 'يريد / يحتاج',
  会: 'يستطيع / سـ (للمستقبل)',
  能: 'يقدر / يستطيع',
  爱: 'يحب / حب',
  笑: 'يضحك / يبتسم',
  哭: 'يبكي',
  玩: 'يلعب',
  唱: 'يغني',
  歌: 'أغنية',
  画: 'يرسم / لوحة',
  睡: 'ينام',
  觉: 'نوم / يشعر',
  请: 'من فضلك / يدعو',
  谢: 'يشكر',
  等: 'ينتظر',
  找: 'يبحث عن',
  穿: 'يرتدي',
  洗: 'يغسل',
  进: 'يدخل',
  出: 'يخرج',
  到: 'يصل إلى',
  过: 'يمر / يعبر',
  生: 'حياة / يولد / طالب',
  活: 'يعيش / حي',
  死: 'يموت',

  // Objects & places
  茶: 'شاي',
  饭: 'أرز / وجبة',
  菜: 'خضار / طبق طعام',
  肉: 'لحم',
  果: 'فاكهة',
  米: 'أرز',
  杯: 'كوب',
  碗: 'وعاء',
  钱: 'مال / نقود',
  车: 'سيارة / مركبة',
  船: 'سفينة / قارب',
  路: 'طريق',
  门: 'باب',
  窗: 'نافذة',
  床: 'سرير',
  桌: 'طاولة',
  椅: 'كرسي',
  纸: 'ورق',
  笔: 'قلم',
  电: 'كهرباء',
  影: 'ظل / فيلم',
  灯: 'مصباح',
  校: 'مدرسة',
  院: 'فناء / معهد',
  店: 'متجر',
  城: 'مدينة',
  市: 'سوق / مدينة',
  村: 'قرية',
  桥: 'جسر',
  海: 'بحر',
  河: 'نهر',
  湖: 'بحيرة',
  永: 'أبدي / دائم',
  安: 'سلام / أمان',
  福: 'بركة / سعادة',
  乐: 'فرح / موسيقى',
  美: 'جميل / جمال',
  力: 'قوة',
  王: 'ملك',
  玉: 'يشم (حجر كريم)',
  宝: 'كنز',
  光: 'ضوء / نور',
  音: 'صوت / نغمة',
  色: 'لون',
  名: 'اسم',
  号: 'رقم / تاريخ',
  岁: 'سنة (من العمر)',
  个: 'أداة عد عامة',
  本: 'أداة عد للكتب / أصل',
  些: 'بعض',
  每: 'كل',
  次: 'مرة',
  件: 'قطعة / أمر',
  块: 'قطعة / يوان',
};

/**
 * Synchronous local lookup for Arabic meaning (exact word -> single char -> composed chars).
 */
export function lookupLocalArabic(rawText: string): string {
  const chars = extractChineseCharacters(rawText);
  if (chars.length === 0) return '';
  const joined = chars.join('');

  if (BUILTIN_ARABIC_WORDS[joined]) {
    return BUILTIN_ARABIC_WORDS[joined];
  }
  if (chars.length === 1 && BUILTIN_ARABIC_CHARS[chars[0]]) {
    return BUILTIN_ARABIC_CHARS[chars[0]];
  }
  return '';
}

function composeCharArabicFallback(chars: string[]): string {
  const parts = chars
    .map((ch) => {
      const ar = BUILTIN_ARABIC_CHARS[ch];
      return ar ? ar.split('/')[0].trim() : '';
    })
    .filter(Boolean);

  if (parts.length === 0) return '';
  return Array.from(new Set(parts)).join(' · ');
}

function containsArabicScript(text: string): boolean {
  return /[\u0600-\u06FF]/.test(text);
}

/**
 * Resolves an Arabic translation for Chinese text:
 * 1. Instant local dictionary match (offline, zero latency)
 * 2. IndexedDB cache from previous lookups
 * 3. Free public MyMemory translation API (zh-CN -> ar, 2.5s timeout, no API key)
 * 4. Local per-character Arabic fallback if offline
 */
export async function translateChineseToArabic(
  rawText: string,
  options?: { allowNetwork?: boolean }
): Promise<string> {
  const chars = extractChineseCharacters(rawText);
  if (chars.length === 0) return '';
  const joined = chars.join('');

  // 1. Instant curated local dictionary
  const local = lookupLocalArabic(joined);
  if (local) return local;

  // 2. IndexedDB cached translation
  try {
    const cached = await repository.getMeta(`${AR_CACHE_PREFIX}${joined}`);
    if (cached && cached.trim()) {
      return cached.trim();
    }
  } catch {
    // IndexedDB unavailable in non-browser context
  }

  // 3. Free public translation fallback (no paid key required)
  if (options?.allowNetwork !== false && typeof fetch === 'function') {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 2500);
    try {
      const url = `https://api.mymemory.translated.net/get?q=${encodeURIComponent(joined)}&langpair=zh-CN|ar`;
      const res = await fetch(url, { signal: controller.signal });
      if (res.ok) {
        const data = (await res.json()) as {
          responseData?: { translatedText?: string };
          responseStatus?: number;
        };
        const candidate = (data.responseData?.translatedText ?? '').trim();
        if (
          candidate &&
          containsArabicScript(candidate) &&
          !candidate.toUpperCase().includes('MYMEMORY WARNING') &&
          candidate.length <= 80
        ) {
          try {
            await repository.setMeta(`${AR_CACHE_PREFIX}${joined}`, candidate);
          } catch {
            // Ignore cache write errors
          }
          return candidate;
        }
      }
    } catch {
      // Network offline or timeout — gracefully fall back below
    } finally {
      clearTimeout(timer);
    }
  }

  // 4. Local per-character breakdown fallback for multi-character words
  if (chars.length > 1) {
    return composeCharArabicFallback(chars);
  }

  return '';
}
