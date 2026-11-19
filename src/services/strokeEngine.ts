import HanziWriter from 'hanzi-writer';
import { repository } from '../storage/db';
import { CharacterStrokeResult, HanziCharStrokeData } from '../types/workbook';

const memoryStrokeCache = new Map<string, HanziCharStrokeData>();
const inFlightRequests = new Map<string, Promise<{ data: HanziCharStrokeData; fromCache: boolean }>>();

const CJK_CHAR_REGEX = /[\u4e00-\u9fff\u3400-\u4dbf]/;

export function extractChineseCharacters(input: string): string[] {
  return Array.from(input.trim()).filter((ch) => CJK_CHAR_REGEX.test(ch));
}

export function getScalingTransform(size: number, padding = 4) {
  return HanziWriter.getScalingTransform(size, size, padding);
}

export async function fetchCharacterStrokeDataWithMeta(
  char: string
): Promise<{ data: HanziCharStrokeData; fromCache: boolean }> {
  const memHit = memoryStrokeCache.get(char);
  if (memHit) return { data: memHit, fromCache: true };

  const existingFlight = inFlightRequests.get(char);
  if (existingFlight) return existingFlight;

  const task = (async () => {
    const dbHit = await repository.getCachedStrokeData(char);
    if (dbHit) {
      memoryStrokeCache.set(char, dbHit);
      return { data: dbHit, fromCache: true };
    }

    const raw = await HanziWriter.loadCharacterData(char);
    if (!raw || !Array.isArray(raw.strokes) || raw.strokes.length === 0) {
      throw new Error(`No stroke-order data available for "${char}".`);
    }

    const parsed: HanziCharStrokeData = {
      char,
      strokes: raw.strokes,
      medians: raw.medians ?? [],
      radStrokes: raw.radStrokes,
      strokeCount: raw.strokes.length,
      source: 'hanzi-writer-data',
      fetchedAt: Date.now(),
    };

    memoryStrokeCache.set(char, parsed);
    await repository.saveCachedStrokeData(parsed);
    return { data: parsed, fromCache: false };
  })();

  inFlightRequests.set(char, task);
  try {
    return await task;
  } finally {
    inFlightRequests.delete(char);
  }
}

export async function fetchCharacterStrokeData(
  char: string
): Promise<HanziCharStrokeData> {
  const { data } = await fetchCharacterStrokeDataWithMeta(char);
  return data;
}

export async function resolveEntryStrokes(
  text: string
): Promise<{
  results: CharacterStrokeResult[];
  cacheHits: number;
  networkFetches: number;
}> {
  const chars = extractChineseCharacters(text);
  const results: CharacterStrokeResult[] = [];
  let cacheHits = 0;
  let networkFetches = 0;

  for (const char of chars) {
    try {
      const { data: strokeData, fromCache } = await fetchCharacterStrokeDataWithMeta(char);
      if (fromCache) cacheHits++;
      else networkFetches++;
      results.push({
        char,
        status: 'ready',
        strokeCount: strokeData.strokeCount,
        strokeData,
      });
    } catch (err) {
      results.push({
        char,
        status: 'unsupported',
        strokeCount: 0,
        error:
          err instanceof Error
            ? err.message
            : `Character "${char}" is missing from Hanzi Writer data.`,
      });
    }
  }

  return { results, cacheHits, networkFetches };
}

export interface StrokeStageRenderOptions {
  pixelSize?: number;
  newStrokeColor?: string;
  completedStrokeColor?: string;
  borderColor?: string;
  backgroundColor?: string;
  numberColor?: string;
}

/**
 * Renders a single cumulative stroke stage directly from real Make Me a Hanzi SVG
 * paths onto a HiDPI HTML5 Canvas and returns PNG bytes for crisp Word/PDF embedding.
 */
export function renderStrokeStageCanvas(
  strokeData: HanziCharStrokeData,
  stageIndex: number,
  options: StrokeStageRenderOptions = {}
): HTMLCanvasElement {
  const pixelSize = options.pixelSize ?? 180;
  const newStrokeColor = options.newStrokeColor ?? '#C62828';
  const completedStrokeColor = options.completedStrokeColor ?? '#C4BEB4';
  const borderColor = options.borderColor ?? '#DDD3C5';
  const backgroundColor = options.backgroundColor ?? '#FDFBF8';
  const numberColor = options.numberColor ?? '#7A7066';

  const canvas = document.createElement('canvas');
  canvas.width = pixelSize;
  canvas.height = pixelSize;
  const ctx = canvas.getContext('2d');
  if (!ctx) return canvas;

  // Paper tile background
  ctx.fillStyle = backgroundColor;
  ctx.fillRect(0, 0, pixelSize, pixelSize);

  // Subtle cross guidelines inside each stage tile
  ctx.save();
  ctx.strokeStyle = '#EAE2D6';
  ctx.lineWidth = Math.max(1, pixelSize * 0.008);
  ctx.setLineDash([pixelSize * 0.03, pixelSize * 0.03]);
  ctx.beginPath();
  ctx.moveTo(pixelSize / 2, 0);
  ctx.lineTo(pixelSize / 2, pixelSize);
  ctx.moveTo(0, pixelSize / 2);
  ctx.lineTo(pixelSize, pixelSize / 2);
  ctx.stroke();
  ctx.restore();

  // Outer tile border
  ctx.strokeStyle = borderColor;
  ctx.lineWidth = Math.max(1.5, pixelSize * 0.014);
  ctx.strokeRect(
    ctx.lineWidth / 2,
    ctx.lineWidth / 2,
    pixelSize - ctx.lineWidth,
    pixelSize - ctx.lineWidth
  );

  // Stage number badge in top-left corner
  ctx.fillStyle = numberColor;
  ctx.font = `600 ${Math.round(pixelSize * 0.13)}px "JetBrains Mono", monospace`;
  ctx.textBaseline = 'top';
  ctx.fillText(String(stageIndex + 1), pixelSize * 0.06, pixelSize * 0.05);

  // Apply exact HanziWriter coordinate transform (1024x1024 y-flipped space).
  // Note: HanziWriter.getScalingTransform returns y in bottom-up space; its SVG
  // transform string is translate(x, height - y) scale(scale, -scale).
  const pad = Math.round(pixelSize * 0.13);
  const { x, y, scale } = HanziWriter.getScalingTransform(pixelSize, pixelSize, pad);

  ctx.save();
  ctx.translate(x, pixelSize - y);
  ctx.scale(scale, -scale);

  // Completed strokes in light gray
  ctx.fillStyle = completedStrokeColor;
  for (let i = 0; i < stageIndex; i++) {
    const path = new Path2D(strokeData.strokes[i]);
    ctx.fill(path);
  }

  // Newly added stroke in red
  if (strokeData.strokes[stageIndex]) {
    ctx.fillStyle = newStrokeColor;
    const activePath = new Path2D(strokeData.strokes[stageIndex]);
    ctx.fill(activePath);
  }

  ctx.restore();
  return canvas;
}

export async function renderStrokeStagePngBytes(
  strokeData: HanziCharStrokeData,
  stageIndex: number,
  options: StrokeStageRenderOptions = {}
): Promise<Uint8Array> {
  const canvas = renderStrokeStageCanvas(strokeData, stageIndex, options);
  const blob = await new Promise<Blob | null>((resolve) =>
    canvas.toBlob((b) => resolve(b), 'image/png')
  );
  if (!blob) {
    throw new Error(`Failed to rasterize stroke stage ${stageIndex + 1} for ${strokeData.char}`);
  }
  const buffer = await blob.arrayBuffer();
  return new Uint8Array(buffer);
}
