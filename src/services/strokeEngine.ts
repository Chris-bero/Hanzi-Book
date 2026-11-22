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

export const CUMULATIVE_COMPLETED_STROKE_COLOR = '#C8C2B8';
export const CUMULATIVE_NEW_STROKE_COLOR = '#C83232';
export const STROKE_STAGE_PADDING_RATIO = 0.1;

export interface CumulativeStrokeLayer {
  strokeIndex: number;
  path: string;
  isNewest: boolean;
  fillColor: string;
}

export interface CumulativeStrokeStage {
  stepNumber: number;
  stageIndex: number;
  totalStrokes: number;
  layers: CumulativeStrokeLayer[];
}

export interface StrokeStageRenderOptions {
  pixelSize?: number;
  newStrokeColor?: string;
  completedStrokeColor?: string;
  borderColor?: string;
  backgroundColor?: string;
  guideLineColor?: string;
  numberColor?: string;
}

/**
 * Canonical cumulative stroke-order builder shared by Live Preview (SVG),
 * Stroke Sequence Inspector, editable Microsoft Word (.docx), and PDF exports.
 *
 * For every step N (1..totalStrokes):
 * - Always renders ALL stroke paths from index 0 to N-1 in exact drawing order.
 * - Previously completed strokes (0..N-2) use completedStrokeColor (#C8C2B8).
 * - Only the newest stroke (N-1) uses newStrokeColor (#C83232).
 * - Never renders an isolated stroke without previous strokes.
 */
export function buildCumulativeStrokeStages(
  strokeData: HanziCharStrokeData,
  options?: {
    newStrokeColor?: string;
    completedStrokeColor?: string;
  }
): CumulativeStrokeStage[] {
  const newColor = options?.newStrokeColor?.trim() || CUMULATIVE_NEW_STROKE_COLOR;
  const completedColor = options?.completedStrokeColor?.trim() || CUMULATIVE_COMPLETED_STROKE_COLOR;
  const strokes = Array.isArray(strokeData.strokes) ? strokeData.strokes : [];
  const totalStrokes = strokes.length;

  return strokes.map((_, stageIndex) => {
    const layers: CumulativeStrokeLayer[] = [];
    for (let i = 0; i <= stageIndex; i++) {
      const path = strokes[i];
      if (!path) continue;
      const isNewest = i === stageIndex;
      layers.push({
        strokeIndex: i,
        path,
        isNewest,
        fillColor: isNewest ? newColor : completedColor,
      });
    }

    return {
      stepNumber: stageIndex + 1,
      stageIndex,
      totalStrokes,
      layers,
    };
  });
}

/**
 * Renders a single cumulative stroke stage directly from real Make Me a Hanzi SVG
 * paths onto a HiDPI HTML5 Canvas and returns PNG bytes for crisp Word/PDF embedding.
 * Uses the exact same cumulative stage builder, padding ratio, and coordinate transform
 * as the Live Preview SVG component.
 */
export function renderStrokeStageCanvas(
  strokeData: HanziCharStrokeData,
  stageIndex: number,
  options: StrokeStageRenderOptions = {}
): HTMLCanvasElement {
  const pixelSize = options.pixelSize ?? 180;
  const newStrokeColor = options.newStrokeColor?.trim() || CUMULATIVE_NEW_STROKE_COLOR;
  const completedStrokeColor =
    options.completedStrokeColor?.trim() || CUMULATIVE_COMPLETED_STROKE_COLOR;
  const borderColor = options.borderColor ?? '#DDD3C5';
  const backgroundColor = options.backgroundColor ?? '#FFFFFF';
  const guideLineColor = options.guideLineColor ?? '#EAE2D6';
  const numberColor = options.numberColor ?? '#7A7066';

  const canvas = document.createElement('canvas');
  canvas.width = pixelSize;
  canvas.height = pixelSize;
  const ctx = canvas.getContext('2d');
  if (!ctx) return canvas;

  const stages = buildCumulativeStrokeStages(strokeData, {
    newStrokeColor,
    completedStrokeColor,
  });
  const clampedIdx = Math.max(0, Math.min(stageIndex, Math.max(0, stages.length - 1)));
  const stage = stages[clampedIdx];

  // Crisp tile background
  ctx.fillStyle = backgroundColor;
  ctx.fillRect(0, 0, pixelSize, pixelSize);

  // Subtle cross guidelines inside each stage tile
  ctx.save();
  ctx.strokeStyle = guideLineColor;
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
  ctx.fillText(String(clampedIdx + 1), pixelSize * 0.06, pixelSize * 0.05);

  // Apply exact HanziWriter coordinate transform (1024x1024 y-flipped space)
  // matching StrokeStageGrid's 10% padding ratio for consistent centering and scale.
  const pad = Math.round(pixelSize * STROKE_STAGE_PADDING_RATIO);
  const { x, y, scale } = HanziWriter.getScalingTransform(pixelSize, pixelSize, pad);

  ctx.save();
  ctx.translate(x, pixelSize - y);
  ctx.scale(scale, -scale);

  // Draw ALL cumulative strokes from index 0 through stageIndex in exact drawing order
  if (stage) {
    for (const layer of stage.layers) {
      ctx.fillStyle = layer.fillColor;
      const path = new Path2D(layer.path);
      ctx.fill(path, 'nonzero');
    }
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
