import React, { useEffect, useRef, useState } from 'react';
import HanziWriter from 'hanzi-writer';
import { Play, RotateCcw, X } from 'lucide-react';
import { WorkbookEntry, WorkbookSettings } from '../../types/workbook';
import { StrokeStageGrid } from '../Stroke/StrokeStageGrid';

interface StrokeSequenceModalProps {
  entry: WorkbookEntry;
  settings: WorkbookSettings;
  onClose: () => void;
}

export const StrokeSequenceModal: React.FC<StrokeSequenceModalProps> = ({
  entry,
  settings,
  onClose,
}) => {
  const writerRefs = useRef<Array<HTMLDivElement | null>>([]);
  const writerInstances = useRef<HanziWriter[]>([]);
  const [selectedStageByChar, setSelectedStageByChar] = useState<Record<number, number | null>>({});
  const [isAnimating, setIsAnimating] = useState(false);

  useEffect(() => {
    writerInstances.current = [];

    entry.charResults.forEach((charRes, idx) => {
      const container = writerRefs.current[idx];
      if (!container || charRes.status !== 'ready' || !charRes.strokeData) return;
      container.innerHTML = '';

      const writer = HanziWriter.create(container, charRes.char, {
        width: 144,
        height: 144,
        padding: 10,
        showOutline: true,
        strokeColor: '#1E1916',
        outlineColor: '#DDD3C5',
        radicalColor: settings.accentColor,
        strokeAnimationSpeed: 1.35,
        delayBetweenStrokes: 240,
        charDataLoader: () => Promise.resolve(charRes.strokeData!),
      });

      writerInstances.current.push(writer);
    });

    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', handleKey);

    return () => {
      window.removeEventListener('keydown', handleKey);
      writerRefs.current.forEach((el) => {
        if (el) el.innerHTML = '';
      });
    };
  }, [entry, settings.accentColor, onClose]);

  const handleAnimateAll = async () => {
    if (isAnimating) return;
    setIsAnimating(true);
    try {
      for (const writer of writerInstances.current) {
        await writer.animateCharacter();
      }
    } finally {
      setIsAnimating(false);
    }
  };

  const handleReset = () => {
    setSelectedStageByChar({});
    for (const writer of writerInstances.current) {
      writer.showCharacter();
    }
  };

  const handleStageSelect = (charIdx: number, stageIndex: number) => {
    setSelectedStageByChar((prev) => ({ ...prev, [charIdx]: stageIndex }));
    const writer = writerInstances.current[charIdx];
    if (writer) {
      writer.animateStroke(stageIndex);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-[#1E1916]/55 backdrop-blur-[1px] flex items-center justify-center p-5"
      role="dialog"
      aria-modal="true"
      aria-label={`Stroke sequence preview for ${entry.text}`}
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="bg-white border border-hairline w-full max-w-4xl max-h-[90vh] overflow-y-auto shadow-[0_24px_60px_rgba(30,25,22,0.2)]">
        {/* Header */}
        <div className="flex flex-wrap items-center justify-between gap-4 px-6 py-5 border-b border-hairline">
          <div>
            <h2 className="font-display text-2xl font-bold text-ink m-0">
              Stroke Sequence Inspector —{' '}
              <span style={{ fontFamily: settings.fontFamily }}>{entry.text}</span>
            </h2>
            <div className="mt-1 flex items-center gap-2 text-xs text-ink-muted tabular-nums">
              <span className="font-semibold text-burgundy">{entry.pinyin || 'No pinyin'}</span>
              <span aria-hidden="true">·</span>
              <span>{entry.meaning || 'No English meaning'}</span>
              <span aria-hidden="true">·</span>
              <span className="font-mono">{entry.totalStrokeCount} strokes</span>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={handleAnimateAll}
              disabled={isAnimating}
              className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold bg-burgundy text-white hover:bg-burgundy-dark disabled:opacity-50 transition-colors cursor-pointer whitespace-nowrap shrink-0"
            >
              <Play className="w-3.5 h-3.5" />
              {isAnimating ? 'Playing…' : 'Animate Sequence'}
            </button>
            <button
              type="button"
              onClick={handleReset}
              className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium border border-hairline-strong text-ink hover:bg-ivory-deep transition-colors cursor-pointer whitespace-nowrap shrink-0"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              Reset
            </button>
            <button
              type="button"
              onClick={onClose}
              aria-label="Close stroke inspector"
              className="p-2 text-ink-muted hover:text-ink hover:bg-ivory-deep transition-colors cursor-pointer"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* Character Sections */}
        <div className="p-6 flex flex-col gap-6">
          {entry.charResults.map((charRes, idx) => (
            <div
              key={`${entry.id}-inspect-${idx}`}
              className="pb-6 border-b border-hairline last:border-b-0 last:pb-0"
            >
              <div className="flex flex-col md:flex-row gap-6 items-start">
                <div className="shrink-0">
                  <div className="text-xs font-medium text-ink-subtle mb-2">
                    Live Vector Player
                  </div>
                  <div className="relative w-36 h-36 border border-hairline-strong bg-paper-grid">
                    <svg
                      viewBox="0 0 100 100"
                      className="absolute inset-0 w-full h-full pointer-events-none"
                    >
                      <line x1="50" y1="0" x2="50" y2="100" stroke="#D3C7B5" strokeDasharray="3 3" />
                      <line x1="0" y1="50" x2="100" y2="50" stroke="#D3C7B5" strokeDasharray="3 3" />
                      <line x1="0" y1="0" x2="100" y2="100" stroke="#DFD5C5" strokeDasharray="3 3" />
                      <line x1="100" y1="0" x2="0" y2="100" stroke="#DFD5C5" strokeDasharray="3 3" />
                    </svg>
                    <div
                      ref={(el) => {
                        writerRefs.current[idx] = el;
                      }}
                      className="relative z-10 w-36 h-36"
                    />
                  </div>
                  <div className="mt-2 text-[11px] text-ink-subtle font-mono tabular-nums">
                    Source: Make Me a Hanzi
                  </div>
                </div>

                <div className="flex-1 min-w-0">
                  <div className="flex items-center justify-between mb-2.5">
                    <span className="text-xs font-semibold text-ink">
                      Cumulative Stroke Stages ({charRes.strokeCount}{' '}
                      {charRes.strokeCount === 1 ? 'stroke' : 'strokes'})
                    </span>
                    <span className="text-xs text-ink-subtle">
                      Click any stage to replay that stroke
                    </span>
                  </div>

                  {charRes.status === 'ready' && charRes.strokeData ? (
                    <StrokeStageGrid
                      strokeData={charRes.strokeData}
                      stageSize={58}
                      newStrokeColor={settings.newStrokeColor}
                      completedStrokeColor={settings.completedStrokeColor}
                      maxColumns={8}
                      highlightStage={selectedStageByChar[idx] ?? null}
                      onSelectStage={(stageIndex) => handleStageSelect(idx, stageIndex)}
                    />
                  ) : (
                    <div className="p-4 border border-[#EBC6C8] bg-[#FDF4F4] text-xs text-[#A51D24]">
                      {charRes.error || `Stroke-order data unavailable for “${charRes.char}”.`}
                    </div>
                  )}
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};
