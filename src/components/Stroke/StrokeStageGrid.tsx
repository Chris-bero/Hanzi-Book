import React, { useMemo } from 'react';
import { getScalingTransform } from '../../services/strokeEngine';
import { HanziCharStrokeData } from '../../types/workbook';

interface StrokeStageGridProps {
  strokeData: HanziCharStrokeData;
  stageSize?: number;
  newStrokeColor?: string;
  completedStrokeColor?: string;
  maxColumns?: number;
  compact?: boolean;
  highlightStage?: number | null;
  onSelectStage?: (stageIndex: number) => void;
}

export const StrokeStageGrid: React.FC<StrokeStageGridProps> = ({
  strokeData,
  stageSize = 40,
  newStrokeColor = '#C62828',
  completedStrokeColor = '#C4BEB4',
  maxColumns,
  compact = false,
  highlightStage = null,
  onSelectStage,
}) => {
  const transformData = useMemo(() => getScalingTransform(100, 10), []);

  const tilePx = Math.max(26, Math.round(stageSize));

  const containerStyle = useMemo<React.CSSProperties>(() => {
    if (compact) {
      return {
        display: 'flex',
        flexWrap: 'wrap',
        gap: '4px',
        alignItems: 'center',
      };
    }
    const count = strokeData.strokeCount;
    const cols = maxColumns
      ? Math.max(1, Math.min(count, maxColumns))
      : Math.min(Math.max(count, 4), 10);
    return {
      display: 'grid',
      gap: '6px',
      gridTemplateColumns: `repeat(${cols}, ${tilePx}px)`,
    };
  }, [compact, strokeData.strokeCount, maxColumns, tilePx]);

  return (
    <div style={containerStyle}>
      {strokeData.strokes.map((_, stageIndex) => {
        const isSelected = highlightStage === stageIndex;
        const interactive = Boolean(onSelectStage);
        return (
          <div
            key={`${strokeData.char}-stage-${stageIndex}`}
            role={interactive ? 'button' : undefined}
            tabIndex={interactive ? 0 : undefined}
            onClick={interactive ? () => onSelectStage?.(stageIndex) : undefined}
            onKeyDown={
              interactive
                ? (e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault();
                      onSelectStage?.(stageIndex);
                    }
                  }
                : undefined
            }
            title={`Stroke ${stageIndex + 1} of ${strokeData.strokeCount}`}
            style={{ width: `${tilePx}px`, height: `${tilePx}px` }}
            className={`relative shrink-0 bg-[#FDFBF8] border flex items-center justify-center p-0.5 transition-colors ${
              isSelected
                ? 'border-burgundy ring-1 ring-burgundy bg-burgundy-tint/40'
                : 'border-hairline'
            } ${interactive ? 'cursor-pointer hover:border-hairline-strong' : ''}`}
          >
            <span
              className="absolute top-[1px] left-[2px] font-mono font-medium text-ink-subtle leading-none tabular-nums select-none"
              style={{ fontSize: tilePx < 34 ? '7.5px' : '8.5px' }}
            >
              {stageIndex + 1}
            </span>
            <svg
              viewBox="0 0 100 100"
              className="w-full h-full"
              aria-label={`${strokeData.char} stroke ${stageIndex + 1} of ${strokeData.strokeCount}`}
            >
              <line
                x1="50"
                y1="0"
                x2="50"
                y2="100"
                stroke="#EAE2D6"
                strokeWidth="0.9"
                strokeDasharray="3 3"
              />
              <line
                x1="0"
                y1="50"
                x2="100"
                y2="50"
                stroke="#EAE2D6"
                strokeWidth="0.9"
                strokeDasharray="3 3"
              />
              <g transform={transformData.transform}>
                {strokeData.strokes.slice(0, stageIndex).map((pathD, idx) => (
                  <path key={`prev-${idx}`} d={pathD} fill={completedStrokeColor} />
                ))}
                <path d={strokeData.strokes[stageIndex]} fill={newStrokeColor} />
              </g>
            </svg>
          </div>
        );
      })}
    </div>
  );
};
