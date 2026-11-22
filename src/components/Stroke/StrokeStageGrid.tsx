import React, { useMemo } from 'react';
import {
  buildCumulativeStrokeStages,
  CUMULATIVE_COMPLETED_STROKE_COLOR,
  CUMULATIVE_NEW_STROKE_COLOR,
  getScalingTransform,
  STROKE_STAGE_PADDING_RATIO,
} from '../../services/strokeEngine';
import { HanziCharStrokeData } from '../../types/workbook';

interface StrokeStageGridProps {
  strokeData: HanziCharStrokeData;
  stageSize?: number;
  newStrokeColor?: string;
  completedStrokeColor?: string;
  tileBorderColor?: string;
  guideLineColor?: string;
  numberColor?: string;
  maxColumns?: number;
  compact?: boolean;
  highlightStage?: number | null;
  onSelectStage?: (stageIndex: number) => void;
}

export const StrokeStageGrid: React.FC<StrokeStageGridProps> = ({
  strokeData,
  stageSize = 40,
  newStrokeColor = CUMULATIVE_NEW_STROKE_COLOR,
  completedStrokeColor = CUMULATIVE_COMPLETED_STROKE_COLOR,
  tileBorderColor,
  guideLineColor = '#EAE2D6',
  numberColor,
  maxColumns,
  compact = false,
  highlightStage = null,
  onSelectStage,
}) => {
  const transformData = useMemo(
    () => getScalingTransform(100, Math.round(100 * STROKE_STAGE_PADDING_RATIO)),
    []
  );

  const stages = useMemo(
    () =>
      buildCumulativeStrokeStages(strokeData, {
        newStrokeColor,
        completedStrokeColor,
      }),
    [strokeData, newStrokeColor, completedStrokeColor]
  );

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
    const count = stages.length;
    const cols = maxColumns
      ? Math.max(1, Math.min(count, maxColumns))
      : Math.min(Math.max(count, 4), 10);
    return {
      display: 'grid',
      gap: '6px',
      gridTemplateColumns: `repeat(${cols}, ${tilePx}px)`,
    };
  }, [compact, stages.length, maxColumns, tilePx]);

  return (
    <div style={containerStyle}>
      {stages.map((stage) => {
        const isSelected = highlightStage === stage.stageIndex;
        const interactive = Boolean(onSelectStage);
        return (
          <div
            key={`${strokeData.char}-stage-${stage.stageIndex}`}
            data-step={stage.stepNumber}
            data-cumulative-strokes={stage.layers.length}
            role={interactive ? 'button' : undefined}
            tabIndex={interactive ? 0 : undefined}
            onClick={interactive ? () => onSelectStage?.(stage.stageIndex) : undefined}
            onKeyDown={
              interactive
                ? (e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault();
                      onSelectStage?.(stage.stageIndex);
                    }
                  }
                : undefined
            }
            title={`Step ${stage.stepNumber} of ${stage.totalStrokes} (cumulative strokes 1–${stage.stepNumber})`}
            style={{
              width: `${tilePx}px`,
              height: `${tilePx}px`,
              borderColor: !isSelected && tileBorderColor ? tileBorderColor : undefined,
            }}
            className={`relative shrink-0 bg-white border flex items-center justify-center p-0.5 transition-colors ${
              isSelected
                ? 'border-burgundy ring-1 ring-burgundy bg-burgundy-tint/30'
                : 'border-hairline'
            } ${interactive ? 'cursor-pointer hover:border-hairline-strong' : ''}`}
          >
            <span
              className="absolute top-[1px] left-[2px] font-mono font-medium text-ink-subtle leading-none tabular-nums select-none"
              style={{
                fontSize: tilePx < 34 ? '7.5px' : '8.5px',
                color: numberColor,
              }}
            >
              {stage.stepNumber}
            </span>
            <svg
              viewBox="0 0 100 100"
              className="w-full h-full"
              shapeRendering="geometricPrecision"
              aria-label={`${strokeData.char} cumulative step ${stage.stepNumber} of ${stage.totalStrokes}`}
            >
              <line
                x1="50"
                y1="0"
                x2="50"
                y2="100"
                stroke={guideLineColor}
                strokeWidth="0.9"
                strokeDasharray="3 3"
              />
              <line
                x1="0"
                y1="50"
                x2="100"
                y2="50"
                stroke={guideLineColor}
                strokeWidth="0.9"
                strokeDasharray="3 3"
              />
              <g transform={transformData.transform}>
                {stage.layers.map((layer) => (
                  <path
                    key={`stroke-${layer.strokeIndex}`}
                    d={layer.path}
                    fill={layer.fillColor}
                  />
                ))}
              </g>
            </svg>
          </div>
        );
      })}
    </div>
  );
};
