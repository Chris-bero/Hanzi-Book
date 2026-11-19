import React from 'react';
import { GridStyle } from '../../types/workbook';

interface WritingGridBoxProps {
  char?: string;
  isTrace?: boolean;
  boxIndex?: number;
  boxSizePx?: number;
  gridStyle: GridStyle;
  gridBgColor: string;
  gridLineColor: string;
  traceGlyphColor: string;
  fontFamily: string;
}

export const WritingGridBox: React.FC<WritingGridBoxProps> = ({
  char,
  isTrace = false,
  boxIndex,
  boxSizePx,
  gridStyle,
  gridBgColor,
  gridLineColor,
  traceGlyphColor,
  fontFamily,
}) => {
  const glyphFontSize = boxSizePx ? Math.max(24, Math.round(boxSizePx * 0.68)) : 34;

  return (
    <div
      className="relative aspect-square border border-[#9E8D7A] overflow-hidden select-none shrink-0"
      style={{
        backgroundColor: gridBgColor,
        width: boxSizePx ? `${boxSizePx}px` : undefined,
        height: boxSizePx ? `${boxSizePx}px` : undefined,
      }}
      aria-label={
        isTrace && char
          ? `Box ${boxIndex ?? ''}: Trace ${char}`
          : `Box ${boxIndex ?? ''}: Empty handwriting square`
      }
    >
      <svg
        viewBox="0 0 100 100"
        className="absolute inset-0 w-full h-full pointer-events-none"
        preserveAspectRatio="none"
      >
        <line
          x1="50"
          y1="0"
          x2="50"
          y2="100"
          stroke={gridLineColor}
          strokeWidth="0.95"
          strokeDasharray={gridStyle === 'cross' ? undefined : '3 3'}
        />
        <line
          x1="0"
          y1="50"
          x2="100"
          y2="50"
          stroke={gridLineColor}
          strokeWidth="0.95"
          strokeDasharray={gridStyle === 'cross' ? undefined : '3 3'}
        />

        {gridStyle === 'mi' && (
          <>
            <line
              x1="0"
              y1="0"
              x2="100"
              y2="100"
              stroke={gridLineColor}
              strokeOpacity="0.65"
              strokeWidth="0.75"
              strokeDasharray="3 3"
            />
            <line
              x1="100"
              y1="0"
              x2="0"
              y2="100"
              stroke={gridLineColor}
              strokeOpacity="0.65"
              strokeWidth="0.75"
              strokeDasharray="3 3"
            />
          </>
        )}
      </svg>

      {isTrace && char && (
        <div
          className="absolute inset-0 flex items-center justify-center leading-none"
          style={{
            fontFamily,
            fontSize: `${glyphFontSize}px`,
            color: traceGlyphColor,
          }}
        >
          {char}
        </div>
      )}
    </div>
  );
};
