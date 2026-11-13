import React from 'react';
import { RichTextLine } from '../lib/reportParser';

interface RichTaskLinesProps {
  lines: RichTextLine[];
  hideStrikethrough?: boolean;
  emptyText?: string;
  emphasisTone?: 'default' | 'issue';
}

function renderWithHighlightedDates(text: string): React.ReactNode {
  const parts = text.split(/(\(~?[\d./-]+\)|\([\d./-]+\))/g);
  return parts.map((part, i) => {
    if (/^\(~?[\d./-]+\)$/.test(part)) {
      return (
        <span
          key={i}
          className="font-mono-tabular text-[12.5px] text-slate-500 ml-0.5"
        >
          {part}
        </span>
      );
    }
    return <React.Fragment key={i}>{part}</React.Fragment>;
  });
}

export const RichTaskLines: React.FC<RichTaskLinesProps> = ({
  lines,
  hideStrikethrough = false,
  emptyText = '작성된 내용이 없습니다.',
  emphasisTone = 'default',
}) => {
  const visibleLines = hideStrikethrough
    ? lines.filter((l) => !l.strikethrough)
    : lines;

  if (visibleLines.length === 0) {
    return (
      <p className="text-sm text-slate-400">
        {lines.length > 0 && hideStrikethrough
          ? '완료(취소선)된 항목 외에 남은 내용이 없습니다.'
          : emptyText}
      </p>
    );
  }

  return (
    <ul className="space-y-1.5">
      {visibleLines.map((line, idx) => {
        const trimmed = line.text.trim();
        const isSubDetail = trimmed.startsWith(':');
        const isMainBullet = trimmed.startsWith('-');
        const isTitleLine = idx === 0 && !isMainBullet && !isSubDetail;

        const baseColor = line.strikethrough
          ? 'line-through text-slate-400'
          : emphasisTone === 'issue'
            ? 'text-red-800'
            : isTitleLine
              ? 'font-semibold text-slate-900'
              : isSubDetail
                ? 'text-slate-600'
                : 'text-slate-800';

        if (isSubDetail) {
          const cleanText = trimmed.replace(/^:\s*/, '');
          return (
            <li
              key={idx}
              className={`pl-5 text-[14px] leading-relaxed flex items-start gap-2 ${baseColor}`}
            >
              <span
                className="text-slate-400 select-none shrink-0 mt-[1px]"
                aria-hidden="true"
              >
                ·
              </span>
              <span className="break-words min-w-0">
                {renderWithHighlightedDates(cleanText)}
              </span>
            </li>
          );
        }

        if (isMainBullet) {
          const cleanText = trimmed.replace(/^-\s*/, '');
          return (
            <li
              key={idx}
              className={`pl-1 text-[14px] leading-relaxed flex items-start gap-2 ${baseColor}`}
            >
              <span
                className="text-slate-400 select-none shrink-0 font-mono-tabular"
                aria-hidden="true"
              >
                -
              </span>
              <span className="break-words min-w-0">
                {renderWithHighlightedDates(cleanText)}
              </span>
            </li>
          );
        }

        return (
          <li
            key={idx}
            className={`text-[14.5px] leading-relaxed break-words ${baseColor}`}
          >
            {renderWithHighlightedDates(trimmed)}
          </li>
        );
      })}
    </ul>
  );
};
