import React from 'react';
import { RichTextLine } from '../lib/reportParser';

interface RichTaskLinesProps {
  lines: RichTextLine[];
  hideStrikethrough?: boolean;
  emptyText?: string;
  emphasisTone?: 'default' | 'issue';
}

interface TaskSection {
  subPrefix: string | null;
  titleLine: RichTextLine | null;
  bodyLines: RichTextLine[];
}

const SUB_NUMBER_LINE_RE = /^(\d{1,2}\s*[-–—]\s*\d{1,2}[)）])\s*(.*)$/;

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

function groupIntoSections(lines: RichTextLine[]): TaskSection[] {
  const sections: TaskSection[] = [];
  let current: TaskSection | null = null;

  for (const line of lines) {
    const trimmed = line.text.trim();
    const subMatch = trimmed.match(SUB_NUMBER_LINE_RE);

    if (subMatch) {
      if (current) sections.push(current);
      const prefix = subMatch[1].replace(/\s+/g, '').replace(/[–—]/g, '-');
      const rest = subMatch[2].trim();
      current = {
        subPrefix: prefix,
        titleLine: rest
          ? { text: rest, strikethrough: line.strikethrough }
          : null,
        bodyLines: [],
      };
      continue;
    }

    if (!current) {
      current = {
        subPrefix: null,
        titleLine: null,
        bodyLines: [line],
      };
    } else if (current.subPrefix && !current.titleLine) {
      // If a sub-number line had no text on the same line, promote the next line to its heading
      current.titleLine = line;
    } else {
      current.bodyLines.push(line);
    }
  }

  if (current) sections.push(current);
  return sections;
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

  const sections = groupIntoSections(visibleLines);

  const renderBodyLine = (
    line: RichTextLine,
    idx: number,
    isFirstOfUnprefixedSection: boolean
  ) => {
    const trimmed = line.text.trim();
    const isSubDetail = trimmed.startsWith(':');
    const isMainBullet = trimmed.startsWith('-');
    const isTitleLine =
      isFirstOfUnprefixedSection && !isMainBullet && !isSubDetail;

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
          className={`pl-4 text-[14px] leading-relaxed flex items-start gap-2 ${baseColor}`}
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
          className={`text-[14px] leading-relaxed flex items-start gap-2 ${baseColor}`}
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
        className={`text-[14px] leading-relaxed break-words ${baseColor}`}
      >
        {renderWithHighlightedDates(trimmed)}
      </li>
    );
  };

  return (
    <div className="space-y-3.5">
      {sections.map((section, sIdx) => {
        if (section.subPrefix) {
          const titleColor = section.titleLine?.strikethrough
            ? 'line-through text-slate-400'
            : emphasisTone === 'issue'
              ? 'font-semibold text-red-800'
              : 'font-semibold text-slate-900';

          return (
            <div key={sIdx} className="flex items-start gap-2">
              <span className="font-mono-tabular text-[13.5px] font-semibold text-slate-600 leading-relaxed shrink-0 min-w-[2.4rem]">
                {section.subPrefix}
              </span>
              <div className="min-w-0 flex-1 space-y-1.5">
                {section.titleLine && (
                  <p className={`text-[14.5px] leading-relaxed break-words ${titleColor}`}>
                    {renderWithHighlightedDates(section.titleLine.text)}
                  </p>
                )}
                {section.bodyLines.length > 0 && (
                  <ul className="space-y-1.5">
                    {section.bodyLines.map((line, lIdx) =>
                      renderBodyLine(line, lIdx, false)
                    )}
                  </ul>
                )}
              </div>
            </div>
          );
        }

        return (
          <ul key={sIdx} className="space-y-1.5">
            {section.bodyLines.map((line, lIdx) =>
              renderBodyLine(line, lIdx, lIdx === 0)
            )}
          </ul>
        );
      })}
    </div>
  );
};
