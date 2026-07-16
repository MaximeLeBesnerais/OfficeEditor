import { useMemo } from 'react';
import type { SourceType, TargetFormat } from '../types.ts';
import { getAvailableFormats, getFormatMeta } from '../utils/format.ts';

interface FormatSelectorProps {
  sourceType: SourceType;
  selectedFormat: TargetFormat;
  onSelectFormat: (format: TargetFormat) => void;
}

export function FormatSelector({
  sourceType,
  selectedFormat,
  onSelectFormat,
}: FormatSelectorProps) {
  const availableFormats = useMemo(
    () => getAvailableFormats(sourceType),
    [sourceType],
  );
  const hasSource = sourceType !== 'unknown';

  return (
    <div
      className={[
        'flex gap-3 overflow-x-auto pb-2 -mx-1 px-1 scrollbar-hide',
        !hasSource && 'opacity-60',
      ].join(' ')}
    >
      {availableFormats.map((format) => {
        const meta = getFormatMeta(format);
        const isSelected = selectedFormat === format;
        return (
          <button
            key={format}
            type="button"
            disabled={!hasSource}
            onClick={() => onSelectFormat(format)}
            className={[
              'group relative flex items-center gap-3 min-w-[8.5rem] px-4 py-3 rounded-xl border text-left transition-all duration-200',
              'hover:-translate-y-0.5 hover:shadow-md',
              isSelected
                ? `bg-white ${meta.border} ring-2 ${meta.ring} shadow-md`
                : 'bg-white border-slate-200 hover:border-slate-300',
              !hasSource && 'cursor-not-allowed hover:translate-y-0 hover:shadow-none',
            ].join(' ')}
          >
            <span
              className={[
                'w-10 h-10 rounded-lg flex items-center justify-center text-sm font-bold shrink-0 transition-colors',
                isSelected ? meta.bgSolid + ' text-white' : meta.bg + ' ' + meta.text,
              ].join(' ')}
            >
              {meta.icon}
            </span>
            <div className="min-w-0">
              <p className="font-semibold text-slate-900">{meta.label}</p>
              <p className="text-xs text-slate-500 truncate">{meta.description}</p>
            </div>
            {isSelected && (
              <span className="absolute top-2 right-2 w-2 h-2 rounded-full bg-indigo-500" />
            )}
          </button>
        );
      })}
    </div>
  );
}
