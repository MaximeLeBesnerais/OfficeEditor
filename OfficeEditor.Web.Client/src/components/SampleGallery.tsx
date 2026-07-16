import type { Sample } from '../types.ts';
import { getFormatMeta } from '../utils/format.ts';

interface SampleGalleryProps {
  samples: Sample[];
  selectedSample: string | null;
  onSelectSample: (name: string | null) => void;
  mode?: 'compact' | 'gallery';
}

export function SampleGallery({
  samples,
  selectedSample,
  onSelectSample,
  mode = 'compact',
}: SampleGalleryProps) {
  if (samples.length === 0) {
    return (
      <p className="text-center text-slate-500 py-8">
        No samples available from the server.
      </p>
    );
  }

  if (mode === 'gallery') {
    return (
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-5">
        {samples.map((sample) => {
          const meta = getFormatMeta(sample.format);
          return (
            <div
              key={sample.name}
              className="group relative flex flex-col rounded-2xl border border-slate-200 bg-white p-5 shadow-sm hover:shadow-lg hover:-translate-y-1 transition-all duration-200"
            >
              <div className="aspect-[16/10] rounded-xl bg-gradient-to-br from-slate-50 to-slate-100 border border-slate-100 flex items-center justify-center mb-4 overflow-hidden">
                <span
                  className={[
                    'w-16 h-16 rounded-2xl flex items-center justify-center text-2xl font-bold shadow-sm',
                    meta.bg,
                    meta.text,
                  ].join(' ')}
                >
                  {meta.icon}
                </span>
              </div>

              <div className="flex-1">
                <div className="flex items-start justify-between gap-3 mb-2">
                  <h3 className="font-bold text-slate-900 leading-tight min-w-0 break-words">
                    {sample.name}
                  </h3>
                  <span
                    className={[
                      'shrink-0 text-[10px] font-bold uppercase tracking-wider px-2 py-1 rounded-md',
                      meta.bg,
                      meta.text,
                    ].join(' ')}
                  >
                    {meta.label}
                  </span>
                </div>
                <p className="text-sm text-slate-600 line-clamp-2 mb-4">
                  {sample.description}
                </p>
              </div>

              <button
                type="button"
                onClick={() => onSelectSample(sample.name)}
                className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl bg-indigo-600 text-white text-sm font-semibold hover:bg-indigo-700 active:scale-[0.98] transition-all duration-200"
              >
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  className="w-4 h-4"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                  strokeWidth={2}
                  aria-hidden="true"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M3.75 13.5l10.5-11.25L12 10.5h8.25L9.75 21.75 12 13.5H3.75Z"
                  />
                </svg>
                Try this file
              </button>
            </div>
          );
        })}
      </div>
    );
  }

  return (
    <div className="flex gap-3 overflow-x-auto pb-2 -mx-1 px-1 scrollbar-hide">
      {samples.map((sample) => {
        const meta = getFormatMeta(sample.format);
        const isSelected = selectedSample === sample.name;
        return (
          <button
            key={sample.name}
            type="button"
            onClick={() => onSelectSample(isSelected ? null : sample.name)}
            className={[
              'group flex items-center gap-3 shrink-0 text-left rounded-xl border px-3 py-2.5 transition-all duration-200',
              'hover:-translate-y-0.5 hover:shadow-md',
              isSelected
                ? 'border-indigo-500 bg-indigo-50 ring-1 ring-indigo-500'
                : 'border-slate-200 bg-white hover:border-indigo-300',
            ].join(' ')}
          >
            <span
              className={[
                'w-8 h-8 rounded-lg flex items-center justify-center text-xs font-bold',
                meta.bg,
                meta.text,
              ].join(' ')}
            >
              {meta.icon}
            </span>
            <div className="min-w-0">
              <p className={['text-sm font-semibold truncate', isSelected ? 'text-indigo-900' : 'text-slate-900'].join(' ')}>
                {sample.name}
              </p>
              <p className="text-xs text-slate-500 truncate">Try this sample</p>
            </div>
            {isSelected && (
              <span className="ml-1 w-2 h-2 rounded-full bg-indigo-500" />
            )}
          </button>
        );
      })}
    </div>
  );
}
