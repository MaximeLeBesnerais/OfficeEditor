export interface TimingStats {
  slides: number;
  totalMs: number;
  perSlideMs?: number;
  label?: string;
  extra?: { label: string; ms: number };
}

interface TimingReceiptProps {
  stats: TimingStats;
}

function formatMs(ms: number): string {
  return ms >= 100 ? ms.toFixed(0) : ms.toFixed(1);
}

function formatPerSlide(ms: number): string {
  return ms.toFixed(1);
}

// Server-provided timings only — never measured client-side.
export function TimingReceipt({ stats }: TimingReceiptProps) {
  const perSlide =
    stats.perSlideMs ?? (stats.slides > 0 ? stats.totalMs / stats.slides : 0);
  const label = stats.label ?? 'rendered';

  return (
    <div className="inline-flex flex-wrap items-center gap-x-2 gap-y-1 rounded-full border border-slate-700/80 bg-slate-900/80 px-4 py-1.5 font-mono text-sm text-slate-300 shadow-lg shadow-black/20">
      <span className="text-slate-100 font-semibold">{stats.slides} slides</span>
      {stats.extra && (
        <>
          <span className="text-slate-600">·</span>
          <span>
            {stats.extra.label} in{' '}
            <span className="text-slate-100">{formatMs(stats.extra.ms)}ms</span>
          </span>
        </>
      )}
      <span className="text-slate-600">·</span>
      <span>
        {label} in <span className="text-emerald-400">{formatMs(stats.totalMs)}ms</span>
      </span>
      <span className="text-slate-600">·</span>
      <span className="text-slate-400">({formatPerSlide(perSlide)}ms/slide)</span>
    </div>
  );
}
