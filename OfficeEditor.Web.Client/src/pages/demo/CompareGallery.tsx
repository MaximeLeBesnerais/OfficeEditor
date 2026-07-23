import { useEffect, useState } from 'react';
import type { GallerySlide } from './SlideGallery.tsx';

export type LoGalleryState = 'pending' | 'empty' | 'ready';

export interface CompareGalleryProps {
  typstSlides: GallerySlide[];
  loSlides: GallerySlide[] | null;
  loState: LoGalleryState;
  loPlaceholder: string | null;
  // Timestamp (Date.now()) of the run start — drives the live chrono shown in
  // the LibreOffice panel while its leg is still in flight.
  loPendingSince: number | null;
  // Panel labels — defaults keep the Typst-vs-LibreOffice duel wording; the Render
  // tab's official-render comparison overrides the right one.
  leftLabel?: string;
  rightLabel?: string;
}

// Ticking elapsed timer for in-flight legs. Purely cosmetic: once the server
// response lands the caller swaps this out for the server-measured value.
export function ElapsedTicker({
  startedAt,
  className,
}: {
  startedAt: number;
  className?: string;
}) {
  const [elapsed, setElapsed] = useState(() => Math.max(0, Date.now() - startedAt));

  useEffect(() => {
    const id = window.setInterval(() => setElapsed(Math.max(0, Date.now() - startedAt)), 100);
    return () => window.clearInterval(id);
  }, [startedAt]);

  return <span className={className}>{(elapsed / 1000).toFixed(1)} s</span>;
}

function SmallSpinner() {
  return (
    <svg
      className="h-5 w-5 animate-spin text-slate-400"
      xmlns="http://www.w3.org/2000/svg"
      fill="none"
      viewBox="0 0 24 24"
      aria-hidden="true"
    >
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
      <path
        className="opacity-90"
        fill="currentColor"
        d="M4 12a8 8 0 0 1 8-8V0C5.373 0 0 5.373 0 12h4Z"
      />
    </svg>
  );
}

function ChevronButton({
  direction,
  disabled,
  onClick,
}: {
  direction: 'prev' | 'next';
  disabled: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      aria-label={direction === 'prev' ? 'Previous slide' : 'Next slide'}
      className="flex h-9 w-9 items-center justify-center rounded-full bg-slate-900 text-slate-200 ring-1 ring-slate-700 transition-colors hover:bg-slate-800 disabled:opacity-30 disabled:hover:bg-slate-900"
    >
      <svg
        xmlns="http://www.w3.org/2000/svg"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth={2}
        className="h-4 w-4"
        aria-hidden="true"
      >
        <path
          strokeLinecap="round"
          strokeLinejoin="round"
          d={direction === 'prev' ? 'M15.75 19.5 8.25 12l7.5-7.5' : 'm8.25 4.5 7.5 7.5-7.5 7.5'}
        />
      </svg>
    </button>
  );
}

// Synchronized side-by-side viewer: one shared slide index drives both the
// TypstBridge render (left) and the LibreOffice render (right).
export function CompareGallery({
  typstSlides,
  loSlides,
  loState,
  loPlaceholder,
  loPendingSince,
  leftLabel = 'OfficeEditor Engine',
  rightLabel = 'LibreOffice',
}: CompareGalleryProps) {
  // LO rasterizes the PDF it produced, so its page count can theoretically
  // differ from the typst slide count — sync the pair on the smaller of the
  // two and assume both sides are 1:1 by slide number up to that point.
  const loReady = loState === 'ready' && loSlides !== null;
  const pairCount = loReady ? Math.min(typstSlides.length, loSlides.length) : typstSlides.length;
  const [index, setIndex] = useState(0);

  // New results restart on the first slide.
  useEffect(() => {
    setIndex(0);
  }, [typstSlides, loSlides]);

  useEffect(() => {
    if (pairCount === 0) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA')) return;
      if (e.key === 'ArrowLeft') {
        e.preventDefault();
        setIndex((i) => Math.max(0, i - 1));
      } else if (e.key === 'ArrowRight') {
        e.preventDefault();
        setIndex((i) => Math.min(pairCount - 1, i + 1));
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [pairCount]);

  if (pairCount === 0) return null;

  const clamped = Math.min(index, pairCount - 1);
  const typst = typstSlides[clamped];
  const lo = loReady ? loSlides[clamped] : null;

  const goTo = (i: number) => setIndex(Math.max(0, Math.min(pairCount - 1, i)));

  return (
    <section className="space-y-4">
      <div className="flex items-center justify-between gap-3">
        <h3 className="text-sm font-semibold uppercase tracking-wider text-slate-400">
          Slide by slide
        </h3>
        <div className="flex items-center gap-3">
          <ChevronButton direction="prev" disabled={clamped === 0} onClick={() => goTo(clamped - 1)} />
          <span className="rounded-full bg-slate-900 px-2.5 py-1 font-mono text-xs text-slate-300 ring-1 ring-slate-700">
            {clamped + 1} / {pairCount}
          </span>
          <ChevronButton
            direction="next"
            disabled={clamped === pairCount - 1}
            onClick={() => goTo(clamped + 1)}
          />
        </div>
      </div>

      <div className="grid gap-4 md:grid-cols-2">
        <div className="space-y-2">
          <span className="inline-block rounded-full bg-emerald-500/10 px-2.5 py-1 text-xs font-semibold text-emerald-300 ring-1 ring-emerald-500/40">
            {leftLabel}
          </span>
          <div className="aspect-video overflow-hidden rounded-2xl bg-slate-900 shadow-2xl shadow-black/40 ring-1 ring-emerald-500/30">
            <img
              src={typst.dataUrl}
              alt={`Slide ${typst.slide} rendered by ${leftLabel}`}
              className="h-full w-full object-contain"
            />
          </div>
        </div>

        <div className="space-y-2">
          <span className="inline-block rounded-full bg-slate-800 px-2.5 py-1 text-xs font-semibold text-slate-300">
            {rightLabel}
          </span>
          <div className="aspect-video overflow-hidden rounded-2xl bg-slate-900 shadow-2xl shadow-black/40 ring-1 ring-slate-700/60">
            {lo ? (
              <img
                src={lo.dataUrl}
                alt={`Slide ${lo.slide} rendered by ${rightLabel}`}
                className="h-full w-full object-contain"
              />
            ) : loState === 'pending' && loPendingSince !== null ? (
              <div className="flex h-full w-full flex-col items-center justify-center gap-3 p-6">
                <SmallSpinner />
                <ElapsedTicker
                  startedAt={loPendingSince}
                  className="font-mono text-lg font-semibold text-slate-300"
                />
                <p className="text-center text-sm text-slate-500">{rightLabel} rendering…</p>
              </div>
            ) : (
              <div className="flex h-full w-full items-center justify-center p-6">
                <p className="text-center text-sm text-slate-500">
                  {loPlaceholder ?? 'Per-slide images unavailable'}
                </p>
              </div>
            )}
          </div>
        </div>
      </div>

      <div className="scrollbar-hide flex gap-2 overflow-x-auto py-1">
        {typstSlides.slice(0, pairCount).map((slide, i) => (
          <button
            key={slide.slide}
            type="button"
            onClick={() => goTo(i)}
            className={[
              'relative aspect-video w-28 shrink-0 overflow-hidden rounded-lg bg-slate-900 transition-all duration-200',
              i === clamped ? 'ring-2 ring-emerald-500' : 'ring-1 ring-slate-800 hover:ring-slate-600',
            ].join(' ')}
          >
            <img
              src={slide.dataUrl}
              alt={`Slide ${slide.slide} thumbnail`}
              className="h-full w-full object-cover"
              loading="lazy"
            />
          </button>
        ))}
      </div>
    </section>
  );
}
