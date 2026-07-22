import { useEffect, useState } from 'react';

export interface GallerySlide {
  slide: number;
  dataUrl: string;
}

interface SlideGalleryProps {
  slides: GallerySlide[];
  title?: string;
}

type GalleryMode = 'wall' | 'focus';

export function SlideGallery({ slides, title }: SlideGalleryProps) {
  const [focusedIndex, setFocusedIndex] = useState(0);
  const [mode, setMode] = useState<GalleryMode>('wall');

  // New results start in the "wall" moment.
  useEffect(() => {
    setFocusedIndex(0);
    setMode('wall');
  }, [slides]);

  useEffect(() => {
    if (slides.length === 0) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA')) return;
      if (e.key === 'ArrowLeft') {
        e.preventDefault();
        setMode('focus');
        setFocusedIndex((i) => Math.max(0, i - 1));
      } else if (e.key === 'ArrowRight') {
        e.preventDefault();
        setMode('focus');
        setFocusedIndex((i) => Math.min(slides.length - 1, i + 1));
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [slides.length]);

  if (slides.length === 0) return null;

  const clampedIndex = Math.min(focusedIndex, slides.length - 1);
  const focused = slides[clampedIndex];

  const goTo = (index: number) => {
    setFocusedIndex(Math.max(0, Math.min(slides.length - 1, index)));
  };

  return (
    <section className="space-y-4">
      <div className="flex items-center justify-between gap-3">
        <h3 className="text-sm font-semibold uppercase tracking-wider text-slate-400">
          {title ?? 'Slides'}
        </h3>
        <div className="flex items-center p-0.5 rounded-full bg-slate-900 border border-slate-800">
          {(['wall', 'focus'] as const).map((m) => (
            <button
              key={m}
              type="button"
              onClick={() => setMode(m)}
              className={[
                'px-3 py-1 rounded-full text-xs font-semibold transition-all duration-200 capitalize',
                mode === m
                  ? 'bg-slate-700 text-white shadow-sm'
                  : 'text-slate-400 hover:text-slate-200',
              ].join(' ')}
            >
              {m === 'wall' ? 'Wall' : 'Focus'}
            </button>
          ))}
        </div>
      </div>

      {mode === 'wall' ? (
        <div className="fade-in-up grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3">
          {slides.map((slide, index) => (
            <button
              key={slide.slide}
              type="button"
              onClick={() => {
                goTo(index);
                setMode('focus');
              }}
              className="group relative aspect-video overflow-hidden rounded-xl ring-1 ring-slate-800 hover:ring-indigo-500/70 transition-all duration-200 bg-slate-900"
            >
              <img
                src={slide.dataUrl}
                alt={`Slide ${slide.slide}`}
                className="h-full w-full object-cover"
                loading="lazy"
              />
              <span className="absolute bottom-1.5 right-1.5 rounded-md bg-slate-950/80 px-1.5 py-0.5 font-mono text-[10px] text-slate-300 opacity-0 group-hover:opacity-100 transition-opacity">
                {slide.slide}
              </span>
            </button>
          ))}
        </div>
      ) : (
        <div className="fade-in-up space-y-3">
          <div className="relative mx-auto max-w-4xl">
            <div className="aspect-video overflow-hidden rounded-2xl bg-slate-900 shadow-2xl shadow-black/40 ring-1 ring-slate-700/60">
              <img
                src={focused.dataUrl}
                alt={`Slide ${focused.slide}`}
                className="h-full w-full object-contain"
              />
            </div>
            <button
              type="button"
              onClick={() => goTo(clampedIndex - 1)}
              disabled={clampedIndex === 0}
              aria-label="Previous slide"
              className="absolute left-3 top-1/2 -translate-y-1/2 flex h-10 w-10 items-center justify-center rounded-full bg-slate-950/70 text-slate-200 ring-1 ring-slate-700 hover:bg-slate-800 disabled:opacity-30 disabled:hover:bg-slate-950/70 transition-colors"
            >
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} className="h-5 w-5" aria-hidden="true">
                <path strokeLinecap="round" strokeLinejoin="round" d="M15.75 19.5 8.25 12l7.5-7.5" />
              </svg>
            </button>
            <button
              type="button"
              onClick={() => goTo(clampedIndex + 1)}
              disabled={clampedIndex === slides.length - 1}
              aria-label="Next slide"
              className="absolute right-3 top-1/2 -translate-y-1/2 flex h-10 w-10 items-center justify-center rounded-full bg-slate-950/70 text-slate-200 ring-1 ring-slate-700 hover:bg-slate-800 disabled:opacity-30 disabled:hover:bg-slate-950/70 transition-colors"
            >
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} className="h-5 w-5" aria-hidden="true">
                <path strokeLinecap="round" strokeLinejoin="round" d="m8.25 4.5 7.5 7.5-7.5 7.5" />
              </svg>
            </button>
            <span className="absolute bottom-3 right-3 rounded-full bg-slate-950/80 px-2.5 py-1 font-mono text-xs text-slate-300 ring-1 ring-slate-700">
              {clampedIndex + 1} / {slides.length}
            </span>
          </div>

          <div className="scrollbar-hide flex gap-2 overflow-x-auto py-1">
            {slides.map((slide, index) => (
              <button
                key={slide.slide}
                type="button"
                onClick={() => goTo(index)}
                className={[
                  'relative aspect-video w-28 shrink-0 overflow-hidden rounded-lg transition-all duration-200 bg-slate-900',
                  index === clampedIndex
                    ? 'ring-2 ring-indigo-500'
                    : 'ring-1 ring-slate-800 hover:ring-slate-600',
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
        </div>
      )}
    </section>
  );
}
