import { useCallback, useEffect, useRef, useState } from 'react';
import { extractApiError, fetchDemoDecks, renderDemoDeck } from '../../api.ts';
import { StatusMessage } from '../../components/StatusMessage.tsx';
import { FormatToggle } from './FormatToggle.tsx';
import { SlideGallery } from './SlideGallery.tsx';
import type { GallerySlide } from './SlideGallery.tsx';
import { TimingReceipt } from './TimingReceipt.tsx';
import type { TimingStats } from './TimingReceipt.tsx';
import type { DemoDeck, PreviewFormat } from '../../types.ts';

interface RenderResult {
  deckName: string;
  slides: GallerySlide[];
  stats: TimingStats;
}

function Spinner() {
  return (
    <svg
      className="h-6 w-6 animate-spin text-indigo-400"
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

export function RenderScreen() {
  const [decks, setDecks] = useState<DemoDeck[] | null>(null);
  const [decksError, setDecksError] = useState<string | null>(null);
  const [renderingDeck, setRenderingDeck] = useState<DemoDeck | null>(null);
  const [result, setResult] = useState<RenderResult | null>(null);
  const [renderError, setRenderError] = useState<string | null>(null);
  // REF decks contain photos, so raster previews are the sensible default.
  const [format, setFormat] = useState<PreviewFormat>('png');

  useEffect(() => {
    let cancelled = false;
    fetchDemoDecks()
      .then((data) => {
        if (!cancelled) setDecks(data);
      })
      .catch((err: unknown) => {
        if (!cancelled) setDecksError(extractApiError(err));
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const handleRender = useCallback(async (deck: DemoDeck, outputFormat: PreviewFormat) => {
    setRenderingDeck(deck);
    setRenderError(null);
    setResult(null);
    try {
      const response = await renderDemoDeck(deck.name, outputFormat);
      setResult({
        deckName: deck.name,
        slides: response.previews.map((p) => ({
          slide: p.slide,
          dataUrl: `data:${p.contentType};base64,${p.contentBase64}`,
        })),
        stats: {
          slides: response.slideCount,
          totalMs: response.totalMilliseconds,
          label: 'rendered',
        },
      });
    } catch (err) {
      setRenderError(extractApiError(err));
    } finally {
      setRenderingDeck(null);
    }
  }, []);

  const autoRenderFired = useRef(false);

  useEffect(() => {
    if (!decks || autoRenderFired.current) return;
    const name = new URLSearchParams(window.location.search).get('deck');
    if (!name) return;
    const match = decks.find((d) => d.name === name);
    if (!match) return;
    autoRenderFired.current = true;
    void handleRender(match, format);
  }, [decks, format, handleRender]);

  return (
    <div className="space-y-8">
      <section className="space-y-2">
        <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight text-slate-50">
          Render an existing deck
        </h1>
        <p className="text-slate-400">
          Pick a deck — the OfficeEditor Engine compiles every slide to a PNG preview.
        </p>
      </section>

      {decksError && (
        <StatusMessage type="error" title="Could not load demo decks" message={decksError} />
      )}

      {!decks && !decksError && (
        <div className="flex items-center gap-3 text-slate-400">
          <Spinner />
          <span>Loading decks…</span>
        </div>
      )}

      {decks && decks.length === 0 && !decksError && (
        <StatusMessage type="info" title="No decks" message="The demo deck list is empty." />
      )}

      {decks && decks.length > 0 && (
        <div className="space-y-3">
          <div className="flex items-center gap-3">
            <span className="text-sm font-semibold text-slate-300">Preview format</span>
            <FormatToggle value={format} onChange={setFormat} disabled={renderingDeck !== null} />
          </div>
          <div className="grid gap-4 sm:grid-cols-2">
            {decks.map((deck) => {
              const isRendering = renderingDeck?.name === deck.name;
              return (
                <button
                  key={deck.name}
                  type="button"
                  onClick={() => void handleRender(deck, format)}
                  disabled={renderingDeck !== null}
                  className={[
                    'group rounded-2xl border p-5 text-left transition-all duration-200',
                    'border-slate-800 bg-slate-900/60 hover:border-indigo-500/60 hover:bg-slate-900',
                    'disabled:cursor-not-allowed disabled:opacity-60',
                    'focus:outline-none focus:ring-2 focus:ring-indigo-500/60',
                  ].join(' ')}
                >
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0 space-y-1">
                      <p className="truncate text-lg font-bold text-slate-100 group-hover:text-white">
                        {deck.name}
                      </p>
                      <p className="text-sm text-slate-400 leading-relaxed">{deck.description}</p>
                    </div>
                    {isRendering ? (
                      <Spinner />
                    ) : (
                      <span className="shrink-0 rounded-full bg-slate-800 px-2.5 py-1 font-mono text-xs text-slate-300">
                        {deck.slideCount > 0 ? `${deck.slideCount} slides` : 'deck'}
                      </span>
                    )}
                  </div>
                  <p className="mt-3 font-mono text-xs text-slate-500">{deck.fileName}</p>
                </button>
              );
            })}
          </div>
        </div>
      )}

      {renderingDeck && (
        <div className="fade-in-up flex items-center gap-4 rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
          <Spinner />
          <div>
            <p className="font-semibold text-slate-100">
              Rendering {renderingDeck.name}…
            </p>
            <p className="text-sm text-slate-400">
              OfficeEditor Engine compiling{' '}
              {renderingDeck.slideCount > 0 ? `${renderingDeck.slideCount} slides` : 'slides'}
            </p>
          </div>
        </div>
      )}

      {renderError && (
        <StatusMessage type="error" title="Render failed" message={renderError} />
      )}

      {result && (
        <div className="fade-in-up space-y-5">
          <TimingReceipt stats={result.stats} />
          <SlideGallery slides={result.slides} title={result.deckName} />
        </div>
      )}
    </div>
  );
}
