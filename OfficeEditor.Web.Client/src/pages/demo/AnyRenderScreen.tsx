import { useCallback, useRef, useState } from 'react';
import { extractApiError, renderAnyDeck } from '../../api.ts';
import { StatusMessage } from '../../components/StatusMessage.tsx';
import { FormatToggle } from './FormatToggle.tsx';
import { SlideGallery } from './SlideGallery.tsx';
import type { GallerySlide } from './SlideGallery.tsx';
import { TimingReceipt } from './TimingReceipt.tsx';
import type { TimingStats } from './TimingReceipt.tsx';
import type { PreviewFormat } from '../../types.ts';

interface RenderResult {
  fileName: string;
  slides: GallerySlide[];
  stats: TimingStats;
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
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

export function AnyRenderScreen() {
  const [file, setFile] = useState<File | null>(null);
  const [pickError, setPickError] = useState<string | null>(null);
  const [rendering, setRendering] = useState(false);
  const [result, setResult] = useState<RenderResult | null>(null);
  const [renderError, setRenderError] = useState<string | null>(null);
  // Arbitrary decks may contain photos, so raster previews are the sensible default.
  const [format, setFormat] = useState<PreviewFormat>('png');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFilePicked = useCallback((picked: File | null) => {
    setResult(null);
    setRenderError(null);
    if (!picked) {
      setFile(null);
      setPickError(null);
      return;
    }
    if (!picked.name.toLowerCase().endsWith('.pptx')) {
      setFile(null);
      setPickError(`"${picked.name}" is not a .pptx file. Pick a PowerPoint deck.`);
      return;
    }
    setPickError(null);
    setFile(picked);
  }, []);

  const handleRender = useCallback(async () => {
    if (!file) return;
    setRendering(true);
    setRenderError(null);
    setResult(null);
    try {
      const response = await renderAnyDeck(file, format);
      setResult({
        fileName: file.name,
        slides: response.previews.map((p) => ({
          slide: p.slide,
          dataUrl: `data:${p.contentType};base64,${p.contentBase64}`,
        })),
        stats: {
          slides: response.slideCount,
          totalMs: response.totalMilliseconds,
          perSlideMs:
            response.slideCount > 0 ? response.totalMilliseconds / response.slideCount : 0,
          label: 'rendered',
        },
      });
    } catch (err) {
      setRenderError(extractApiError(err));
    } finally {
      setRendering(false);
    }
  }, [file, format]);

  const handleReset = useCallback(() => {
    setFile(null);
    setPickError(null);
    setResult(null);
    setRenderError(null);
    // Clear the input so re-picking the same file still fires onChange.
    if (fileInputRef.current) fileInputRef.current.value = '';
  }, []);

  return (
    <div className="space-y-8">
      <section className="space-y-2">
        <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight text-slate-50">
          Render any deck
        </h1>
        <p className="text-slate-400">
          Drop in your own .pptx — the OfficeEditor Engine compiles every slide to a preview.
        </p>
      </section>

      <div className="space-y-4">
        <label
          className={[
            'flex cursor-pointer flex-col items-center justify-center gap-3 rounded-2xl border-2 border-dashed p-10 text-center transition-all duration-200',
            'border-slate-700 bg-slate-900/60 hover:border-indigo-500/60 hover:bg-slate-900',
            rendering ? 'pointer-events-none opacity-60' : '',
          ].join(' ')}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept=".pptx"
            className="sr-only"
            disabled={rendering}
            onChange={(e) => handleFilePicked(e.target.files?.[0] ?? null)}
          />
          <svg
            xmlns="http://www.w3.org/2000/svg"
            fill="none"
            viewBox="0 0 24 24"
            strokeWidth={1.5}
            stroke="currentColor"
            className="h-10 w-10 text-slate-500"
            aria-hidden="true"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5m-13.5-9L12 3m0 0 4.5 4.5M12 3v13.5"
            />
          </svg>
          {file ? (
            <div className="space-y-1">
              <p className="truncate text-lg font-bold text-slate-100">{file.name}</p>
              <p className="font-mono text-xs text-slate-500">{formatBytes(file.size)}</p>
            </div>
          ) : (
            <div className="space-y-1">
              <p className="text-lg font-bold text-slate-200">Choose a .pptx file</p>
              <p className="text-sm text-slate-500">Click to browse your files</p>
            </div>
          )}
        </label>

        {pickError && (
          <StatusMessage type="error" title="Invalid file" message={pickError} />
        )}

        <div className="flex flex-wrap items-center gap-3">
          <span className="text-sm font-semibold text-slate-300">Preview format</span>
          <FormatToggle value={format} onChange={setFormat} disabled={rendering} />
        </div>

        <div className="flex flex-wrap items-center gap-4">
          <button
            type="button"
            onClick={() => void handleRender()}
            disabled={!file || rendering}
            className={[
              'inline-flex items-center justify-center gap-2 rounded-xl px-6 py-3',
              'bg-gradient-to-r from-indigo-500 to-violet-600 font-bold text-white shadow-lg shadow-indigo-500/25',
              'hover:from-indigo-400 hover:to-violet-500 active:scale-[0.98]',
              'focus:outline-none focus:ring-2 focus:ring-indigo-400 focus:ring-offset-2 focus:ring-offset-slate-950',
              'disabled:cursor-not-allowed disabled:opacity-50',
              'transition-all duration-200',
            ].join(' ')}
          >
            Render deck
          </button>
          {(result || renderError) && !rendering && (
            <button
              type="button"
              onClick={handleReset}
              className="text-sm font-semibold text-slate-400 transition-colors hover:text-slate-200"
            >
              Render another deck
            </button>
          )}
        </div>
      </div>

      {rendering && file && (
        <div className="fade-in-up flex items-center gap-4 rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
          <Spinner />
          <div>
            <p className="font-semibold text-slate-100">Rendering {file.name}…</p>
            <p className="text-sm text-slate-400">OfficeEditor Engine compiling slides</p>
          </div>
        </div>
      )}

      {renderError && (
        <StatusMessage type="error" title="Render failed" message={renderError} />
      )}

      {result && (
        <div className="fade-in-up space-y-5">
          <TimingReceipt stats={result.stats} />
          <SlideGallery slides={result.slides} title={result.fileName} />
        </div>
      )}
    </div>
  );
}
