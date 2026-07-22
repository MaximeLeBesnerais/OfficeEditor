import { useCallback, useEffect, useRef, useState } from 'react';
import {
  extractApiError,
  fetchCompareCapabilities,
  fetchDemoDecks,
  resolveDownloadUrl,
  runLibreOfficeCompare,
  runLibreOfficeCompareUpload,
  runTypstCompare,
  runTypstCompareUpload,
} from '../../api.ts';
import { StatusMessage } from '../../components/StatusMessage.tsx';
import { CompareGallery, ElapsedTicker } from './CompareGallery.tsx';
import type { LoGalleryState } from './CompareGallery.tsx';
import type { GallerySlide } from './SlideGallery.tsx';
import type {
  CompareCapabilities,
  DemoDeck,
  LibreOfficeLegResult,
  SlidePreviewDto,
  TypstLegResult,
} from '../../types.ts';

type LegStatus = 'idle' | 'pending' | 'done' | 'error';

interface TypstLegState {
  status: LegStatus;
  result: TypstLegResult | null;
  error: string | null;
}

interface LoLegState {
  status: LegStatus;
  result: LibreOfficeLegResult | null;
  error: string | null;
}

const IDLE_TYPST: TypstLegState = { status: 'idle', result: null, error: null };
const IDLE_LO: LoLegState = { status: 'idle', result: null, error: null };

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

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

// Long durations read better in seconds; short ones stay in milliseconds.
function formatDuration(ms: number): string {
  return ms >= 4000 ? `${(ms / 1000).toFixed(1)} s` : `${Math.round(ms)} ms`;
}

// Big final timing for a resolved leg. Server-measured only — the live chrono
// shown while pending is swapped out for this the moment the response lands.
function DurationFigure({
  ms,
  valueClass,
  detailClass,
}: {
  ms: number;
  valueClass: string;
  detailClass: string;
}) {
  return (
    <div>
      <p className={`font-mono text-4xl font-bold ${valueClass}`}>{formatDuration(ms)}</p>
      {ms >= 4000 && (
        <p className={`mt-1 font-mono text-xs ${detailClass}`}>
          ({Math.round(ms).toLocaleString('en-US')} ms)
        </p>
      )}
    </div>
  );
}

function toGallerySlides(previews: SlidePreviewDto[]): GallerySlide[] {
  return previews.map((p) => ({
    slide: p.slide,
    dataUrl: `data:${p.contentType};base64,${p.contentBase64}`,
  }));
}

// External-link button to a leg's PDF output. Tint varies per card; the rest
// of the styling is shared so both links read as the same affordance.
function PdfLink({ url, label, tintClass }: { url: string; label: string; tintClass: string }) {
  return (
    <a
      href={resolveDownloadUrl(url)}
      target="_blank"
      rel="noreferrer"
      className={`inline-flex items-center gap-2 rounded-full border px-5 py-2 text-sm font-semibold transition-colors ${tintClass}`}
    >
      {label}
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
          d="M13.5 6H5.25A2.25 2.25 0 0 0 3 8.25v10.5A2.25 2.25 0 0 0 5.25 21h10.5A2.25 2.25 0 0 0 18 18.75V10.5m-10.5 6L21 3m0 0h-5.25M21 3v5.25"
        />
      </svg>
    </a>
  );
}

// Short footnote shown once results are visible — explains why the two
// methods produce such different timings without competing with the cards.
function MethodologyNote() {
  return (
    <div className="mx-auto max-w-3xl space-y-1.5 text-center">
      <p className="text-[11px] font-semibold uppercase tracking-wide text-slate-400">
        Why the methods differ
      </p>
      <p className="text-xs leading-relaxed text-slate-500">
        Both sides receive the same .pptx and must produce the same two artifacts: per-slide PNGs
        and a PDF. The OfficeEditor Engine translates the deck once — OOXML → model → layout — then
        emits PNG and PDF natively, in parallel: neither output depends on the other. LibreOffice
        boots a full office suite to convert the deck to PDF, then rasterizes that PDF into PNGs as
        a second, dependent step. All timings are server-measured wall-clock; the engine's total is
        the parallel pair, not a sum.
      </p>
    </div>
  );
}

function PendingLeg({ startedAt, label }: { startedAt: number; label: string }) {
  return (
    <div className="flex items-center gap-4">
      <Spinner />
      <div className="space-y-0.5">
        <p className="font-semibold text-slate-100">{label}</p>
        <ElapsedTicker startedAt={startedAt} className="font-mono text-sm text-slate-400" />
      </div>
    </div>
  );
}

function LegErrorPanel({ message }: { message: string }) {
  return (
    <div className="rounded-xl border border-red-500/40 bg-red-500/10 p-4 text-sm leading-relaxed text-red-200">
      {message}
    </div>
  );
}

function TypstCard({ leg, startedAt }: { leg: TypstLegState; startedAt: number | null }) {
  const result = leg.result;
  return (
    <div className="space-y-3 rounded-2xl border border-emerald-500/40 bg-emerald-500/5 p-5">
      <h3 className="text-sm font-semibold uppercase tracking-wider text-emerald-300">
        OfficeEditor Engine
      </h3>
      {leg.status === 'pending' && startedAt !== null && (
        <PendingLeg startedAt={startedAt} label="OfficeEditor Engine rendering…" />
      )}
      {leg.status === 'error' && (
        <LegErrorPanel message={leg.error ?? 'OfficeEditor Engine render failed.'} />
      )}
      {leg.status === 'done' && result && (
        <>
          <dl className="space-y-1 font-mono text-sm text-slate-400">
            <div className="flex justify-between">
              <dt>Slide PNGs</dt>
              <dd>{formatDuration(result.pngMilliseconds)}</dd>
            </div>
            <div className="flex justify-between">
              <dt>PDF</dt>
              <dd>
                {result.pdfMilliseconds !== null ? formatDuration(result.pdfMilliseconds) : '—'}
              </dd>
            </div>
          </dl>
          {result.pdfError && (
            <p className="rounded-lg border border-amber-500/40 bg-amber-500/10 px-3 py-2 text-xs leading-relaxed text-amber-200">
              PDF step failed: {result.pdfError}
            </p>
          )}
          <DurationFigure
            ms={result.totalMilliseconds}
            valueClass="text-emerald-300"
            detailClass="text-emerald-500/70"
          />
          <p className="text-xs text-slate-500">PNG + PDF rendered in parallel</p>
          <p className="font-mono text-sm text-slate-400">
            {result.slideCount > 0
              ? `${(result.totalMilliseconds / result.slideCount).toFixed(1)} ms/slide`
              : '—'}
            {' · '}
            {result.slideCount} slides
          </p>
          {result.pdfDownloadUrl && (
            <div>
              <PdfLink
                url={result.pdfDownloadUrl}
                label="Open the OfficeEditor PDF"
                tintClass="border-emerald-500/50 bg-emerald-500/10 text-emerald-200 hover:border-emerald-400 hover:bg-emerald-500/20"
              />
            </div>
          )}
        </>
      )}
    </div>
  );
}

function LibreOfficeCard({ leg, startedAt }: { leg: LoLegState; startedAt: number | null }) {
  const lo = leg.result;
  const loFailed = lo !== null && (!lo.success || !lo.available || lo.error !== null);
  return (
    <div className="space-y-3 rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
      <div className="flex items-center justify-between gap-3">
        <h3 className="text-sm font-semibold uppercase tracking-wider text-slate-400">
          LibreOffice
        </h3>
        {lo?.version && (
          <span className="truncate font-mono text-xs text-slate-500">{lo.version}</span>
        )}
      </div>
      {leg.status === 'pending' && startedAt !== null && (
        <div className="space-y-2">
          <PendingLeg startedAt={startedAt} label="LibreOffice is thinking…" />
          <p className="text-sm text-slate-500">
            This usually takes 5–20 seconds — that's the point.
          </p>
        </div>
      )}
      {leg.status === 'error' && (
        <LegErrorPanel message={leg.error ?? 'LibreOffice render failed.'} />
      )}
      {leg.status === 'done' && lo && loFailed && (
        <div className="rounded-xl border border-amber-500/40 bg-amber-500/10 p-4 text-sm leading-relaxed text-amber-200">
          {lo.error ?? 'LibreOffice is not available on this machine.'}
        </div>
      )}
      {leg.status === 'done' && lo && !loFailed && (
        <>
          <dl className="space-y-1 font-mono text-sm text-slate-400">
            <div className="flex justify-between">
              <dt>PDF conversion</dt>
              <dd>
                {lo.conversionMilliseconds !== null ? formatDuration(lo.conversionMilliseconds) : '—'}
              </dd>
            </div>
            <div className="flex justify-between">
              <dt>Slide PNGs</dt>
              <dd>
                {lo.rasterizationMilliseconds !== null
                  ? formatDuration(lo.rasterizationMilliseconds)
                  : '—'}
              </dd>
            </div>
          </dl>
          {lo.totalMilliseconds !== null ? (
            <DurationFigure
              ms={lo.totalMilliseconds}
              valueClass="text-slate-200"
              detailClass="text-slate-500"
            />
          ) : (
            <p className="font-mono text-4xl font-bold text-slate-200">—</p>
          )}
          {lo.totalMilliseconds !== null && lo.slideCount > 0 && (
            <p className="font-mono text-sm text-slate-400">
              {`${(lo.totalMilliseconds / lo.slideCount).toFixed(1)} ms/slide`}
              {' · '}
              {lo.slideCount} slides
            </p>
          )}
          {lo.pdfDownloadUrl && (
            <div>
              <PdfLink
                url={lo.pdfDownloadUrl}
                label="Open the LibreOffice PDF"
                tintClass="border-slate-700 bg-slate-900 text-slate-200 hover:border-slate-500 hover:bg-slate-800"
              />
            </div>
          )}
        </>
      )}
    </div>
  );
}

export function CompareScreen() {
  const [decks, setDecks] = useState<DemoDeck[] | null>(null);
  const [decksError, setDecksError] = useState<string | null>(null);
  const [capabilities, setCapabilities] = useState<CompareCapabilities | null>(null);
  const [selectedDeck, setSelectedDeck] = useState<DemoDeck | null>(null);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [pickError, setPickError] = useState<string | null>(null);
  const [typstLeg, setTypstLeg] = useState<TypstLegState>(IDLE_TYPST);
  const [loLeg, setLoLeg] = useState<LoLegState>(IDLE_LO);
  const [runStartedAt, setRunStartedAt] = useState<number | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

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

  useEffect(() => {
    let cancelled = false;
    fetchCompareCapabilities()
      .then((data) => {
        if (!cancelled) setCapabilities(data);
      })
      .catch(() => {
        // Capabilities are an advisory hint only — the LO leg endpoint is
        // the source of truth, so a failed probe just hides the notice.
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const clearRun = useCallback(() => {
    setTypstLeg(IDLE_TYPST);
    setLoLeg(IDLE_LO);
    setRunStartedAt(null);
  }, []);

  const handleSelectDeck = useCallback(
    (deck: DemoDeck) => {
      setSelectedDeck(deck);
      setSelectedFile(null);
      setPickError(null);
      // Clear the input so re-picking the same file still fires onChange.
      if (fileInputRef.current) fileInputRef.current.value = '';
      clearRun();
    },
    [clearRun],
  );

  const handleFilePicked = useCallback(
    (picked: File | null) => {
      if (!picked) return;
      if (!picked.name.toLowerCase().endsWith('.pptx')) {
        setSelectedFile(null);
        setPickError(`"${picked.name}" is not a .pptx file. Pick a PowerPoint deck.`);
        return;
      }
      setPickError(null);
      setSelectedFile(picked);
      setSelectedDeck(null);
      clearRun();
    },
    [clearRun],
  );

  // Both legs fire in parallel with independent .then chains so each card
  // fills in the instant its own call resolves — typst (~0.5s) never waits
  // on LibreOffice (~9s), and one leg's failure doesn't sink the other.
  const handleRun = useCallback(() => {
    const deck = selectedDeck;
    const file = selectedFile;
    if (!deck && !file) return;

    const startedAt = Date.now();
    setRunStartedAt(startedAt);
    setTypstLeg({ status: 'pending', result: null, error: null });
    setLoLeg({ status: 'pending', result: null, error: null });

    const typstPromise = file ? runTypstCompareUpload(file) : runTypstCompare(deck!.name);
    typstPromise
      .then((result) => setTypstLeg({ status: 'done', result, error: null }))
      .catch((err: unknown) =>
        setTypstLeg({ status: 'error', result: null, error: extractApiError(err) }),
      );

    const loPromise = file ? runLibreOfficeCompareUpload(file) : runLibreOfficeCompare(deck!.name);
    loPromise
      .then((result) => setLoLeg({ status: 'done', result, error: null }))
      .catch((err: unknown) =>
        setLoLeg({ status: 'error', result: null, error: extractApiError(err) }),
      );
  }, [selectedDeck, selectedFile]);

  const running = typstLeg.status === 'pending' || loLeg.status === 'pending';
  const runActive = typstLeg.status !== 'idle';

  const typstResult = typstLeg.result;
  const lo = loLeg.result;

  // Both totals now include PNG + PDF work, so the ratio is apples-to-apples.
  const speedup =
    typstResult !== null &&
    lo !== null &&
    lo.totalMilliseconds !== null &&
    typstResult.totalMilliseconds > 0
      ? lo.totalMilliseconds / typstResult.totalMilliseconds
      : null;

  const loGalleryState: LoGalleryState =
    loLeg.status === 'pending'
      ? 'pending'
      : loLeg.status === 'done' && lo?.previews != null && lo.previews.length > 0
        ? 'ready'
        : 'empty';

  const loPlaceholder =
    loLeg.status === 'error'
      ? (loLeg.error ?? 'LibreOffice render failed')
      : lo !== null
        ? (lo.error ??
          (!lo.pdftoppmAvailable
            ? 'Per-slide images unavailable — pdftoppm not found'
            : 'Per-slide images unavailable'))
        : null;

  return (
    <div className="space-y-8">
      <section className="space-y-2">
        <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight text-slate-50">
          The duel: OfficeEditor Engine vs LibreOffice
        </h1>
        <p className="text-slate-400">
          Same deck, two renderers — compare their speed and output side by side.
        </p>
      </section>

      {capabilities && !capabilities.available && (
        <StatusMessage
          type="info"
          message="LibreOffice not detected on this machine — the comparison will show the OfficeEditor Engine side only"
        />
      )}

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
        <div className="grid gap-4 sm:grid-cols-2">
          {decks.map((deck) => {
            const isSelected = selectedFile === null && selectedDeck?.name === deck.name;
            return (
              <button
                key={deck.name}
                type="button"
                onClick={() => handleSelectDeck(deck)}
                disabled={running}
                className={[
                  'group rounded-2xl border p-5 text-left transition-all duration-200',
                  isSelected
                    ? 'border-indigo-500 bg-indigo-500/10 ring-1 ring-indigo-500/50'
                    : 'border-slate-800 bg-slate-900/60 hover:border-indigo-500/60 hover:bg-slate-900',
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
                  {isSelected ? (
                    <span className="shrink-0 rounded-full bg-indigo-500/20 px-2.5 py-1 font-mono text-xs font-semibold text-indigo-300 ring-1 ring-indigo-500/50">
                      Selected
                    </span>
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
      )}

      <div className="space-y-3">
        <label
          className={[
            'flex cursor-pointer items-center justify-center gap-4 rounded-2xl border-2 border-dashed p-5 text-center transition-all duration-200',
            selectedFile
              ? 'border-indigo-500 bg-indigo-500/10 ring-1 ring-indigo-500/50'
              : 'border-slate-700 bg-slate-900/60 hover:border-indigo-500/60 hover:bg-slate-900',
            running ? 'pointer-events-none opacity-60' : '',
          ].join(' ')}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept=".pptx"
            className="sr-only"
            disabled={running}
            onChange={(e) => handleFilePicked(e.target.files?.[0] ?? null)}
          />
          <svg
            xmlns="http://www.w3.org/2000/svg"
            fill="none"
            viewBox="0 0 24 24"
            strokeWidth={1.5}
            stroke="currentColor"
            className="h-8 w-8 shrink-0 text-slate-500"
            aria-hidden="true"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5m-13.5-9L12 3m0 0 4.5 4.5M12 3v13.5"
            />
          </svg>
          {selectedFile ? (
            <div className="min-w-0 space-y-0.5 text-left">
              <p className="truncate text-lg font-bold text-slate-100">{selectedFile.name}</p>
              <p className="font-mono text-xs text-slate-500">{formatBytes(selectedFile.size)}</p>
            </div>
          ) : (
            <div className="space-y-0.5 text-left">
              <p className="text-lg font-bold text-slate-200">Or compare your own deck</p>
              <p className="text-sm text-slate-500">Click to browse for a .pptx file</p>
            </div>
          )}
        </label>

        {pickError && <StatusMessage type="error" title="Invalid file" message={pickError} />}
      </div>

      <div>
        <button
          type="button"
          onClick={handleRun}
          disabled={(!selectedDeck && !selectedFile) || running}
          className="rounded-full bg-indigo-500 px-6 py-2.5 text-sm font-semibold text-white shadow-lg shadow-indigo-500/25 transition-all duration-200 hover:bg-indigo-400 disabled:cursor-not-allowed disabled:opacity-50 disabled:hover:bg-indigo-500"
        >
          {running ? 'Comparing…' : 'Run comparison'}
        </button>
      </div>

      {runActive && (
        <div className="fade-in-up space-y-6">
          <div className="grid gap-4 md:grid-cols-2">
            <TypstCard leg={typstLeg} startedAt={runStartedAt} />
            <LibreOfficeCard leg={loLeg} startedAt={runStartedAt} />
          </div>

          {speedup !== null && (
            <div className="flex justify-center">
              <span className="rounded-full border border-emerald-500/50 bg-emerald-500/10 px-5 py-2 font-mono text-sm font-semibold text-emerald-300 shadow-lg shadow-emerald-500/10">
                OfficeEditor Engine is {speedup.toFixed(1)}× faster
              </span>
            </div>
          )}

          <MethodologyNote />

          {typstResult !== null && typstResult.previews.length > 0 && (
            <CompareGallery
              typstSlides={toGallerySlides(typstResult.previews)}
              loSlides={loGalleryState === 'ready' && lo?.previews ? toGallerySlides(lo.previews) : null}
              loState={loGalleryState}
              loPlaceholder={loPlaceholder}
              loPendingSince={runStartedAt}
            />
          )}

          {!running && (
            <div>
              <button
                type="button"
                onClick={clearRun}
                className="text-sm font-semibold text-indigo-400 underline underline-offset-4 hover:text-indigo-300"
              >
                Run another comparison
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
