import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  extractApiError,
  fetchDeckTemplate,
  generateDeck,
  resolveDownloadUrl,
} from '../../api.ts';
import { StatusMessage } from '../../components/StatusMessage.tsx';
import { FormatToggle } from './FormatToggle.tsx';
import { SlideGallery } from './SlideGallery.tsx';
import type { GallerySlide } from './SlideGallery.tsx';
import { TimingReceipt } from './TimingReceipt.tsx';
import { DEFAULT_THEME_ID, THEME_PRESETS } from './themes.ts';
import type { ThemePreset } from './themes.ts';
import type { GenerateDeckResponse, GenerationDocument, PreviewFormat } from '../../types.ts';

const TITLE_SENTINEL = '{{DECK_TITLE}}';

type PaletteRole = 'primary' | 'accent' | 'ink';

type PaletteOverrides = Record<PaletteRole, string>;

const PALETTE_ROLES: { role: PaletteRole; label: string }[] = [
  { role: 'primary', label: 'Primary' },
  { role: 'accent', label: 'Accent' },
  { role: 'ink', label: 'Ink' },
];

function presetColors(theme: ThemePreset): PaletteOverrides {
  return {
    primary: theme.design.palette.primary,
    accent: theme.design.palette.accent,
    ink: theme.design.palette.ink,
  };
}

function Spinner() {
  return (
    <svg
      className="h-5 w-5 animate-spin"
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

function getIssueMessages(response: GenerateDeckResponse): string[] {
  const issues = response.errors
    ?.map((issue) => issue.message)
    .filter((m): m is string => Boolean(m));
  if (issues && issues.length > 0) return issues;
  return [response.errorMessage ?? 'Generation failed.'];
}

function buildDocument(
  template: GenerationDocument,
  title: string,
  theme: ThemePreset,
  palette: PaletteOverrides,
): GenerationDocument {
  const patched = JSON.parse(
    JSON.stringify(template).replaceAll(TITLE_SENTINEL, title),
  ) as GenerationDocument;
  // Deep-copy the preset design so picker overrides never mutate THEME_PRESETS.
  const design = JSON.parse(JSON.stringify(theme.design)) as ThemePreset['design'];
  design.palette.primary = palette.primary;
  design.palette.accent = palette.accent;
  design.palette.ink = palette.ink;
  patched.design = design as unknown as Record<string, unknown>;
  return patched;
}

export function GenerateScreen() {
  const [template, setTemplate] = useState<GenerationDocument | null>(null);
  const [templateError, setTemplateError] = useState<string | null>(null);
  const [title, setTitle] = useState('Northwind Labs');
  const [themeId, setThemeId] = useState(DEFAULT_THEME_ID);
  const [palette, setPalette] = useState<PaletteOverrides>(() =>
    presetColors(THEME_PRESETS.find((t) => t.id === DEFAULT_THEME_ID) ?? THEME_PRESETS[0]),
  );
  // Vector deck = crisp, so SVG previews are the default here; ?format=png overrides.
  const [format, setFormat] = useState<PreviewFormat>(() =>
    new URLSearchParams(window.location.search).get('format') === 'png' ? 'png' : 'svg',
  );
  const [showJson, setShowJson] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [response, setResponse] = useState<GenerateDeckResponse | null>(null);
  const [generateError, setGenerateError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    fetchDeckTemplate()
      .then((doc) => {
        if (!cancelled) setTemplate(doc);
      })
      .catch((err: unknown) => {
        if (!cancelled) setTemplateError(extractApiError(err));
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const theme = THEME_PRESETS.find((t) => t.id === themeId) ?? THEME_PRESETS[0];

  // Switching preset resets the color pickers to that preset's palette.
  useEffect(() => {
    setPalette(presetColors(theme));
  }, [theme]);

  const paletteDirty =
    palette.primary !== theme.design.palette.primary ||
    palette.accent !== theme.design.palette.accent ||
    palette.ink !== theme.design.palette.ink;

  const document = useMemo<GenerationDocument | null>(() => {
    if (!template) return null;
    return buildDocument(template, title, theme, palette);
  }, [template, title, theme, palette]);

  const documentJson = useMemo(
    () => (document ? JSON.stringify(document, null, 2) : ''),
    [document],
  );

  const handleGenerate = useCallback(
    async (doc: GenerationDocument | null) => {
      if (!doc) return;
      setGenerating(true);
      setGenerateError(null);
      setResponse(null);
      try {
        setResponse(await generateDeck(doc, format));
      } catch (err) {
        setGenerateError(extractApiError(err));
      } finally {
        setGenerating(false);
      }
    },
    [format],
  );

  const paramsApplied = useRef(false);

  useEffect(() => {
    if (!template || paramsApplied.current) return;
    paramsApplied.current = true;
    const params = new URLSearchParams(window.location.search);
    const paramTitle = params.get('title');
    const preset = THEME_PRESETS.find((t) => t.id === params.get('theme'));

    const nextTitle = paramTitle ?? title;
    const nextTheme = preset ?? theme;

    if (paramTitle !== null) setTitle(paramTitle);
    if (preset) setThemeId(preset.id);

    if (params.get('autogen') === '1') {
      // Build the patched document directly from the param values so the
      // auto-generate never depends on state that hasn't flushed yet.
      void handleGenerate(buildDocument(template, nextTitle, nextTheme, presetColors(nextTheme)));
    }
  }, [template, title, theme, handleGenerate]);

  const slides: GallerySlide[] =
    response?.success && response.previews
      ? response.previews.map((p) => ({
          slide: p.slide,
          dataUrl: `data:${p.contentType};base64,${p.contentBase64}`,
        }))
      : [];

  const renderedMs = response
    ? response.totalMilliseconds - response.generationMilliseconds
    : 0;

  const warnings: string[] = response?.success
    ? [
        ...(response.previewError ? [response.previewError] : []),
        ...response.warnings,
        ...response.pipelineWarnings,
      ]
    : [];

  return (
    <div className="flex flex-col gap-8 lg:flex-row">
      {/* Left panel */}
      <div className="space-y-6 lg:w-2/5">
        <section className="space-y-2">
          <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight text-slate-50">
            Generate a deck from JSON
          </h1>
          <p className="text-slate-400">
            Edit the title, swap the theme, generate — watch the pipeline run.
          </p>
        </section>

        {templateError && (
          <StatusMessage
            type="error"
            title="Template unavailable"
            message={templateError}
          />
        )}

        <div className="space-y-2">
          <label htmlFor="deck-title" className="block text-sm font-semibold text-slate-300">
            Deck title
          </label>
          <input
            id="deck-title"
            type="text"
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            disabled={!template || generating}
            className={[
              'w-full rounded-xl border border-slate-700 bg-slate-900 px-4 py-2.5',
              'text-slate-100 placeholder-slate-500',
              'focus:border-indigo-500 focus:outline-none focus:ring-2 focus:ring-indigo-500/40',
              'disabled:opacity-50',
            ].join(' ')}
            placeholder="Northwind Labs"
          />
        </div>

        <div className="space-y-2">
          <p className="text-sm font-semibold text-slate-300">Theme</p>
          <div className="flex flex-wrap gap-2">
            {THEME_PRESETS.map((preset) => {
              const active = preset.id === theme.id;
              const palette = preset.design.palette;
              return (
                <button
                  key={preset.id}
                  type="button"
                  onClick={() => setThemeId(preset.id)}
                  disabled={!template || generating}
                  className={[
                    'flex items-center gap-2.5 rounded-xl border px-3.5 py-2 text-sm font-semibold transition-all duration-200',
                    active
                      ? 'border-indigo-500 bg-indigo-500/10 text-slate-100 ring-1 ring-indigo-500/50'
                      : 'border-slate-700 bg-slate-900 text-slate-400 hover:border-slate-500 hover:text-slate-200',
                    'disabled:opacity-50',
                  ].join(' ')}
                >
                  <span className="flex -space-x-1">
                    {[palette.primary, palette.accent, palette.ink].map((color) => (
                      <span
                        key={color}
                        className="h-3 w-3 rounded-full ring-1 ring-slate-950"
                        style={{ backgroundColor: color }}
                      />
                    ))}
                  </span>
                  {preset.name}
                </button>
              );
            })}
          </div>
        </div>

        <div className="space-y-2">
          <div className="flex items-center justify-between">
            <p className="text-sm font-semibold text-slate-300">Custom colors</p>
            {paletteDirty && (
              <button
                type="button"
                onClick={() => setPalette(presetColors(theme))}
                disabled={!template || generating}
                className="text-xs font-semibold text-slate-400 transition-colors hover:text-slate-200 disabled:opacity-50"
              >
                Reset to preset
              </button>
            )}
          </div>
          <div className="flex flex-wrap gap-3">
            {PALETTE_ROLES.map(({ role, label }) => (
              <label
                key={role}
                className={[
                  'flex items-center gap-2 rounded-xl border border-slate-700 bg-slate-900 px-3 py-2',
                  !template || generating ? 'opacity-50' : '',
                ].join(' ')}
              >
                <input
                  type="color"
                  value={palette[role]}
                  onChange={(e) => setPalette((p) => ({ ...p, [role]: e.target.value }))}
                  disabled={!template || generating}
                  aria-label={`${label} color`}
                  className="h-6 w-8 cursor-pointer rounded border-0 bg-transparent p-0 disabled:cursor-not-allowed"
                />
                <span className="text-sm text-slate-300">{label}</span>
                <span className="font-mono text-xs uppercase text-slate-500">
                  {palette[role]}
                </span>
              </label>
            ))}
          </div>
        </div>

        <div className="flex items-center gap-3">
          <span className="text-sm font-semibold text-slate-300">Preview format</span>
          <FormatToggle value={format} onChange={setFormat} disabled={!template || generating} />
        </div>

        <div className="space-y-2">
          <button
            type="button"
            onClick={() => setShowJson((v) => !v)}
            className="inline-flex items-center gap-1.5 text-sm font-semibold text-slate-400 hover:text-slate-200 transition-colors"
          >
            <svg
              xmlns="http://www.w3.org/2000/svg"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth={2}
              className={[
                'h-4 w-4 transition-transform duration-200',
                showJson ? 'rotate-90' : '',
              ].join(' ')}
              aria-hidden="true"
            >
              <path strokeLinecap="round" strokeLinejoin="round" d="m8.25 4.5 7.5 7.5-7.5 7.5" />
            </svg>
            {showJson ? 'Hide JSON' : 'Show JSON'}
          </button>
          {showJson && (
            <textarea
              readOnly
              value={documentJson}
              spellCheck={false}
              className={[
                'h-72 w-full resize-none overflow-auto rounded-xl border border-slate-800',
                'bg-slate-900/80 p-4 font-mono text-xs leading-relaxed text-slate-300',
                'focus:outline-none',
              ].join(' ')}
            />
          )}
        </div>

        <button
          type="button"
          onClick={() => void handleGenerate(document)}
          disabled={!template || generating}
          className={[
            'inline-flex w-full items-center justify-center gap-2 rounded-xl px-6 py-3',
            'bg-gradient-to-r from-indigo-500 to-violet-600 font-bold text-white shadow-lg shadow-indigo-500/25',
            'hover:from-indigo-400 hover:to-violet-500 active:scale-[0.98]',
            'focus:outline-none focus:ring-2 focus:ring-indigo-400 focus:ring-offset-2 focus:ring-offset-slate-950',
            'disabled:cursor-not-allowed disabled:opacity-50',
            'transition-all duration-200',
          ].join(' ')}
        >
          {generating && <Spinner />}
          {generating ? 'Generating…' : 'Generate'}
        </button>

        {generateError && (
          <StatusMessage type="error" title="Generation failed" message={generateError} />
        )}

        {response && !response.success && (
          <StatusMessage
            type="error"
            title="Generation failed"
            message={response.errorMessage ?? 'The document was rejected.'}
            messages={getIssueMessages(response)}
          />
        )}

        {response?.success && warnings.length > 0 && (
          <StatusMessage
            type="warning"
            title="Generated with warnings"
            message={
              response.previewError
                ? 'The deck was built, but slide previews failed.'
                : 'The pipeline reported warnings.'
            }
            messages={warnings}
          />
        )}
      </div>

      {/* Right stage */}
      <div className="lg:w-3/5">
        {response?.success ? (
          <div className="fade-in-up space-y-5">
            <TimingReceipt
              stats={{
                slides: response.slideCount,
                totalMs: renderedMs,
                label: 'rendered',
                extra: { label: 'built', ms: response.generationMilliseconds },
              }}
            />
            <SlideGallery slides={slides} title="Generated deck" />
            {response.downloadUrl && (
              <a
                href={resolveDownloadUrl(response.downloadUrl)}
                download="northwind-labs-demo.pptx"
                className={[
                  'inline-flex items-center justify-center gap-2 rounded-xl px-6 py-3',
                  'bg-slate-100 font-bold text-slate-900 shadow-lg',
                  'hover:bg-white active:scale-[0.98]',
                  'focus:outline-none focus:ring-2 focus:ring-slate-400 focus:ring-offset-2 focus:ring-offset-slate-950',
                  'transition-all duration-200',
                ].join(' ')}
              >
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  className="h-5 w-5"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                  strokeWidth={2}
                  aria-hidden="true"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5M7.5 12 12 16.5m0 0L16.5 12M12 16.5V3"
                  />
                </svg>
                Download .pptx
              </a>
            )}
          </div>
        ) : (
          <div className="flex h-full min-h-72 flex-col items-center justify-center gap-4 rounded-3xl border border-dashed border-slate-800 bg-slate-900/40 p-10 text-center">
            <div className="grid grid-cols-3 gap-2 opacity-60" aria-hidden="true">
              {['bg-indigo-500/40', 'bg-violet-500/40', 'bg-fuchsia-500/40', 'bg-slate-700', 'bg-indigo-500/40', 'bg-slate-700', 'bg-slate-700', 'bg-violet-500/40', 'bg-slate-700'].map(
                (color, index) => (
                  <span key={index} className={`aspect-video w-14 rounded-md ${color}`} />
                ),
              )}
            </div>
            <div>
              <p className="text-lg font-bold text-slate-200">Generate to see your deck</p>
              <p className="text-sm text-slate-500">
                Slides appear here the moment the pipeline finishes.
              </p>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
