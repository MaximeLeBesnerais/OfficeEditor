import type {
  CompareCapabilities,
  DemoDeck,
  DemoDecksResponse,
  DemoRenderResponse,
  GenerateDeckResponse,
  GenerationDocument,
  LibreOfficeLegResult,
  OfficialSlidesResponse,
  PreviewFormat,
  TypstLegResult,
} from './types.ts';

// API calls use relative paths so Vite's dev proxy can forward them to
// the ASP.NET Core backend at http://localhost:5001. See vite.config.ts.
const API_BASE = '';

function resolveUrl(path: string): string {
  return path.startsWith('http') ? path : `${API_BASE}${path}`;
}

async function handleResponse<T>(response: Response): Promise<T> {
  if (!response.ok) {
    const text = await response.text().catch(() => 'Unknown error');
    throw new Error(text || `Request failed with status ${response.status}`);
  }
  return response.json() as Promise<T>;
}

export function resolveDownloadUrl(downloadUrl: string): string {
  return resolveUrl(downloadUrl);
}

// handleResponse throws the raw response text; demo endpoints return JSON
// error bodies, so try to extract a human-readable message from them.
export function extractApiError(err: unknown): string {
  const raw = err instanceof Error ? err.message : String(err);
  try {
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed === 'object' && parsed !== null) {
      const body = parsed as Record<string, unknown>;
      if (typeof body.error === 'string' && body.error) return body.error;
      if (typeof body.errorMessage === 'string' && body.errorMessage) return body.errorMessage;
      if (Array.isArray(body.errors)) {
        const messages = body.errors
          .map((issue: unknown) =>
            typeof issue === 'object' && issue !== null
              ? (issue as Record<string, unknown>).message
              : null,
          )
          .filter((m): m is string => typeof m === 'string' && m.length > 0);
        if (messages.length > 0) return messages.join('\n');
      }
    }
  } catch {
    // Not a JSON body — fall back to the raw text.
  }
  return raw || 'An unexpected error occurred.';
}

export async function fetchDemoDecks(): Promise<DemoDeck[]> {
  const response = await fetch(`${API_BASE}/api/demo/decks`);
  const data = await handleResponse<DemoDecksResponse>(response);
  return data.decks;
}

export async function renderDemoDeck(
  name: string,
  format: PreviewFormat,
): Promise<DemoRenderResponse> {
  const response = await fetch(`${API_BASE}/api/demo/render`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, format }),
  });
  return handleResponse<DemoRenderResponse>(response);
}

// Uploads an arbitrary user-provided .pptx; same response shape as renderDemoDeck.
export async function renderAnyDeck(
  file: File,
  format: PreviewFormat,
): Promise<DemoRenderResponse> {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('format', format);
  const response = await fetch(`${API_BASE}/api/demo/render-upload`, {
    method: 'POST',
    body: formData,
  });
  return handleResponse<DemoRenderResponse>(response);
}

// Official render (PowerPoint ground truth) of a whitelisted REF deck. Never throws:
// a missing official render or unavailable pdftoppm degrades to null, and callers fall
// back to the plain engine gallery with a muted note — never an error panel.
export async function fetchOfficialSlides(name: string): Promise<OfficialSlidesResponse | null> {
  try {
    const response = await fetch(
      `${API_BASE}/api/demo/decks/${encodeURIComponent(name)}/official-slides`,
    );
    if (!response.ok) return null;
    return (await response.json()) as OfficialSlidesResponse;
  } catch {
    return null;
  }
}

export async function fetchCompareCapabilities(): Promise<CompareCapabilities> {  const response = await fetch(`${API_BASE}/api/demo/compare/capabilities`);
  return handleResponse<CompareCapabilities>(response);
}

export async function runTypstCompare(name: string): Promise<TypstLegResult> {
  const response = await fetch(`${API_BASE}/api/demo/compare/typst`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name }),
  });
  return handleResponse<TypstLegResult>(response);
}

// Uploads an arbitrary user-provided .pptx for the TypstBridge leg of the duel.
export async function runTypstCompareUpload(file: File): Promise<TypstLegResult> {
  const formData = new FormData();
  formData.append('file', file);
  const response = await fetch(`${API_BASE}/api/demo/compare/typst-upload`, {
    method: 'POST',
    body: formData,
  });
  return handleResponse<TypstLegResult>(response);
}

export async function runLibreOfficeCompare(name: string): Promise<LibreOfficeLegResult> {
  const response = await fetch(`${API_BASE}/api/demo/compare/libreoffice`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name }),
  });
  return handleResponse<LibreOfficeLegResult>(response);
}

// Uploads an arbitrary user-provided .pptx for the LibreOffice leg of the duel.
export async function runLibreOfficeCompareUpload(file: File): Promise<LibreOfficeLegResult> {
  const formData = new FormData();
  formData.append('file', file);
  const response = await fetch(`${API_BASE}/api/demo/compare/libreoffice-upload`, {
    method: 'POST',
    body: formData,
  });
  return handleResponse<LibreOfficeLegResult>(response);
}

export async function fetchDeckTemplate(): Promise<GenerationDocument> {
  const response = await fetch(`${API_BASE}/api/demo/deck-template`);
  return handleResponse<GenerationDocument>(response);
}

export async function generateDeck(
  document: GenerationDocument,
  previewFormat: PreviewFormat,
): Promise<GenerateDeckResponse> {
  const response = await fetch(`${API_BASE}/api/decks/generate`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ document, previewFormat, ppi: 150 }),
  });
  if (!response.ok) {
    // 400 responses carry the full result shape (success:false + errors).
    const text = await response.text().catch(() => '');
    if (text) {
      try {
        return JSON.parse(text) as GenerateDeckResponse;
      } catch {
        // Not JSON — throw the raw text below.
      }
    }
    throw new Error(text || `Request failed with status ${response.status}`);
  }
  return response.json() as Promise<GenerateDeckResponse>;
}
