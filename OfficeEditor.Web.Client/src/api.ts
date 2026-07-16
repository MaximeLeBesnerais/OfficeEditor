import type { ConversionResult, Sample } from './types.ts';

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

export async function fetchHealth(): Promise<{ status: string }> {
  const response = await fetch(`${API_BASE}/api/health`);
  return handleResponse<{ status: string }>(response);
}

export async function fetchSamples(): Promise<Sample[]> {
  const response = await fetch(`${API_BASE}/api/samples`);
  return handleResponse<Sample[]>(response);
}

export async function convertFile(formData: FormData): Promise<ConversionResult> {
  const response = await fetch(`${API_BASE}/api/convert`, {
    method: 'POST',
    body: formData,
  });
  return handleResponse<ConversionResult>(response);
}

export function getDownloadUrl(result: ConversionResult): string {
  return resolveUrl(result.downloadUrl);
}

export function getPreviewUrl(result: ConversionResult): string | null {
  return result.previewUrl ? resolveUrl(result.previewUrl) : null;
}
