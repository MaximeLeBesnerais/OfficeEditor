import type { ConversionResult } from '../types.ts';
import { getPreviewUrl } from '../api.ts';

interface PreviewPanelProps {
  result: ConversionResult;
}

export function PreviewPanel({ result }: PreviewPanelProps) {
  const previewUrl = getPreviewUrl(result);
  const contentType = result.contentType.toLowerCase();

  if (!previewUrl) {
    return (
      <div className="rounded-2xl border border-slate-200 bg-slate-50 p-10 text-center">
        <div className="mx-auto w-14 h-14 rounded-full bg-slate-100 text-slate-500 flex items-center justify-center mb-3">
          <svg
            xmlns="http://www.w3.org/2000/svg"
            className="w-7 h-7"
            fill="none"
            viewBox="0 0 24 24"
            stroke="currentColor"
            strokeWidth={1.5}
            aria-hidden="true"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M2.036 12.322a1.012 1.012 0 0 1 0-.639C3.423 7.51 7.36 4.5 12 4.5c4.638 0 8.573 3.007 9.963 7.178.07.207.07.431 0 .639C20.577 16.49 16.64 19.5 12 19.5c-4.638 0-8.573-3.007-9.963-7.178Z"
            />
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0Z"
            />
          </svg>
        </div>
        <p className="text-slate-600">
          Preview is not available for{' '}
          <span className="font-semibold text-slate-900">{result.outputFileName}</span>.
          Use the download button to open it locally.
        </p>
      </div>
    );
  }

  if (contentType.includes('pdf')) {
    return (
      <div className="rounded-2xl border border-slate-200 bg-white shadow-sm overflow-hidden">
        <iframe
          src={previewUrl}
          title={`Preview of ${result.outputFileName}`}
          className="w-full h-[32rem] bg-white"
        />
      </div>
    );
  }

  if (contentType.includes('image') || contentType.includes('svg')) {
    return (
      <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
        <img
          src={previewUrl}
          alt={`Preview of ${result.outputFileName}`}
          className="max-h-[32rem] w-full object-contain rounded-xl"
        />
      </div>
    );
  }

  return (
    <div className="rounded-2xl border border-slate-200 bg-slate-50 p-10 text-center">
      <p className="text-slate-600">
        Browser preview is not supported for{' '}
        <span className="font-semibold text-slate-900">{result.outputFileName}</span>.
        Download the file to view it.
      </p>
    </div>
  );
}
