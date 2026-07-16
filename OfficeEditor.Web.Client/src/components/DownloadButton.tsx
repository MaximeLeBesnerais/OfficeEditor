import type { ConversionResult } from '../types.ts';
import { getDownloadUrl } from '../api.ts';
import { getFormatMeta } from '../utils/format.ts';

interface DownloadButtonProps {
  result: ConversionResult;
}

export function DownloadButton({ result }: DownloadButtonProps) {
  const meta = getFormatMeta(result.outputFileName.split('.').pop() ?? '');

  return (
    <a
      href={getDownloadUrl(result)}
      download={result.outputFileName}
      className={[
        'inline-flex items-center justify-center gap-2',
        'px-6 py-3 rounded-xl text-white font-bold shadow-lg',
        'bg-slate-900 hover:bg-slate-800',
        'focus:outline-none focus:ring-2 focus:ring-slate-500 focus:ring-offset-2',
        'active:scale-[0.98]',
        'transition-all duration-200',
      ].join(' ')}
    >
      <span
        className={[
          'inline-flex items-center justify-center w-7 h-7 rounded-lg text-xs font-bold',
          meta.bgSolid,
        ].join(' ')}
      >
        {meta.label}
      </span>
      <svg
        xmlns="http://www.w3.org/2000/svg"
        className="w-5 h-5"
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
      Download {result.outputFileName}
    </a>
  );
}
