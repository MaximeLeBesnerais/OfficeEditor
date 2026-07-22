import type { PreviewFormat } from '../../types.ts';

interface FormatToggleProps {
  value: PreviewFormat;
  onChange: (format: PreviewFormat) => void;
  disabled?: boolean;
}

const OPTIONS: { id: PreviewFormat; label: string }[] = [
  { id: 'svg', label: 'SVG' },
  { id: 'png', label: 'PNG' },
];

export function FormatToggle({ value, onChange, disabled }: FormatToggleProps) {
  return (
    <div className="inline-flex items-center rounded-full border border-slate-800 bg-slate-900 p-1">
      {OPTIONS.map((option) => (
        <button
          key={option.id}
          type="button"
          onClick={() => onChange(option.id)}
          disabled={disabled}
          className={[
            'rounded-full px-4 py-1.5 text-sm font-semibold transition-all duration-200',
            value === option.id
              ? 'bg-slate-700 text-white shadow-sm'
              : 'text-slate-400 hover:bg-slate-800 hover:text-slate-200',
            'disabled:cursor-not-allowed disabled:opacity-50',
          ].join(' ')}
        >
          {option.label}
        </button>
      ))}
    </div>
  );
}
