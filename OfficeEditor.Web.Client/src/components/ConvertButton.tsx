interface ConvertButtonProps {
  isLoading: boolean;
  disabled: boolean;
  onClick: () => void;
}

export function ConvertButton({ isLoading, disabled, onClick }: ConvertButtonProps) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled || isLoading}
      className={[
        'w-full inline-flex items-center justify-center gap-2.5',
        'px-6 py-4 rounded-xl text-white font-bold text-lg shadow-lg shadow-indigo-500/25',
        'bg-gradient-to-r from-indigo-600 to-violet-600',
        'hover:from-indigo-700 hover:to-violet-700 hover:shadow-xl hover:shadow-indigo-500/30',
        'active:scale-[0.99]',
        'focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:ring-offset-2',
        'disabled:opacity-50 disabled:cursor-not-allowed disabled:shadow-none disabled:hover:from-indigo-600 disabled:hover:to-violet-600 disabled:active:scale-100',
        'transition-all duration-200',
      ].join(' ')}
    >
      {isLoading ? (
        <>
          <svg
            className="animate-spin h-5 w-5 text-white"
            xmlns="http://www.w3.org/2000/svg"
            fill="none"
            viewBox="0 0 24 24"
            aria-hidden="true"
          >
            <circle
              className="opacity-25"
              cx="12"
              cy="12"
              r="10"
              stroke="currentColor"
              strokeWidth="4"
            />
            <path
              className="opacity-75"
              fill="currentColor"
              d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"
            />
          </svg>
          <span>Converting your file…</span>
        </>
      ) : (
        <>
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
              d="M16.023 9.348h4.992v-.001M2.985 19.644v-4.992m0 0h4.992m-4.993 4.992 6.364-6.364m0 0 6.364 6.364m-6.364-6.364v-9.03"
            />
          </svg>
          <span>Convert now</span>
        </>
      )}
    </button>
  );
}
