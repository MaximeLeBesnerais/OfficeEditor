interface HeaderProps {
  currentPage: 'convert' | 'samples';
  onNavigate: (page: 'convert' | 'samples') => void;
}

export function Header({ currentPage, onNavigate }: HeaderProps) {
  return (
    <header className="sticky top-0 z-50 bg-white/80 backdrop-blur-xl border-b border-slate-200/80">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="relative w-9 h-9 rounded-xl bg-gradient-to-br from-indigo-500 via-violet-600 to-fuchsia-600 flex items-center justify-center text-white font-bold text-sm shadow-lg shadow-indigo-500/20 overflow-hidden">
            <span className="relative z-10">OE</span>
            <div className="absolute inset-0 bg-gradient-to-tr from-white/20 to-transparent" />
          </div>
          <span className="text-lg font-bold text-slate-900 tracking-tight">
            OfficeEditor
          </span>
        </div>

        <nav className="flex items-center p-1 bg-slate-100/80 rounded-full">
          <button
            type="button"
            onClick={() => onNavigate('convert')}
            className={[
              'px-4 py-1.5 rounded-full text-sm font-semibold transition-all duration-200',
              currentPage === 'convert'
                ? 'bg-white text-indigo-700 shadow-sm'
                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-200/50',
            ].join(' ')}
          >
            Convert
          </button>
          <button
            type="button"
            onClick={() => onNavigate('samples')}
            className={[
              'px-4 py-1.5 rounded-full text-sm font-semibold transition-all duration-200',
              currentPage === 'samples'
                ? 'bg-white text-indigo-700 shadow-sm'
                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-200/50',
            ].join(' ')}
          >
            Samples
          </button>
        </nav>
      </div>
    </header>
  );
}
