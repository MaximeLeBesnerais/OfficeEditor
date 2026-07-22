import { useState } from 'react';
import { AnyRenderScreen } from './AnyRenderScreen.tsx';
import { CompareScreen } from './CompareScreen.tsx';
import { GenerateScreen } from './GenerateScreen.tsx';
import { RenderScreen } from './RenderScreen.tsx';

type DemoTab = 'render' | 'generate' | 'any' | 'compare';

export function DemoPage() {
  const [tab, setTab] = useState<DemoTab>(() => {
    const screen = new URLSearchParams(window.location.search).get('screen');
    if (screen === 'generate') return 'generate';
    if (screen === 'any') return 'any';
    if (screen === 'compare') return 'compare';
    return 'render';
  });

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100">
      <header className="sticky top-0 z-50 border-b border-slate-800/80 bg-slate-950/80 backdrop-blur-xl">
        <div className="mx-auto flex h-16 max-w-7xl items-center justify-between gap-4 px-4 sm:px-6 lg:px-8">
          <div className="flex items-center gap-3">
            <div className="relative flex h-9 w-9 items-center justify-center overflow-hidden rounded-xl bg-gradient-to-br from-indigo-500 via-violet-600 to-fuchsia-600 text-sm font-bold text-white shadow-lg shadow-indigo-500/20">
              <span className="relative z-10">OE</span>
              <div className="absolute inset-0 bg-gradient-to-tr from-white/20 to-transparent" />
            </div>
            <span className="text-lg font-bold tracking-tight">
              OfficeEditor
              <span className="ml-2 font-semibold text-slate-500">— Live Demo</span>
            </span>
          </div>

          <nav className="flex items-center gap-3">
            <div className="flex items-center rounded-full border border-slate-800 bg-slate-900 p-1">
              {(
                [
                  { id: 'render', label: '1 · Render' },
                  { id: 'generate', label: '2 · Generate' },
                  { id: 'any', label: '3 · Any render' },
                  { id: 'compare', label: '4 · Compare' },
                ] as const
              ).map((item) => (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => setTab(item.id)}
                  className={[
                    'rounded-full px-4 py-1.5 text-sm font-semibold transition-all duration-200',
                    tab === item.id
                      ? 'bg-slate-700 text-white shadow-sm'
                      : 'text-slate-400 hover:bg-slate-800 hover:text-slate-200',
                  ].join(' ')}
                >
                  {item.label}
                </button>
              ))}
            </div>
          </nav>
        </div>
      </header>

      <main className="mx-auto max-w-7xl px-4 py-8 sm:px-6 lg:px-8">
        {tab === 'render' ? (
          <RenderScreen />
        ) : tab === 'generate' ? (
          <GenerateScreen />
        ) : tab === 'any' ? (
          <AnyRenderScreen />
        ) : (
          <CompareScreen />
        )}
      </main>
    </div>
  );
}
