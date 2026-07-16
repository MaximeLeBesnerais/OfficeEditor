import { useEffect, useState } from 'react';
import { fetchSamples } from './api.ts';
import { Header } from './components/Header.tsx';
import { ConvertPage } from './pages/ConvertPage.tsx';
import { SamplesPage } from './pages/SamplesPage.tsx';
import { StatusMessage } from './components/StatusMessage.tsx';
import type { Sample } from './types.ts';

export default function App() {
  const [currentPage, setCurrentPage] = useState<'convert' | 'samples'>('convert');
  const [samples, setSamples] = useState<Sample[]>([]);
  const [samplesError, setSamplesError] = useState<string | null>(null);
  const [preselectedSample, setPreselectedSample] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    fetchSamples()
      .then((data) => {
        if (!cancelled) setSamples(data);
      })
      .catch((err: unknown) => {
        if (!cancelled) {
          setSamplesError(
            err instanceof Error
              ? err.message
              : 'Could not load samples from the API.',
          );
        }
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const handleTrySample = (name: string) => {
    setPreselectedSample(name);
    setCurrentPage('convert');
  };

  const handleConsumePreselection = () => {
    setPreselectedSample(null);
  };

  return (
    <div className="min-h-screen bg-gradient-to-b from-slate-50 to-slate-100 flex flex-col">
      <Header currentPage={currentPage} onNavigate={setCurrentPage} />
      {samplesError && (
        <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 pt-6 w-full">
          <StatusMessage
            type="warning"
            title="Samples unavailable"
            message={samplesError}
          />
        </div>
      )}
      <main className="flex-1">
        {currentPage === 'convert' ? (
          <ConvertPage
            samples={samples}
            preselectedSample={preselectedSample}
            onConsumePreselection={handleConsumePreselection}
          />
        ) : (
          <SamplesPage samples={samples} onTrySample={handleTrySample} />
        )}
      </main>
    </div>
  );
}
