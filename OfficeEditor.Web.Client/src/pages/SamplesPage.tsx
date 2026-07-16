import { SampleGallery } from '../components/SampleGallery.tsx';
import type { Sample } from '../types.ts';

interface SamplesPageProps {
  samples: Sample[];
  onTrySample: (name: string) => void;
}

export function SamplesPage({ samples, onTrySample }: SamplesPageProps) {
  return (
    <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-10 space-y-8">
      <section className="text-center space-y-3">
        <h1 className="text-3xl sm:text-4xl font-extrabold text-slate-900 tracking-tight">
          Sample gallery
        </h1>
        <p className="text-base sm:text-lg text-slate-600 max-w-2xl mx-auto">
          Browse the built-in samples. Click "Try this file" on any card to use it as a starting point.
        </p>
      </section>

      <section>
        <SampleGallery
          samples={samples}
          selectedSample={null}
          onSelectSample={(name) => name && onTrySample(name)}
          mode="gallery"
        />
      </section>
    </div>
  );
}
