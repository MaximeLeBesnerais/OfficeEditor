import { useCallback, useEffect, useMemo, useState } from 'react';
import { convertFile } from '../api.ts';
import { ConvertButton } from '../components/ConvertButton.tsx';
import { DownloadButton } from '../components/DownloadButton.tsx';
import { FileDropZone } from '../components/FileDropZone.tsx';
import { FormatSelector } from '../components/FormatSelector.tsx';
import { PreviewPanel } from '../components/PreviewPanel.tsx';
import { SampleGallery } from '../components/SampleGallery.tsx';
import { StatusMessage } from '../components/StatusMessage.tsx';
import { getFormatMeta } from '../utils/format.ts';
import type { ConversionResult, Sample, SourceType, TargetFormat } from '../types.ts';

interface ConvertPageProps {
  samples: Sample[];
  preselectedSample?: string | null;
  onConsumePreselection?: () => void;
}

function detectSourceType(file: File | null, sampleFormat: string | null): SourceType {
  const ext = file?.name.split('.').pop()?.toLowerCase() ?? sampleFormat?.toLowerCase() ?? '';
  if (['docx', 'pptx', 'xlsx'].includes(ext)) return 'office';
  if (ext === 'pdf') return 'pdf';
  if (['png', 'jpg', 'jpeg', 'svg', 'gif', 'webp', 'bmp'].includes(ext)) return 'image';
  return 'unknown';
}

export function ConvertPage({ samples, preselectedSample, onConsumePreselection }: ConvertPageProps) {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [selectedSample, setSelectedSample] = useState<string | null>(null);
  const [targetFormat, setTargetFormat] = useState<TargetFormat>('pdf');
  const [result, setResult] = useState<ConversionResult | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const sourceType = useMemo(
    () => detectSourceType(selectedFile, selectedSample ? getSampleFormat(samples, selectedSample) : null),
    [selectedFile, selectedSample, samples],
  );

  useEffect(() => {
    if (preselectedSample) {
      setSelectedSample(preselectedSample);
      setSelectedFile(null);
      setError(null);
      setResult(null);
      onConsumePreselection?.();
    }
  }, [preselectedSample, onConsumePreselection]);

  useEffect(() => {
    if (sourceType === 'office' && !['pdf', 'png', 'svg'].includes(targetFormat)) {
      setTargetFormat('pdf');
    }
    if (sourceType === 'pdf' && !['png', 'svg', 'docx'].includes(targetFormat)) {
      setTargetFormat('png');
    }
    if (sourceType === 'image' && !['pdf', 'png', 'svg'].includes(targetFormat)) {
      setTargetFormat('pdf');
    }
  }, [sourceType, targetFormat]);

  const handleFileSelected = useCallback((file: File | null) => {
    setSelectedFile(file);
    if (file) setSelectedSample(null);
    setError(null);
    setResult(null);
  }, []);

  const handleSelectSample = useCallback((name: string | null) => {
    setSelectedSample(name);
    if (name) setSelectedFile(null);
    setError(null);
    setResult(null);
  }, []);

  const handleConvert = useCallback(async () => {
    setError(null);
    setResult(null);
    setIsLoading(true);

    try {
      const formData = new FormData();
      if (selectedFile) {
        formData.append('file', selectedFile);
      } else if (selectedSample) {
        formData.append('sampleName', selectedSample);
      }
      formData.append('targetFormat', targetFormat);

      const conversionResult = await convertFile(formData);
      setResult(conversionResult);

      if (!conversionResult.success) {
        setError(conversionResult.errorMessage ?? 'Conversion failed.');
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An unexpected error occurred.');
    } finally {
      setIsLoading(false);
    }
  }, [selectedFile, selectedSample, targetFormat]);

  const handleReset = useCallback(() => {
    setSelectedFile(null);
    setSelectedSample(null);
    setTargetFormat('pdf');
    setResult(null);
    setError(null);
  }, []);

  const canConvert = Boolean(selectedFile || selectedSample);
  const resultMeta = result?.success
    ? getFormatMeta(result.outputFileName.split('.').pop() ?? '')
    : null;

  return (
    <div className="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8 py-10 space-y-8">
      <section className="text-center space-y-3">
        <h1 className="text-3xl sm:text-4xl font-extrabold text-slate-900 tracking-tight">
          Convert Office documents
        </h1>
        <p className="text-base sm:text-lg text-slate-600 max-w-2xl mx-auto">
          Upload a file or pick a sample, choose your target format, and download the result in seconds.
        </p>
      </section>

      {error && (
        <StatusMessage
          type="error"
          title="Conversion failed"
          message={error}
          messages={result?.messages}
        />
      )}

      <section className="bg-white rounded-3xl border border-slate-200 shadow-xl shadow-slate-200/50 overflow-hidden">
        <div className="p-6 sm:p-8 space-y-8">
          {/* Step 1: Source */}
          <div className="space-y-4">
            <div className="flex items-center gap-3">
              <span className="flex items-center justify-center w-7 h-7 rounded-full bg-indigo-100 text-indigo-700 text-sm font-bold">
                1
              </span>
              <h2 className="text-lg font-bold text-slate-900">Source</h2>
            </div>

            <FileDropZone selectedFile={selectedFile} onFileSelected={handleFileSelected} />

            {samples.length > 0 && (
              <div className="space-y-2">
                <p className="text-sm font-semibold text-slate-700">Or try a sample</p>
                <SampleGallery
                  samples={samples}
                  selectedSample={selectedSample}
                  onSelectSample={handleSelectSample}
                  mode="compact"
                />
              </div>
            )}
          </div>

          <hr className="border-slate-100" />

          {/* Step 2: Format */}
          <div className="space-y-4">
            <div className="flex items-center gap-3">
              <span
                className={[
                  'flex items-center justify-center w-7 h-7 rounded-full text-sm font-bold',
                  canConvert
                    ? 'bg-indigo-100 text-indigo-700'
                    : 'bg-slate-100 text-slate-400',
                ].join(' ')}
              >
                2
              </span>
              <h2 className="text-lg font-bold text-slate-900">Target format</h2>
            </div>
            <FormatSelector
              sourceType={sourceType}
              selectedFormat={targetFormat}
              onSelectFormat={setTargetFormat}
            />
          </div>

          <hr className="border-slate-100" />

          {/* Step 3: Convert */}
          <div className="space-y-3">
            <div className="flex items-center gap-3">
              <span
                className={[
                  'flex items-center justify-center w-7 h-7 rounded-full text-sm font-bold',
                  canConvert
                    ? 'bg-indigo-100 text-indigo-700'
                    : 'bg-slate-100 text-slate-400',
                ].join(' ')}
              >
                3
              </span>
              <h2 className="text-lg font-bold text-slate-900">Convert</h2>
            </div>
            <ConvertButton
              isLoading={isLoading}
              disabled={!canConvert}
              onClick={handleConvert}
            />
            {!canConvert && !isLoading && (
              <p className="text-sm text-slate-500 text-center">
                Upload a file or select a sample to start.
              </p>
            )}
          </div>
        </div>
      </section>

      {/* Result */}
      {result?.success && (
        <section className="fade-in-up bg-white rounded-3xl border border-slate-200 shadow-xl shadow-slate-200/50 overflow-hidden"
        >
          <div className="p-6 sm:p-8 space-y-6">
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
              <div className="flex items-center gap-4">
                {resultMeta && (
                  <span
                    className={[
                      'w-12 h-12 rounded-xl flex items-center justify-center text-sm font-bold shadow-sm',
                      resultMeta.bgSolid,
                      'text-white',
                    ].join(' ')}
                  >
                    {resultMeta.icon}
                  </span>
                )}
                <div className="min-w-0">
                  <h2 className="text-xl font-bold text-slate-900 truncate">
                    {result.outputFileName}
                  </h2>
                  <div className="flex items-center gap-2 text-sm text-slate-500">
                    <span
                      className={[
                        'inline-flex items-center px-2 py-0.5 rounded-md text-xs font-bold uppercase tracking-wider',
                        resultMeta?.bg,
                        resultMeta?.text,
                      ].join(' ')}
                    >
                      {resultMeta?.label ?? targetFormat}
                    </span>
                    <span>Ready for download</span>
                  </div>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <DownloadButton result={result} />
              </div>
            </div>

            <PreviewPanel result={result} />

            <div className="flex justify-center">
              <button
                type="button"
                onClick={handleReset}
                className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl border border-slate-200 text-slate-600 font-semibold hover:bg-slate-50 hover:text-slate-900 transition-colors"
              >
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  className="w-4 h-4"
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
                Convert another
              </button>
            </div>
          </div>
        </section>
      )}
    </div>
  );
}

function getSampleFormat(samples: Sample[], name: string): string {
  return samples.find((s) => s.name === name)?.format ?? '';
}
