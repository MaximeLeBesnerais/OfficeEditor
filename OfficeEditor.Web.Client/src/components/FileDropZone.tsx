import { useCallback, useState } from 'react';
import { formatFileSize } from '../utils/format.ts';

interface FileDropZoneProps {
  selectedFile: File | null;
  onFileSelected: (file: File | null) => void;
}

export function FileDropZone({ selectedFile, onFileSelected }: FileDropZoneProps) {
  const [isDragging, setIsDragging] = useState(false);

  const handleDragOver = useCallback((event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback(
    (event: React.DragEvent<HTMLDivElement>) => {
      event.preventDefault();
      setIsDragging(false);
      const file = event.dataTransfer.files?.[0] ?? null;
      onFileSelected(file);
    },
    [onFileSelected],
  );

  const handleInputChange = useCallback(
    (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0] ?? null;
      onFileSelected(file);
    },
    [onFileSelected],
  );

  const handleRemove = useCallback(
    (event: React.MouseEvent) => {
      event.stopPropagation();
      onFileSelected(null);
    },
    [onFileSelected],
  );

  return (
    <div
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      className={[
        'relative group rounded-2xl border-2 border-dashed text-center transition-all duration-200',
        'bg-slate-50/50 hover:bg-indigo-50/60',
        isDragging
          ? 'border-indigo-500 bg-indigo-50 scale-[1.01] shadow-xl shadow-indigo-500/10'
          : 'border-slate-300 hover:border-indigo-400',
        selectedFile ? 'p-6' : 'p-10 sm:p-12',
      ].join(' ')}
    >
      <input
        type="file"
        id="file-upload"
        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
        onChange={handleInputChange}
        aria-label="Upload a file"
      />

      {selectedFile ? (
        <div className="relative flex items-center gap-4 pointer-events-none">
          <div className="w-14 h-14 rounded-xl bg-indigo-100 text-indigo-600 flex items-center justify-center shrink-0">
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
                d="M19.5 14.25v-2.625a3.375 3.375 0 0 0-3.375-3.375h-1.5A1.125 1.125 0 0 1 13.5 7.125v-1.5a3.375 3.375 0 0 0-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 0 0-9-9Z"
              />
            </svg>
          </div>
          <div className="flex-1 text-left min-w-0">
            <p className="text-slate-900 font-semibold truncate">{selectedFile.name}</p>
            <p className="text-slate-500 text-sm">
              {formatFileSize(selectedFile.size)} · Click or drop to replace
            </p>
          </div>
          <button
            type="button"
            onClick={handleRemove}
            className="pointer-events-auto relative z-10 shrink-0 inline-flex items-center justify-center w-9 h-9 rounded-lg bg-white border border-slate-200 text-slate-500 hover:text-red-600 hover:border-red-200 hover:bg-red-50 transition-colors shadow-sm"
            aria-label="Remove selected file"
          >
            <svg
              xmlns="http://www.w3.org/2000/svg"
              className="w-5 h-5"
              fill="none"
              viewBox="0 0 24 24"
              stroke="currentColor"
              strokeWidth={1.5}
              aria-hidden="true"
            >
              <path strokeLinecap="round" strokeLinejoin="round" d="M6 18 18 6M6 6l12 12" />
            </svg>
          </button>
        </div>
      ) : (
        <div className="pointer-events-none">
          <div className="mx-auto w-16 h-16 rounded-2xl bg-gradient-to-br from-indigo-100 to-violet-100 text-indigo-600 flex items-center justify-center mb-4 group-hover:scale-110 transition-transform duration-200">
            <svg
              xmlns="http://www.w3.org/2000/svg"
              className="w-8 h-8"
              fill="none"
              viewBox="0 0 24 24"
              stroke="currentColor"
              strokeWidth={1.5}
              aria-hidden="true"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M12 16.5V9.75m0 0-3 3m3-3 3 3M6.75 19.5h10.5a2.25 2.25 0 0 0 2.25-2.25V6.75a2.25 2.25 0 0 0-2.25-2.25H6.75A2.25 2.25 0 0 0 4.5 6.75v10.5a2.25 2.25 0 0 0 2.25 2.25Z"
              />
            </svg>
          </div>
          <p className="text-slate-900 font-semibold text-lg">
            Drop a file here, or click to browse
          </p>
          <p className="text-slate-500 text-sm mt-1">
            Supports DOCX, PPTX, XLSX, PDF, PNG, SVG
          </p>
        </div>
      )}
    </div>
  );
}
