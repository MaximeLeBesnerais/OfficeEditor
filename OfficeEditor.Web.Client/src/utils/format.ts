import type { SourceType, TargetFormat } from '../types.ts';

export interface FormatMeta {
  id: string;
  label: string;
  description: string;
  color: string;
  bg: string;
  bgSolid: string;
  text: string;
  ring: string;
  border: string;
  icon: string;
}

const FORMAT_REGISTRY: Record<string, FormatMeta> = {
  pptx: {
    id: 'pptx',
    label: 'PPTX',
    description: 'PowerPoint',
    color: 'orange',
    bg: 'bg-orange-50',
    bgSolid: 'bg-orange-500',
    text: 'text-orange-700',
    ring: 'ring-orange-200',
    border: 'border-orange-200',
    icon: 'P',
  },
  docx: {
    id: 'docx',
    label: 'DOCX',
    description: 'Word',
    color: 'blue',
    bg: 'bg-blue-50',
    bgSolid: 'bg-blue-500',
    text: 'text-blue-700',
    ring: 'ring-blue-200',
    border: 'border-blue-200',
    icon: 'W',
  },
  xlsx: {
    id: 'xlsx',
    label: 'XLSX',
    description: 'Excel',
    color: 'green',
    bg: 'bg-green-50',
    bgSolid: 'bg-green-500',
    text: 'text-green-700',
    ring: 'ring-green-200',
    border: 'border-green-200',
    icon: 'X',
  },
  pdf: {
    id: 'pdf',
    label: 'PDF',
    description: 'Document',
    color: 'red',
    bg: 'bg-red-50',
    bgSolid: 'bg-red-500',
    text: 'text-red-700',
    ring: 'ring-red-200',
    border: 'border-red-200',
    icon: 'P',
  },
  png: {
    id: 'png',
    label: 'PNG',
    description: 'Image',
    color: 'purple',
    bg: 'bg-purple-50',
    bgSolid: 'bg-purple-500',
    text: 'text-purple-700',
    ring: 'ring-purple-200',
    border: 'border-purple-200',
    icon: 'I',
  },
  svg: {
    id: 'svg',
    label: 'SVG',
    description: 'Vector',
    color: 'violet',
    bg: 'bg-violet-50',
    bgSolid: 'bg-violet-500',
    text: 'text-violet-700',
    ring: 'ring-violet-200',
    border: 'border-violet-200',
    icon: 'V',
  },
  md: {
    id: 'md',
    label: 'Markdown',
    description: 'Text',
    color: 'slate',
    bg: 'bg-slate-100',
    bgSolid: 'bg-slate-500',
    text: 'text-slate-700',
    ring: 'ring-slate-200',
    border: 'border-slate-200',
    icon: 'M',
  },
};

export const allFormats: TargetFormat[] = ['pdf', 'png', 'svg', 'docx', 'pptx', 'xlsx'];

export function getFormatMeta(format: string): FormatMeta {
  return (
    FORMAT_REGISTRY[format.toLowerCase()] ?? {
      id: format,
      label: format.toUpperCase(),
      description: 'File',
      color: 'slate',
      bg: 'bg-slate-100',
      bgSolid: 'bg-slate-500',
      text: 'text-slate-700',
      ring: 'ring-slate-200',
      border: 'border-slate-200',
      icon: format.slice(0, 1).toUpperCase(),
    }
  );
}

export function getAvailableFormats(sourceType: SourceType): TargetFormat[] {
  switch (sourceType) {
    case 'office':
      return ['pdf', 'png', 'svg'];
    case 'pdf':
      return ['png', 'svg', 'docx'];
    case 'image':
      return ['pdf', 'png', 'svg'];
    default:
      return allFormats;
  }
}

export function formatFileSize(bytes: number): string {
  if (bytes === 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  const index = Math.min(
    Math.floor(Math.log10(bytes) / Math.log10(1024)),
    units.length - 1,
  );
  const value = bytes / Math.pow(1024, index);
  return `${value.toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
}
