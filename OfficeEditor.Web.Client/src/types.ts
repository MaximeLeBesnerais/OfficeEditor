export interface Sample {
  name: string;
  format: string;
  description: string;
}

export interface ConversionResult {
  success: boolean;
  outputFileName: string;
  contentType: string;
  downloadUrl: string;
  previewUrl: string | null;
  errorMessage: string | null;
  messages: string[];
}

export type TargetFormat = 'pdf' | 'png' | 'svg' | 'docx' | 'pptx' | 'xlsx';

export type SourceType = 'office' | 'pdf' | 'image' | 'unknown';
