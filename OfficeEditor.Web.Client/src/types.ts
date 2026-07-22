export type PreviewFormat = 'svg' | 'png';

export interface DemoDeck {
  name: string;
  fileName: string;
  description: string;
  slideCount: number;
}

export interface DemoDecksResponse {
  decks: DemoDeck[];
}

export interface SlidePreviewDto {
  slide: number;
  format: string;
  contentType: string;
  contentBase64: string;
}

export interface DemoRenderResponse {
  success: boolean;
  deckId: string;
  slideCount: number;
  totalMilliseconds: number;
  previews: SlidePreviewDto[];
}

export interface CompareCapabilities {
  available: boolean;
  version: string | null;
  pdftoppmAvailable: boolean;
  skipReason: string | null;
}

export interface LibreOfficeLegResult {
  success: boolean;
  deck: string;
  slideCount: number;
  available: boolean;
  version: string | null;
  pdftoppmAvailable: boolean;
  conversionMilliseconds: number | null;
  rasterizationMilliseconds: number | null;
  totalMilliseconds: number | null;
  previews: SlidePreviewDto[] | null;
  pdfDownloadUrl: string | null;
  error: string | null;
}

export interface TypstLegResult {
  success: boolean;
  deck: string;
  slideCount: number;
  pngMilliseconds: number;
  pdfMilliseconds: number | null;
  totalMilliseconds: number;
  previews: SlidePreviewDto[];
  pdfDownloadUrl: string | null;
  pdfError: string | null;
}

export type GenerationDocument = Record<string, unknown>;

export interface GenerationIssue {
  path: string;
  message: string;
  suggestion: string | null;
  severity: string;
}

export interface GenerateDeckResponse {
  success: boolean;
  deckId: string | null;
  slideCount: number;
  downloadUrl: string | null;
  previews: SlidePreviewDto[];
  warnings: string[];
  pipelineWarnings: string[];
  previewError: string | null;
  errorMessage: string | null;
  errors?: GenerationIssue[];
  generationMilliseconds: number;
  totalMilliseconds: number;
}
