// Vendored theme presets for the /demo generate screen. These mirror
// demo/themes.json at the repo root — keep the values in sync when it changes.
// Switching theme replaces the `design` key of the generation document
// wholesale; slides are untouched.

export interface ThemeDesign {
  palette: {
    primary: string;
    accent: string;
    ink: string;
    paper: string;
    muted: string;
    mist: string;
    glow: string;
  };
  fonts: {
    display: string;
    body: string;
  };
  shape: {
    cornerRadius: number;
    cardStyle: 'outline' | 'filled';
  };
  metrics: {
    marginPt: number;
    gutterPt: number;
    titleSizePt: number;
    bodySizePt: number;
  };
}

export interface ThemePreset {
  id: string;
  name: string;
  design: ThemeDesign;
}

const SHARED_FONTS: ThemeDesign['fonts'] = {
  display: 'Avenir Next',
  body: 'Helvetica Neue',
};

const SHARED_SHAPE: ThemeDesign['shape'] = {
  cornerRadius: 0,
  cardStyle: 'outline',
};

const SHARED_METRICS: ThemeDesign['metrics'] = {
  marginPt: 48,
  gutterPt: 18,
  titleSizePt: 42,
  bodySizePt: 16,
};

export const THEME_PRESETS: ThemePreset[] = [
  {
    id: 'northwind',
    name: 'Northwind Teal',
    design: {
      palette: {
        ink: '#1E2A32',
        paper: '#FFFFFF',
        primary: '#0E7C7B',
        accent: '#D9A441',
        muted: '#7A8288',
        mist: '#EDF2F2',
        glow: '#35B5B3',
      },
      fonts: SHARED_FONTS,
      shape: SHARED_SHAPE,
      metrics: SHARED_METRICS,
    },
  },
  {
    id: 'corporate-blue',
    name: 'Corporate Blue',
    design: {
      palette: {
        ink: '#17263E',
        paper: '#FFFFFF',
        primary: '#1C5C9E',
        accent: '#B07E28',
        muted: '#6E7B8A',
        mist: '#EDF1F7',
        glow: '#6FA8DC',
      },
      fonts: SHARED_FONTS,
      shape: SHARED_SHAPE,
      metrics: SHARED_METRICS,
    },
  },
  {
    id: 'heritage',
    name: 'Pres-pro Heritage',
    design: {
      palette: {
        ink: '#2C3932',
        paper: '#FFFFFF',
        primary: '#9C8A54',
        accent: '#169C9A',
        muted: '#7A7567',
        mist: '#F1EEE4',
        glow: '#2FB3B1',
      },
      fonts: SHARED_FONTS,
      shape: SHARED_SHAPE,
      metrics: SHARED_METRICS,
    },
  },
];

export const DEFAULT_THEME_ID = 'northwind';
