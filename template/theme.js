// Classmethod PPTX Theme - derived from classmethod-marp-theme
// All colors are 6-char hex WITHOUT "#" prefix (PptxGenJS requirement)

const COLORS = {
  // Brand accent colors
  BLUE: "2C67E5",
  RED: "DF3756",

  // Grayscale (Classmethod brand spec)
  BLACK: "000000",
  GRAY_DARK: "262626",
  GRAY_MEDIUM: "434343", // primary text
  GRAY_LIGHT: "595959", // secondary text
  GRAY_LIGHTER: "646464",
  GRAY_LIGHTEST: "999999", // page numbers, muted elements
  WHITE: "FFFFFF",

  // Backgrounds
  BG_DEFAULT: "FFFFFF",
  BG_ACCENT: "F3F3F3",
  BG_CODE: "E1E1E1",
  BORDER: "D9D9D9",

  // Derived / supplementary
  BLUE_LIGHT: "EBF0FC",
  RED_LIGHT: "FDEEF1",
};

const FONTS = {
  HEADING: "Hiragino Sans",
  BODY: "Hiragino Sans",
  CODE: "Consolas",
};

// Font sizes in points (adapted from Marp CSS px → PowerPoint pt)
const SIZE = {
  TITLE: 36,
  H2: 28,
  H3: 24,
  BODY: 16,
  SMALL: 14,
  CAPTION: 12,
  NOTE: 10,
  STAT_NUMBER: 60,
};

// Slide dimensions: LAYOUT_16x9 = 10" × 5.625"
const SLIDE = {
  W: 10,
  H: 5.625,
};

// Standard layout measurements (inches)
const LAYOUT = {
  MARGIN_X: 0.6,
  MARGIN_Y: 0.5,
  HEADER_H: 0.7,
  HEADER_LINE_Y: 0.72,
  CONTENT_TOP: 0.95,
  CONTENT_W: 10 - 0.6 * 2, // 8.8"
  CONTENT_H: 5.625 - 0.95 - 0.35, // ~4.325"
  FOOTER_Y: 5.25,
  GAP: 0.3,
};

// Reusable shadow factory (fresh object each call to avoid PptxGenJS mutation bug)
function makeShadow() {
  return {
    type: "outer",
    color: "000000",
    blur: 4,
    offset: 2,
    angle: 135,
    opacity: 0.1,
  };
}

function makeCardShadow() {
  return {
    type: "outer",
    color: "000000",
    blur: 6,
    offset: 2,
    angle: 135,
    opacity: 0.12,
  };
}

module.exports = {
  COLORS,
  FONTS,
  SIZE,
  SLIDE,
  LAYOUT,
  makeShadow,
  makeCardShadow,
};
