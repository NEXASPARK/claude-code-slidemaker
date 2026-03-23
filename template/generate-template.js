const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const path = require("path");

const {
  COLORS,
  FONTS,
  SIZE,
  SLIDE,
  LAYOUT,
  makeShadow,
  makeCardShadow,
} = require("./theme");

// ---------------------------------------------------------------------------
// Icon helper
// ---------------------------------------------------------------------------
function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) }),
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, "#" + color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

// ---------------------------------------------------------------------------
// Shared slide helpers
// ---------------------------------------------------------------------------

function addHeaderBar(pres, slide, title) {
  slide.addText(title, {
    x: 0.5,
    y: 0.08,
    w: 9,
    h: 0.6,
    fontSize: SIZE.TITLE - 4,
    fontFace: FONTS.HEADING,
    color: COLORS.GRAY_MEDIUM,
    bold: true,
    margin: 0,
  });
  slide.addShape(pres.shapes.LINE, {
    x: 0,
    y: LAYOUT.HEADER_LINE_Y,
    w: SLIDE.W,
    h: 0,
    line: { color: COLORS.BLACK, width: 1 },
  });
}

function addPageNumber(slide, num) {
  slide.addText(String(num), {
    x: SLIDE.W - 0.8,
    y: 0.12,
    w: 0.5,
    h: 0.5,
    fontSize: SIZE.CAPTION,
    fontFace: FONTS.BODY,
    color: COLORS.GRAY_LIGHTEST,
    align: "right",
    margin: 0,
  });
}

function addFooterNote(slide, text) {
  slide.addText(text, {
    x: LAYOUT.MARGIN_X,
    y: LAYOUT.FOOTER_Y,
    w: LAYOUT.CONTENT_W,
    h: 0.3,
    fontSize: SIZE.NOTE,
    fontFace: FONTS.BODY,
    color: COLORS.GRAY_LIGHT,
    align: "left",
    margin: 0,
  });
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
async function main() {
  const {
    FaRocket,
    FaChartLine,
    FaUsers,
    FaCog,
    FaLightbulb,
    FaShieldAlt,
    FaBolt,
    FaGlobe,
    FaCheckCircle,
    FaArrowRight,
    FaQuoteLeft,
    FaUser,
  } = require("react-icons/fa");

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "NEXASPARK";
  pres.title = "NEXASPARK Slide Template";

  const SHAPES = pres.shapes;

  // Pre-render icons
  const icons = {};
  const iconList = [
    ["rocket", FaRocket, COLORS.BLUE],
    ["chart", FaChartLine, COLORS.BLUE],
    ["users", FaUsers, COLORS.BLUE],
    ["cog", FaCog, COLORS.BLUE],
    ["lightbulb", FaLightbulb, COLORS.BLUE],
    ["shield", FaShieldAlt, COLORS.BLUE],
    ["bolt", FaBolt, COLORS.BLUE],
    ["globe", FaGlobe, COLORS.BLUE],
    ["check", FaCheckCircle, COLORS.BLUE],
    ["arrow", FaArrowRight, COLORS.GRAY_LIGHTEST],
    ["quote", FaQuoteLeft, COLORS.GRAY_LIGHTER],
    ["rocket_white", FaRocket, COLORS.WHITE],
    ["chart_white", FaChartLine, COLORS.WHITE],
    ["users_white", FaUsers, COLORS.WHITE],
    ["check_white", FaCheckCircle, COLORS.WHITE],
    ["lightbulb_white", FaLightbulb, COLORS.WHITE],
    ["shield_white", FaShieldAlt, COLORS.WHITE],
    ["globe_white", FaGlobe, COLORS.WHITE],
    ["cog_white", FaCog, COLORS.WHITE],
    ["user", FaUser, COLORS.WHITE],
  ];
  for (const [name, Comp, color] of iconList) {
    icons[name] = await iconToBase64Png(Comp, color);
  }

  let pageNum = 0;

  // =======================================================================
  // SLIDE 1: Title Slide
  // =======================================================================
  {
    const s = pres.addSlide();
    s.background = { color: COLORS.BG_ACCENT };

    // Company name (text-based logo)
    s.addText("株式会社 NEXASPARK", {
      x: 1,
      y: 1.22,
      w: 8,
      h: 0.6,
      fontSize: 20,
      fontFace: "Meiryo UI",
      color: COLORS.GRAY_LIGHTEST,
      bold: true,
      align: "center",
      valign: "middle",
    });

    s.addText("XXXX プレゼンテーションタイトル", {
      x: 0.5,
      y: 1.7,
      w: 9,
      h: 1.4,
      fontSize: SIZE.TITLE + 2,
      fontFace: FONTS.HEADING,
      color: COLORS.GRAY_DARK,
      bold: true,
      align: "center",
      valign: "middle",
    });

    s.addText("XXXX サブタイトル・説明文をここに入力", {
      x: 1,
      y: 3.1,
      w: 8,
      h: 0.6,
      fontSize: SIZE.BODY,
      fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHT,
      align: "center",
    });

    s.addShape(SHAPES.LINE, {
      x: 3.5,
      y: 3.9,
      w: 3,
      h: 0,
      line: { color: COLORS.BORDER, width: 1 },
    });

    s.addText("XXXX 20XX/XX/XX  発表者名", {
      x: 2,
      y: 4.1,
      w: 6,
      h: 0.5,
      fontSize: SIZE.SMALL,
      fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHTEST,
      align: "center",
    });
  }

  // =======================================================================
  // SLIDE 2: Section Divider
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    s.background = { color: COLORS.BG_ACCENT };

    s.addShape(SHAPES.RECTANGLE, {
      x: 0.6,
      y: 2.0,
      w: 0.07,
      h: 1.5,
      fill: { color: COLORS.BLUE },
    });

    s.addText("XXXX セクションタイトル", {
      x: 0.9,
      y: 2.0,
      w: 8.5,
      h: 0.8,
      fontSize: SIZE.H2 + 4,
      fontFace: FONTS.HEADING,
      color: COLORS.GRAY_DARK,
      bold: true,
      valign: "middle",
      margin: 0,
    });

    s.addText("XXXX セクションの説明テキスト", {
      x: 0.9,
      y: 2.9,
      w: 8.5,
      h: 0.6,
      fontSize: SIZE.BODY,
      fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHT,
      margin: 0,
    });
  }

  // =======================================================================
  // SLIDE 3: Standard Content (Header + Bullets)
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX スライドタイトル");
    addPageNumber(s, pageNum);

    s.addText(
      [
        {
          text: "XXXX 見出しテキスト",
          options: { bold: true, fontSize: SIZE.H3, breakLine: true },
        },
        { text: "", options: { fontSize: 8, breakLine: true } },
        {
          text: "XXXX 箇条書き項目1：説明テキストをここに入力",
          options: { bullet: true, breakLine: true },
        },
        {
          text: "XXXX 箇条書き項目2：説明テキストをここに入力",
          options: { bullet: true, breakLine: true },
        },
        {
          text: "XXXX 箇条書き項目3：説明テキストをここに入力",
          options: { bullet: true, breakLine: true },
        },
        {
          text: "XXXX 箇条書き項目4：説明テキストをここに入力",
          options: { bullet: true },
        },
      ],
      {
        x: LAYOUT.MARGIN_X,
        y: LAYOUT.CONTENT_TOP,
        w: LAYOUT.CONTENT_W,
        h: LAYOUT.CONTENT_H,
        fontSize: SIZE.BODY,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_MEDIUM,
        valign: "top",
        paraSpaceAfter: 4,
      },
    );
  }

  // =======================================================================
  // SLIDE 4: No Header (Full Content)
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();

    s.addText("XXXX フルコンテンツのタイトル", {
      x: LAYOUT.MARGIN_X,
      y: 0.3,
      w: LAYOUT.CONTENT_W,
      h: 0.8,
      fontSize: SIZE.TITLE,
      fontFace: FONTS.HEADING,
      color: COLORS.GRAY_DARK,
      bold: true,
      align: "center",
      margin: 0,
    });

    s.addText(
      [
        {
          text: "XXXX このレイアウトはヘッダーバーがなく、コンテンツ領域を広く使えます。情報量が多いスライドや、独自のレイアウトが必要な場合に適しています。",
          options: { breakLine: true },
        },
        { text: "", options: { fontSize: 8, breakLine: true } },
        {
          text: "XXXX 箇条書き項目1",
          options: { bullet: true, breakLine: true },
        },
        {
          text: "XXXX 箇条書き項目2",
          options: { bullet: true, breakLine: true },
        },
        { text: "XXXX 箇条書き項目3", options: { bullet: true } },
      ],
      {
        x: LAYOUT.MARGIN_X,
        y: 1.2,
        w: LAYOUT.CONTENT_W,
        h: 4.0,
        fontSize: SIZE.BODY,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_MEDIUM,
        valign: "top",
        paraSpaceAfter: 4,
      },
    );

    addPageNumber(s, pageNum);
  }

  // =======================================================================
  // SLIDE 5: Image Center
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX 図表のタイトル");
    addPageNumber(s, pageNum);

    const imgW = 7,
      imgH = 3.6;
    const imgX = (SLIDE.W - imgW) / 2;
    const imgY = LAYOUT.CONTENT_TOP + 0.15;
    s.addShape(SHAPES.RECTANGLE, {
      x: imgX,
      y: imgY,
      w: imgW,
      h: imgH,
      fill: { color: COLORS.BG_ACCENT },
      line: { color: COLORS.BORDER, width: 1, dashType: "dash" },
    });
    s.addText("XXXX 画像・図表プレースホルダー", {
      x: imgX,
      y: imgY,
      w: imgW,
      h: imgH,
      fontSize: SIZE.BODY,
      fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHTEST,
      align: "center",
      valign: "middle",
    });

    addFooterNote(s, "XXXX 出典・注釈テキスト");
  }

  // =======================================================================
  // SLIDE 6: Content + Image Right (50/50)
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX テキスト＋図（右）");
    addPageNumber(s, pageNum);

    const splitX = 5.1;
    s.addText(
      [
        {
          text: "XXXX 見出し",
          options: { bold: true, fontSize: SIZE.H3, breakLine: true },
        },
        { text: "", options: { fontSize: 8, breakLine: true } },
        {
          text: "XXXX 左側にテキスト説明を配置します。箇条書きや段落テキストで内容を表現します。",
          options: { breakLine: true },
        },
        { text: "", options: { fontSize: 8, breakLine: true } },
        {
          text: "XXXX ポイント1の説明",
          options: { bullet: true, breakLine: true },
        },
        {
          text: "XXXX ポイント2の説明",
          options: { bullet: true, breakLine: true },
        },
        { text: "XXXX ポイント3の説明", options: { bullet: true } },
      ],
      {
        x: LAYOUT.MARGIN_X,
        y: LAYOUT.CONTENT_TOP,
        w: splitX - LAYOUT.MARGIN_X - 0.3,
        h: LAYOUT.CONTENT_H,
        fontSize: SIZE.BODY,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_MEDIUM,
        valign: "top",
        paraSpaceAfter: 4,
      },
    );

    s.addShape(SHAPES.RECTANGLE, {
      x: splitX,
      y: LAYOUT.HEADER_LINE_Y,
      w: SLIDE.W - splitX,
      h: SLIDE.H - LAYOUT.HEADER_LINE_Y,
      fill: { color: COLORS.BG_ACCENT },
    });
    s.addText("XXXX 画像エリア", {
      x: splitX,
      y: LAYOUT.HEADER_LINE_Y,
      w: SLIDE.W - splitX,
      h: SLIDE.H - LAYOUT.HEADER_LINE_Y,
      fontSize: SIZE.BODY,
      fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHTEST,
      align: "center",
      valign: "middle",
    });
  }

  // =======================================================================
  // SLIDE 7: Content + Image Left (50/50)
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX テキスト＋図（左）");
    addPageNumber(s, pageNum);

    const splitX = 4.9;
    s.addShape(SHAPES.RECTANGLE, {
      x: 0,
      y: LAYOUT.HEADER_LINE_Y,
      w: splitX,
      h: SLIDE.H - LAYOUT.HEADER_LINE_Y,
      fill: { color: COLORS.BG_ACCENT },
    });
    s.addText("XXXX 画像エリア", {
      x: 0,
      y: LAYOUT.HEADER_LINE_Y,
      w: splitX,
      h: SLIDE.H - LAYOUT.HEADER_LINE_Y,
      fontSize: SIZE.BODY,
      fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHTEST,
      align: "center",
      valign: "middle",
    });

    s.addText(
      [
        {
          text: "XXXX 見出し",
          options: { bold: true, fontSize: SIZE.H3, breakLine: true },
        },
        { text: "", options: { fontSize: 8, breakLine: true } },
        {
          text: "XXXX 右側にテキストを配置するレイアウトです。",
          options: { breakLine: true },
        },
        { text: "", options: { fontSize: 8, breakLine: true } },
        {
          text: "XXXX ポイント1の説明",
          options: { bullet: true, breakLine: true },
        },
        {
          text: "XXXX ポイント2の説明",
          options: { bullet: true, breakLine: true },
        },
        { text: "XXXX ポイント3の説明", options: { bullet: true } },
      ],
      {
        x: splitX + 0.3,
        y: LAYOUT.CONTENT_TOP,
        w: SLIDE.W - splitX - 0.3 - LAYOUT.MARGIN_X,
        h: LAYOUT.CONTENT_H,
        fontSize: SIZE.BODY,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_MEDIUM,
        valign: "top",
        paraSpaceAfter: 4,
      },
    );
  }

  // =======================================================================
  // SLIDE 8: Two Column Layout
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX 2カラムレイアウト");
    addPageNumber(s, pageNum);

    const colW = (LAYOUT.CONTENT_W - LAYOUT.GAP) / 2;
    const colY = LAYOUT.CONTENT_TOP;

    s.addText(
      [
        {
          text: "XXXX 左カラム見出し",
          options: {
            bold: true,
            fontSize: SIZE.H3,
            breakLine: true,
            color: COLORS.GRAY_DARK,
          },
        },
        { text: "", options: { fontSize: 8, breakLine: true } },
        {
          text: "XXXX 左カラムの内容1",
          options: { bullet: true, breakLine: true },
        },
        {
          text: "XXXX 左カラムの内容2",
          options: { bullet: true, breakLine: true },
        },
        { text: "XXXX 左カラムの内容3", options: { bullet: true } },
      ],
      {
        x: LAYOUT.MARGIN_X,
        y: colY,
        w: colW,
        h: LAYOUT.CONTENT_H,
        fontSize: SIZE.BODY,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_MEDIUM,
        valign: "top",
        paraSpaceAfter: 4,
      },
    );

    s.addText(
      [
        {
          text: "XXXX 右カラム見出し",
          options: {
            bold: true,
            fontSize: SIZE.H3,
            breakLine: true,
            color: COLORS.GRAY_DARK,
          },
        },
        { text: "", options: { fontSize: 8, breakLine: true } },
        {
          text: "XXXX 右カラムの内容1",
          options: { bullet: true, breakLine: true },
        },
        {
          text: "XXXX 右カラムの内容2",
          options: { bullet: true, breakLine: true },
        },
        { text: "XXXX 右カラムの内容3", options: { bullet: true } },
      ],
      {
        x: LAYOUT.MARGIN_X + colW + LAYOUT.GAP,
        y: colY,
        w: colW,
        h: LAYOUT.CONTENT_H,
        fontSize: SIZE.BODY,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_MEDIUM,
        valign: "top",
        paraSpaceAfter: 4,
      },
    );

    s.addShape(SHAPES.LINE, {
      x: LAYOUT.MARGIN_X + colW + LAYOUT.GAP / 2,
      y: colY,
      w: 0,
      h: LAYOUT.CONTENT_H,
      line: { color: COLORS.BORDER, width: 1 },
    });
  }

  // =======================================================================
  // SLIDE 9: Three Column Layout
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX 3カラムレイアウト");
    addPageNumber(s, pageNum);

    const gap = 0.25;
    const colW = (LAYOUT.CONTENT_W - gap * 2) / 3;
    const colY = LAYOUT.CONTENT_TOP;
    const labels = ["XXXX カラム1", "XXXX カラム2", "XXXX カラム3"];

    for (let i = 0; i < 3; i++) {
      const x = LAYOUT.MARGIN_X + i * (colW + gap);
      s.addText(
        [
          {
            text: labels[i],
            options: {
              bold: true,
              fontSize: SIZE.H3,
              breakLine: true,
              color: COLORS.GRAY_DARK,
            },
          },
          { text: "", options: { fontSize: 8, breakLine: true } },
          { text: "XXXX 項目A", options: { bullet: true, breakLine: true } },
          { text: "XXXX 項目B", options: { bullet: true, breakLine: true } },
          { text: "XXXX 項目C", options: { bullet: true } },
        ],
        {
          x,
          y: colY,
          w: colW,
          h: LAYOUT.CONTENT_H,
          fontSize: SIZE.BODY,
          fontFace: FONTS.BODY,
          color: COLORS.GRAY_MEDIUM,
          valign: "top",
          paraSpaceAfter: 4,
        },
      );

      if (i < 2) {
        s.addShape(SHAPES.LINE, {
          x: x + colW + gap / 2,
          y: colY,
          w: 0,
          h: LAYOUT.CONTENT_H,
          line: { color: COLORS.BORDER, width: 1 },
        });
      }
    }
  }

  // =======================================================================
  // SLIDE 10: Key Stats / Numbers
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX 主要指標");
    addPageNumber(s, pageNum);

    const stats = [
      { num: "98%", label: "XXXX 指標ラベル1", icon: "rocket" },
      { num: "2.5x", label: "XXXX 指標ラベル2", icon: "chart" },
      { num: "500+", label: "XXXX 指標ラベル3", icon: "users" },
    ];

    const cardW = 2.6,
      gap = 0.35;
    const totalW = cardW * 3 + gap * 2;
    const startX = (SLIDE.W - totalW) / 2;
    const cardY = LAYOUT.CONTENT_TOP + 0.15;
    const cardH = 3.8;

    for (let i = 0; i < stats.length; i++) {
      const x = startX + i * (cardW + gap);

      // Card background
      s.addShape(SHAPES.RECTANGLE, {
        x,
        y: cardY,
        w: cardW,
        h: cardH,
        fill: { color: COLORS.WHITE },
        line: { color: COLORS.BORDER, width: 0.5 },
        shadow: makeCardShadow(),
      });
      // Blue top accent
      s.addShape(SHAPES.RECTANGLE, {
        x,
        y: cardY,
        w: cardW,
        h: 0.06,
        fill: { color: COLORS.BLUE },
      });
      // Icon in circle
      const circleSize = 0.6;
      s.addShape(SHAPES.OVAL, {
        x: x + (cardW - circleSize) / 2,
        y: cardY + 0.3,
        w: circleSize,
        h: circleSize,
        fill: { color: COLORS.BLUE },
      });
      const whiteKey = stats[i].icon + "_white";
      if (icons[whiteKey]) {
        s.addImage({
          data: icons[whiteKey],
          x: x + (cardW - 0.34) / 2,
          y: cardY + 0.3 + (circleSize - 0.34) / 2,
          w: 0.34,
          h: 0.34,
        });
      }
      // Big number
      s.addText(stats[i].num, {
        x,
        y: cardY + 1.15,
        w: cardW,
        h: 1.1,
        fontSize: SIZE.STAT_NUMBER - 8,
        fontFace: FONTS.HEADING,
        color: COLORS.BLUE,
        bold: true,
        align: "center",
        valign: "middle",
        margin: 0,
      });
      // Label
      s.addText(stats[i].label, {
        x: x + 0.2,
        y: cardY + 2.5,
        w: cardW - 0.4,
        h: 1.0,
        fontSize: SIZE.BODY,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_LIGHT,
        align: "center",
        valign: "top",
        margin: 0,
      });
    }
  }

  // =======================================================================
  // SLIDE 11: Timeline / Process
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX プロセス・手順");
    addPageNumber(s, pageNum);

    const steps = [
      { num: "01", title: "XXXX ステップ1", desc: "XXXX 説明文をここに入力" },
      { num: "02", title: "XXXX ステップ2", desc: "XXXX 説明文をここに入力" },
      { num: "03", title: "XXXX ステップ3", desc: "XXXX 説明文をここに入力" },
      { num: "04", title: "XXXX ステップ4", desc: "XXXX 説明文をここに入力" },
    ];

    const stepW = 2.0,
      gapX = 0.35;
    const totalW = stepW * 4 + gapX * 3;
    const startX = (SLIDE.W - totalW) / 2;
    const baseY = LAYOUT.CONTENT_TOP + 0.6;

    // Connecting line behind circles
    const lineY = baseY + 0.35;
    s.addShape(SHAPES.LINE, {
      x: startX + stepW / 2,
      y: lineY,
      w: totalW - stepW,
      h: 0,
      line: { color: COLORS.BORDER, width: 2 },
    });

    for (let i = 0; i < steps.length; i++) {
      const x = startX + i * (stepW + gapX);
      const circleCx = x + stepW / 2;
      const circleSize = 0.7;

      s.addShape(SHAPES.OVAL, {
        x: circleCx - circleSize / 2,
        y: baseY,
        w: circleSize,
        h: circleSize,
        fill: { color: COLORS.BLUE },
      });
      s.addText(steps[i].num, {
        x: circleCx - circleSize / 2,
        y: baseY,
        w: circleSize,
        h: circleSize,
        fontSize: SIZE.BODY,
        fontFace: FONTS.HEADING,
        color: COLORS.WHITE,
        bold: true,
        align: "center",
        valign: "middle",
        margin: 0,
      });

      s.addText(steps[i].title, {
        x,
        y: baseY + 0.9,
        w: stepW,
        h: 0.5,
        fontSize: SIZE.BODY,
        fontFace: FONTS.HEADING,
        color: COLORS.GRAY_DARK,
        bold: true,
        align: "center",
        valign: "top",
        margin: 0,
      });
      s.addText(steps[i].desc, {
        x,
        y: baseY + 1.4,
        w: stepW,
        h: 1.2,
        fontSize: SIZE.SMALL,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_LIGHT,
        align: "center",
        valign: "top",
        margin: 0,
      });
    }
  }

  // =======================================================================
  // SLIDE 12: Comparison / Before-After
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX 比較");
    addPageNumber(s, pageNum);

    const colW = 4.1,
      gap = 0.4;
    const startX = (SLIDE.W - colW * 2 - gap) / 2;
    const cardY = LAYOUT.CONTENT_TOP + 0.1;
    const cardH = 3.8;

    const sides = [
      {
        label: "XXXX Before / 現状",
        color: COLORS.GRAY_LIGHT,
        items: [
          "XXXX 課題・問題点1",
          "XXXX 課題・問題点2",
          "XXXX 課題・問題点3",
        ],
      },
      {
        label: "XXXX After / 改善後",
        color: COLORS.BLUE,
        items: ["XXXX 改善効果1", "XXXX 改善効果2", "XXXX 改善効果3"],
      },
    ];

    for (let i = 0; i < 2; i++) {
      const x = startX + i * (colW + gap);
      const side = sides[i];

      s.addShape(SHAPES.RECTANGLE, {
        x,
        y: cardY,
        w: colW,
        h: cardH,
        fill: { color: COLORS.WHITE },
        line: { color: COLORS.BORDER, width: 1 },
        shadow: makeCardShadow(),
      });
      s.addShape(SHAPES.RECTANGLE, {
        x,
        y: cardY,
        w: colW,
        h: 0.07,
        fill: { color: side.color },
      });
      s.addText(side.label, {
        x: x + 0.3,
        y: cardY + 0.3,
        w: colW - 0.6,
        h: 0.5,
        fontSize: SIZE.H3,
        fontFace: FONTS.HEADING,
        color: side.color,
        bold: true,
        margin: 0,
      });
      const items = side.items.map((item, idx) => ({
        text: item,
        options: { bullet: true, breakLine: idx < side.items.length - 1 },
      }));
      s.addText(items, {
        x: x + 0.3,
        y: cardY + 1.0,
        w: colW - 0.6,
        h: cardH - 1.3,
        fontSize: SIZE.BODY,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_MEDIUM,
        valign: "top",
        paraSpaceAfter: 6,
      });
    }
  }

  // =======================================================================
  // SLIDE 13: Icon Grid 2x3
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX 特徴・機能一覧");
    addPageNumber(s, pageNum);

    const gridIcons = [
      { key: "rocket", title: "XXXX 機能1", desc: "XXXX 説明テキスト" },
      { key: "chart", title: "XXXX 機能2", desc: "XXXX 説明テキスト" },
      { key: "shield", title: "XXXX 機能3", desc: "XXXX 説明テキスト" },
      { key: "lightbulb", title: "XXXX 機能4", desc: "XXXX 説明テキスト" },
      { key: "globe", title: "XXXX 機能5", desc: "XXXX 説明テキスト" },
      { key: "cog", title: "XXXX 機能6", desc: "XXXX 説明テキスト" },
    ];

    const cols = 3,
      rows = 2;
    const cellW = 2.7,
      cellH = 1.8;
    const gapX = 0.25,
      gapY = 0.2;
    const totalW = cellW * cols + gapX * (cols - 1);
    const totalH = cellH * rows + gapY * (rows - 1);
    const startX = (SLIDE.W - totalW) / 2;
    const startY = LAYOUT.CONTENT_TOP + (LAYOUT.CONTENT_H - totalH) / 2;

    for (let idx = 0; idx < gridIcons.length; idx++) {
      const row = Math.floor(idx / cols);
      const col = idx % cols;
      const x = startX + col * (cellW + gapX);
      const y = startY + row * (cellH + gapY);

      const circleSize = 0.55;
      s.addShape(SHAPES.OVAL, {
        x: x + 0.15,
        y: y + 0.1,
        w: circleSize,
        h: circleSize,
        fill: { color: COLORS.BLUE },
      });
      const whiteIconKey = gridIcons[idx].key + "_white";
      if (icons[whiteIconKey]) {
        s.addImage({
          data: icons[whiteIconKey],
          x: x + 0.15 + (circleSize - 0.3) / 2,
          y: y + 0.1 + (circleSize - 0.3) / 2,
          w: 0.3,
          h: 0.3,
        });
      }

      s.addText(gridIcons[idx].title, {
        x: x + 0.15,
        y: y + 0.75,
        w: cellW - 0.3,
        h: 0.35,
        fontSize: SIZE.BODY,
        fontFace: FONTS.HEADING,
        color: COLORS.GRAY_DARK,
        bold: true,
        margin: 0,
      });
      s.addText(gridIcons[idx].desc, {
        x: x + 0.15,
        y: y + 1.1,
        w: cellW - 0.3,
        h: 0.6,
        fontSize: SIZE.SMALL,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_LIGHT,
        margin: 0,
      });
    }
  }

  // =======================================================================
  // SLIDE 14: Quote / Callout
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    s.background = { color: COLORS.BG_ACCENT };

    s.addImage({
      data: icons.quote,
      x: 1.2,
      y: 1.0,
      w: 0.6,
      h: 0.6,
    });

    s.addText(
      "XXXX ここに引用文やキーメッセージを入力します。印象的なフレーズや重要な発言を大きく表示するレイアウトです。",
      {
        x: 1.2,
        y: 1.7,
        w: 7.5,
        h: 2.2,
        fontSize: SIZE.H2,
        fontFace: FONTS.HEADING,
        color: COLORS.GRAY_DARK,
        italic: true,
        valign: "top",
        margin: 0,
      },
    );

    s.addShape(SHAPES.LINE, {
      x: 1.2,
      y: 4.1,
      w: 1.5,
      h: 0,
      line: { color: COLORS.BLUE, width: 2 },
    });
    s.addText("XXXX 発言者名  /  XXXX 役職・所属", {
      x: 1.2,
      y: 4.3,
      w: 7,
      h: 0.4,
      fontSize: SIZE.BODY,
      fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHT,
      margin: 0,
    });
  }

  // =======================================================================
  // SLIDE 15: Agenda / Table of Contents
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();

    s.addText("XXXX アジェンダ", {
      x: LAYOUT.MARGIN_X,
      y: 0.4,
      w: LAYOUT.CONTENT_W,
      h: 0.8,
      fontSize: SIZE.TITLE,
      fontFace: FONTS.HEADING,
      color: COLORS.GRAY_DARK,
      bold: true,
      margin: 0,
    });

    const agendaItems = [
      "XXXX アジェンダ項目1",
      "XXXX アジェンダ項目2",
      "XXXX アジェンダ項目3",
      "XXXX アジェンダ項目4",
      "XXXX アジェンダ項目5",
    ];

    const itemH = 0.7,
      startY = 1.4;

    for (let i = 0; i < agendaItems.length; i++) {
      const y = startY + i * (itemH + 0.12);
      const circleSize = 0.45;

      s.addShape(SHAPES.OVAL, {
        x: LAYOUT.MARGIN_X,
        y: y + 0.12,
        w: circleSize,
        h: circleSize,
        fill: { color: COLORS.BLUE },
      });
      s.addText(String(i + 1), {
        x: LAYOUT.MARGIN_X,
        y: y + 0.12,
        w: circleSize,
        h: circleSize,
        fontSize: SIZE.BODY - 2,
        fontFace: FONTS.HEADING,
        color: COLORS.WHITE,
        bold: true,
        align: "center",
        valign: "middle",
        margin: 0,
      });

      s.addText(agendaItems[i], {
        x: LAYOUT.MARGIN_X + 0.65,
        y: y,
        w: LAYOUT.CONTENT_W - 0.65,
        h: itemH,
        fontSize: SIZE.H3 - 2,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_MEDIUM,
        valign: "middle",
        margin: 0,
      });

      if (i < agendaItems.length - 1) {
        s.addShape(SHAPES.LINE, {
          x: LAYOUT.MARGIN_X + 0.65,
          y: y + itemH + 0.03,
          w: LAYOUT.CONTENT_W - 0.65,
          h: 0,
          line: { color: COLORS.BORDER, width: 0.5 },
        });
      }
    }
  }

  // =======================================================================
  // SLIDE 16: Table (data presentation)
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX テーブルレイアウト");
    addPageNumber(s, pageNum);

    const headerRow = [
      {
        text: "XXXX 項目",
        options: {
          bold: true,
          color: COLORS.GRAY_MEDIUM,
          fill: { color: COLORS.BG_ACCENT },
        },
      },
      {
        text: "XXXX 列2",
        options: {
          bold: true,
          color: COLORS.GRAY_MEDIUM,
          fill: { color: COLORS.BG_ACCENT },
        },
      },
      {
        text: "XXXX 列3",
        options: {
          bold: true,
          color: COLORS.GRAY_MEDIUM,
          fill: { color: COLORS.BG_ACCENT },
        },
      },
      {
        text: "XXXX 列4",
        options: {
          bold: true,
          color: COLORS.GRAY_MEDIUM,
          fill: { color: COLORS.BG_ACCENT },
        },
      },
    ];

    const dataRows = [];
    for (let r = 0; r < 5; r++) {
      dataRows.push([
        {
          text: `XXXX データ${r + 1}-1`,
          options: { color: COLORS.GRAY_MEDIUM },
        },
        {
          text: `XXXX データ${r + 1}-2`,
          options: { color: COLORS.GRAY_MEDIUM },
        },
        {
          text: `XXXX データ${r + 1}-3`,
          options: { color: COLORS.GRAY_MEDIUM },
        },
        {
          text: `XXXX データ${r + 1}-4`,
          options: { color: COLORS.GRAY_MEDIUM },
        },
      ]);
    }

    s.addTable([headerRow, ...dataRows], {
      x: LAYOUT.MARGIN_X,
      y: LAYOUT.CONTENT_TOP + 0.1,
      w: LAYOUT.CONTENT_W,
      h: 3.5,
      fontSize: SIZE.SMALL,
      fontFace: FONTS.BODY,
      border: { pt: 0.5, color: COLORS.BORDER },
      colW: [2.2, 2.2, 2.2, 2.2],
      rowH: [0.5, 0.5, 0.5, 0.5, 0.5, 0.5],
      align: "left",
      valign: "middle",
    });

    addFooterNote(s, "XXXX 出典・注釈テキスト");
  }

  // =======================================================================
  // SLIDE 17: Blue accent section
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    s.background = { color: COLORS.BLUE };

    s.addText("XXXX キーメッセージ", {
      x: 1.2,
      y: 1.4,
      w: 7.6,
      h: 1.2,
      fontSize: SIZE.TITLE + 6,
      fontFace: FONTS.HEADING,
      color: COLORS.WHITE,
      bold: true,
      valign: "middle",
      margin: 0,
    });

    s.addText(
      "XXXX 強調したいメッセージやセクションの導入文をここに入力します。青背景のスライドは視覚的なアクセントになります。",
      {
        x: 1.2,
        y: 2.8,
        w: 7.6,
        h: 1.5,
        fontSize: SIZE.BODY + 2,
        fontFace: FONTS.BODY,
        color: COLORS.WHITE,
        valign: "top",
        margin: 0,
      },
    );
  }

  // =======================================================================
  // SLIDE 18: Team Introduction
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX チーム紹介");
    addPageNumber(s, pageNum);

    const members = [
      {
        name: "XXXX 名前1",
        role: "XXXX 役職・部署",
        desc: "XXXX 担当領域や一言紹介",
      },
      {
        name: "XXXX 名前2",
        role: "XXXX 役職・部署",
        desc: "XXXX 担当領域や一言紹介",
      },
      {
        name: "XXXX 名前3",
        role: "XXXX 役職・部署",
        desc: "XXXX 担当領域や一言紹介",
      },
      {
        name: "XXXX 名前4",
        role: "XXXX 役職・部署",
        desc: "XXXX 担当領域や一言紹介",
      },
    ];

    const cardW = 2.0,
      gap = 0.27;
    const totalW = cardW * 4 + gap * 3;
    const startX = (SLIDE.W - totalW) / 2;
    const baseY = LAYOUT.CONTENT_TOP + 0.15;
    const cardH = 3.8;
    const photoSize = 1.2;

    for (let i = 0; i < members.length; i++) {
      const x = startX + i * (cardW + gap);

      // Card background
      s.addShape(SHAPES.ROUNDED_RECTANGLE, {
        x,
        y: baseY,
        w: cardW,
        h: cardH,
        rectRadius: 0.1,
        fill: { color: COLORS.BG_ACCENT },
        shadow: makeCardShadow(),
      });

      // Photo placeholder circle
      const circleX = x + (cardW - photoSize) / 2;
      const circleY = baseY + 0.35;
      s.addShape(SHAPES.OVAL, {
        x: circleX,
        y: circleY,
        w: photoSize,
        h: photoSize,
        fill: { color: COLORS.BORDER },
      });

      // User icon inside circle
      const iconSize = 0.5;
      s.addImage({
        data: icons["user"],
        x: circleX + (photoSize - iconSize) / 2,
        y: circleY + (photoSize - iconSize) / 2,
        w: iconSize,
        h: iconSize,
      });

      // Name
      s.addText(members[i].name, {
        x,
        y: circleY + photoSize + 0.25,
        w: cardW,
        h: 0.45,
        fontSize: SIZE.BODY,
        fontFace: FONTS.HEADING,
        color: COLORS.GRAY_DARK,
        bold: true,
        align: "center",
        valign: "middle",
        margin: 0,
      });

      // Role
      s.addText(members[i].role, {
        x,
        y: circleY + photoSize + 0.7,
        w: cardW,
        h: 0.35,
        fontSize: SIZE.SMALL,
        fontFace: FONTS.BODY,
        color: COLORS.BLUE,
        align: "center",
        valign: "middle",
        margin: 0,
      });

      // Description
      s.addText(members[i].desc, {
        x: x + 0.15,
        y: circleY + photoSize + 1.1,
        w: cardW - 0.3,
        h: 0.8,
        fontSize: SIZE.CAPTION,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_LIGHT,
        align: "center",
        valign: "top",
        margin: 0,
      });
    }
  }

  // =======================================================================
  // SLIDE 19: Thank You / Closing
  // =======================================================================
  {
    const s = pres.addSlide();
    s.background = { color: COLORS.BG_ACCENT };

    // Company name (text-based logo)
    s.addText("株式会社 NEXASPARK", {
      x: 1,
      y: 1.2,
      w: 8,
      h: 0.6,
      fontSize: 20,
      fontFace: "Meiryo UI",
      color: COLORS.GRAY_LIGHTEST,
      bold: true,
      align: "center",
      valign: "middle",
    });

    s.addText("XXXX ご清聴ありがとうございました", {
      x: 0.5,
      y: 2.1,
      w: 9,
      h: 1.0,
      fontSize: SIZE.TITLE,
      fontFace: FONTS.HEADING,
      color: COLORS.GRAY_DARK,
      bold: true,
      align: "center",
      valign: "middle",
    });

    s.addShape(SHAPES.LINE, {
      x: 3.5,
      y: 3.3,
      w: 3,
      h: 0,
      line: { color: COLORS.BORDER, width: 1 },
    });

    s.addText(
      [
        { text: "XXXX お名前", options: { bold: true, breakLine: true } },
        { text: "XXXX 部署・所属", options: { breakLine: true } },
        { text: "XXXX email@example.com", options: {} },
      ],
      {
        x: 2,
        y: 3.5,
        w: 6,
        h: 1.4,
        fontSize: SIZE.BODY,
        fontFace: FONTS.BODY,
        color: COLORS.GRAY_LIGHT,
        align: "center",
        paraSpaceAfter: 4,
      },
    );
  }

  // =======================================================================
  // Write file
  // =======================================================================
  const outputPath = path.join(__dirname, "template.pptx");
  await pres.writeFile({ fileName: outputPath });
  console.log("Template generated:", outputPath);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
