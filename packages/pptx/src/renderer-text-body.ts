// PPTX text-body painting: bullets and auto-numbering, stacked text, and the
// horizontal text-body renderer. Split from renderer.ts (#1857).

import {
  FONT_BASELINE_SENTINEL,
  FONT_SPACE_SENTINEL,
} from '@silurus/ooxml-core/internal/font-measurement-sentinels';
import type { TextBody, Paragraph } from './types';
import { asBullet } from './types';
import {
  resolveFill as resolveFillCore,
  EMU_PER_PT as PT_TO_EMU,
  peekCachedBitmapByPath,
  symbolFontToUnicode,
  hasTextWarp,
} from '@silurus/ooxml-core';
import { segmentsHaveRtl, computeLineVisualOrder, type LineVisualOrder } from './bidi-line';
import { justifyLine, type Justified } from './text-justify';
import { justifiedPiecePositions } from '@silurus/ooxml-core';
import { resolveTabWidths } from './tab-layout.js';
import {
  powerPointCompatOffNaturalLine,
  powerPointExactLinePoints,
  powerPointNaturalLine,
  type PowerPointFaceMetrics,
  POWERPOINT_FONT_ALGN_UNIT_PT,
  powerPointFontAlgnOffset,
  powerPointFontAlgnReference,
} from './powerpoint-line-metrics.js';
import { drawEaVertRun } from './vertical-text.js';
import { renderStackedText, type StackedParagraphInput } from './stacked-text.js';
import {
  drawingMlLineHeight,
  drawingMlSpacedLineBox,
  drawingMlLineX,
  drawingMlLineShouldJustify,
  drawingMlBlockTop,
  drawingMlTextRect,
  type DrawingMlLineSpacing,
} from '@silurus/ooxml-core/internal/drawingml-text';
import {
  emuToPx,
  hexToRgba,
  type RenderContext,
  type TextRunCallback,
  type FetchImage,
} from './renderer-context';
import {
  layoutParagraph,
  codePointCount,
  resolveSegmentTextPaint,
  buildFont,
  normalizeFontFamily,
  hasNamedFontFamily,
  pictureBulletSizePt,
  paragraphInputRuns,
  lineMetricFor,
  naturalWidthExceedsBbox,
  paragraphHasBullet,
  firstLineIndentPxFor,
  measureTextAdvance,
  hasNativeLetterSpacing,
  applyTextRunReflection,
  type LayoutSegment,
  type LayoutLine,
} from './renderer-text-layout';
import {
  paintRunUnderline,
  paintRunStrike,
  paintEquation,
  paintHighlight,
  paintRunDecorations,
} from './renderer-text-paint';
import { renderWarpedText } from './renderer-warp-text';

/**
 * Format an autoNum bullet label from a counter value and OOXML numType.
 * Spec: ECMA-376 §20.1.10.61 (ST_TextAutonumberScheme).
 *
 * Supports the symmetric Plain / Period / ParenR / ParenBoth variants for
 * the four core scripts (arabic, alphaLc/Uc, romanLc/Uc) plus arabicDb
 * (full-width Arabic digits). CJK / Thai / Hebrew / Arabic-Abjad schemes
 * are intentionally not handled here — they fall through to the default
 * `arabicPeriod` rendering rather than emitting a wrong-looking glyph.
 */
function formatAutoNum(counter: number, numType: string): string {
  const arabic = `${counter}`;
  const alphaLc = counter >= 1 && counter <= 26
    ? String.fromCharCode(96 + counter)
    : arabic;
  const alphaUc = counter >= 1 && counter <= 26
    ? String.fromCharCode(64 + counter)
    : arabic;
  const romanLc = toRoman(counter).toLowerCase();
  const romanUc = toRoman(counter);
  // ECMA-376 §20.1.10.61 lists `arabicDb*` as full-width Arabic digits
  // (U+FF10–U+FF19), used in East Asian numbered lists.
  const arabicDb = arabic.replace(/[0-9]/g, (d) =>
    String.fromCharCode(0xff10 + (d.charCodeAt(0) - 0x30)),
  );

  switch (numType) {
    case 'arabicPlain':       return arabic;
    case 'arabicPeriod':      return `${arabic}.`;
    case 'arabicParenR':      return `${arabic})`;
    case 'arabicParenBoth':   return `(${arabic})`;
    case 'arabicDbPlain':     return arabicDb;
    case 'arabicDbPeriod':    return `${arabicDb}.`;
    case 'alphaLcPlain':      return alphaLc;
    case 'alphaLcPeriod':     return `${alphaLc}.`;
    case 'alphaLcParenR':     return `${alphaLc})`;
    case 'alphaLcParenBoth':  return `(${alphaLc})`;
    case 'alphaUcPlain':      return alphaUc;
    case 'alphaUcPeriod':     return `${alphaUc}.`;
    case 'alphaUcParenR':     return `${alphaUc})`;
    case 'alphaUcParenBoth':  return `(${alphaUc})`;
    case 'romanLcPlain':      return romanLc;
    case 'romanLcPeriod':     return `${romanLc}.`;
    case 'romanLcParenR':     return `${romanLc})`;
    case 'romanLcParenBoth':  return `(${romanLc})`;
    case 'romanUcPlain':      return romanUc;
    case 'romanUcPeriod':     return `${romanUc}.`;
    case 'romanUcParenR':     return `${romanUc})`;
    case 'romanUcParenBoth':  return `(${romanUc})`;
    default:                  return `${arabic}.`;
  }
}

function toRoman(n: number): string {
  const vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
  const syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
  let result = '';
  for (let i = 0; i < vals.length; i++) {
    while (n >= vals[i]) { result += syms[i]; n -= vals[i]; }
  }
  return result;
}

/**
 * True when a paragraph has renderable content — at least one non-empty text
 * run, or any equation run. Line breaks alone (a paragraph the user created by
 * pressing Enter on an empty line) do NOT count as content.
 *
 * PowerPoint does not draw a bullet/number marker on an empty paragraph and
 * does not advance an autoNum counter for it (a blank line between "1." and
 * "2." stays unnumbered and the sequence continues). ECMA-376 §21.1.2.4.x
 * (CT_Bullet / CT_TextCharBullet / CT_TextAutoNumberBullet) gives no numeric
 * rule for this, so PowerPoint's behaviour is the reference: suppress the
 * marker whenever the paragraph carries no real content. This is the plain
 * "is there content?" test, not a sample-specific heuristic.
 */
export function paragraphHasRenderableContent(para: Paragraph): boolean {
  for (const r of para.runs) {
    if (r.type === 'text' && r.text !== '') return true;
    if (r.type === 'math') return true;
  }
  return false;
}

/**
 * Resolve the bullet/number marker *label string* for a paragraph, mutating
 * the per-level autoNum counter map as a side effect. Returns '' (= no marker)
 * for `none`/`inherit` bullets and for empty paragraphs.
 *
 * Empty-paragraph handling (PowerPoint reference behaviour, see
 * `paragraphHasRenderableContent`): when the paragraph has no renderable
 * content we draw no marker AND do not advance the autoNum counter, so a blank
 * line between numbered items keeps the sequence going (… "1." / "" / "2." …).
 *
 * The counter-reset semantics (clear on char bullets / non-list paragraphs)
 * are intentionally kept here so the same map walk the renderer relied on is
 * exercised by the unit tests. Font and colour resolution stay in the renderer
 * because they depend on canvas/theme context.
 */
export function resolveBulletLabel(
  para: Paragraph,
  autoNumCounters: Map<number, number>,
): string {
  const hasContent = paragraphHasRenderableContent(para);
  if (para.bullet.type === 'char') {
    // Reset counters when switching to char bullets.
    autoNumCounters.clear();
    return hasContent ? symbolFontToUnicode(para.bullet.char, para.bullet.fontFamily ?? null) : '';
  }
  if (para.bullet.type === 'autoNum') {
    if (!hasContent) return '';
    const lvl = para.lvl;
    if (!autoNumCounters.has(lvl)) {
      autoNumCounters.set(lvl, para.bullet.startAt ?? 1);
    } else {
      autoNumCounters.set(lvl, autoNumCounters.get(lvl)! + 1);
    }
    return formatAutoNum(autoNumCounters.get(lvl)!, para.bullet.numType);
  }
  // none / inherit — not a list paragraph; reset autoNum counters.
  autoNumCounters.clear();
  return '';
}

/** Shared marker content, font, color and size resolution (§21.1.2.4).
 * Layout owns placement; the marker's authored properties do not depend on
 * whether its paragraph advances horizontally or down a stacked column. */
function paragraphBullet(body: TextBody, para: Paragraph, paraDefaultColor: string,
  scale: number, fontScale: number, rc: RenderContext, autoNumCounters: Map<number, number>) {
  // Per ECMA-376 §21.1.2.4.13: when no buSz* is declared, the bullet takes
  // the first run's font size. Using paraDefaultFontSizePx here (the layout
  // lvl1pPr defRPr fallback, typically 18pt) oversizes the bullet so a
  // hanging indent calibrated against the run (12pt) can't contain it —
  // that's why the em-dash was overlapping the text.
  const bulletBaseSizePt = (() => {
    for (const r of para.runs) {
      if (r.type === 'text' && r.fontSize != null) return r.fontSize;
    }
    return para.defFontSize ?? body.defaultFontSize ?? 18;
  })();
  const bulletBaseSizePx = bulletBaseSizePt * PT_TO_EMU * scale * fontScale;

  // ECMA-376 §21.1.2.4.4 (CT_TextCharBullet) / §21.1.2.4.10 (buClrTx): when
  // no explicit `<a:buClr>` is present, the bullet inherits the *first run*'s
  // color, NOT the shape-level default text color. The two diverge on
  // templates where `<p:style><a:fontRef>` resolves to white (lt1) — runs
  // override that with their own `<a:rPr><a:solidFill>`, but bullets without
  // a buClr would otherwise pick up the white default and become invisible
  // (the slide-13 sample-2 regression).
  const firstRunColorHex = (() => {
    for (const r of para.runs) {
      if (r.type === 'text' && r.color) return r.color;
    }
    return null;
  })();
  const bulletInheritedColor = firstRunColorHex
    ? hexToRgba(firstRunColorHex)
    : paraDefaultColor;
  const firstRunFontFamily = (() => {
    for (const r of para.runs) {
      if (r.type === 'text' && r.fontFamily) return r.fontFamily;
    }
    return para.defFontFamily ?? null;
  })();

  let bulletLabel  = '';
  let bulletSizePx = bulletBaseSizePx;
  let bulletFamily = 'sans-serif';
  let bulletFont   = buildFont(false, false, bulletSizePx, bulletFamily, rc);
  let bulletColor  = bulletInheritedColor;
  let bulletFollowsText = false;
  // Picture bullet (`<a:buBlip>`, §21.1.2.4.2). Resolved to its image + drawn
  // size below; stays null for char/number/none bullets.
  let bulletImage: { imagePath: string; mimeType: string; sizePx: number } | null = null;

  // Resolve the marker label and advance the autoNum counter. Empty
  // paragraphs (Enter on a blank line) draw no marker and do not advance the
  // counter — PowerPoint keeps a blank line between numbered items unnumbered
  // while continuing the sequence. ECMA-376 §21.1.2.4.x gives no numeric
  // rule, so PowerPoint's behaviour is the reference (see resolveBulletLabel).
  bulletLabel = resolveBulletLabel(para, autoNumCounters);

  // The parser may emit the picture-bullet variant (`type: "blip"`), which the
  // shared core `Bullet` type doesn't list — narrow once via asBullet.
  const bullet = asBullet(para.bullet);
  if (bullet.type === 'char') {
    const b = bullet;
    // ECMA-376 §21.1.2.4.10 (buSzPts): an ABSOLUTE point size sizes the marker
    // independent of the run — convert pt → px like any run size
    // (PT_TO_EMU·scale·fontScale). It wins over §21.1.2.4.9 (buSzPct, percent of
    // the run size); with neither, the marker takes the run size (§21.1.2.4.13).
    const bSizePx = b.sizePts != null
      ? b.sizePts * PT_TO_EMU * scale * fontScale
      : b.sizePct != null
        ? bulletBaseSizePx * (b.sizePct / 100)
        : bulletBaseSizePx;
    // If the char was mapped to a Unicode symbol, use sans-serif for reliable rendering.
    // Otherwise use the specified font (e.g. Wingdings on systems that have it).
    const convertedFamily = bulletLabel !== b.char ? 'sans-serif' : normalizeFontFamily(b.fontFamily ?? null, rc);
    bulletSizePx = bSizePx;
    bulletFamily = convertedFamily;
    bulletFont  = buildFont(false, false, bSizePx, convertedFamily, rc, bulletLabel,
      hasNamedFontFamily(b.fontFamily));
    bulletColor = b.color ? hexToRgba(b.color) : bulletInheritedColor;
    bulletFollowsText = !b.color;
  } else if (bullet.type === 'autoNum') {
    const b = bullet;
    const bSizePx = b.sizePts != null
      ? b.sizePts * PT_TO_EMU * scale * fontScale
      : b.sizePct != null
        ? bulletBaseSizePx * (b.sizePct / 100)
        : bulletBaseSizePx;
    bulletSizePx = bSizePx;
    bulletFamily = normalizeFontFamily(b.fontFamily ?? firstRunFontFamily, rc);
    bulletFont = buildFont(
      false,
      false,
      bSizePx,
      normalizeFontFamily(b.fontFamily ?? firstRunFontFamily, rc),
      rc,
      bulletLabel,
      hasNamedFontFamily(b.fontFamily ?? firstRunFontFamily),
    );
    // ECMA-376 §21.1.2.4.4 (buClr): an explicit `<a:buClr>` colours the
    // auto-number marker, mirroring the char-bullet branch above. Only when it
    // is absent does the marker fall back to the buClrTx default
    // (§21.1.2.4.5 — the inherited first-run colour).
    bulletColor = bullet.color ? hexToRgba(bullet.color) : bulletInheritedColor;
    bulletFollowsText = !bullet.color;
  } else if (bullet.type === 'blip') {
    // ECMA-376 §21.1.2.4.2 picture bullet. Its height follows the text's em
    // box, scaled by `<a:buSzPct>`, while paint preserves the source aspect
    // (§21.1.2.4.9; default 100%). It's not a glyph, so there is no label —
    // an empty paragraph still draws no marker (bulletLabel stays '' and the
    // draw site gates the image on the first line having content).
    // §21.1.2.4.10 (buSzPts): an absolute point size overrides the §21.1.2.4.9
    // percentage, matching the char-bullet branch above.
    const b = bullet;
    const sizePx = pictureBulletSizePt(body, para, b) * PT_TO_EMU * scale * fontScale;
    bulletImage = { imagePath: b.imagePath, mimeType: b.mimeType, sizePx };
  }

  return { bulletLabel, bulletFont, bulletSizePx, bulletFamily, bulletColor, bulletFollowsText, bulletImage };
}

/**
 * PowerPoint input for stacked vertical text. Paragraph defaults, the stored
 * normAutofit fontScale / lnSpcReduction and paragraph spacing follow the
 * horizontal path; spAutoFit keeps the stored geometry (PowerPoint does not
 * refit an unedited shape). Marker properties share the horizontal resolver;
 * margins and indents are measured in the column reading direction.
 */
function renderStackedTextBody(
  ctx: CanvasRenderingContext2D,
  body: TextBody,
  vert: 'wordArtVert' | 'wordArtVertRtl',
  bx: number,
  by: number,
  bw: number,
  bh: number,
  scale: number,
  bodyDefaultColor: string,
  shapeRotation: number,
  slideNumber: number | undefined,
  rc: RenderContext,
  onTextRun?: TextRunCallback,
  fetchImage?: FetchImage,
): void {
  const fontScale = body.autoFit === 'norm' && body.fontScale != null && body.fontScale > 0 && body.fontScale < 1
    ? body.fontScale : 1;
  const pxPerPt = PT_TO_EMU * scale;
  const inner = drawingMlTextRect(bw, bh, { lIns: body.lIns, rIns: body.rIns, tIns: body.tIns, bIns: body.bIns }, scale);
  const bodyDefaultFontSizePx = (body.defaultFontSize ?? 18) * pxPerPt * fontScale;
  const bodyDefaultBold = body.defaultBold ?? false;
  const bodyDefaultItalic = body.defaultItalic ?? false;
  let sameStyle: ((a: LayoutSegment, b: LayoutSegment) => boolean) | undefined;
  const autoNumCounters = new Map<number, number>();
  const markers: ReturnType<typeof paragraphBullet>[] = [];
  const paragraphs: StackedParagraphInput<LayoutSegment>[] = body.paragraphs.map((para) => {
    const sizePx = para.defFontSize != null ? para.defFontSize * pxPerPt * fontScale : bodyDefaultFontSizePx;
    const color = para.defColor ? hexToRgba(para.defColor) : bodyDefaultColor;
    const marker = paragraphBullet(body, para, color, scale, fontScale, rc, autoNumCounters);
    markers.push(marker);
    const built = paragraphInputRuns(para, sizePx, color, scale, bodyDefaultBold, bodyDefaultItalic, fontScale,
      slideNumber, rc, false, ctx);
    sameStyle ??= built.sameStyle;
    const firstText = built.input.find((item) => item.type === 'text');
    const family = normalizeFontFamily(para.defFontFamily ?? null, rc);
    const bold = para.defBold ?? bodyDefaultBold;
    const italic = para.defItalic ?? bodyDefaultItalic;
    const markStyle: LayoutSegment = firstText?.type === 'text' ? firstText.style : {
      text: '', font: buildFont(bold, italic, sizePx, family, rc, ''), sizePx, color,
      underline: false, strikethrough: false, lineMetric: lineMetricFor(family, bold, italic, rc), faceFamily: family,
    };
    // §21.1.2.2.2: the insertion position has its own run properties. In
    // particular an explicit noStrike must stop inherited final-mark strokes.
    const endRun = para.endRunProperties;
    const endInput = endRun ? paragraphInputRuns({ ...para, runs: [{ ...endRun, text: ' ' }] },
      sizePx, color, scale, bodyDefaultBold, bodyDefaultItalic, fontScale, slideNumber, rc, false, ctx)
      .input.find(item => item.type === 'text') : undefined;
    const spacing = (pts: number | null, pct: number | undefined): DrawingMlLineSpacing =>
      pct != null ? { type: 'pct', val: pct }
        : pts != null ? { type: 'pts', val: (pts / 100) * fontScale } : undefined;
    return {
      runs: built.input,
      marginStart: emuToPx(para.marL, scale),
      marginEnd: emuToPx(para.marR, scale),
      firstLineIndent: emuToPx(para.indent, scale),
      marker: !paragraphHasRenderableContent(para) ? undefined : marker.bulletLabel ? {
        text: marker.bulletLabel,
        style: { text: marker.bulletLabel, font: marker.bulletFont, sizePx: marker.bulletSizePx,
          color: marker.bulletColor, underline: false, strikethrough: false,
          faceFamily: marker.bulletFamily, lineMetric: lineMetricFor(marker.bulletFamily, false, false, rc) },
      } : marker.bulletImage ? {
        advance: marker.bulletImage.sizePx, thickness: marker.bulletImage.sizePx,
      } : undefined,
      alignment: para.alignment,
      lineSpacing: para.spaceLine?.type === 'pts'
        ? { type: 'pts', val: powerPointExactLinePoints(para.spaceLine.val) } : para.spaceLine,
      spaceBefore: spacing(para.spaceBefore, para.spaceBeforePct),
      spaceAfter: spacing(para.spaceAfter, para.spaceAfterPct),
      markStyle,
      endMarkStyle: endInput?.type === 'text' ? endInput.style : undefined,
      eastAsianLineBreak: para.eaLnBrk !== false,
    };
  });
  const runs = renderStackedText(ctx, {
    vert,
    rect: { left: bx + inner.left, top: by + inner.top, width: inner.width, height: inner.height },
    anchor: body.verticalAnchor,
    anchorCtr: body.anchorCtr === true,
    wrap: body.wrap !== 'none',
    spcFirstLastPara: body.spcFirstLastPara === true,
    lnSpcReduction: body.autoFit === 'norm' ? body.lnSpcReduction ?? 0 : 0,
    pxPerPt,
    paragraphs,
    sameStyle,
    drawMarker: (index, axisX, top) => {
      const para = body.paragraphs[index];
      if (!paragraphHasRenderableContent(para)) return;
      const marker = markers[index];
      const y = top;
      // §21.1.2.4.2: keep a picture marker per paragraph, sized by the shared
      // buSzPct/buSzPts resolver. Electronic PDF export has shown a single
      // image at the box upper-left for stacked text, but that observation is
      // not isolated across image classes/export paths: no compatibility rule.
      ctx.save();
      const image = marker.bulletImage;
      const bitmap = image ? rc.pictureBulletImages?.has(image.imagePath)
        ? rc.pictureBulletImages.get(image.imagePath) : fetchImage
          ? peekCachedBitmapByPath(image.imagePath, fetchImage) : null : null;
      if (image && bitmap) {
        const h = image.sizePx;
        const w = bitmap.height > 0 ? h * bitmap.width / bitmap.height : h;
        ctx.drawImage(bitmap, axisX - w / 2, y, w, h);
      }
      ctx.restore();
    },
    markerPaint: (index, axisX, baseline) => {
      const marker = markers[index];
      const first = paragraphs[index].runs.find(run => run.type === 'text');
      return marker.bulletFollowsText && first?.type === 'text' && first.style.patternFill
        ? resolveSegmentTextPaint(ctx, first.style, axisX, baseline, scale) : marker.bulletColor;
    },
    drawUnderline: (style, width, columnSizePx) => {
      paintRunUnderline(ctx, style, 0, 0, width, style.color, scale, rc.dpr, columnSizePx);
    },
    sameUnderline: (a, b) => a.underline === b.underline && a.underlineStyle === b.underlineStyle
      && a.underlineColor === b.underlineColor && a.underlineFill === b.underlineFill
      && a.underlineLine === b.underlineLine && a.underlineLineNoFill === b.underlineLineNoFill
      && a.color === b.color && a.noFill === b.noFill,
    drawStrike: (style, x, baseline, width) => {
      paintRunStrike(ctx, style, x, baseline, width, style.color, rc.dpr);
    },
    drawObject: (style, x, y) => {
      paintEquation(ctx, style, x, y);
    },
  });
  if (!onTextRun) return;
  for (const run of runs) {
    onTextRun({
      text: run.text, inShapeX: run.x - bx, inShapeY: run.y - by, w: run.w, h: run.h,
      fontSize: run.fontSize, font: run.font, shapeX: bx, shapeY: by, shapeW: bw, shapeH: bh,
      rotation: shapeRotation,
      ...(run.hyperlink ? { hyperlink: run.hyperlink } : {}),
    });
  }
}

// Exported (like `layoutParagraph` / `paintHighlight`) so the picture-bullet
// draw path can be unit-tested against a mock 2D context without standing up a
// full canvas. Not re-exported from index.ts — module-internal otherwise.
export function renderTextBody(
  ctx: CanvasRenderingContext2D,
  body: TextBody,
  bx: number,
  by: number,
  bw: number,
  bh: number,
  scale: number,
  shapeDefaultTextColor: string | null = null,
  shapeRotation = 0,
  shapeFlipH = false,
  shapeFlipV = false,
  themeDefaultColor = '#000000',
  slideNumber?: number,
  rc: RenderContext = { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
  onTextRun?: TextRunCallback,
  measureOnly = false,
  fetchImage?: FetchImage,
  // ECMA-376 §20.1.10.83 `eaVert` ("East Asian Vertical"): CJK glyphs stand
  // UPRIGHT while Latin rotates. Set true ONLY when re-entering horizontally
  // inside the +90°-rotated `eaVert` frame (below), so the run-draw loop applies
  // per-glyph UAX#50 orientation via `drawEaVertRun`. False for `horz` and for
  // `vert`/`vert270` (whose spec meaning IS to rotate every glyph), keeping those
  // paths byte-identical.
  eaVertUpright = false,
  // A zero-height table row asks PowerPoint to derive its height from the
  // natural line box. Office treats a positive a:tr@h as an authored minimum
  // ([MS-OE376] §2.1.1347): use the glyph-size box when checking whether content
  // actually exceeds that minimum, so implicit leading alone does not enlarge
  // an already-sufficient row.
  measureNaturalLineSpacing = measureOnly,
): number | void {
  // Stacked vertical text (wordArtVert / wordArtVertRtl): upright glyphs one
  // above another in columns; see stacked-text.ts and core layoutStackedText.
  if (body.vert === 'wordArtVert' || body.vert === 'wordArtVertRtl') {
    // Like the rotated modes, a table row measures a vertical body by its box.
    if (measureOnly) return bw;
    renderStackedTextBody(ctx, body, body.vert, bx, by, bw, bh, scale, shapeDefaultTextColor ?? themeDefaultColor,
      shapeRotation, slideNumber, rc, onTextRun, fetchImage);
    return;
  }

  // Vertical text: rotate rendering context so text flows top-to-bottom.
  // "vert" and "eaVert" both approximate to 90° clockwise rotation.
  // "vert270" rotates 270° (= 90° counterclockwise).
  const isVert    = body.vert === 'vert' || body.vert === 'eaVert';
  const isVert270 = body.vert === 'vert270';

  if (isVert || isVert270) {
    // Set up a rotated coordinate space:
    // Centre of the bounding box remains fixed; swap w and h for the text layout.
    const cx = bx + bw / 2;
    const cy = by + bh / 2;
    const vertRot = isVert ? 90 : -90;

    // Wrap onTextRun to convert from the rotated sub-frame back to the original
    // shape frame so that _buildTextLayer can apply a single CSS rotation.
    //
    // In the recursive call the origin is (-bh/2, -bw/2) with axes (bh, bw).
    // For a run at canvas (penX, cursorY) in that sub-frame:
    //   inShapeX_rec = penX + bh/2,  inShapeY_rec = cursorY + bw/2
    //
    // We need the position in the *original* shape frame so that after
    // CSS rotate(shapeRotation + vertRot) the span lands on the same pixel:
    //   inShapeX_span = penX + bw/2 = inShapeX_rec - bh/2 + bw/2
    //   inShapeY_span = cursorY + bh/2 = inShapeY_rec - bw/2 + bh/2
    const wrappedOnTextRun: TextRunCallback | undefined = onTextRun
      ? (run) => onTextRun({
          ...run,
          inShapeX: run.inShapeX - bh / 2 + bw / 2,
          inShapeY: run.inShapeY - bw / 2 + bh / 2,
          shapeX: bx,
          shapeY: by,
          shapeW: bw,
          shapeH: bh,
          rotation: shapeRotation,
          textBodyRotation: vertRot,
        })
      : undefined;

    if (measureOnly) {
      // The rotated sub-frame's content height runs along the original width
      // axis; for a table row the vertical extent is the original box width.
      return bw;
    }
    ctx.save();
    ctx.translate(cx, cy);
    ctx.rotate(isVert270 ? -Math.PI / 2 : Math.PI / 2);
    // After rotation the "width" direction of the new frame is the original height.
    // `eaVert` re-enters with per-glyph uprighting (UAX#50): CJK stands up, Latin
    // stays sideways. `vert`/`vert270` re-enter WITHOUT it — their spec meaning is
    // to rotate every glyph — so those stay byte-identical to before.
    renderTextBody(ctx, { ...body, vert: 'horz' }, -bh / 2, -bw / 2, bh, bw, scale, shapeDefaultTextColor, 0, false, false, themeDefaultColor, slideNumber, rc, wrappedOnTextRun, false, fetchImage, body.vert === 'eaVert');
    ctx.restore();
    return;
  }

  // ── WordArt text warp (ECMA-376 §20.1.9.19, prstTxWarp) ──────────────────
  // When the body carries a known warp preset, each glyph is mapped through the
  // preset's envelope rather than laid out flat. This is a *separate* draw path
  // gated on the warp's presence, so every unwarped body renders exactly as
  // before (byte-identical). measureOnly bodies fall through to the flat pass —
  // a warped shape is fixed-size, not a table cell auto-sizing to its text.
  const warp = body.textWarp;
  if (!measureOnly && warp && hasTextWarp(warp.preset)) {
    renderWarpedText(
      ctx,
      body,
      warp.preset,
      warp.adj ?? [],
      bx,
      by,
      bw,
      bh,
      scale,
      shapeDefaultTextColor ?? themeDefaultColor,
      rc,
    );
    return;
  }

  const textRect = drawingMlTextRect(bw, bh, {
    lIns: body.lIns, rIns: body.rIns, tIns: body.tIns, bIns: body.bIns,
  }, scale);
  const lPad = textRect.left;
  const rPad = emuToPx(body.rIns, scale);
  const tPad = emuToPx(body.tIns, scale);
  const bPad = emuToPx(body.bIns, scale);
  // ECMA-376 §20.1.10.5 spAutoFit: "the shape's height and width should resize
  // to accommodate the text". The bbox is no longer a wrap boundary in that
  // mode — the shape grows to fit. Treating spAutoFit as wrap=none on the
  // horizontal axis keeps mixed-size sequences (e.g. "20代", "YoY+11.9%")
  // on one line, matching PowerPoint. Vertical growth is handled below by
  // the spAutoFit branch in the height calculation.
  // ECMA-376 §20.1.10.5 spAutoFit + §20.1.10.7 wrap interaction:
  // - `wrap="none"`: text never wraps (regardless of autoFit).
  // - `wrap="square"` + `<a:spAutoFit/>`: PowerPoint auto-fits the SHAPE to
  //   the text. If the text's natural single-line width fits the bbox, the
  //   shape stays at its current width and the text doesn't wrap (sample-2
  //   slide-13's "20代" textbox: ~50px text in ~70px bbox → no wrap). If
  //   the natural width exceeds the bbox, text wraps and the shape grows
  //   vertically (sample-2 slide-16's "1Q業績 要因" callouts: a single
  //   long Japanese paragraph that has to wrap within a fixed-width bbox).
  // - `wrap="square"` (default) without spAutoFit: always wrap.
  //
  // The pre-pass below measures the paragraphs' total natural width; if
  // every paragraph fits the bbox without wrapping, we keep the spAutoFit
  // "no-wrap" semantics. Otherwise we wrap normally so the long text
  // doesn't run off the side of the shape.
  const baseDoWrap = body.wrap !== 'none';
  const isSpAutoFit = body.autoFit === 'sp';
  const doWrap = isSpAutoFit
    ? (baseDoWrap && naturalWidthExceedsBbox(ctx, body, bw, lPad, rPad, scale, rc, slideNumber))
    : baseDoWrap;
  // ECMA-376 §20.1.10.34 numCol — distribute paragraphs across N text columns.
  const numCol = Math.max(1, body.numCol ?? 1);
  const spcColPx = emuToPx(body.spcCol ?? 0, scale);

  const bodyDefaultBold   = body.defaultBold   ?? false;
  const bodyDefaultItalic = body.defaultItalic ?? false;
  const bodyDefaultColor = shapeDefaultTextColor ?? themeDefaultColor;

  // ── Pass 1: lay out all paragraphs ──────────────────────────────────────

  interface LineEntry {
    line: LayoutLine;
    linePx: number;       // spacing advancement (lineHeight + spaceAfter for last line)
    lineHeight: number;   // pure line height used for baseline positioning (without spaceAfter)
    baselineLineHeight: number; // authored box used only to seat this line's glyph baseline
    topGapPx: number;     // spaceBefore for first line of paragraph
    textXOffset: number;  // additional X offset for first-line indent (non-bullet)
    bulletLabel: string;  // text to render as bullet ('' = none)
    bulletFont: string;
    bulletColor: string;
    bulletFollowsText: boolean;
    bulletX: number;      // canvas X for bullet
    // Picture bullet (`<a:buBlip>`, §21.1.2.4.2): the resolved image + its
    // authored height in px (scaled by buSzPct); painting preserves the source
    // aspect ratio. null when this paragraph has no picture bullet. Only set on
    // the paragraph's first line.
    bulletImage: { imagePath: string; mimeType: string; sizePx: number } | null;
    textX: number;        // canvas X for text
    textMaxW: number;     // max wrap width
    alignment: string;
    isLastLine: boolean;
    para: Paragraph;
    /** This spAutoFit line replaces an authored-font design floor with metrics
     * from the font Canvas actually resolved. */
    useResolvedFontMetrics: boolean;
    /** PowerPoint's metric ascent of the spaced line (px), when every run of
     * the body has a known face; undefined for the ordinary 0.8 × line model. */
    metricAscent?: number;
    /** The line's natural (unspaced) descent in the metric model. */
    metricNaturalDescent?: number;
  }

  // buildLayout runs Pass 1 at a given font scale (1.0 = normal; <1 = normAutoFit shrink)
  const buildLayout = (fontScale: number, metric: boolean): {
    allLines: LineEntry[];
    totalHeight: number;
    requiredHeight: number;
    metricOk: boolean;
  } => {
  const bodyDefaultFontSizePx = (body.defaultFontSize ?? 18) * PT_TO_EMU * scale * fontScale;
  const allLines: LineEntry[] = [];
  let totalHeight = 0;
  // Visual bounds may be taller than the baseline advance. This is especially
  // important for spAutoFit: the live Canvas font box must enlarge the SHAPE
  // enough to contain the last line, but must not silently become the pitch of
  // every preceding line when a:lnSpc is omitted (#1473).
  let requiredHeight = 0;
  // Every line must resolve PowerPoint's metric split (see
  // powerpoint-line-metrics.ts) for the body to use it; one body never mixes
  // the two line models.
  let metricOk = metric;

  // AutoNum counters per list level
  const autoNumCounters = new Map<number, number>();

  for (let paraIdx = 0; paraIdx < body.paragraphs.length; paraIdx++) {
    const para = body.paragraphs[paraIdx];
    const marLPx   = emuToPx(para.marL,   scale);
    const marRPx   = emuToPx(para.marR,   scale);
    const indentPx = emuToPx(para.indent, scale);

    // Para-level defaults (cascade: para defRPr → body default)
    const paraDefaultFontSizePx = para.defFontSize != null
      ? para.defFontSize * PT_TO_EMU * scale * fontScale : bodyDefaultFontSizePx;
    const paraDefaultColor = para.defColor
      ? hexToRgba(para.defColor) : bodyDefaultColor;

    // Bullet resolution. A picture bullet (`blip`) occupies the gutter exactly
    // like a char/autoNum marker, so it must suppress the first-line hanging
    // indent too — otherwise the first line of a hanging-indent list starts at
    // the bullet's x and renders ON TOP of the picture (cf. the char-bullet em-
    // dash overlap noted below, and docx PR #476).
    const hasBullet = paragraphHasBullet(para);

    const { bulletLabel, bulletFont, bulletColor, bulletFollowsText, bulletImage } =
      paragraphBullet(body, para, paraDefaultColor, scale, fontScale, rc, autoNumCounters);

    // Text start X and wrap width.
    //
    // ECMA-376 §20.1.10.34 numCol — the bbox is split into N equal columns
    // separated by `spcCol` gutters, and paragraphs flow column-by-column.
    // For Pass 1 we lay out at the column width so wrapping & line counts are
    // correct; Pass 2 reuses the same lines and just shifts the X origin per
    // column. textX/bulletX below are relative to column 0; the column shift
    // is added later when we walk `allLines`.
    const colWidth = numCol > 1
      ? (bw - lPad - rPad - (numCol - 1) * spcColPx) / numCol
      : bw - lPad - rPad;
    const textX    = bx + lPad + marLPx;
    // The marker seats at the RAW (signed) indent — a hanging gutter is negative,
    // so this is deliberately NOT routed through firstLineIndentPxFor (which is
    // the non-bullet first-line TEXT shift). Keep them distinct.
    const bulletX  = bx + lPad + marLPx + indentPx;
    const textMaxW = colWidth - marLPx - marRPx;

    const maxW = doWrap ? textMaxW : Infinity;
    // A signed first-line indent changes ONLY the first line's wrap budget
    // (continuation lines keep the full width). firstLineIndentPxFor keeps this
    // in lockstep with the draw-side `textXOffset` (§below) and the
    // naturalWidthExceedsBbox measurement; a bullet's gutter contributes 0.
    const firstLineIndentPx = firstLineIndentPxFor(hasBullet, indentPx);
    const lines = layoutParagraph(ctx, para, maxW, paraDefaultFontSizePx, paraDefaultColor, scale, marLPx, bodyDefaultBold, bodyDefaultItalic, fontScale, slideNumber, rc, firstLineIndentPx);

    // spaceBefore/After are in hundredths of a point → convert to canvas px.
    // Percentage forms (ECMA-376 §21.1.2.3.11 spcPct) are resolved per line
    // below, against the text size of the line the spacing is attached to.
    const spaceBeforePx = para.spaceBefore != null ? (para.spaceBefore / 100) * PT_TO_EMU * scale * fontScale : 0;
    const spaceAfterPx  = para.spaceAfter  != null ? (para.spaceAfter  / 100) * PT_TO_EMU * scale * fontScale : 0;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const isFirst = i === 0;
      const isLast  = i === lines.length - 1;

      // Line height: use the max font size among rendered segments. The layout
      // default (paraDefaultFontSizePx) is used only as a fallback for empty
      // paragraphs — otherwise PowerPoint slide-layouts with placeholder
      // defaults like `defRPr sz="30000"` (300pt prompt-text marker) would
      // inflate lineHeight and push real 24pt runs far below the anchor.
      let maxSizePx = 0;
      const metricRuns: { sizePx: number; face: PowerPointFaceMetrics }[] = [];
      // Measure the fonts Canvas actually resolved (including browser
      // substitutions). A tall live font box is retained for containment under
      // spAutoFit, but PowerPoint does not repeat that box as the implicit
      // baseline pitch when `<a:lnSpc>` is omitted.
      let resolvedFontLine = 0;
      for (const seg of line.segments) {
        // For an equation, the line must be at least as tall as its own font
        // size (so a short label like "y"/"p"/"z" gets the normal font-ascent
        // baseline and sits where PowerPoint puts it — not floated up by its
        // tight ink box), AND tall enough for its visual box (fractions, norms,
        // big operators) to fit inside the 1.2× leading.
        const effSize = seg.math
          ? Math.max(seg.sizePx, (seg.math.ascent + seg.math.descent) / 1.2)
          : seg.sizePx;
        if (effSize > maxSizePx) maxSizePx = effSize;
        if (seg.math) metricOk = false;
        else if (!seg.isTab) {
          if (seg.lineMetric === undefined) metricOk = false;
          else metricRuns.push({ sizePx: seg.sizePx, face: seg.lineMetric });
          // A run's latin face sizes its line even where an East Asian or
          // symbol segment draws no latin glyph; an unused ea or cs face does
          // not (#1610 supplements 2 and 3).
          if (seg.lineMetricLatin === null) metricOk = false;
          else if (seg.lineMetricLatin !== undefined
            && seg.lineMetricLatin !== seg.lineMetric) {
            metricRuns.push({ sizePx: seg.sizePx, face: seg.lineMetricLatin });
          }
        }
        if (!seg.math) {
          if (isSpAutoFit) {
            ctx.font = seg.font;
            const metrics = ctx.measureText(seg.text || FONT_BASELINE_SENTINEL);
            const fontAscent = metrics.fontBoundingBoxAscent ?? 0;
            const fontDescent = metrics.fontBoundingBoxDescent ?? 0;
            const resolved = fontAscent + fontDescent;
            if (resolved > resolvedFontLine) resolvedFontLine = resolved;
          }
        }
      }
      // Break / end-of-paragraph marks after text: their face at the size of
      // the run they follow (layoutParagraph `followingMark`). A mark whose
      // face has no reference metrics (a generic or embedded family) draws
      // nothing, so it is left out rather than taking the whole body off the
      // metric model.
      for (const mark of line.metricMarks ?? []) {
        if (mark.lineMetric !== undefined) metricRuns.push({ sizePx: mark.sizePx, face: mark.lineMetric });
      }
      if (maxSizePx === 0) maxSizePx = paraDefaultFontSizePx;
      const textMaxSizePx = maxSizePx;
      // Bullet font size also counts
      if (isFirst && bulletLabel) {
        ctx.font = bulletFont;
        const bm = ctx.measureText(FONT_BASELINE_SENTINEL);
        const bSizeApprox = bm.actualBoundingBoxAscent + bm.actualBoundingBoxDescent;
        if (bSizeApprox > maxSizePx) maxSizePx = bSizeApprox;
      }
      // A picture bullet's box also counts toward the line height so a tall
      // bitmap marker isn't clipped by a short first line.
      if (isFirst && bulletImage && bulletImage.sizePx > maxSizePx) {
        maxSizePx = bulletImage.sizePx;
      }
      // A marker taller than the text is outside the #1610 controls.
      if (maxSizePx > textMaxSizePx) metricOk = false;

      // PowerPoint's natural single-line pitch is 120% of the authored text size.
      // An Office-produced boundary deck confirms that this is independent of
      // the chosen font and of spAutoFit: Meiryo and Arial, fixed and spAutoFit
      // shapes all keep the same omitted-lnSpc pitch at a given point size.
      // Percentage spacing scales this renderer's PowerPoint-compatible natural
      // line box (ECMA-376 §21.1.2.2.5/.11 defines the authored percentage;
      // Office output supplies the line-box compatibility behaviour). A positive
      // a:tr@h remains a minimum. A PowerPoint table control with 16pt text,
      // 120% lnSpc and 8.5pt vertical cell margins grows a 36.85pt row;
      // changing only lnSpc to 100% or the margins to zero does not. Explicit
      // percentage spacing therefore consumes the painted natural line box
      // even for one final line. Omitted lnSpc keeps its glyph-box exception
      // in a positive row. Neither path uses a substituted font's design box.
      const naturalSingle = maxSizePx * 1.2;
      const useResolvedFontMetrics = isSpAutoFit && resolvedFontLine > naturalSingle;
      // A live resolved font box describes containment, not baseline advance.
      // Keeping the two values separate prevents a tall Meiryo design box from
      // being repeated between every pair of lines under spAutoFit (#1473).
      const implicitSingle = naturalSingle;
      // Paint and table measurement must agree for explicitly authored
      // percentage spacing. Positive rows with omitted lnSpc retain the
      // existing glyph-box containment rule: the implicit leading is not by
      // itself evidence that PowerPoint grows the authored minimum. An explicit
      // percentage, however, is part of the authored content extent, and every
      // line in a multi-line body consumes the painted line box.
      const paintedLineHeight = drawingMlLineHeight(
        implicitSingle, para.spaceLine, PT_TO_EMU * scale,
      );
      let lineHeight = paintedLineHeight;
      if (measureOnly && !isSpAutoFit && !measureNaturalLineSpacing) {
        if (!para.spaceLine) {
          lineHeight = maxSizePx;
        }
      }
      // PowerPoint retains its established percentage-line advance, but seats
      // glyphs from a tall resolved fallback inside the authored percentage
      // box. Keeping that box separate from `lineHeight` fixes the first-line
      // origin without tightening every continuation line in the paragraph.
      const baselineLineHeight = para.spaceLine?.type === 'pct'
        ? maxSizePx * (para.spaceLine.val / 100000)
        : lineHeight;
      // normAutofit lnSpcReduction (ECMA-376 §21.1.2.1.3): PowerPoint reduces
      // each paragraph's line spacing by this fraction alongside the font
      // shrink. Apply it only when normAutofit stored a value AND the paragraph
      // has PERCENTAGE line spacing — the spec's normative note reads "This
      // attribute applies only to paragraphs with percentage line spacing." pct
      // and the implicit single (= 100 % percentage) qualify; an absolute
      // spcPts height does not.
      if (body.autoFit === 'norm' && body.lnSpcReduction != null && para.spaceLine?.type !== 'pts') {
        lineHeight *= 1 - body.lnSpcReduction;
      }
      // PowerPoint's metric line (#1610, powerpoint-line-metrics.ts): the
      // same 1.2 × size box, split at the baseline by the runs' faces and
      // re-divided for lnSpc by the shared DrawingML rule. An empty line (no
      // glyphs) takes a 0.8 split; its split moves no visible glyph and the
      // next line's top depends only on the line height. spcPts is rounded to
      // whole points first.
      let metricAscent: number | undefined;
      let metricNaturalDescent: number | undefined;
      // A percentage spcBef/spcAft is a fraction of one natural line: 1.2 ×
      // size in the #1610 model, the natural box under compatLnSpc="0".
      let pctSpacingUnit = naturalSingle;
      // An explicit compatLnSpc="0" selects the #1604 natural box
      // (powerPointCompatOffNaturalLine). Its face participation is the
      // #1610 one (the run's own face and its latin face). A line with no
      // glyph run has no measured box in that model, and neither has a face
      // without Excel tables; either keeps the ordinary model for the body.
      const compatOff = body.compatLnSpc === false;
      if (metric && metricOk && compatOff
        && (metricRuns.length === 0 || metricRuns.some((r) => r.face.excel === undefined))) {
        metricOk = false;
      }
      // pPr@fontAlgn t / ctr / b (powerPointFontAlgnReference): the line box
      // is split at its alignment reference and each glyph run is offset from
      // it. A line without runs has nothing to align and keeps the baseline
      // rule.
      const fontAlgn = para.fontAlgn;
      let alignedLine: { ascent: number; descent: number } | undefined;
      if (metric && metricOk && fontAlgn && metricRuns.length > 0) {
        const unitPx = POWERPOINT_FONT_ALGN_UNIT_PT * PT_TO_EMU * scale;
        alignedLine = powerPointFontAlgnReference(fontAlgn, metricRuns, compatOff, unitPx);
        if (!alignedLine) metricOk = false;
        for (const seg of line.segments) {
          if (!alignedLine) break;
          if (seg.isTab || !seg.text) continue;
          const offset = seg.lineMetric
            && powerPointFontAlgnOffset(fontAlgn, { sizePx: seg.sizePx, face: seg.lineMetric }, compatOff, unitPx);
          if (offset === undefined) {
            metricOk = false;
            break;
          }
          seg.fontAlgnOffsetPx = offset;
        }
      }
      if (metric && metricOk && !(measureOnly && !isSpAutoFit && !measureNaturalLineSpacing && !para.spaceLine)) {
        const natural = alignedLine ?? (compatOff
          ? powerPointCompatOffNaturalLine(metricRuns.map((r) => ({ sizePx: r.sizePx, box: r.face.excel! })))
          : metricRuns.length > 0
            ? powerPointNaturalLine(metricRuns.map((r) => ({ sizePx: r.sizePx, share: r.face.share })))
            : { ascent: naturalSingle * 0.8, descent: naturalSingle * 0.2 });
        if (compatOff) pctSpacingUnit = natural.ascent + natural.descent;
        const spacing = para.spaceLine?.type === 'pts'
          ? { type: 'pts' as const, val: powerPointExactLinePoints(para.spaceLine.val) }
          : para.spaceLine;
        const box = drawingMlSpacedLineBox(natural, spacing, PT_TO_EMU * scale,
          body.autoFit === 'norm' && body.lnSpcReduction != null ? body.lnSpcReduction : 0);
        lineHeight = box.ascent + box.descent;
        metricAscent = box.ascent;
        metricNaturalDescent = natural.descent;
      }
      // ECMA-376 §21.1.2.2.9-.10 with §21.1.2.3.11: a percentage spcBef /
      // spcAft is a fraction of the text size, 100000 being one line. It is
      // taken from the first line for space before and from the last line for
      // space after. PowerPoint's unit is one single line of that text: its
      // largest size × 1.2. It does not depend on the paragraph's lnSpc,
      // lnSpcReduction, or a substituted font's design line. This comes from
      // PowerPoint's PDF of a spacing control deck: 100% before and after
      // 40 pt Arial or Meiryo adds 48 pt with lnSpc 100% and with lnSpc 80%.
      // A 20 pt line adds 24 pt, a 20 + 40 pt line adds 48 pt, and a two-line
      // paragraph uses its first line for before and its last line for after.
      // Paint and table measurement use the same base.
      const lineSpaceAfterPx = isLast && para.spaceAfterPct != null
        ? pctSpacingUnit * (para.spaceAfterPct / 100000)
        : spaceAfterPx;
      const lineSpaceBeforePx = isFirst && para.spaceBeforePct != null
        ? pctSpacingUnit * (para.spaceBeforePct / 100000)
        : spaceBeforePx;
      // ECMA-376 §21.1.2.1.1 bodyPr@spcFirstLastPara (default false): the
      // first paragraph's space before and the last paragraph's space after
      // are not respected at the edges of the text body. PowerPoint PDF
      // controls agree for t/ctr/b anchors in both pts and pct: with the
      // attribute absent or 0 the edge spacing moves no glyph, and only 1
      // applies it. A trailing empty paragraph turns the previous spcAft into
      // inter-paragraph spacing, which is respected.
      const respectEdges = body.spcFirstLastPara === true;
      const lastParagraph = paraIdx === body.paragraphs.length - 1;
      const linePx  = lineHeight
        + (isLast && (respectEdges || !lastParagraph) ? lineSpaceAfterPx : 0);
      const topGap  = isFirst && (respectEdges || paraIdx > 0) ? lineSpaceBeforePx : 0;
      // Preserve the signed non-bullet first-line indent in draw, wrapping and
      // spAutoFit measurement alike. Continuation lines remain at marL.
      const textXOffset = isFirst ? firstLineIndentPxFor(hasBullet, indentPx) : 0;

      // Picture bullets, like char/number markers, are drawn only on the
      // paragraph's first line and only when that line carries content (an
      // empty paragraph gets no marker — PowerPoint behaviour, mirroring
      // resolveBulletLabel's empty-paragraph handling).
      const lineHasContent = line.segments.some(
        (s) => (s.text && s.text.length > 0) || s.math != null,
      );
      const entryBulletImage = isFirst && lineHasContent ? bulletImage : null;

      allLines.push({
        line, linePx, lineHeight, baselineLineHeight, topGapPx: topGap,
        textXOffset,
        bulletLabel: isFirst ? bulletLabel : '',
        bulletFont, bulletColor, bulletFollowsText, bulletX,
        bulletImage: entryBulletImage,
        textX, textMaxW,
        alignment: para.alignment,
        isLastLine: isLast,
        para,
        useResolvedFontMetrics: metricAscent === undefined && useResolvedFontMetrics,
        metricAscent,
        metricNaturalDescent,
      });
      const lineTop = totalHeight + topGap;
      totalHeight += linePx + topGap;
      const requiredLineHeight = useResolvedFontMetrics && metricAscent === undefined
        ? Math.max(lineHeight, resolvedFontLine)
        : lineHeight;
      requiredHeight = Math.max(
        requiredHeight,
        totalHeight,
        // The space after actually applied to this line: percentage and
        // spcFirstLastPara edge rules included (linePx above).
        lineTop + requiredLineHeight + (linePx - lineHeight),
      );
    }
  }

  return { allLines, totalHeight, requiredHeight, metricOk };
  }; // end buildLayout

  // Try PowerPoint's metric line model first; a body with any face outside
  // it is laid out again with the ordinary model.
  let useMetricLines = true;
  const layoutAt = (fontScale: number) => {
    let layout = buildLayout(fontScale, useMetricLines);
    if (useMetricLines && !layout.metricOk) {
      useMetricLines = false;
      layout = buildLayout(fontScale, false);
    }
    return layout;
  };
  let { allLines, requiredHeight } = layoutAt(1.0);

  // ── normAutoFit ──────────────────────────────────────────────────────────
  // ECMA-376 §21.1.2.1.3 gives an omitted fontScale the value 100% and an
  // omitted lnSpcReduction the value 0%. PowerPoint did not synthesize a
  // fontScale while opening, exporting, or saving an unchanged file.
  // Mac PowerPoint PDF controls kept 24 pt text at 100% with no stored scale
  // across fitting and overflowing boxes (44–140 pt), 90–120% line spacing,
  // point/percentage paragraph spacing, a trailing paragraph, installed and
  // substituted fonts, and slide-local/layout-inherited placeholders. Stored
  // 50% and 80% scales were honored even when they did not fit the content.
  // Therefore our own text metrics must not manufacture a new scale here.
  if (body.autoFit === 'norm') {
    if (body.fontScale != null && body.fontScale > 0) {
      if (body.fontScale < 1.0) ({ allLines, requiredHeight } = layoutAt(body.fontScale));
    }
  }

  // ── measure-only: return the content height the text body needs ─────────
  // Used by renderTable to grow rows to fit their tallest cell (ECMA-376
  // §21.1.3.18: a:tr@h is a minimum). Returns padding + laid-out text height.
  if (measureOnly) {
    return tPad + requiredHeight + bPad;
  }

  // ── anchor="b" with bh=0: auto-height growing upward from by ────────────
  // When cy=0 and anchor="b", off_y is the bottom anchor; shape grows upward.
  const anchor = body.verticalAnchor ?? 't';
  let effectiveBy = by;
  let effectiveBh: number;
  if (bh === 0 && anchor === 'b') {
    effectiveBh = tPad + requiredHeight + bPad;
    effectiveBy = by - effectiveBh;
  } else {
    // spAutoFit (§21.1.2.1.4) asks the editor to resize the shape to its text,
    // and PowerPoint writes the resized extents when it saves. Rendering and
    // PDF export keep the authored extents (#1610 controls: 60 pt and 300 pt
    // boxes holding 100 pt text, anchor t and ctr, are laid out exactly as
    // with noAutofit; a centred block overflows both edges). Only a zero
    // height, which has no box to anchor in, still takes the content height.
    effectiveBh = body.autoFit === 'sp' && bh === 0
      ? tPad + requiredHeight + bPad
      : bh;
  }

  // ── Vertical anchor ─────────────────────────────────────────────────────
  // PowerPoint anchors the metric block by its last line's natural descent,
  // not by the spaced one: an lnSpc/spcPts change below the last baseline
  // does not move a bottom-anchored block (#1610 controls: two 100 pt lines
  // at 80/90/150 % and 100/140 pt, ctr and b, Arial and Meiryo; normAutofit
  // lnSpcReduction likewise).
  const lastMetric = allLines[allLines.length - 1];
  const anchorHeight = lastMetric?.metricAscent !== undefined
    && lastMetric.metricNaturalDescent !== undefined
    ? requiredHeight - (lastMetric.lineHeight - lastMetric.metricAscent) + lastMetric.metricNaturalDescent
    : requiredHeight;
  let cursorY: number;
  const contentH = Math.max(0, effectiveBh - tPad - bPad);
  if (anchor === 'ctr') {
    cursorY = effectiveBy + drawingMlBlockTop('ctr', {
      left: lPad, top: tPad, width: bw - lPad - rPad, height: contentH,
    }, anchorHeight);
  } else if (anchor === 'b') {
    cursorY = effectiveBy + effectiveBh - bPad - anchorHeight;
  } else {
    cursorY = effectiveBy + tPad;
  }

  // ── Pass 2: render ───────────────────────────────────────────────────────
  ctx.save();
  // penX / baseline are computed manually below, so the canvas text origin
  // must be normalized before fillText() or alignment/anchor math will drift.
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  // ECMA-376 §20.1.2.3.6 (a:bodyPr): PowerPoint does NOT clip text that
  // overflows its shape — the text simply renders past the shape bounds and
  // can overlap with adjacent elements. Our previous behavior clipped, which
  // cropped long text in fixed-height boxes. Only clip when the caller has
  // opted into wrap=none AND a finite x-axis rectangle (rare), which we
  // approximate here by skipping clipping entirely for the default bodyPr.
  // `body.wrap === "none"` means horizontal non-wrap; it doesn't affect
  // clipping per spec either, so we just don't clip.

  // Multi-column flow state. PowerPoint's actual behaviour with `numCol`:
  //
  // - If the total content height fits within a single column, ALL paragraphs
  //   stay in column 0 even when `numCol > 1`. (Sample-2 slide-13's "従来"
  //   box has 4 paragraphs in a `numCol="2"` body; PowerPoint stacks them
  //   vertically because they fit, not 2+2.)
  // - Otherwise, paragraphs are distributed *balanced* (each column gets
  //   roughly ⌈N/numCol⌉ entries) so parallel rows across columns line up.
  //   (Sample-2 slide-13's "新機能" box has 9 paragraphs that overflow a
  //   single column; the right column starts with a blank paragraph so the
  //   blue 利用履歴/価値観/属性詳細 sit at the same y as 年齢/性別/居住地.)
  //
  // A pure overflow-driven advance would either pack column 0 maximally and
  // shift the right column up (breaking row alignment) or never collapse
  // multi-col to single when content fits.
  //
  // We pre-compute, for each entry, which column it belongs to. Column 0
  // starts at the initial cursorY; each subsequent column resets the cursor
  // and shifts X by colWidth + spcCol.
  const colTopY = cursorY;
  const colWidthShift = numCol > 1
    ? (bw - lPad - rPad - (numCol - 1) * spcColPx) / numCol + spcColPx
    : 0;
  const colHeightCapacity = Math.max(0, effectiveBh - tPad - bPad);
  // When `bh === 0` (caller relies on auto-height) the capacity is unbounded
  // and content always fits, so a multi-col directive collapses to single col.
  // `spcAft` is spacing AFTER a paragraph. On the final line there is no
  // following content, so that trailing gap cannot by itself fill the column
  // and force otherwise-fitting text into the next overflow container
  // (ECMA-376 §21.1.2.1.1 bodyPr@numCol). Keep it for drawing/anchoring, but
  // exclude it from the one-column overflow decision. A common boundary case
  // is a heading plus three short bullets whose final 3pt spcAft alone crosses
  // a rounded-rectangle text-rect capacity.
  const lastEntry = allLines[allLines.length - 1];
  const trailingSpaceAfter = lastEntry
    ? Math.max(0, lastEntry.linePx - lastEntry.lineHeight)
    : 0;
  const occupiedHeight = requiredHeight - trailingSpaceAfter;
  const fitsInOneCol = bh === 0 || occupiedHeight <= colHeightCapacity + 0.5;
  const useMultiCol = numCol > 1 && !fitsInOneCol;
  const linesPerCol = useMultiCol ? Math.ceil(allLines.length / numCol) : allLines.length;
  let colIdx = 0;
  let entriesInCol = 0;

  for (const entry of allLines) {
    const { line, linePx, lineHeight, topGapPx, textXOffset, bulletLabel, bulletFont, bulletColor, bulletFollowsText, bulletImage, alignment, isLastLine, useResolvedFontMetrics, metricAscent } = entry;
    // Balanced column advance: when the current column has reached its share
    // of paragraphs, jump to the next one. PowerPoint never breaks a single
    // line across columns and never spills past the last column — anything
    // beyond the last column simply runs past the bbox, matching PPT.
    if (numCol > 1 && colIdx < numCol - 1 && entriesInCol >= linesPerCol) {
      colIdx++;
      entriesInCol = 0;
      cursorY = colTopY;
    }
    cursorY += topGapPx;
    entriesInCol++;
    // §21.1.2.1.1 bodyPr@rtlCol: columns fill right-to-left — mirror the
    // visual column index. LTR bodies (rtlCol false/absent) are unchanged.
    const visCol = body.rtlCol ? numCol - 1 - colIdx : colIdx;
    const xShift = visCol * colWidthShift;
    const textX = entry.textX + xShift;
    const bulletX = entry.bulletX + xShift;
    const textMaxW = entry.textMaxW;

    // Bidi: base direction from a:pPr@rtl (the parser already flips the default
    // alignment l→r for rtl paragraphs). Engage only when the base is RTL or a
    // line carries strong-RTL characters, so LTR slides keep their exact path.
    const baseRtl = entry.para.rtl === true;
    const paraNeedsBidi = baseRtl || segmentsHaveRtl(line.segments);
    const hasTab = line.segments.some((seg) => seg.isTab);

    if (hasTab) {
      // §21.1.2.1.x: resolve every inline tab in the logical READING frame.
      // UAX#9 L2 later mirrors the resulting cells physically for an RTL base.
      const marLPxE = emuToPx(entry.para.marL, scale);
      const marRPxE = emuToPx(entry.para.marR, scale);
      // The reading pen starts at the leading indent PLUS the draw-side
      // first-line indent (textXOffset): a stop is an ABSOLUTE distance from the
      // leading text-inset edge (§21.1.2.1), so the indent moves where content
      // STARTS, not where a stop sits — omitting it would widen every gap by the
      // indent and slide the cells past their stops. textXOffset only shifts the
      // LTR pen (the RTL draw right-anchors and never applies it), so it joins
      // the LTR frame only.
      const leadingIndentPx = baseRtl ? marRPxE : marLPxE + textXOffset;
      const limitPx = textMaxW + marLPxE + marRPxE;
      const tabFontSeg = line.segments.find((seg) => seg.isTab) as LayoutSegment;
      ctx.font = tabFontSeg.font;
      const spaceW = ctx.measureText(FONT_SPACE_SENTINEL).width;
      const items = line.segments.map((seg) => {
        if (seg.isTab) return { isTab: true, width: 0 };
        if (seg.math) return { isTab: false, width: seg.math.width };
        ctx.font = seg.font;
        const ls = seg.letterSpacingPx ?? 0;
        return {
          isTab: false,
          width: seg.text
            ? (seg.leadingLetterSpacingPx ?? 0) + measureTextAdvance(ctx, seg.text, ls)
            : 0,
        };
      });
      const stops = (entry.para.tabStops ?? []).map((stop) => ({
        pos: emuToPx(stop.pos, scale),
        algn: stop.algn,
      }));
      // Default tab grid (§21.1.2.2.7): explicit pPr value, else PowerPoint's
      // universal 1-inch default so an unreached tab snaps to the grid (matching
      // the wrap pass) rather than collapsing to a space (issue #1006).
      const defTabSzPx = emuToPx(entry.para.defTabSz ?? 914400, scale);
      const widths = resolveTabWidths(items, stops, leadingIndentPx, limitPx, spaceW, defTabSzPx);
      for (let i = 0; i < line.segments.length; i++) {
        if (line.segments[i].isTab) line.segments[i].tabWidthPx = widths[i];
      }
    }

    // Measure line width and the metrics of the fonts Canvas actually resolved.
    // spAutoFit is a live recalculation, so its baseline must use the same
    // resolved metrics as the line-height pass above. Ordinary text retains the
    // established PowerPoint-compatible baseline.
    let lineWidth = 0;
    let maxAscent = 0;
    let resolvedFontAscent = 0;
    let resolvedFontHeight = 0;
    for (const seg of line.segments) {
      if (seg.isTab) {
        lineWidth += seg.tabWidthPx ?? 0;
        continue;
      }
      if (seg.math) {
        lineWidth += seg.math.width;
        maxAscent = Math.max(maxAscent, seg.math.ascent);
        continue;
      }
      ctx.font = seg.font;
      const m = ctx.measureText(seg.text || FONT_BASELINE_SENTINEL);
      const ls = seg.letterSpacingPx ?? 0;
      lineWidth += seg.leadingLetterSpacingPx ?? 0;
      lineWidth += seg.text ? measureTextAdvance(ctx, seg.text, ls) : 0;
      if (m.actualBoundingBoxAscent > 0) {
        maxAscent = Math.max(maxAscent, m.actualBoundingBoxAscent);
      }
      if (useResolvedFontMetrics) {
        const fontAscent = m.fontBoundingBoxAscent ?? 0;
        const fontHeight = fontAscent + (m.fontBoundingBoxDescent ?? 0);
        if (fontHeight > resolvedFontHeight) {
          resolvedFontHeight = fontHeight;
          resolvedFontAscent = fontAscent;
        }
      }
    }
    const baselineOffset = useResolvedFontMetrics && resolvedFontHeight > 0
      ? anchor === 't' && maxAscent > 0
        // PowerPoint's spAutoFit recalculation normally seats the visible ink
        // at the top inset. A top-anchored percentage a:lnSpc is the one
        // exception confirmed by the Office boundary matrix: it retains the
        // authored percentage box and centers a taller resolved fallback in
        // that box. Absolute-point spacing and non-top anchors retain the
        // established behavior; applying this correction there moves them
        // away from the PowerPoint output.
        ? entry.para.spaceLine?.type === 'pct'
          ? Math.max(
              maxAscent,
              resolvedFontAscent + (entry.baselineLineHeight - resolvedFontHeight) / 2,
            )
          : maxAscent
        : Math.max(
            maxAscent,
            resolvedFontAscent + Math.max(0, lineHeight - resolvedFontHeight) / 2,
          )
      : Math.max(lineHeight * 0.8, maxAscent);
    // Under pPr@fontAlgn t / ctr / b, metricAscent is the line's alignment
    // reference and each run sits at its own offset from it. A list marker
    // follows the line's first run.
    const referenceY = metricAscent !== undefined
      ? cursorY + metricAscent
      : cursorY + baselineOffset;
    const baseline = referenceY
      + (line.segments.find((seg) => !seg.isTab && !seg.math && !!seg.text)?.fontAlgnOffsetPx ?? 0);

    // Reading-frame marker placement under an RTL base (issue #930, same class as
    // the docx #830 / pptx #913 leading-edge mirroring). PowerPoint seats a list
    // marker's LEADING edge — the RIGHT edge in RTL — at the line's leading
    // (right) edge and lays the text CONTIGUOUSLY to its left, so `text right =
    // leadingEdge − markerAdvance` (verified on the sample PDF: the "1." and "•"
    // markers share one right edge, each text right edge = that edge minus the
    // marker's own width). The prior code mirrored the LTR hanging gutter about
    // the text's right edge — `+ (textX − bulletX)` = `+|indent|` — which pushed
    // the marker |indent| PAST the leading edge into the right margin (the
    // far-frame over-indent this fixes), most visibly for a NARROW bullet whose
    // gap to the text then equalled the whole hanging indent.
    //
    // `leadingEdgeX` is the RTL reading start (= the `algn:'r'` right edge). The
    // marker advance is reserved so the pen below right-aligns the text to
    // `leadingEdgeX − reserve`; 0 for LTR / marker-less lines keeps them
    // byte-identical.
    const leadingEdgeX = textX + textMaxW;
    const bulletBmp = bulletImage && fetchImage
      ? rc.pictureBulletImages?.has(bulletImage.imagePath)
        ? rc.pictureBulletImages.get(bulletImage.imagePath)
        : peekCachedBitmapByPath(bulletImage.imagePath, fetchImage)
      : undefined;
    let rtlMarkerReservePx = 0;
    if (paraNeedsBidi && baseRtl) {
      if (bulletLabel) {
        ctx.font = bulletFont;
        rtlMarkerReservePx = ctx.measureText(bulletLabel).width;
      } else if (bulletImage && bulletBmp) {
        const h = bulletImage.sizePx;
        rtlMarkerReservePx = bulletBmp.height > 0
          ? h * (bulletBmp.width / bulletBmp.height)
          : h;
      }
    }

    // Draw bullet.
    if (bulletLabel) {
      ctx.font = bulletFont;
      // §21.1.2.4.5 buClrTx follows the first run's patterned fill in the
      // PowerPoint control. Keep the existing solid-colour inheritance path
      // for runs without pattFill; their bullet colour can come from a level
      // style rather than that first run. An explicit buClr stays solid.
      const firstTextSegment = line.segments.find(seg => !seg.isTab && !seg.math && !!seg.text);
      ctx.fillStyle = bulletFollowsText && firstTextSegment?.patternFill
        ? resolveSegmentTextPaint(ctx, firstTextSegment, bulletX, baseline, scale)
        : bulletColor;
      if (paraNeedsBidi && baseRtl) {
        const prevDir = ctx.direction;
        ctx.direction = 'rtl';
        // Marker RIGHT edge at the leading edge; its left edge tucks the text
        // that the pen shifts left by `rtlMarkerReservePx` (= this width).
        ctx.fillText(bulletLabel, leadingEdgeX - rtlMarkerReservePx, baseline);
        ctx.direction = prevDir;
      } else {
        ctx.fillText(bulletLabel, bulletX, baseline);
      }
    }

    // Picture bullet (`<a:buBlip>`, ECMA-376 §21.1.2.4.2). The bitmap sits on
    // the text baseline at the same gutter x a char bullet uses. The image was
    // warmed by renderSlide's prefetch pass; if its decode hasn't resolved yet
    // (or fetchImage is absent), draw nothing — the marker simply appears once
    // the image source is ready, never blocking the frame.
    if (bulletImage && fetchImage) {
      if (bulletBmp) {
        // The bullet HEIGHT is the text-derived size (× buSzPct); the WIDTH is
        // derived from the decoded bitmap's intrinsic aspect ratio so a
        // non-square marker isn't squished. §21.1.2.4.2 is silent on the exact
        // dimensions; this mirrors the PowerPoint runtime, which scales the
        // picture to the line text height while preserving its aspect ratio.
        const h = bulletImage.sizePx;
        const w = bulletBmp.height > 0 ? h * (bulletBmp.width / bulletBmp.height) : h;
        const imgY = baseline - h; // bottom-aligned to the baseline
        if (paraNeedsBidi && baseRtl) {
          // Marker RIGHT edge at the leading edge (matching the char-bullet
          // reading-frame placement above); the text pen reserves this width.
          ctx.drawImage(bulletBmp, leadingEdgeX - w, imgY, w, h);
        } else {
          ctx.drawImage(bulletBmp, bulletX, imgY, w, h);
        }
      }
    }

    const effectiveTextX = textX + textXOffset;
    let penX: number;
    if (hasTab) {
      // Tab stops are absolute from the leading text-inset edge; paragraph
      // alignment must not add a second offset (#913/#916).
      penX = baseRtl
        ? textX + textMaxW - rtlMarkerReservePx - lineWidth
        : effectiveTextX;
    } else {
      if (alignment === 'ctr') {
        penX = drawingMlLineX('ctr', effectiveTextX,
          textMaxW - textXOffset, lineWidth, false, true);
      } else if (alignment === 'r') {
        // Reading-frame (#930): an RTL marker leads at the right edge, so the text
        // right-aligns to `leadingEdge − markerAdvance` (contiguous with the
        // marker). `rtlMarkerReservePx` is 0 for non-list / marker-less lines, so
        // plain RTL paragraphs keep `leadingEdge − lineWidth` (byte-identical).
        penX = drawingMlLineX('r', textX - rtlMarkerReservePx,
          textMaxW, lineWidth, false, true);
      } else {
        penX = effectiveTextX;
      }
    }

    // Justified alignment (ECMA-376 §20.1.10.59 ST_TextAlignType): just/justLow
    // fill the column by widening inter-word + inter-CJK gaps with the
    // paragraph's last line left natural; dist/thaiDist do the same on every
    // line. Segments are merged by style in layout, so justifyLine splits each
    // line inside its text at the gaps and returns draw pieces carrying `jext`
    // (px to advance after each piece). penX already starts at the left edge
    // (the `else` above). Disabled under bidi (visual≠logical order).
    const justifyMode =
      alignment === 'just' || alignment === 'justLow' ? 'just' as const
      : alignment === 'thaiDist' ? 'thaiDist' as const
      : alignment === 'dist' ? 'dist' as const
      : null;
    // Tab-delimited cells keep their natural widths; the inline gaps provide
    // their stop alignment and must not participate in justification.
    // A `just` line ended by a manual <a:br> is left-aligned like the last line
    // (§20.1.10.59 + §21.1.2.2.1); `dist` ignores this and fills every line
    // (justifyLine only suppresses the last line for `just`).
    const endsLogicalLine = isLastLine || (line.endsWithBreak ?? false);
    const drawSegs = justifyMode && !paraNeedsBidi && !hasTab
      && drawingMlLineShouldJustify(alignment, isLastLine, line.endsWithBreak ?? false)
      ? justifyLine(line.segments, textMaxW - textXOffset, lineWidth, justifyMode, endsLogicalLine)
      : null;
    const segs: (LayoutSegment & Partial<Justified>)[] = drawSegs ?? line.segments;

    // Visual draw order: under bidi, reorder segments per UAX#9 (rule L2) and
    // draw each with ctx.direction matching its resolved direction. textAlign
    // is already 'left', so penX stays the segment's left edge.
    const visual: LineVisualOrder | null = paraNeedsBidi
      ? computeLineVisualOrder(line.segments, baseRtl)
      : null;
    const visualBoundarySpacing = (previousLi: number, currentLi: number): number => {
      if (Math.abs(previousLi - currentLi) !== 1) return 0;
      const before = line.segments[Math.min(previousLi, currentLi)];
      const after = line.segments[Math.max(previousLi, currentLi)];
      if (
        before.isTab || before.math || after.isTab || after.math
        || before.sourceRunId == null
        || before.sourceRunId !== after.sourceRunId
      ) return 0;
      // `after` owns the logical boundary metadata. Place that advance between
      // the two segments in VISUAL order, even when bidi reverses them.
      return after.leadingLetterSpacingPx ?? 0;
    };
    const segCount = segs.length;
    for (let vi = 0; vi < segCount; vi++) {
      const li = visual ? visual.order[vi] : vi;
      const seg = segs[li];
      const segRtl = visual ? visual.rtl[li] : false;
      if (paraNeedsBidi) ctx.direction = segRtl ? 'rtl' : 'ltr';
      if (vi > 0) {
        const previousLi = visual ? visual.order[vi - 1] : vi - 1;
        penX += visualBoundarySpacing(previousLi, li);
      }
      if (seg.isTab) {
        penX += seg.tabWidthPx ?? 0;
        continue;
      }
      // Justification advance after this piece (0 when not justifying). Added to
      // the pen, and folded into the trailing edge of underline / strikethrough
      // and the reported onTextRun width, so decorations and the text layer span
      // the widened gap instead of leaving a hole. The last content piece always
      // has jext 0 (justifyLine never stretches after the final glyph).
      const jext = seg.jext ?? 0;
      const splitBefore = seg.splitBefore;
      const segPerGap = seg.perGap ?? 0;
      const internalStretch =
        splitBefore && splitBefore.length > 0 ? splitBefore.length * segPerGap : 0;

      // ── Equation segment: draw the cached image instead of text ──────────
      if (seg.math) {
        const w = seg.math.width;
        paintEquation(ctx, seg, penX, baseline - seg.math.ascent);
        penX += w;
        penX += jext;
        continue;
      }
      ctx.font = seg.font;
      const drawSizePx = seg.drawSizePx ?? seg.sizePx;
      // baseline shift: OOXML baseline in thousandths of a point; positive = superscript (up)
      const baselineShift = seg.baseline ? -(seg.baseline / 100000) * seg.sizePx : 0;
      const segBaseline = referenceY + (seg.fontAlgnOffsetPx ?? 0) + baselineShift;
      const glyphPaint = resolveSegmentTextPaint(ctx, seg, penX, segBaseline, scale);
      ctx.fillStyle = glyphPaint;
      const ls = seg.letterSpacingPx ?? 0;

      // Run-level text highlight (rPr > a:highlight, ECMA-376 §21.1.2.3.4).
      // Box advance = glyph measure + letter spacing + justification stretch
      // (internal CJK pitch + trailing `jext`).
      if (seg.highlight && seg.text) {
        const hlW = measureTextAdvance(ctx, seg.text, ls)
          + internalStretch
          + jext;
        paintHighlight(ctx, penX, segBaseline, hlW, drawSizePx, seg.highlight, glyphPaint);
      }

      // noFill is the GLYPH fill choice (§21.1.2.3.9).  It does not erase an
      // independently authored a:ln outline, uFill underline, highlight, or
      // the text-selection geometry.  Keep those consumers in the common path.

      const segShadow = seg.shadow;

      // Draw `text` at `atX` honouring this segment's letter-spacing / RTL
      // semantics. Lifted into a closure so the fill and stroke (outline) paths
      // share it, and so the split-CJK branch below can call it per piece.
      const drawWithFont = (
        target: CanvasRenderingContext2D,
        text: string,
        atX: number,
        op: 'fill' | 'stroke',
      ): void => {
        const paint = op === 'fill' ? target.fillText.bind(target) : target.strokeText.bind(target);
        if (ls !== 0 && codePointCount(text) > 1) {
          // rPr @spc (§21.1.2.3.9; ST_TextPoint §20.1.10.74): apply the
          // normalized spacing via Canvas letterSpacing and draw the whole
          // CONTEXTUALLY-shaped string in ONE paint. The layout measures this
          // same state via measureTextAdvance, so contextual CJK packing and
          // multi-code-point shaping clusters use the browser's exact advance.
          // This keeps 約物半角 contextual half-width collapse honoured, so a
          // piece's glyphs stay aligned with its contextual `dx` origin and the
          // next piece/run never overlaps — the pptx analog of docx PR #626), AND
          // preserves Arabic cursive joining for RTL. (The old LTR branch summed
          // ISOLATED measure(ch)+ls, which overran the contextual box at 約物
          // punctuation.)
          const lctx = target as CanvasRenderingContext2D & { letterSpacing?: string };
          const prev = lctx.letterSpacing;
          if (hasNativeLetterSpacing(target)) {
            lctx.letterSpacing = `${ls}px`;
            paint(text, atX, segBaseline);
            try { lctx.letterSpacing = prev; } catch { /* inert implementation */ }
          } else {
            // Pre-letterSpacing Canvas fallback. It cannot preserve contextual
            // shaping across the split, but its manual advance matches the
            // code-point approximation in measureTextAdvance.
            let x = atX;
            const cps = [...text];
            for (let index = 0; index < cps.length; index++) {
              const ch = cps[index];
              paint(ch, x, segBaseline);
              if (index < cps.length - 1) x += target.measureText(ch).width + ls;
            }
          }
        } else {
          paint(text, atX, segBaseline);
        }
      };

      // `splitBefore` is set when justifyLine annotated this segment with
      // internal CJK gaps. Anchor each piece to the whole-string prefix advance
      // (via core's `justifiedPiecePositions`) to absorb 約物半角 drift — see
      // packages/core/src/text/justify-positions.ts for the invariant proof.
      const measureCb = (s: string): number => measureTextAdvance(ctx, s, ls);
      const pieces =
        splitBefore && splitBefore.length > 0
          ? justifiedPiecePositions([...seg.text], splitBefore, segPerGap, measureCb)
          : null;

      // A run is FULLY distributed when justifyLine opened a gap at EVERY internal
      // inter-glyph boundary (splitBefore.length === cps.length - 1) — e.g. a
      // pure-CJK `algn="just"/"dist"` line. Drawing one single-glyph piece per code
      // point (the `pieces` loop) routes through drawWithFont's `else → paint(ch)`:
      // each piece has one code point, so the `ls !== 0 && codePointCount(text) > 1`
      // letterSpacing branch never fires, and every glyph is painted ISOLATED even
      // when ls is non-zero. That loses the browser's JIS X 4051 約物連続 packing — a
      // closing-class punct immediately followed by an opening bracket ("：［",
      // "、［", "）（") packs ~half-em in measureText (a bracket next to a plain
      // kanji/kana does NOT pack). The isolated full-width bracket then overruns its
      // successor (the justify analog of docx PR #630). The fix: a gap at EVERY
      // boundary ⇒ uniform per-glyph pitch (ls + segPerGap), so draw the whole
      // CONTEXTUALLY-shaped run in ONE paint with ctx.letterSpacing = ls+segPerGap:
      // glyph i lands at measure(prefix_i)+i·(ls+segPerGap) — the exact justified
      // position — so the packing is honoured and nothing overlaps; the final glyph
      // reaches the segment box edge (segW). measureCb already ran (pieces built
      // above) at the natural advance, so no measureText sees this letterSpacing.
      const cps = [...seg.text];
      const fullyDistributed =
        !!splitBefore && splitBefore.length === cps.length - 1 && cps.length > 1;
      const drawRun = (target: CanvasRenderingContext2D, op: 'fill' | 'stroke'): void => {
        if (eaVertUpright) {
          // ECMA-376 §20.1.10.83 eaVert: paint each glyph with its UAX#50 vertical
          // orientation (CJK/kana upright, Latin sideways, brackets/comma
          // substituted, ー rotated). The segment's cell positions are already set
          // by the horizontal layout (penX / segW), so only the glyph orientation
          // changes. A FULLY-DISTRIBUTED line (algn just/dist opening a gap at every
          // inter-glyph boundary — e.g. 均等割り付け of a pure-CJK column) carries its
          // uniform justification pitch in `segPerGap`; fold it into the per-glyph
          // advance so the column SPREADS to fill the length instead of bunching at
          // the top (mirrors the horizontal fully-distributed fast-path below).
          // Partial `just` distribution (mixed Latin/CJK pieces) is a rare eaVert
          // edge and is not redistributed here.
          const eaPitch = fullyDistributed ? ls + segPerGap : ls;
          drawEaVertRun(target, seg.text, penX, segBaseline, drawSizePx, eaPitch, op);
          return;
        }
        if (fullyDistributed) {
          const paint = op === 'fill'
            ? target.fillText.bind(target)
            : target.strokeText.bind(target);
          const pitch = ls + segPerGap;
          if (hasNativeLetterSpacing(target)) {
            const lctx = target as CanvasRenderingContext2D & { letterSpacing: string };
            const prev = lctx.letterSpacing;
            lctx.letterSpacing = `${pitch}px`;
            paint(seg.text, penX, segBaseline);
            try { lctx.letterSpacing = prev; } catch { /* inert implementation */ }
          } else {
            // Old/inert Canvas: use the same code-point pitch approximation as
            // measureTextAdvance so distributed layout and paint stay aligned.
            let x = penX;
            for (let index = 0; index < cps.length; index++) {
              const ch = cps[index];
              paint(ch, x, segBaseline);
              if (index < cps.length - 1) x += target.measureText(ch).width + pitch;
            }
          }
        } else if (pieces) {
          for (const { text: pieceText, dx } of pieces) {
            drawWithFont(target, pieceText, penX + dx, op);
          }
        } else {
          drawWithFont(target, seg.text, penX, op);
        }
      };

      // A run-level reflection is painted before the glyphs themselves, using
      // the same device-space effect pipeline as shape reflections. Keeping the
      // draw closure target-agnostic preserves letter spacing, justification,
      // vertical glyph orientation, and outlines without rasterizing the whole
      // text body as one effect image.
      const segReflection = seg.reflection;
      if (segReflection && seg.text && !seg.noFill) {
        const deviceW = (ctx.canvas as { width: number }).width || 0;
        const deviceH = (ctx.canvas as { height: number }).height || 0;
        if (deviceW > 0 && deviceH > 0) {
          ctx.font = seg.font;
          const metrics = ctx.measureText(seg.text);
          const ascent = Number.isFinite(metrics.actualBoundingBoxAscent)
            ? metrics.actualBoundingBoxAscent
            : drawSizePx * 0.8;
          // Zero is a valid descent for all-uppercase runs. Treating it as
          // "missing" inserts a synthetic 0.2em blank band below the glyph;
          // the reflection's strongest stA region then lands on transparency
          // and the visible strokes are almost fully faded.
          const descent = Number.isFinite(metrics.actualBoundingBoxDescent)
            ? metrics.actualBoundingBoxDescent
            : drawSizePx * 0.2;
          const left = Number.isFinite(metrics.actualBoundingBoxLeft)
            ? metrics.actualBoundingBoxLeft
            : 0;
          const right = Number.isFinite(metrics.actualBoundingBoxRight)
            ? metrics.actualBoundingBoxRight
            : metrics.width;
          const liveTransform = ctx.getTransform();
          const det = Math.abs(
            liveTransform.a * liveTransform.d - liveTransform.b * liveTransform.c,
          );
          const devScale = det > 0 ? Math.sqrt(det) : 1;
          const bbox = {
            x: (penX - left) * devScale,
            y: (segBaseline - ascent) * devScale,
            w: Math.max(1, left + right) * devScale,
            h: Math.max(1, ascent + descent) * devScale,
          };
          applyTextRunReflection(
            ctx,
            (target) => {
              target.font = seg.font;
              target.fillStyle = resolveSegmentTextPaint(target, seg, penX, segBaseline, scale);
              drawRun(target, 'fill');
            },
            bbox,
            segReflection,
            scale * devScale,
            liveTransform,
            deviceW,
            deviceH,
          );
          ctx.font = seg.font;
          ctx.fillStyle = glyphPaint;
        }
      }

      // Run-level text shadow (rPr > effectLst > outerShdw). Apply it only to
      // the primary glyph paint; the reflection above must not cast a second
      // shadow of its already-mirrored pixels.
      if (segShadow && !seg.noFill) {
        const dirRad = (segShadow.dir * Math.PI) / 180;
        const dist = emuToPx(segShadow.dist, scale);
        ctx.save();
        ctx.shadowColor = hexToRgba(segShadow.color, segShadow.alpha);
        ctx.shadowBlur = emuToPx(segShadow.blur, scale);
        ctx.shadowOffsetX = Math.cos(dirRad) * dist;
        ctx.shadowOffsetY = Math.sin(dirRad) * dist;
      }

      if (!seg.noFill) drawRun(ctx, 'fill');

      if (segShadow && !seg.noFill) ctx.restore();

      ctx.font = seg.font;
      const segW = measureTextAdvance(ctx, seg.text, ls) + internalStretch;

      // Run-level text outline (rPr > a:ln). Strokes each glyph in addition
      // to the fill so the text reads as a thin lined character. ECMA-376
      // §20.1.2.2.24: `w` is EMU; convert to px via the same scale used for
      // fonts. Min 0.5 px so a 1-pt outline at small zoom levels stays
      // visible. Skip when width <= 0 (parser may emit width=0 for
      // explicitly empty <a:ln/>).
      const segOutline = seg.outline;
      if (segOutline && segOutline.width > 0) {
        ctx.save();
        ctx.lineWidth = Math.max(0.5, emuToPx(segOutline.width, scale));
        ctx.strokeStyle = segOutline.fill
          ? resolveFillCore(
              segOutline.fill, ctx, penX, segBaseline - drawSizePx,
              Math.max(1, segW), drawSizePx, 0, scale * PT_TO_EMU,
            ) ?? 'rgba(0,0,0,0)'
          : segOutline.color ? `#${segOutline.color}` : seg.noFill ? 'rgba(0,0,0,0)' : glyphPaint;
        if (!seg.noFill || segOutline.fill || segOutline.color) {
          ctx.lineJoin = 'round';
          drawRun(ctx, 'stroke');
        }
        ctx.restore();
      }

      if (onTextRun && seg.text) {
        onTextRun({
          text: seg.text,
          inShapeX: penX - bx,
          inShapeY: cursorY - by,
          w: segW + jext,
          h: lineHeight,
          fontSize: drawSizePx,
          font: seg.font,
          shapeX: bx,
          shapeY: by,
          shapeW: bw,
          shapeH: bh,
          rotation: shapeRotation,
          hyperlink: seg.hyperlink,
        });
      }

      paintRunDecorations(ctx, seg, penX, segBaseline, segW + jext, glyphPaint, scale, rc.dpr);
      penX += segW;
      penX += jext;
    }
    if (paraNeedsBidi) ctx.direction = 'ltr';

    cursorY += linePx;
  }

  ctx.restore();
}
