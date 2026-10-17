// XLSX slicer painting. Split from renderer.ts (#1857).

import type { Worksheet, SlicerItem, SlicerStyle, SlicerElementStyle } from './types.js';
import { EMU_PER_PX } from '@silurus/ooxml-core';
import type { GridAxisGeometry } from './internal/grid-axis-geometry.js';
import { sheetAnchoredRectX } from './renderer-font-metrics.js';
import { sheetXForCol, sheetYForRow } from './renderer-drawings.js';

// ── renderSlicers ───────────────────────────────────────────────────────────
//
// Office 2010+ pivot / table slicer. We don't own a slicer engine, so this
// renders a static button bank: header with the slicer caption, then one
// button per saved item using the selection flags from the slicerCache. The
// visual language (pale blue outline, white "selected" buttons on a darker
// background, gray "deselected" buttons) intentionally mirrors Excel's
// default slicer style. Custom styles resolve from the workbook's standard
// table-style dxfs (frame/header) plus x14 slicer-style dxfs (item states).

const SLICER_HEADER_FONT = '600 12px "Meiryo UI", "Segoe UI", sans-serif';
const SLICER_ITEM_FONT   = '11px "Meiryo UI", "Segoe UI", sans-serif';
const SLICER_BG           = '#FFFFFF';
const SLICER_BORDER       = '#BFBFBF';
const SLICER_HEADER_BG    = '#F2F2F2';
const SLICER_HEADER_FG    = '#404040';
const SLICER_ITEM_SEL_BG  = '#FFFFFF';
const SLICER_ITEM_SEL_FG  = '#000000';
const SLICER_ITEM_SEL_BD  = '#A5A5A5';
const SLICER_ITEM_OFF_BG  = '#E7E6E6';
const SLICER_ITEM_OFF_FG  = '#A6A6A6';
const SLICER_ITEM_OFF_BD  = '#C6C6C6';

export function renderSlicers(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  colAxis: GridAxisGeometry,
  rowAxis: GridAxisGeometry,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
  rtl: boolean,
  canvasW: number,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;
  const slicers = ws.slicers;
  if (!slicers) return;

  const scrollOriginSheetX = sheetXForCol(colAxis, startCol);
  const scrollOriginSheetY = sheetYForRow(rowAxis, startRow);
  const clipX = sheetAnchoredRectX(scrollAreaX, scrollAreaW, canvasW, rtl);

  for (const anchor of slicers) {
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;
    const toCol1   = anchor.toCol   + 1;
    const toRow1   = anchor.toRow   + 1;

    const shX1 = sheetXForCol(colAxis, fromCol1) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const shY1 = sheetYForRow(rowAxis, fromRow1) + (anchor.fromRowOff * cs) / EMU_PER_PX;
    const shX2 = sheetXForCol(colAxis, toCol1) + (anchor.toColOff * cs) / EMU_PER_PX;
    const shY2 = sheetYForRow(rowAxis, toRow1) + (anchor.toRowOff * cs) / EMU_PER_PX;

    const w = shX2 - shX1;
    const h = shY2 - shY1;
    if (w <= 0 || h <= 0) continue;

    const logicalX = scrollAreaX + (shX1 - scrollOriginSheetX) - scrollOffsetX;
    const x = sheetAnchoredRectX(logicalX, w, canvasW, rtl);
    const y = scrollAreaY + (shY1 - scrollOriginSheetY) - scrollOffsetY;

    if (x + w < clipX || x > clipX + scrollAreaW) continue;
    if (y + h < scrollAreaY || y > scrollAreaY + scrollAreaH) continue;

    ctx.save();
    ctx.beginPath();
    ctx.rect(clipX, scrollAreaY, scrollAreaW, scrollAreaH);
    ctx.clip();

    drawSlicerFrame(ctx, anchor.caption, anchor.items, x, y, w, h, cs, anchor.style);

    ctx.restore();
  }
}

export function drawSlicerFrame(
  ctx: CanvasRenderingContext2D,
  caption: string,
  items: SlicerItem[],
  x: number,
  y: number,
  w: number,
  h: number,
  cs: number,
  style?: SlicerStyle,
): void {
  const custom = style != null;
  // Outer frame. For a custom style, an absent border/fill contributes no
  // formatting (§18.5.1.2); do not invent the fallback gray hairline.
  ctx.fillStyle = style?.whole?.fillColor ?? SLICER_BG;
  ctx.fillRect(x, y, w, h);
  const outerBorder = custom ? style.whole?.borderColor : SLICER_BORDER;
  if (outerBorder) {
    ctx.strokeStyle = outerBorder;
    ctx.lineWidth = 1;
    ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
  }

  // Header band with caption.
  const headerH = Math.max(20 * cs, 14);
  const headerFill = custom ? style.header?.fillColor : SLICER_HEADER_BG;
  if (headerFill) {
    ctx.fillStyle = headerFill;
    ctx.fillRect(x + 1, y + 1, w - 2, headerH);
  }
  ctx.fillStyle = style?.header?.fontColor ?? SLICER_HEADER_FG;
  ctx.font = slicerFont(style?.header, SLICER_HEADER_FONT, cs);
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'left';
  const headerPad = 6 * cs;
  drawClippedText(ctx, caption, x + headerPad, y + headerH / 2 + 1, w - 2 * headerPad);

  // Item buttons. Items expand to fill the available height (up to a
  // minimum button size) and are clipped by the slicer rect. We don't
  // implement scroll arrows because this renderer is non-interactive.
  if (items.length === 0) return;
  const gap = Math.max(1, Math.round(2 * cs));
  const innerPad = 4 * cs;
  const listX = x + innerPad;
  const listY = y + headerH + innerPad;
  const listW = w - 2 * innerPad;
  const listH = h - headerH - 2 * innerPad;
  if (listW <= 0 || listH <= 0) return;

  // Prefer Excel's rough row height (~20 sheet-px) but compress if the
  // slicer is shallow so at least the first items fit.
  const preferredItemH = Math.max(18 * cs, 16);
  const maxVisibleByH = Math.max(1, Math.floor((listH + gap) / (preferredItemH + gap)));
  const visible = Math.min(items.length, maxVisibleByH);
  const itemH = Math.min(preferredItemH, (listH - gap * (visible - 1)) / visible);
  if (itemH <= 0) return;

  ctx.font = scaleFont(SLICER_ITEM_FONT, cs);
  const itemPad = 8 * cs;
  for (let i = 0; i < visible; i++) {
    const item = items[i];
    const iy = listY + i * (itemH + gap);
    const selected = item.selected;
    const itemStyle = selected
      ? style?.selectedItemWithData
      : style?.unselectedItemWithData;
    const fill = itemStyle?.fillColor ??
      (selected ? SLICER_ITEM_SEL_BG : SLICER_ITEM_OFF_BG);
    const border = custom
      ? itemStyle?.borderColor
      : (selected ? SLICER_ITEM_SEL_BD : SLICER_ITEM_OFF_BD);
    ctx.fillStyle = fill;
    if (custom) {
      roundedRectPath(ctx, listX, iy, listW, itemH, Math.min(4 * cs, itemH / 4));
      ctx.fill();
      if (border) {
        ctx.strokeStyle = border;
        ctx.lineWidth = 1;
        ctx.stroke();
      }
    } else {
      ctx.fillRect(listX, iy, listW, itemH);
      if (border) {
        ctx.strokeStyle = border;
        ctx.lineWidth = 1;
        ctx.strokeRect(listX + 0.5, iy + 0.5, listW - 1, itemH - 1);
      }
    }
    ctx.font = slicerFont(itemStyle, SLICER_ITEM_FONT, cs);
    ctx.fillStyle = itemStyle?.fontColor ??
      (selected ? SLICER_ITEM_SEL_FG : SLICER_ITEM_OFF_FG);
    drawClippedText(ctx, item.name, listX + itemPad, iy + itemH / 2 + 1, listW - 2 * itemPad);
  }
}

function roundedRectPath(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  radius: number,
): void {
  const r = Math.max(0, Math.min(radius, w / 2, h / 2));
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

function slicerFont(
  style: SlicerElementStyle | undefined,
  fallback: string,
  cs: number,
): string {
  if (!style) return scaleFont(fallback, cs);
  const size = Math.round((style.fontSize ?? 11) * cs);
  const weight = style.fontBold ? 'bold ' : '';
  const family = style.fontFamily
    ? `"${style.fontFamily}", "Segoe UI", sans-serif`
    : '"Meiryo UI", "Segoe UI", sans-serif';
  return `${weight}${size}px ${family}`;
}

function scaleFont(css: string, cs: number): string {
  // Re-scale the leading `<size>px` token by `cs`. Safe fallback: leave the
  // string as-is so the slicer remains readable when parsing fails.
  return css.replace(/(\d+(?:\.\d+)?)px/, (_, n) => `${Math.round(Number(n) * cs)}px`);
}

function drawClippedText(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  maxWidth: number,
): void {
  if (maxWidth <= 0) return;
  let s = text;
  if (ctx.measureText(s).width > maxWidth) {
    const ellipsis = '…';
    while (s.length > 0 && ctx.measureText(s + ellipsis).width > maxWidth) {
      s = s.slice(0, -1);
    }
    s = s.length > 0 ? s + ellipsis : '';
  }
  ctx.fillText(s, x, y);
}
