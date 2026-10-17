// PPTX table (graphicFrame a:tbl) painting. Split from renderer.ts (#1857).

import { clipCanvasHorizontally } from '@silurus/ooxml-core/internal/canvas-clip';
import type { TableElement, TableCell, Fill, Stroke, TextBody } from './types';
import { crispOffset, EMU_PER_PT as PT_TO_EMU, fillDoubleBorder } from '@silurus/ooxml-core';
import { resolveTableBorderConflict } from './table-border-conflict.js';
import {
  emuToPx,
  hexToRgba,
  applyFrameTransform,
  type RenderContext,
  type TextRunCallback,
} from './renderer-context';
import {
  applyStroke,
  resolveShapeFill,
  shapeFillKey,
  paintPreparedShapeImageFill,
} from './renderer-shape-fill';
import { renderTextBody } from './renderer-text-body';

export function renderTable(
  ctx: CanvasRenderingContext2D,
  el: TableElement,
  scale: number,
  slideNumber?: number,
  rc: RenderContext = { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
  onTextRun?: TextRunCallback,
) {
  ctx.save();
  applyFrameTransform(ctx, el, scale);
  const x0 = emuToPx(el.x, scale);
  const y0 = emuToPx(el.y, scale);

  // ECMA-376 §21.1.2.1.1 describes spcFirstLastPara for a text body, but
  // PowerPoint does not apply its edge exception inside a:tc. PDF controls
  // with centred and top-anchored table cells retain identical glyph and row
  // positions for absent/0/1, with 12pt spcPts and 50% spcPct before or
  // after. A second paragraph does receive spcAft as an interior gap. Keep the
  // Office table-cell rule in both measurement and paint.
  const tableTextBody = (body: TextBody): TextBody => body.spcFirstLastPara
    ? { ...body, spcFirstLastPara: false }
    : body;

  // Convert col widths to pixels.
  const colWidths = el.cols.map(c => emuToPx(c, scale));
  const numCols = colWidths.length;

  // Spanned width starting at column `ci` for `span` columns.
  const spannedWidth = (ci: number, span: number): number => {
    let w = 0;
    for (let s = 0; s < span; s++) w += colWidths[ci + s] ?? 0;
    return w;
  };

  // ── Row heights: Office minimum-row semantics ─────────────────────────────
  // ECMA-376 §21.1.3.18 defines a:tr@h as the row height; [MS-OE376]
  // §2.1.1347 additionally constrains it to zero or at least the minimum row
  // height. PowerPoint grows a row to fit its tallest cell's laid-out text. A
  // literal h=0 therefore becomes
  // content-driven. We measure each cell's text body at its spanned width
  // (reusing the same renderTextBody machinery via measureOnly) and take
  // max(tr@h, tallest single-row cell content). A rowSpan cell distributes
  // its content height across the rows it covers so it doesn't inflate the
  // first row.
  const authoredRowHeights = el.rows.map(r => emuToPx(r.height, scale));
  const rowHeights = [...authoredRowHeights];
  const authoredRowsTotalEmu = el.rows.reduce((sum, row) => sum + row.height, 0);
  // The graphic-frame extent is the table's authored outer height. PowerPoint
  // can leave slack between that extent and the sum of positive a:tr@h minima;
  // that discrepancy signals that text-driven row growth was present when the
  // table was authored. When the positive minima already fill the frame, do not
  // invent growth merely because browser font metrics wrap differently.
  const hasAuthoredRowGrowthSignal = el.height > authoredRowsTotalEmu;

  // First pass: single-row (rowSpan ≤ 1) cells set their own row's minimum.
  for (let ri = 0; ri < el.rows.length; ri++) {
    const row = el.rows[ri];
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      if (cell.hMerge || cell.vMerge) continue;
      if ((cell.rowSpan || 1) > 1) continue;
      if (!cell.textBody) continue;
      if (row.height > 0 && !hasAuthoredRowGrowthSignal) continue;
      const cellW = spannedWidth(ci, cell.gridSpan || 1);
      const needed = (renderTextBody(
        ctx, tableTextBody(cell.textBody), 0, 0, cellW, 0, scale, null, 0, false, false,
        '#000000', slideNumber, rc, undefined, true, undefined, false, row.height === 0,
      ) as number) || 0;
      if (needed > rowHeights[ri]) rowHeights[ri] = needed;
    }
  }

  // Second pass: rowSpan cells. If the content needs more than the sum of the
  // rows it spans, distribute the deficit across those rows so the merged area
  // grows without inflating any single row beyond what its own content needs.
  for (let ri = 0; ri < el.rows.length; ri++) {
    const row = el.rows[ri];
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      if (cell.hMerge || cell.vMerge) continue;
      const span = cell.rowSpan || 1;
      if (span <= 1 || !cell.textBody) continue;
      const cellW = spannedWidth(ci, cell.gridSpan || 1);
      const hasAutoHeightRow = el.rows
        .slice(ri, Math.min(el.rows.length, ri + span))
        .some((spannedRow) => spannedRow.height === 0);
      if (!hasAutoHeightRow && !hasAuthoredRowGrowthSignal) continue;
      const needed = (renderTextBody(
        ctx, tableTextBody(cell.textBody), 0, 0, cellW, 0, scale, null, 0, false, false,
        '#000000', slideNumber, rc, undefined, true, undefined, false, hasAutoHeightRow,
      ) as number) || 0;
      let have = 0;
      for (let s = 0; s < span && ri + s < rowHeights.length; s++) have += rowHeights[ri + s];
      if (needed > have) {
        const extra = (needed - have) / span;
        for (let s = 0; s < span && ri + s < rowHeights.length; s++) rowHeights[ri + s] += extra;
      }
    }
  }

  // ── Column x-positions ────────────────────────────────────────────────────
  // ECMA-376 §21.1.3.13 (a:tblPr@rtl): a right-to-left table places column 0
  // at the right edge, columns advancing leftward. We precompute the left
  // pixel edge of each column so RTL is a coordinate flip (no ctx.scale(-1,1)
  // which would mirror text and borders).
  const tableW = colWidths.reduce((a, b) => a + b, 0);
  const colLeft: number[] = new Array(numCols);
  if (el.rtl) {
    let right = x0 + tableW;
    for (let ci = 0; ci < numCols; ci++) {
      right -= colWidths[ci];
      colLeft[ci] = right;
    }
  } else {
    let left = x0;
    for (let ci = 0; ci < numCols; ci++) {
      colLeft[ci] = left;
      left += colWidths[ci];
    }
  }

  // Left pixel edge of a cell spanning columns [ci, ci+span). Under RTL the
  // span grows leftward, so the merged cell's left edge is the leftmost column.
  const spannedLeft = (ci: number, span: number): number =>
    el.rtl ? colLeft[ci + span - 1] : colLeft[ci];

  const rowTop: number[] = new Array(el.rows.length);
  {
    let y = y0;
    for (let ri = 0; ri < el.rows.length; ri++) { rowTop[ri] = y; y += rowHeights[ri]; }
  }

  // Text-run geometry is needed only by selection/find consumers. Ordinary
  // paint-only renders avoid frame trigonometry. A single callback reads the
  // current synchronous cell state, avoiding one closure allocation per cell.
  let tableTextRun: TextRunCallback | undefined;
  let cellHorzOverflow: 'clip' | 'overflow' | undefined;
  let textRunCell: Readonly<{ row: number; column: number }> = { row: 0, column: 0 };
  if (onTextRun) {
    // The canvas rotates/flips the complete graphic frame around the authored
    // table centre. The overlays group runs by their cell-sized text frame, so
    // move each cell centre through that same outer transform and carry the
    // rotation/flips to the DOM group. Applying the transform around the moved
    // cell centre is geometrically equivalent to applying it around the table
    // centre while preserving the local frame for vertical-text rotation.
    const frameW = emuToPx(el.width, scale);
    const frameH = emuToPx(el.height, scale);
    const frameCx = x0 + frameW / 2;
    const frameCy = y0 + frameH / 2;
    const rotationRad = (el.rotation * Math.PI) / 180;
    const rotationCos = Math.cos(rotationRad);
    const rotationSin = Math.sin(rotationRad);
    tableTextRun = (run) => {
      let dx = run.shapeX + run.shapeW / 2 - frameCx;
      let dy = run.shapeY + run.shapeH / 2 - frameCy;
      if (el.flipH) dx = -dx;
      if (el.flipV) dy = -dy;
      const cellCx = frameCx + rotationCos * dx - rotationSin * dy;
      const cellCy = frameCy + rotationSin * dx + rotationCos * dy;
      onTextRun({
        ...run,
        ...(el.id === undefined ? {} : { shapeId: el.id }),
        shapeX: cellCx - run.shapeW / 2,
        shapeY: cellCy - run.shapeH / 2,
        rotation: el.rotation,
        ...(el.flipH ? { shapeFlipH: true } : {}),
        ...(el.flipV ? { shapeFlipV: true } : {}),
        tableCell: textRunCell,
        ...(cellHorzOverflow === undefined ? {} : { cellHorzOverflow }),
      });
    };
  }

  // ECMA-376 border-collapse: paint every cell's fill + text body FIRST, then
  // every cell's borders, so a neighbouring cell's background fill can never
  // overpaint the half of a shared gridline an adjacent cell already drew.
  // (Mirrors the docx table renderer's two-pass order.) Collect each
  // non-continuation cell's paint box once; rowSpan/gridSpan and the RTL
  // coordinate flip are already baked into colX/rowY/cellW/cellH, so both passes
  // replay identical geometry.
  // Each job also records its GRID footprint (top-left slot ci/ri, its column
  // span and last row) so the border pass can find, for every interior gridline,
  // the neighbouring cell that shares it — DrawingML tables populate every grid
  // slot (a spanning `<a:tc>` is followed by `hMerge`/`vMerge` continuation
  // `<a:tc>`s), so `ci` is the true grid column index.
  const jobs: Array<{
    cell: TableCell; colX: number; rowY: number; cellW: number; cellH: number;
    ci: number; ri: number; span: number; lastRi: number;
  }> = [];
  // Grid-slot → job index occupancy, so an interior edge can look up the adjacent
  // cell. A spanning anchor fills every slot it covers; the neighbour query on any
  // covered slot resolves back to that anchor job.
  const occupancy: number[][] = el.rows.map(() => new Array<number>(numCols).fill(-1));
  for (let ri = 0; ri < el.rows.length; ri++) {
    const row = el.rows[ri];
    const rowY = rowTop[ri];
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      // Merged-cell continuations are drawn by their anchor cell.
      if (cell.hMerge || cell.vMerge) continue;
      const span = cell.gridSpan || 1;
      const rspan = cell.rowSpan || 1;
      const cellW = spannedWidth(ci, span);
      let cellH = 0;
      for (let s = 0; s < rspan; s++) cellH += rowHeights[ri + s] ?? 0;
      const colX = spannedLeft(ci, span);
      const lastRi = Math.min(ri + rspan - 1, el.rows.length - 1);
      const jobIndex = jobs.length;
      jobs.push({ cell, colX, rowY, cellW, cellH, ci, ri, span, lastRi });
      // Record this cell's grid footprint so interior-edge neighbour lookups
      // resolve to it (a spanning cell owns every slot it covers).
      for (let rj = ri; rj <= lastRi; rj++) {
        for (let cj = ci; cj < ci + span && cj < numCols; cj++) {
          occupancy[rj][cj] = jobIndex;
        }
      }
    }
  }

  // tblBg uses the full grown table rectangle, once, before alpha cell
  // fills. A per-cell background would restart theme gradients at every row.
  const tableHeight = rowHeights.reduce((sum, height) => sum + height, 0);
  // Keep the theme paint and band alpha intact. Solid/gradient matrix
  // controls agree with Office PDF vectors and Poppler rasterization; MuPDF
  // can differ by 1–2 RGB levels on non-white alpha composites. That PDF
  // rasterizer difference is not a colour compensation rule for Canvas.
  // Image fills share shape decoding, effects, crop/tile placement and cache
  // ownership; ordinary fills keep the canvas-aware gradient/pattern resolver.
  const paintTableFill = (fill: Fill | null, x: number, y: number, w: number, h: number) => {
    if (fill?.fillType === 'image') {
      ctx.beginPath();
      ctx.rect(x, y, w, h);
      paintPreparedShapeImageFill(ctx, fill, rc.shapeFillImages?.get(shapeFillKey(fill)),
        { x, y, w, h }, scale, false, el);
    } else {
      const paint = resolveShapeFill(fill, ctx, x, y, w, h, el.rotation, scale * PT_TO_EMU);
      if (paint) {
        ctx.fillStyle = paint;
        ctx.fillRect(x, y, w, h);
      }
    }
  };
  paintTableFill(el.background ?? null, x0, y0, tableW, tableHeight);

  // Pass 1: fills + text bodies.
  for (const { cell, colX, rowY, cellW, cellH, ci, ri } of jobs) {
    paintTableFill(cell.fill, colX, rowY, cellW, cellH);
    // Text body — default run colour comes from the table style's tcTxStyle
    // (e.g. white header text on an accent fill); a run's explicit colour wins.
    // ECMA-376 Part 1 DrawingML CT_TableCell permits txBody to be absent. Empty
    // cells therefore emit no fabricated text run; the logical address on real
    // runs is enough for consumers to stop search matches at cell boundaries.
    if (cell.textBody) {
      if (tableTextRun) {
        textRunCell = { row: ri, column: ci };
      }
      const cellDefaultColor = cell.textColor ? hexToRgba(cell.textColor) : null;
      // ECMA-376 §21.1.3.17/§20.1.10.69 and the schema default: cells clip
      // horizontal overflow unless explicitly enabled. The interval belongs to
      // the physical merged-cell frame, inside the graphic-frame transform. Keep
      // the existing vertical overflow, row measurements and shape text policy.
      // Library policy: admit only an effective horizontal body (tcPr@vert,
      // already folded into body.vert, omitted or `horz`). How the physical x
      // interval maps onto rotated, stacked or unknown directions is not settled
      // here, so those bodies keep their previous unclipped paint and overlays.
      // DrawingML ST_TextVerticalType restricts xsd:token in both strict and
      // transitional schemas. For this single-word enum, collapse accepts only
      // U+0020/TAB/LF/CR at the ends; NBSP and internal whitespace cannot mean
      // `horz`. Apply this only to admission, leaving the body/paint policy intact.
      // Keep this tolerant for caller-supplied and previously parsed models too;
      // normalizing the shared direction would also change orientation dispatch.
      const rawVert = cell.textBody.vert;
      const isHorizontal = rawVert === undefined || typeof rawVert === 'string'
        && /^[ \t\r\n]*horz[ \t\r\n]*$/.test(rawVert);
      cellHorzOverflow = !isHorizontal ? undefined
        : cell.horzOverflow === 'overflow' ? 'overflow' : 'clip';
      ctx.save();
      try {
        if (cellHorzOverflow === 'clip') clipCanvasHorizontally(ctx, colX, cellW);
        renderTextBody(
          ctx,
          tableTextBody(cell.textBody),
          colX,
          rowY,
          cellW,
          cellH,
          scale,
          cellDefaultColor,
          0,
          false,
          false,
          '#000000',
          slideNumber,
          rc,
          tableTextRun,
        );
      } finally {
        ctx.restore();
      }
    }
  }

  // Pass 2: borders, painted on top of every fill.
  //
  // SPEC IS SILENT on adjacent-cell border conflict for DrawingML/PresentationML
  // tables (unlike WordprocessingML §17.4.66). When cell spacing is zero, two
  // neighbouring cells each contribute a line for the SHARED interior gridline;
  // the OLD code drew both, so the later-painted cell won by paint order (and a
  // translucent line doubled its ink). We now draw each gridline EXACTLY ONCE via
  // a single-ownership convention (the pptx leg / structural mirror of docx
  // PR #811, §17.4.66):
  //   • outer table edges → the single bordering cell draws its own line;
  //   • interior VERTICAL line → the physically-LEFT cell owns it (drawn as its
  //     physical-right edge), resolved against the physically-right neighbour;
  //   • interior HORIZONTAL line → the physically-ABOVE cell owns it (drawn as its
  //     bottom edge), resolved against the below neighbour.
  // The winner between two conflicting facing lines is chosen by our DEFINED,
  // deterministic rule (spec silent): null loses → wider wins → darker wins →
  // owner (reading-order-first) wins. See `resolveTableBorderConflict`.
  //
  // Crispness nudge (see crispOffset): an axis-aligned thin grid line whose
  // device-pixel width is odd straddles two device rows/cols on a DPR=1 display
  // (each ~50% ink → blurry). Snapping the cell edge perpendicular to the line
  // onto the nearest crisp device position centers an odd-width stroke on one
  // device row/col → crisp. `applyStroke` sets ctx.lineWidth to the logical
  // width, so we read it back to pick the right offset, and the snap delta is
  // derived from the edge's own coordinate (fractional-safe). Horizontal edges
  // (T/B) shift Y from their y; vertical edges (L/R) shift X from their x.
  // Diagonals can't be pixel-aligned, so untouched.
  const dpr = rc.dpr;

  // Resolve the job whose grid footprint covers slot (ri, ci), or null when the
  // slot is empty / out of range (an occupancy value of -1 or an index off-grid).
  const jobAt = (ri: number, ci: number): typeof jobs[number] | null => {
    if (ri < 0 || ri >= occupancy.length) return null;
    if (ci < 0 || ci >= numCols) return null;
    const idx = occupancy[ri][ci];
    return idx < 0 ? null : jobs[idx];
  };

  const strokeSeg = (
    stroke: Stroke,
    x1: number, y1: number, x2: number, y2: number,
  ) => {
    applyStroke(ctx, stroke, scale, {
      x: Math.min(x1, x2),
      y: Math.min(y1, y2),
      w: Math.max(1, Math.abs(x2 - x1)),
      h: Math.max(1, Math.abs(y2 - y1)),
    }, el.rotation);
    // ECMA-376 §20.1.8.42 `cmpd="dbl"`: table borders use the same
    // device-pixel-safe rail/gap/rail painter as the other OOXML renderers.
    // A fill-based painter is important for thin rules because two independently
    // snapped canvas strokes can collapse onto one device row. Preserve dashed
    // compound lines on the existing stroke path until their rail-wise dash
    // phase can be represented without changing authored semantics.
    if (
      stroke.cmpd === 'dbl'
      && !stroke.dashStyle
      && !(stroke.customDash?.length)
      && (x1 === x2 || y1 === y2)
    ) {
      ctx.fillStyle = ctx.strokeStyle;
      fillDoubleBorder(
        ctx,
        x1,
        y1,
        x2,
        y2,
        Math.max(0.5, emuToPx(stroke.width, scale)),
        dpr,
      );
      return;
    }
    // Vertical edge (x1===x2) nudges X; horizontal edge (y1===y2) nudges Y.
    const dx = x1 === x2 ? crispOffset(x1, ctx.lineWidth, dpr) : 0;
    const dy = y1 === y2 ? crispOffset(y1, ctx.lineWidth, dpr) : 0;
    ctx.beginPath();
    ctx.moveTo(x1 + dx, y1 + dy);
    ctx.lineTo(x2 + dx, y2 + dy);
    ctx.stroke();
  };

  for (const j of jobs) {
    const { cell, colX, rowY, cellW, cellH } = j;
    ctx.save();

    // Map this cell's LOGICAL border specs onto PHYSICAL sides. Under a
    // right-to-left table (§21.1.3.13) logical column 0 is at the physical RIGHT,
    // so a cell's logical-left border (`borderL`) paints on its physical-right
    // edge and its logical-right border (`borderR`) on its physical-left edge;
    // the physically-right NEIGHBOUR sits at a LOWER logical column. LTR keeps the
    // natural mapping.
    const physLeftSpec = el.rtl ? cell.borderR : cell.borderL;
    const physRightSpec = el.rtl ? cell.borderL : cell.borderR;
    const physLeftOuter = el.rtl ? j.ci + j.span === numCols : j.ci === 0;
    const physRightOuter = el.rtl ? j.ci === 0 : j.ci + j.span === numCols;
    // Grid slot of the physically-right neighbour (whose facing edge conflicts
    // with THIS cell's physical-right line). LTR: the column just past this span.
    // RTL: the column just before this cell's leading logical column.
    const physRightNbrCi = el.rtl ? j.ci - 1 : j.ci + j.span;
    // The neighbour's facing edge is the one on ITS physical-left side, i.e. its
    // logical-right spec under RTL, logical-left spec otherwise.
    const nbrPhysLeftSpec = (nb: TableCell): Stroke | null => (el.rtl ? nb.borderR : nb.borderL);

    // TOP: only the outer top row draws its own top; interior tops are owned by
    // the cell above (its bottom). Vertical direction is unaffected by RTL.
    if (j.ri === 0 && cell.borderT) {
      strokeSeg(cell.borderT, colX, rowY, colX + cellW, rowY);
    }
    // PHYSICAL LEFT: only the outer-left column draws its own left; interior
    // lefts are owned by the physically-left neighbour (its right edge).
    if (physLeftOuter && physLeftSpec) {
      strokeSeg(physLeftSpec, colX, rowY, colX, rowY + cellH);
    }
    // BOTTOM: outer bottom → own spec; interior → resolve THIS cell's bottom
    // against the below neighbour's top and draw the single winner.
    {
      if (j.lastRi === el.rows.length - 1) {
        const spec = cell.borderB;
        if (spec) strokeSeg(spec, colX, rowY + cellH, colX + cellW, rowY + cellH);
      } else {
        // #824 — the shared horizontal edge below this cell may face SEVERAL finer
        // below-cells with differing borderT (this cell is wider via gridSpan).
        // Subdivide the edge at those cells' column boundaries and resolve EACH
        // sub-segment against its OWN below neighbour, drawing a per-segment winner
        // rather than resolving the whole span against the origin slot alone.
        const belowRi = j.lastRi + 1;
        const y = rowY + cellH;
        const cEndMax = Math.min(j.ci + j.span, numCols);
        let cj = j.ci;
        while (cj < cEndMax) {
          const idx = occupancy[belowRi][cj];
          let cEnd = cj + 1;
          while (cEnd < cEndMax && occupancy[belowRi][cEnd] === idx) cEnd++;
          const below = jobAt(belowRi, cj);
          const spec = resolveTableBorderConflict(cell.borderB, below ? below.cell.borderT : null);
          if (spec) {
            const segLeft = spannedLeft(cj, cEnd - cj);
            const segW = spannedWidth(cj, cEnd - cj);
            strokeSeg(spec, segLeft, y, segLeft + segW, y);
          }
          cj = cEnd;
        }
      }
    }
    // PHYSICAL RIGHT: outer-right → own spec; interior → resolve THIS cell's
    // physical-right line against the physically-right neighbour's facing edge
    // (so each shared vertical line is drawn once).
    {
      if (physRightOuter) {
        const spec = physRightSpec;
        if (spec) strokeSeg(spec, colX + cellW, rowY, colX + cellW, rowY + cellH);
      } else {
        // #824 — a rowSpan cell's physical-right edge may face SEVERAL finer
        // right-neighbours down the rows it spans, with differing facing specs.
        // Subdivide the edge at those neighbours' row boundaries and resolve EACH
        // sub-segment against its OWN neighbour's facing (physical-left) edge.
        const x = colX + cellW;
        let rj = j.ri;
        while (rj <= j.lastRi) {
          const idx = occupancy[rj][physRightNbrCi];
          let rEnd = rj;
          while (rEnd + 1 <= j.lastRi && occupancy[rEnd + 1][physRightNbrCi] === idx) rEnd++;
          const right = jobAt(rj, physRightNbrCi);
          const spec = resolveTableBorderConflict(physRightSpec, right ? nbrPhysLeftSpec(right.cell) : null);
          if (spec) strokeSeg(spec, x, rowTop[rj], x, rowTop[rEnd] + rowHeights[rEnd]);
          rj = rEnd + 1;
        }
      }
    }

    // Diagonal borders: top-left→bottom-right and bottom-left→top-right. These
    // are never shared between cells, so they are always drawn per-cell (and are
    // not pixel-aligned).
    if (cell.diagonalTL) {
      applyStroke(ctx, cell.diagonalTL, scale, {
        x: colX, y: rowY, w: cellW, h: cellH,
      }, el.rotation);
      ctx.beginPath();
      ctx.moveTo(colX, rowY);
      ctx.lineTo(colX + cellW, rowY + cellH);
      ctx.stroke();
    }
    if (cell.diagonalTR) {
      applyStroke(ctx, cell.diagonalTR, scale, {
        x: colX, y: rowY, w: cellW, h: cellH,
      }, el.rotation);
      ctx.beginPath();
      ctx.moveTo(colX + cellW, rowY);
      ctx.lineTo(colX, rowY + cellH);
      ctx.stroke();
    }
    ctx.restore();
  }
  ctx.restore();
}
