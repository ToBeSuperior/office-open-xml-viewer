import { expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';
import type { DocumentLayout, NoteLayout, PaintNode } from './layout/types.js';

// WORD_PER_COLUMN_FOOTNOTE_BANDS: below compatibility mode 15 each column of a
// multi-column section owns the notes referenced in it. Exact 10pt lines and
// a 10pt formatted separator story make every band an exact line count:
// 180pt text width in two 80pt columns (gap 20pt), 100pt body, 10 lines each.
const context = {
  font: '10px Arial', letterSpacing: '0px', fontKerning: 'none',
  measureText: (text: string) => ({
    width: [...text].length * 4,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
  }),
} as unknown as CanvasRenderingContext2D;

function textRun(text: string, extra: Record<string, unknown> = {}) {
  return {
    type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Arial', fontFamilyEastAsia: '', isLink: false,
    background: null, vertAlign: null, hyperlink: null, ...extra,
  };
}

function paragraph(text: string, noteId?: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: { value: 10, rule: 'exact', explicit: true },
    numbering: null, tabStops: [],
    runs: [textRun(text), ...(noteId ? [textRun(noteId, { noteRef: { kind: 'footnote', id: noteId } })] : [])],
    defaultFontSize: 10, defaultFontFamily: 'Arial', widowControl: false,
  } as unknown as DocParagraph;
}

const label = (prefix: string, index: number) => prefix + String(index).padStart(2, '0');

function model(
  compatibilityMode: number,
  references: Readonly<Record<number, string>>,
  notes: Readonly<Record<string, number>>,
  bodyLines = 40,
): DocxDocumentModel {
  return {
    section: {
      pageWidth: 324, pageHeight: 244, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false, sectionStart: 'nextPage',
      columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
    } as unknown as SectionProps,
    body: Array.from({ length: bodyLines }, (_, index) => paragraph(label('L', index + 1), references[index + 1])) as unknown as BodyElement[],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { Arial: 'swiss' },
    settings: { compatibilityMode },
    footnotes: Object.entries(notes).map(([id, count]) => ({
      id,
      content: Array.from({ length: count }, (_, index) => paragraph(label(`N${id}-`, index + 1))) as unknown as BodyElement[],
    })),
    __noteLayoutSettings: {
      footnoteSeparator: 'short',
      footnoteContinuationSeparator: 'full',
      footnoteSeparatorParagraph: { ...paragraph(''), runs: [] },
      footnoteContinuationSeparatorParagraph: { ...paragraph(''), runs: [] },
    },
  } as unknown as DocxDocumentModel;
}

const text = (block: PaintNode): string => block.kind === 'paragraph'
  ? block.lines.flatMap((line) => line.placements.flatMap((run) => run.kind === 'text' ? [run.text] : [])).join('')
  : '';

/** Page 1 per column: body labels and the note lines painted in it. */
function columns(layout: DocumentLayout) {
  const page = layout.pages[0]!;
  const columnOf = (xPt: number) => (xPt < 162 ? 0 : 1);
  const result = [{ body: [] as string[], notes: [] as string[] }, { body: [] as string[], notes: [] as string[] }];
  for (const node of page.layers.body) result[columnOf(node.flowBounds.xPt)]!.body.push(text(node).slice(0, 3));
  for (const note of page.layers.notes.filter((node): node is NoteLayout => node.kind === 'note')) {
    result[columnOf(note.flowBounds.xPt)]!.notes.push(...note.story.blocks.map(text));
  }
  return result.map(({ body, notes }) => ({ body: `${body[0]}..${body.at(-1)}`, notes }));
}

const render = (document: DocxDocumentModel): DocumentLayout => layoutDocument(document,
  createLayoutServices(document, { measureContext: context, allowFootnoteContinuation: true }),
  { currentDateMs: 0 });

it('gives each column below mode 15 its own band, leaving a column without references whole', () => {
  // Note 1 (2 lines) on L03: column 1 keeps 7 lines above its 30pt band,
  // column 2 keeps all 10.
  expect(columns(render(model(11, { 3: '1' }, { 1: 2 })))).toEqual([
    { body: 'L01..L07', notes: ['N1-01', 'N1-02'] },
    { body: 'L08..L17', notes: [] },
  ]);
  // A reference in column 2 shortens only column 2.
  expect(columns(render(model(11, { 12: '1' }, { 1: 2 })))).toEqual([
    { body: 'L01..L10', notes: [] },
    { body: 'L11..L17', notes: ['N1-01', 'N1-02'] },
  ]);
});

it('continues a note at the top of the next column band, ahead of that column\'s own notes', () => {
  // Note 1 (6 lines) on L07 keeps its reference with two lines; four
  // continue in column 2 (with the continuation band) before note 2.
  expect(columns(render(model(11, { 7: '1', 9: '2' }, { 1: 6, 2: 1 })))).toEqual([
    { body: 'L01..L07', notes: ['N1-01', 'N1-02'] },
    { body: 'L08..L11', notes: ['N1-03', 'N1-04', 'N1-05', 'N1-06', 'N2-01'] },
  ]);
});

it('keeps the page-wide band in compatibility mode 15', () => {
  const [first, second] = columns(render(model(15, { 3: '1' }, { 1: 2 })));
  expect(first!.body).toBe('L01..L07');
  expect(second!.body).toBe('L08..L14');
});

it('places document-end notes after the text in the column where it ends', () => {
  // 14 lines end in column 2 above the full column 1; the endnote follows L14
  // there instead of being refused below the deeper column 1.
  const document = {
    ...model(11, {}, {}),
    body: Array.from({ length: 14 }, (_, index) => {
      const paragraphModel = paragraph(label('L', index + 1));
      return index === 2
        ? { ...paragraphModel, runs: [...paragraphModel.runs, textRun('e', { noteRef: { kind: 'endnote', id: '1' } })] }
        : paragraphModel;
    }) as unknown as BodyElement[],
    endnotes: [{ id: '1', content: [paragraph('E1-01')] as unknown as BodyElement[] }],
  } as unknown as DocxDocumentModel;
  const layout = render(document);
  expect(layout.diagnostics).toEqual([]);
  const endnotes = layout.pages[0]!.layers.notes.filter((node): node is NoteLayout => node.kind === 'note');
  expect(endnotes.map((note) => [note.story.blocks.map(text).join(''), note.flowBounds.xPt, note.flowBounds.yPt]))
    .toEqual([['E1-01', 172, 112]]);
});

it('keeps a continuation notice with the column band that reserved it', () => {
  // Reserved 10pt separator stories and a 10pt continuation notice. Note 1
  // (6 lines) on L07 continues from column 1, whose band reserves the notice;
  // note 2 on L09 joins column 2's band. The notice stays in column 1.
  const control = textRun('');
  const storyParagraph = { ...paragraph(''), runs: [], paragraphMarkFontFacts: { fontFamily: 'Arial', fontSize: 10 } };
  const story = { paragraph: { paragraph: storyParagraph, contentMark: { run: control } } };
  const rule = (mark: 'short' | 'full') => ({ rule: { mark, control: { run: control } }, ...story });
  const document = model(11, { 7: '1', 9: '2' }, { 1: 6, 2: 1 });
  (document as unknown as Record<string, unknown>).__noteLayoutSettings = {
    footnoteSeparator: 'short', footnoteContinuationSeparator: 'full',
    reservedSeparators: { footnote: {
      separator: rule('short'), continuationSeparator: rule('full'), continuationNotice: story,
    } },
  };
  const layout = render(document);
  const notes = layout.pages[0]!.layers.notes.filter((node): node is NoteLayout => node.kind === 'note');
  expect(notes.map((note) => [note.story.blocks.map(text).join(''), note.flowBounds.xPt < 162, note.trailing?.role]))
    .toEqual([
      ['N1-01', true, 'continuationNotice'],
      ['N1-02N1-03N1-04N1-05N1-06', false, undefined],
      ['N2-01', false, undefined],
    ]);
});

it('continues a note outliving the body into the next column before a new page', () => {
  // Three body lines end in column 1 with a 20-line note on L03: column 1
  // keeps a six-line head, column 2's band the next nine lines and page 2
  // the last five.
  const layout = render(model(11, { 3: '1' }, { 1: 20 }, 3));
  const lineCounts = layout.pages.map((page) => page.layers.notes
    .filter((node): node is NoteLayout => node.kind === 'note')
    .map((note) => [note.flowBounds.xPt < 162 ? 1 : 2, note.story.blocks.length]));
  expect(lineCounts).toEqual([[[1, 6], [2, 9]], [[1, 5]]]);
});

it('limits a column-2 endnote only by column 2 footnotes', () => {
  // Column 1 holds L01..L03 above note 1's band (top 102pt); the text ends
  // in column 2 at L08 (122pt) and the endnote fits below it there.
  const document = model(11, { 1: '1' }, { 1: 6 }, 8);
  const body = document.body as unknown as DocParagraph[];
  body[7] = { ...body[7]!, runs: [...body[7]!.runs, textRun('e', { noteRef: { kind: 'endnote', id: '1' } })] } as DocParagraph;
  (document as unknown as { endnotes: unknown }).endnotes = [{ id: '1', content: [paragraph('E1-01')] }];
  const layout = render(document);
  expect(layout.diagnostics).toEqual([]);
  const endnotes = layout.pages.flatMap((page) => page.layers.notes)
    .filter((node): node is NoteLayout => node.kind === 'note' && node.source.story === 'endnote');
  expect(endnotes.map((note) => note.story.blocks.map(text).join(''))).toEqual(['E1-01']);
});
