// Renders the guide content (content1-4.js) to a styled .docx and a matching GitHub Markdown file.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, LevelFormat,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, PageBreak,
  Header, Footer, PageNumber, Bookmark, InternalHyperlink, ExternalHyperlink,
  TabStopType, VerticalAlign,
} = require('docx');

const OUT_DIR = process.argv[2] || path.resolve(__dirname, '..');
const BASENAME = 'Season-15-Warlock-Blazing-Scream-Guide';

const blocks = [
  ...require('./content1'),
  ...require('./content2'),
  ...require('./content3'),
  ...require('./content4'),
];

// ---------- theme ----------
const C = {
  accent: '7A1F1F', accentDark: '4A1010', text: '222222', muted: '666666',
  headFill: '5C1A1B', zebra: 'FAF5F2', grid: 'D9CFC7', kvFill: 'F3E9E4',
  tip: ['E9F5EC', '2E7D32'], warn: ['FFF4E5', 'D9822B'], info: ['EAF2FB', '2F6DB5'],
};
const FONT = 'Calibri';
const PAGE_W = 12240, PAGE_H = 15840, MARGIN = 1080;
const CONTENT_W = PAGE_W - 2 * MARGIN; // 10080

// ---------- inline markup: **bold**, *italic* ----------
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const tok = m[0];
    if (tok.startsWith('**')) out.push(new TextRun({ text: tok.slice(2, -2), bold: true, ...base }));
    else out.push(new TextRun({ text: tok.slice(1, -1), italics: true, ...base }));
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out;
}

// ---------- numbering ----------
let listInstance = 0;
const numbering = {
  config: [
    { reference: 'bullets', levels: [
      { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 240 } } } },
      { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 240 } } } },
    ] },
    { reference: 'numbers', levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 400, hanging: 320 } }, run: { bold: true, color: C.accent } } },
      { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 800, hanging: 240 } } } },
    ] },
    { reference: 'checks', levels: [
      { level: 0, format: LevelFormat.BULLET, text: '☐', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 400, hanging: 320 } }, run: { font: 'Segoe UI Symbol', size: 22 } } },
    ] },
  ],
};

// ---------- block renderers (docx) ----------
const SP = { after: 80 };

function para(text, opts = {}) {
  return new Paragraph({ spacing: SP, ...opts, children: runs(text, opts.runBase || {}) });
}

function bulletParas(items, extra = {}) {
  const out = [];
  for (const it of items) {
    if (typeof it === 'string') {
      out.push(new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 60 }, ...extra, children: runs(it) }));
    } else if (it.sub) {
      for (const s of it.sub) out.push(new Paragraph({ numbering: { reference: 'bullets', level: 1 }, spacing: { after: 40 }, ...extra, children: runs(s) }));
    } else if (it.text) {
      out.push(new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 60 }, ...extra, children: runs(it.text) }));
    }
  }
  return out;
}

function numberedParas(items, extra = {}) {
  const inst = ++listInstance;
  const out = [];
  for (const it of items) {
    if (typeof it === 'string') {
      out.push(new Paragraph({ numbering: { reference: 'numbers', level: 0, instance: inst }, spacing: { after: 70 }, ...extra, children: runs(it) }));
    } else if (it.sub) {
      for (const s of it.sub) out.push(new Paragraph({ numbering: { reference: 'numbers', level: 1, instance: inst }, spacing: { after: 40 }, ...extra, children: runs(s) }));
    }
  }
  return out;
}

function checkParas(items) {
  return items.map(it => new Paragraph({ numbering: { reference: 'checks', level: 0 }, spacing: { after: 70 }, children: runs(it) }));
}

const cellBorders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: C.grid },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: C.grid },
  left: { style: BorderStyle.SINGLE, size: 4, color: C.grid },
  right: { style: BorderStyle.SINGLE, size: 4, color: C.grid },
};

function fixWidths(widths) {
  const sum = widths.reduce((a, b) => a + b, 0);
  if (sum === CONTENT_W) return widths;
  const scaled = widths.map(w => Math.floor(w * CONTENT_W / sum));
  scaled[scaled.length - 1] += CONTENT_W - scaled.reduce((a, b) => a + b, 0);
  return scaled;
}

function cellParas(text, size, bold, color, keepNext = false) {
  const parts = Array.isArray(text) ? text : [text];
  return parts.map(t => new Paragraph({ spacing: { after: 0 }, keepNext, children: runs(String(t), { size, ...(bold ? { bold: true } : {}), ...(color ? { color } : {}) }) }));
}

function table(b, keep = false) {
  const widths = fixWidths(b.widths);
  const size = b.small ? 17 : 19; // half-points
  const rows = [];
  if (b.head) {
    rows.push(new TableRow({ tableHeader: true, cantSplit: true, children: b.head.map((h, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA }, borders: cellBorders,
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: C.headFill },
      margins: { top: 70, bottom: 70, left: 100, right: 100 }, verticalAlign: VerticalAlign.CENTER,
      children: cellParas(h, size, true, 'FFFFFF', keep),
    })) }));
  }
  b.rows.forEach((r, ri) => {
    rows.push(new TableRow({ cantSplit: true, children: r.map((c, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA }, borders: cellBorders,
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: ri % 2 === 1 ? C.zebra : 'FFFFFF' },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: cellParas(c, size, false, undefined, keep),
    })) }));
  });
  return [new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths, rows }),
          new Paragraph({ spacing: { after: 120 }, keepNext: keep, children: [] })];
}

function kv(b) {
  const widths = [2000, CONTENT_W - 2000];
  const last = b.rows.length - 1;
  const rows = b.rows.map(([k, v], ri) => new TableRow({ cantSplit: true, children: [
    new TableCell({ width: { size: widths[0], type: WidthType.DXA }, borders: cellBorders,
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: C.kvFill },
      margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: cellParas(k, 19, true, C.accentDark, ri < last) }),
    new TableCell({ width: { size: widths[1], type: WidthType.DXA }, borders: cellBorders,
      margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: cellParas(v, 19, false, undefined, ri < last) }),
  ] }));
  return [new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths, rows }),
          new Paragraph({ spacing: { after: 120 }, children: [] })];
}

function callout(b) {
  const [fill, line] = C[b.kind] || C.info;
  const children = [];
  if (b.title) children.push(new Paragraph({ spacing: { after: 80 }, keepNext: true,
    children: runs('**' + b.title + '**', { color: line, size: 22 }) }));
  for (const item of b.body) {
    if (typeof item === 'string') children.push(new Paragraph({ spacing: { after: 60 }, children: runs(item) }));
    else if (item.bullets) children.push(...bulletParas(item.bullets));
    else if (item.numbered) children.push(...numberedParas(item.numbered));
  }
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const cell = new TableCell({
    width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill },
    borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 36, color: line } },
    margins: { top: 140, bottom: 100, left: 220, right: 220 },
    children,
  });
  return [new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
            rows: [new TableRow({ cantSplit: true, children: [cell] })] }),
          new Paragraph({ spacing: { after: 140 }, children: [] })];
}

function links(b) {
  return b.items.map(([title, url]) => new Paragraph({
    numbering: { reference: 'bullets', level: 0 }, spacing: { after: 40 },
    children: [
      new TextRun({ text: title + ' — ', size: 19 }),
      new ExternalHyperlink({ link: url, children: [new TextRun({ text: url, style: 'Hyperlink', size: 17 })] }),
    ],
  }));
}

function partBanner(text) {
  return new Paragraph({
    spacing: { before: 0, after: 200 },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: C.headFill },
    indent: { left: 0, right: 0 },
    children: [new TextRun({ text: '  ' + text.toUpperCase(), bold: true, color: 'FFFFFF', size: 26, characterSpacing: 20 })],
  });
}

// ---------- title + contents ----------
function titlePage() {
  const out = [];
  out.push(new Paragraph({ spacing: { before: 1800, after: 120 }, children: [new TextRun({ text: 'DIABLO IV  ·  SEASON 15  ·  SEASON OF HELL\'S LEGACY', bold: true, color: C.accent, size: 22, characterSpacing: 30 })] }));
  out.push(new Paragraph({ heading: HeadingLevel.TITLE, spacing: { after: 120 }, children: [new TextRun({ text: 'Warlock — Blazing Scream' })] }));
  out.push(new Paragraph({ spacing: { after: 480 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.accent, space: 8 } },
    children: [new TextRun({ text: 'Level 70 → Endgame Playbook', size: 36, color: C.muted })] }));
  out.push(...kv({ rows: [
    ['Patch', '3.2.1 (+ Hotfix 4) · updated 26 September 2026'],
    ['Starting point', 'Level 70 · Lord of Hatred campaign complete · **Rare items only** · **Splinter of Destruction**'],
    ['Build reference', 'mobalytics.gg/diablo-4/builds/warlock-blazing-scream'],
    ['Data used', 'Maxroll\'s decoded Blazing Scream planner (5 variants), D4Guides\' S-tier build, the Season 15 game data (per-slot stat pools, aspect slot rules, temper manuals), and Season 15 system guides'],
  ] }));
  out.push(new Paragraph({ spacing: { before: 240, after: 80 }, children: runs('**How to use this document**', { color: C.accent, size: 24 }) }));
  out.push(...numberedParas([
    'Read **§1** (the short answer) and **§4** (the step-by-step plan). They tell you exactly what to do next.',
    'Every time an item drops, check it against **§5** (stat priority by slot).',
    'Use **§6–21** as reference: season systems, the build phases, Paragon, farming and crafting.',
  ]));
  out.push(new Paragraph({ children: [new PageBreak()] }));
  return out;
}

function contentsPage() {
  const out = [new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Contents')] })];
  for (const b of blocks) {
    if (b.t === 'part') out.push(new Paragraph({ spacing: { before: 160, after: 40 }, children: [new TextRun({ text: b.text, bold: true, color: C.accent, size: 21 })] }));
    if (b.t === 'h1') out.push(new Paragraph({ spacing: { after: 30 }, indent: { left: 280 }, children: [
      new InternalHyperlink({ anchor: b.id, children: [new TextRun({ text: b.text, style: 'Hyperlink', size: 20 })] }),
    ] }));
  }
  out.push(new Paragraph({ children: [new PageBreak()] }));
  return out;
}

// ---------- body ----------
function renderBody() {
  const out = [];
  blocks.forEach((b, i) => {
    const next = blocks[i + 1] || {};
    switch (b.t) {
      case 'part': out.push(partBanner(b.text)); break;
      case 'h1': out.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new Bookmark({ id: b.id, children: [new TextRun(b.text)] })] })); break;
      case 'h2': out.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(b.text)] })); break;
      case 'h3': out.push(new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(b.text)] })); break;
      case 'p': out.push(para(b.text)); break;
      case 'bullets': out.push(...bulletParas(b.items)); break;
      case 'numbered': out.push(...numberedParas(b.items)); break;
      case 'check': out.push(...checkParas(b.items)); break;
      case 'table': out.push(...table(b, next.t === 'kv')); break;
      case 'kv': out.push(...kv(b)); break;
      case 'callout': out.push(...callout(b)); break;
      case 'links': out.push(...links(b)); break;
      case 'pagebreak': out.push(new Paragraph({ children: [new PageBreak()] })); break;
      case 'spacer': out.push(new Paragraph({ spacing: { after: 160 }, children: [] })); break;
      default: throw new Error('unknown block ' + b.t);
    }
  });
  return out;
}

const doc = new Document({
  creator: 'Claude',
  title: 'Season 15 Warlock – Blazing Scream: Level 70 → Endgame Playbook',
  description: 'Step-by-step Diablo IV Season 15 guide for the Blazing Scream Warlock',
  styles: {
    default: { document: { run: { font: FONT, size: 21, color: C.text } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', next: 'Normal', run: { font: FONT, size: 64, bold: true, color: C.accentDark } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 32, bold: true, color: C.accent },
        paragraph: { spacing: { before: 320, after: 140 }, keepNext: true, keepLines: true, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.grid, space: 4 } } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 26, bold: true, color: C.accentDark },
        paragraph: { spacing: { before: 240, after: 100 }, keepNext: true, keepLines: true, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 23, bold: true, color: C.text },
        paragraph: { spacing: { before: 200, after: 80 }, keepNext: true, keepLines: true, outlineLevel: 2 } },
    ],
    characterStyles: [
      { id: 'Hyperlink', name: 'Hyperlink', basedOn: 'DefaultParagraphFont', run: { color: '1F5FA8', underline: {} } },
    ],
  },
  numbering,
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: 'S15 Warlock – Blazing Scream Playbook', size: 16, color: C.muted })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: 'Page ', size: 16, color: C.muted }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: C.muted }),
        new TextRun({ text: ' of ', size: 16, color: C.muted }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: C.muted }),
      ] })] }) },
    children: [...titlePage(), ...contentsPage(), ...renderBody()],
  }],
});

// ---------- markdown ----------
function slug(h) {
  return h.trim().toLowerCase().replace(/[^\p{L}\p{N}\- ]/gu, '').replace(/ /g, '-');
}
function mdEsc(s) { return String(s).replace(/\|/g, '\\|'); }
function mdList(items, numbered) {
  const lines = []; let n = 0;
  for (const it of items) {
    if (typeof it === 'string') lines.push(numbered ? `${++n}. ${it}` : `- ${it}`);
    else if (it.sub) it.sub.forEach(s => lines.push(`${numbered ? '   ' : '  '}- ${s}`));
  }
  return lines.join('\n');
}
function renderMarkdown() {
  const L = [];
  L.push('# Season 15 Warlock – Blazing Scream: Level 70 → Endgame Playbook', '');
  L.push('**Patch 3.2.1 (+ Hotfix 4) · updated 26 Sep 2026** · Starting point: level 70, Lord of Hatred campaign done, **Rare items only**, **Splinter of Destruction**.', '');
  L.push('Build reference: <https://mobalytics.gg/diablo-4/builds/warlock-blazing-scream>. A formatted Word version of this guide is in this repo: `Season-15-Warlock-Blazing-Scream-Guide.docx`.', '');
  L.push('## Contents', '');
  for (const b of blocks) {
    if (b.t === 'part') L.push(`- **${b.text}**`);
    if (b.t === 'h1') L.push(`  - [${b.text}](#${slug(b.text)})`);
  }
  L.push('');
  for (const b of blocks) {
    switch (b.t) {
      case 'part': L.push('---', '', `# ${b.text}`, ''); break;
      case 'h1': L.push(`## ${b.text}`, ''); break;
      case 'h2': L.push(`### ${b.text}`, ''); break;
      case 'h3': L.push(`#### ${b.text}`, ''); break;
      case 'p': L.push(b.text, ''); break;
      case 'bullets': L.push(mdList(b.items, false), ''); break;
      case 'numbered': L.push(mdList(b.items, true), ''); break;
      case 'check': L.push(b.items.map(i => `- [ ] ${i}`).join('\n'), ''); break;
      case 'table': {
        L.push('| ' + b.head.map(mdEsc).join(' | ') + ' |');
        L.push('|' + b.head.map(() => '---').join('|') + '|');
        b.rows.forEach(r => L.push('| ' + r.map(c => mdEsc(Array.isArray(c) ? c.join('<br>') : c)).join(' | ') + ' |'));
        L.push(''); break;
      }
      case 'kv': L.push(b.rows.map(([k, v]) => `- **${k}:** ${v}`).join('\n'), ''); break;
      case 'callout': {
        if (b.title) L.push(`> **${b.title}**`, '>');
        for (const item of b.body) {
          if (typeof item === 'string') L.push(`> ${item}`, '>');
          else if (item.bullets) { L.push(mdList(item.bullets, false).split('\n').map(x => `> ${x}`).join('\n'), '>'); }
          else if (item.numbered) { L.push(mdList(item.numbered, true).split('\n').map(x => `> ${x}`).join('\n'), '>'); }
        }
        if (L[L.length - 1] === '>') L.pop();
        L.push(''); break;
      }
      case 'links': L.push(b.items.map(([t, u]) => `- [${t}](${u})`).join('\n'), ''); break;
      case 'pagebreak': case 'spacer': break;
    }
  }
  return L.join('\n').replace(/\n{3,}/g, '\n\n') + '\n';
}

// docx-js gives every Bookmark the same w:id; renumber start/end pairs so ids are unique.
async function fixBookmarkIds(buf) {
  const JSZip = require('jszip');
  const zip = await JSZip.loadAsync(buf);
  let xml = await zip.file('word/document.xml').async('string');
  let n = 0, current = 0;
  xml = xml.replace(/<w:bookmark(Start|End)\b([^>]*?)w:id="\d+"([^>]*)\/>/g, (m, kind, a, b) => {
    if (kind === 'Start') current = ++n;
    return `<w:bookmark${kind}${a}w:id="${current}"${b}/>`;
  });
  zip.file('word/document.xml', xml);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

Packer.toBuffer(doc).then(fixBookmarkIds).then(buf => {
  const docxPath = path.join(OUT_DIR, BASENAME + '.docx');
  fs.writeFileSync(docxPath, buf);
  const mdPath = path.join(OUT_DIR, BASENAME + '.md');
  fs.writeFileSync(mdPath, renderMarkdown());
  console.log('wrote', docxPath, buf.length, 'bytes');
  console.log('wrote', mdPath);
});
