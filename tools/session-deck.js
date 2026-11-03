// Session decks 2+ · CELPIP Prep Course · Elevate with Ms. Maria · Learnovate Qatar
// The slides are read from index.html (S<N>_SLIDES, S<N>_GUIDE and the live-tool data), so the
// PowerPoint always matches the built-in website slides.
// Build:  NODE_PATH=<folder with pptxgenjs, react-icons, react, react-dom, sharp>/node_modules \
//         APPLY_THEME=<pptx skill>/scripts/apply_theme.js node tools/session-deck.js 2 decks/session-02.pptx
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const pptxgen = require('pptxgenjs');
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const sharp = require('sharp');
const Lu = require('react-icons/lu');
const { applyTheme } = require(process.env.APPLY_THEME || './apply_theme.js');

const N = +process.argv[2];
const OUT = process.argv[3] || `session-${String(N).padStart(2, '0')}.pptx`;

// ---------- read the slide data from index.html ----------
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const script = html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>'));
const grab = (startMark, endMark) => { const a = script.indexOf(startMark), b = script.indexOf(endMark, a); if (a < 0 || b < 0) throw new Error('Missing ' + startMark); return script.slice(a, b); };
const sessionsSrc = grab('const SESSIONS=[', '\nconst QUIZ=');
const toolsSrc = grab('/* ---------- Sessions 2 & 3 · live widgets', '\n/* ---------- diagnostic summary');
const ctx = { LIVE_WIDGETS: {}, esc: s => String(s) };
vm.createContext(ctx);
vm.runInContext(sessionsSrc.replace('const SESSIONS=', 'this.SESSIONS=') + '\n' + toolsSrc + `
this.D={SLIDES:typeof S${N}_SLIDES!=='undefined'?S${N}_SLIDES:null,GUIDE:typeof S${N}_GUIDE!=='undefined'?S${N}_GUIDE:null,
THINK_CLIP,ANNO_CLIP,POLL_ITEMS,NEWS_ITEM,NEWS_CATS,TRACK_CLIP,ATT_TEXT,ATT_PENS};`, ctx);
const D = ctx.D, SESSION = ctx.SESSIONS[N - 1];
if (!D.SLIDES) throw new Error(`No S${N}_SLIDES in index.html`);

const THEME = { name: 'Elevate Pastel', headFontFace: 'Poppins', bodyFontFace: 'Poppins',
  colors: { dk1: '14213D', lt1: 'FFFFFF', dk2: '5B6782', lt2: 'F6F7FB', accent1: '1A5FCC', accent2: '6A2BC4', accent3: 'C93D1A',
    accent4: '007F7F', accent5: '008F55', accent6: 'E0A800', hlink: '1A5FCC', folHlink: '6A2BC4' } };
const SK = { L: ['EAF2FF', '1A5FCC'], R: ['F3ECFE', '6A2BC4'], W: ['FFEFEA', 'C93D1A'], S: ['E5F7F6', '007F7F'], O: ['E7F7EF', '008F55'], G: ['FFF7DB', '8A6A00'] };
const ROT = ['L', 'R', 'W', 'S', 'O'];
const ICON = { L: 'LuHeadphones', R: 'LuBookOpen', W: 'LuPenLine', S: 'LuMic', O: 'LuCompass' };
const NAVY = '14213D', MUTED = '5B6782', LINE = 'E6E9F2';
const MARK = { d: ['FDE2E2', '9B1C1C'], c: ['DDF3E4', '17603A'], s: ['FFF1BF', '7A5A00'], p: ['DDF3E4', '17603A'], n: ['FDE2E2', '9B1C1C'], h: ['FFF1BF', '7A5A00'] };
const ATT_MARK = { p: MARK.p, n: MARK.n, h: MARK.h, c: ['E3EEFF', '1A4FA8'] };
const clean = t => String(t).replace(/<[^>]+>/g, '');

async function icon(name, color, size = 256) {
  const svg = renderToStaticMarkup(React.createElement(Lu[name], { color: '#' + color, size: String(size), strokeWidth: 1.75 }));
  return 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
}
async function gradient(w, h, stops) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">${stops.map(([o, c]) => `<stop offset="${o}" stop-color="#${c}"/>`).join('')}</linearGradient>
    <radialGradient id="r" cx="0.85" cy="0.15" r="0.6"><stop offset="0" stop-color="#ffffff" stop-opacity="0.55"/><stop offset="1" stop-color="#ffffff" stop-opacity="0"/></radialGradient></defs>
    <rect width="100%" height="100%" fill="url(#g)"/><rect width="100%" height="100%" fill="url(#r)"/></svg>`;
  return 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
}
const shadow = () => ({ type: 'outer', color: '14213D', opacity: 0.08, blur: 14, offset: 3, angle: 90 });

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.author = 'Ms. Maria · Learnovate Qatar';
  pres.title = `Session ${N} · ${SESSION.title}`;
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const titleBg = await gradient(1920, 1080, [[0, 'EAF2FF'], [0.5, 'F3ECFE'], [1, 'E7F7EF']]);
  const closeBg = await gradient(1920, 1080, [[0, 'EAF2FF'], [1, 'E7F7EF']]);
  pres.defineSlideMaster({ title: 'TITLE', background: { data: titleBg },
    objects: [
      { placeholder: { options: { name: 'title', type: 'title', x: 0.75, y: 2.3, w: 7.6, h: 2.3, fontSize: 42, bold: true, color: C.text1, valign: 'top', align: 'left', margin: 0 }, text: '' } },
      { placeholder: { options: { name: 'body', type: 'body', x: 0.75, y: 4.75, w: 7.6, h: 0.6, fontSize: 18, color: C.text2, align: 'left', margin: 0 }, text: '' } },
      { text: { text: 'ELEVATE WITH MS. MARIA  ·  LEARNOVATE QATAR', options: { x: 0.75, y: 6.6, w: 8, h: 0.35, fontSize: 10, bold: true, color: C.text2, charSpacing: 3, margin: 0 } } }] });
  const contentObjects = () => [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.75, y: 0.88, w: 11.8, h: 0.8, fontSize: 32, bold: true, color: C.text1, valign: 'top', align: 'left', margin: 0 }, text: '' } },
    { text: { text: `CELPIP Prep Course  ·  Session ${N}`, options: { x: 0.75, y: 6.95, w: 6, h: 0.3, fontSize: 9, color: C.text2, margin: 0 } } }];
  const num = { x: 12.08, y: 6.95, w: 0.5, h: 0.3, fontSize: 9, color: '5B6782', align: 'right' };
  pres.defineSlideMaster({ title: 'CONTENT', background: { color: 'FFFFFF' }, objects: contentObjects(), slideNumber: num });
  pres.defineSlideMaster({ title: 'CLOSING', background: { data: closeBg }, objects: contentObjects(), slideNumber: { ...num } });

  // ---------- helpers ----------
  const T = (s, text, o) => s.addText(text, { color: NAVY, margin: 0, isTextBox: true, valign: 'middle', ...o });
  const title = (s, t) => s.addText(t, { placeholder: 'title' });
  const eyebrow = (s, text, k) => T(s, text.toUpperCase(), { x: 0.75, y: 0.5, w: 11, h: 0.32, fontSize: 11, bold: true, color: SK[k][1], charSpacing: 3, valign: 'top' });
  const card = (s, x, y, w, h, fill, name, sh, line) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.2, fill: { color: fill }, line: { color: line || fill, width: line ? 1 : 0.75 }, objectName: name, ...(sh ? { shadow: shadow() } : {}) });
  const chip = async (s, x, y, d, k, ic, name, onWhite) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: d, h: d, rectRadius: d * 0.3, fill: { color: onWhite ? SK[k][0] : 'FFFFFF' }, line: { color: onWhite ? SK[k][0] : 'FFFFFF' }, objectName: name + ' chip' });
    s.addImage({ data: await icon(ic, SK[k][1]), x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52, altText: name });
  };
  const badge = (s, x, y, d, k, label, name, fs = 12) => {
    const h = Math.min(d, 0.5);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: d, h, rectRadius: h / 2, fill: { color: SK[k][1] }, line: { color: SK[k][1] }, objectName: name });
    T(s, String(label), { x, y, w: d, h, align: 'center', fontSize: fs, bold: true, color: 'FFFFFF' });
  };
  const tipsBox = async (s, x, y, w, h, list, name, headline = 'Remember') => {
    card(s, x, y, w, h, 'FFFFFF', name, true, LINE);
    if (h < 1.6) {
      await chip(s, x + 0.25, y + (h - 0.55) / 2, 0.55, 'G', 'LuLightbulb', 'Tip', true);
      T(s, [{ text: headline + '   ', options: { bold: true } }, { text: list.join('  ') }], { x: x + 1.0, y: y + 0.08, w: w - 1.25, h: h - 0.16, fontSize: 14 });
      return;
    }
    await chip(s, x + 0.3, y + 0.3, 0.55, 'G', 'LuLightbulb', 'Tip', true);
    T(s, headline, { x: x + 1.0, y: y + 0.3, w: w - 1.3, h: 0.55, fontSize: 15, bold: true });
    T(s, list.map((t, i) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < list.length - 1, paraSpaceAfter: 8 } })), { x: x + 0.3, y: y + 1.0, w: w - 0.6, h: h - 1.2, fontSize: 14, valign: 'top' });
  };
  const notesFor = (sl, extra) => (clean(sl.notes || '') + (extra ? '\n\n' + extra : '')).trim();
  const section = name => pres.addSection({ title: name });
  let sec = 'Welcome';
  const newSlide = (master = 'CONTENT') => pres.addSlide({ masterName: master, sectionTitle: sec });
  const SKILLNAME = { L: 'Listening', R: 'Reading', W: 'Writing', S: 'Speaking', O: 'Orientation' };
  const ebFor = (sl, i) => (i === D.SLIDES.length - 1 ? `Study Hub · Session ${N}` : sl.eyebrow || (sl.k === 'O' ? 'Strategy' : SKILLNAME[SESSION.k]));
  const nobreak = t => String(t).replace(/(\S)–(\S)/g, '$1\u2060–\u2060$2');

  // ---------- layouts (same visual language as the Session 1 deck) ----------
  const ICON_RULES = [[/word|vocab/i, 'LuBookOpenCheck'], [/predict|context/i, 'LuEye'], [/distractor|trap/i, 'LuAlertTriangle'], [/signal|change|mind/i, 'LuRefreshCw'],
    [/paraphrase/i, 'LuRepeat'], [/condition/i, 'LuFilter'], [/main idea|headline/i, 'LuNewspaper'], [/detail/i, 'LuListChecks'], [/impact/i, 'LuTrendingUp'],
    [/look back|score/i, 'LuBarChart3'], [/cause/i, 'LuSearch'], [/fix/i, 'LuWrench'], [/goal|focus|target/i, 'LuTarget'], [/contrast/i, 'LuArrowLeftRight'],
    [/disagree|polite no|negative/i, 'LuThumbsDown'], [/partly/i, 'LuScale'], [/decision|^agree/i, 'LuThumbsUp'], [/reluctant/i, 'LuMeh'], [/surprise/i, 'LuZap'],
    [/relief|positive|attitude/i, 'LuSmile'], [/hedge/i, 'LuHelpCircle'], [/viewpoint/i, 'LuMessageSquare'], [/imply/i, 'LuMessageCircle'], [/infer/i, 'LuLightbulb'],
    [/tone/i, 'LuVolume2'], [/^what$/i, 'LuClipboardList'], [/^when$/i, 'LuCalendar'], [/bring/i, 'LuLaptop']];
  const iconFor = (t, k) => { for (const [r, ic] of ICON_RULES) if (r.test(String(t).replace(/^\d+\s*·\s*/, ''))) return ic; return ICON[k] || 'LuSparkles'; };
  const oval = (s, x, y, d, k, label, name) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: SK[k][1] }, line: { color: SK[k][1], width: 0 }, objectName: name });
    T(s, String(label), { x, y, w: d, h: d, align: 'center', fontSize: 12, bold: true, color: 'FFFFFF' });
  };
  const softPill = (s, x, y, w, h, k, label, name) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2, fill: { color: SK[k][0] }, line: { color: SK[k][0] }, objectName: name });
    T(s, String(label), { x, y, w, h, fontSize: 12, bold: true, color: SK[k][1], align: 'center' });
  };
  // a row label: a round number badge for short labels, a soft pill for words
  const rowLabel = (s, x, y, rowH, k, a, name) => {
    const short = String(a).length <= 2;
    if (short) { oval(s, x, y + (rowH - 0.46) / 2, 0.46, k, a, name); return 0.46; }
    const w = Math.max(0.8, Math.min(1.9, 0.3 + String(a).length * 0.11));
    softPill(s, x, y + (rowH - 0.4) / 2, w, 0.4, k, a, name); return w;
  };

  async function titleSlide(sl) {
    const s = newSlide('TITLE');
    T(s, 'SESSION ' + String(N).padStart(2, '0'), { x: 0.75, y: 1.75, w: 4, h: 0.4, fontSize: 13, bold: true, color: SK[sl.k][1], charSpacing: 4 });
    s.addText(nobreak(sl.h), { placeholder: 'title' });
    s.addText('CELPIP Prep Course  ·  2-hour live class', { placeholder: 'body' });
    const four = [['L', 'LuHeadphones', 'Listening'], ['R', 'LuBookOpen', 'Reading'], ['W', 'LuPenLine', 'Writing'], ['S', 'LuMic', 'Speaking']];
    for (let i = 0; i < 4; i++) {
      const [k, ic, n] = four[i], x = 9.05 + (i % 2) * 1.85, y = 2.0 + Math.floor(i / 2) * 2.05, on = k === SESSION.k;
      card(s, x, y, 1.65, 1.85, 'FFFFFF', n + ' tile', true, on ? SK[k][1] : null);
      await chip(s, x + 0.43, y + 0.3, 0.8, k, ic, n, true);
      T(s, n, { x, y: y + 1.2, w: 1.65, h: 0.4, fontSize: 12, bold: true, align: 'center', color: on ? SK[k][1] : NAVY });
    }
    const g = D.GUIDE && D.GUIDE[0];
    s.addNotes(notesFor(sl, g ? 'DELIVERY TIPS FOR THIS CLASS\n' + [...g.say, ...g.ask, ...g.watch, g.tip].map(x => '• ' + x).join('\n') : ''));
  }
  function outcomesSlide(sl) {
    const s = newSlide();
    eyebrow(s, "Today's outcomes", 'R'); title(s, nobreak(sl.h));
    sl.list.forEach(([a, b], i) => {
      const y = 1.95 + i * 0.86, k = ['L', 'R', 'O', 'W', 'S'][i % 5];
      oval(s, 0.75, y + 0.12, 0.46, k, a, 'Outcome ' + a);
      T(s, b, { x: 1.45, y, w: 6.6, h: 0.7, fontSize: 18 });
    });
    card(s, 8.6, 1.95, 4.0, 4.3, 'F6F7FB', 'Agenda');
    T(s, '2 HOURS', { x: 8.95, y: 2.2, w: 3.3, h: 0.3, fontSize: 11, bold: true, color: MUTED, charSpacing: 3 });
    SESSION.live.forEach(([a, , b], i) => {
      const y = 2.65 + i * 0.68;
      T(s, a, { x: 8.95, y, w: 0.8, h: 0.5, fontSize: 14, bold: true, color: '6A2BC4' });
      T(s, b, { x: 9.8, y, w: 2.6, h: 0.5, fontSize: 14 });
    });
    s.addNotes(notesFor(sl));
  }
  async function cardsSlide(sl, i, master) {
    const s = newSlide(master);
    eyebrow(s, ebFor(sl, i), sl.k); title(s, nobreak(sl.h));
    const cs = sl.cards, n = cs.length, tips = sl.tips || [];
    if (n >= 6) { // word cards, 4 per row (as in Session 1)
      cs.forEach(([a, b], j) => {
        const k = ROT[j % 5], x = 0.75 + (j % 4) * 3.0, y = 1.95 + Math.floor(j / 4) * 2.3;
        card(s, x, y, 2.8, 2.1, SK[k][0], a);
        T(s, a, { x: x + 0.3, y: y + 0.25, w: 2.3, h: 0.5, fontSize: 19, bold: true, color: SK[k][1], valign: 'top' });
        T(s, b, { x: x + 0.3, y: y + 0.8, w: 2.3, h: 1.15, fontSize: 14, valign: 'top' });
      });
    } else { // icon cards with 01 / 02 numbering (as in Session 1)
      const gap = 0.25, w = (11.8 - gap * (n - 1)) / n, h = tips.length ? 3.2 : 4.3;
      for (let j = 0; j < n; j++) {
        const [a, b] = cs[j], k = ROT[(j + ROT.indexOf(sl.k)) % 5], x = 0.75 + j * (w + gap);
        const head = String(a).replace(/^\d+\s*·\s*/, '');
        card(s, x, 1.95, w, h, SK[k][0], head);
        await chip(s, x + 0.3, 2.25, 0.75, k, iconFor(head, k), head);
        T(s, String(j + 1).padStart(2, '0'), { x: x + w - 0.9, y: 2.3, w: 0.6, h: 0.4, fontSize: 12, bold: true, color: SK[k][1], align: 'right' });
        T(s, head, { x: x + 0.3, y: 3.15, w: w - 0.6, h: 0.6, fontSize: n > 3 ? 17 : 19, bold: true, valign: 'top' });
        T(s, b, { x: x + 0.3, y: 3.75, w: w - 0.6, h: h - 1.95, fontSize: n > 3 ? 14 : 15, valign: 'top' });
      }
      if (tips.length) await tipsBox(s, 0.75, 5.4, 11.8, 1.25, tips, 'Tips', 'Remember');
    }
    s.addNotes(notesFor(sl));
  }
  async function closingSlide(sl) {
    const s = newSlide('CLOSING');
    eyebrow(s, `Study Hub · Session ${N}`, 'O'); title(s, nobreak(sl.h));
    const todo = sl.list.map(x => x[1]), rowH = Math.min(0.7, (4.6 - (todo.length - 1) * 0.12) / todo.length);
    todo.forEach((t, i) => {
      const y = 1.95 + i * (rowH + 0.12);
      card(s, 0.75, y, 7.4, rowH, 'FFFFFF', 'To-do ' + (i + 1), true);
      oval(s, 0.97, y + (rowH - 0.42) / 2, 0.42, 'O', '✓', 'Check ' + (i + 1));
      T(s, t, { x: 1.6, y, w: 6.4, h: rowH, fontSize: 16 });
    });
    card(s, 8.6, 1.95, 3.9, 4.15, 'FFFFFF', 'Next class', true);
    await chip(s, 8.95, 2.3, 0.85, 'O', 'LuFlag', 'Next class', true);
    T(s, [{ text: 'Your results open the next class.', options: { bold: true, fontSize: 18, breakLine: true } }, { text: `Finish the Study Hub before Session ${N + 1} so we can work on your real mistakes.`, options: { fontSize: 14, color: MUTED } }],
      { x: 8.95, y: 3.4, w: 3.25, h: 2.4, valign: 'top', paraSpaceAfter: 10 });
    s.addNotes(notesFor(sl));
  }
  async function listSlide(sl, i, master) {
    const s = newSlide(master);
    eyebrow(s, ebFor(sl, i), sl.k); title(s, nobreak(sl.h));
    const L = sl.list, tips = sl.tips || [];
    const said = L.every(([a]) => a === 'Said' || a === 'Meant');
    if (said) { // said vs meant pairs
      for (let j = 0; j < L.length; j += 2) {
        const y = 1.95 + (j / 2) * 1.05;
        card(s, 0.75, y, 6.6, 0.9, 'F6F7FB', 'Said ' + j);
        await chip(s, 0.95, y + 0.17, 0.56, 'L', 'LuMessageCircle', 'Said', true);
        T(s, L[j][1], { x: 1.7, y, w: 5.5, h: 0.9, fontSize: 16 });
        T(s, '→', { x: 7.4, y, w: 0.5, h: 0.9, fontSize: 20, bold: true, color: MUTED, align: 'center' });
        card(s, 7.95, y, 4.6, 0.9, SK.O[0], 'Meant ' + j);
        await chip(s, 8.12, y + 0.17, 0.56, 'O', 'LuLightbulb', 'Meant');
        T(s, L[j + 1][1], { x: 8.85, y, w: 3.55, h: 0.9, fontSize: 16, bold: true });
      }
      if (tips.length) await tipsBox(s, 0.75, 5.35, 11.8, 1.3, tips, 'Tips', 'Remember');
    } else if (tips.length) { // rows left, tips right (as in Session 1's skill slides)
      const rowH = Math.min(0.72, (4.75 - (L.length - 1) * 0.1) / L.length);
      L.forEach(([a, b], j) => {
        const y = 1.95 + j * (rowH + 0.1);
        card(s, 0.75, y, 7.3, rowH, j % 2 ? 'FFFFFF' : 'F6F7FB', 'Row ' + (j + 1), false, j % 2 ? LINE : null);
        const lw = rowLabel(s, 0.95, y, rowH, sl.k, a, 'Label ' + (j + 1));
        T(s, b, { x: 1.2 + lw, y, w: 6.65 - lw, h: rowH, fontSize: 15 });
      });
      await tipsBox(s, 8.35, 1.95, 4.15, Math.max(2.6, L.length * (rowH + 0.1) - 0.1), tips, 'Strategies', 'How to score higher');
    } else { // two columns of rows
      const rowsN = Math.ceil(L.length / 2), rowH = Math.min(0.95, (4.7 - (rowsN - 1) * 0.15) / rowsN);
      L.forEach(([a, b], j) => {
        const col = j % 2, row = Math.floor(j / 2), x = 0.75 + col * 6.0, y = 1.95 + row * (rowH + 0.15), k = ROT[j % 5];
        card(s, x, y, 5.8, rowH, (row + col) % 2 ? 'FFFFFF' : 'F6F7FB', 'Item ' + (j + 1), false, (row + col) % 2 ? LINE : null);
        const lw = rowLabel(s, x + 0.25, y, rowH, k, a, 'Label ' + (j + 1));
        T(s, b, { x: x + 0.5 + lw, y, w: 5.1 - lw, h: rowH, fontSize: 16 });
      });
    }
    s.addNotes(notesFor(sl));
  }

  // ---------- live tools as slides ----------
  const W = {
    async THINK(sl) {
      const c = D.THINK_CLIP;
      let s = newSlide(); eyebrow(s, 'Listening · think aloud', 'L'); title(s, sl.h);
      T(s, c.context, { x: 0.75, y: 1.7, w: 11.8, h: 0.4, fontSize: 13, color: MUTED });
      c.cols.forEach((name, ci) => {
        const x = 0.75 + ci * 6.0;
        card(s, x, 2.2, 5.8, 3.0, ci ? SK.R[0] : SK.L[0], name);
        T(s, name, { x: x + 0.3, y: 2.35, w: 5.2, h: 0.45, fontSize: 16, bold: true, color: ci ? SK.R[1] : SK.L[1] });
        const notes = c.lines.filter(l => l[2] === ci).map(l => l[3]);
        T(s, notes.map((t, j) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: j < notes.length - 1, paraSpaceAfter: 6 } })), { x: x + 0.3, y: 2.9, w: 5.2, h: 2.2, fontSize: 15, fontFace: 'Consolas', valign: 'top' });
      });
      await tipsBox(s, 0.75, 5.4, 11.8, 1.25, sl.tips || [], 'Tips', 'Remember');
      s.addNotes(notesFor(sl, 'READ THE CLIP ONE LINE AT A TIME (pause and think aloud after each line):\n' + c.lines.map(l => `${l[0]}: ${l[1]}\n   → note: ${l[3]}   | think aloud: ${l[4]}`).join('\n') + '\n\nQUESTIONS:\n' + c.qs.map(([q, a]) => `• ${q} → ${a}`).join('\n')));
    },
    async ANNOTATE(sl) {
      const c = D.ANNO_CLIP;
      for (const key of [false, true]) {
        const s = newSlide(); eyebrow(s, key ? 'Answer key' : 'Listening · find the trap', 'R'); title(s, key ? sl.h + ': answer key' : sl.h);
        card(s, 0.75, 1.95, 8.0, 4.7, 'F6F7FB', 'Transcript');
        const runs = [];
        c.lines.forEach(([who, segs], j) => {
          runs.push({ text: who + ': ', options: { bold: true } });
          segs.forEach(([t, k], si) => {
            const last = si === segs.length - 1 && j < c.lines.length - 1;
            const o = key && k ? { highlight: MARK[k][0], color: MARK[k][1], bold: true } : (!key && k ? { underline: { style: 'dotted' } } : {});
            runs.push({ text: t, options: { ...o, ...(last ? { breakLine: true } : {}) } });
          });
        });
        T(s, runs, { x: 1.05, y: 2.15, w: 7.4, h: 4.3, fontSize: 15.5, valign: 'top', paraSpaceAfter: 6 });
        const legend = [['d', 'Distractor', 'Said first, then changed'], ['c', 'Answer cue', 'The final detail'], ['s', 'Signal word', 'Warns of a change']];
        legend.forEach(([k, n, d], j) => {
          const y = 1.95 + j * 1.0;
          card(s, 9.05, y, 3.5, 0.85, MARK[k][0], n);
          T(s, [{ text: n, options: { bold: true, color: MARK[k][1], breakLine: true } }, { text: d, options: { fontSize: 12, color: NAVY } }], { x: 9.3, y, w: 3.1, h: 0.85, fontSize: 15 });
        });
        card(s, 9.05, 5.0, 3.5, 1.65, 'FFFFFF', 'Questions', true, LINE);
        T(s, c.qs.map(([q, a], j) => ({ text: key ? `${q} → ${a}` : q, options: { breakLine: j < c.qs.length - 1, paraSpaceAfter: 4 } })), { x: 9.25, y: 5.08, w: 3.15, h: 1.5, fontSize: 11, valign: 'top' });
        s.addNotes(key ? 'Reveal after the class has marked the transcript. Ask: "Which word warned you each time?"' : notesFor(sl, 'Students say which colour each underlined phrase is: red = distractor, green = answer cue, yellow = signal word. The next slide is the answer key.'));
      }
    },
    async POLL(sl) {
      for (let j = 0; j < D.POLL_ITEMS.length; j++) {
        const p = D.POLL_ITEMS[j], s = newSlide();
        eyebrow(s, `Prediction poll · ${j + 1} of ${D.POLL_ITEMS.length}`, 'L'); title(s, p.q);
        card(s, 0.75, 1.95, 11.8, 1.0, SK.L[0], 'Context');
        T(s, p.ctx, { x: 1.05, y: 1.95, w: 11.2, h: 1.0, fontSize: 17, italic: true });
        p.o.forEach((o, k) => {
          const col = k % 2, row = Math.floor(k / 2), x = 0.75 + col * 6.0, y = 3.2 + row * 1.15;
          card(s, x, y, 5.8, 0.95, 'FFFFFF', 'Option ' + 'ABCD'[k], true, LINE);
          badge(s, x + 0.25, y + 0.22, 0.5, ROT[k], 'ABCD'[k], 'Letter ' + 'ABCD'[k], 14);
          T(s, o, { x: x + 0.95, y, w: 4.7, h: 0.95, fontSize: 17 });
        });
        await tipsBox(s, 0.75, 5.6, 11.8, 1.05, ['Hands up for A, B, C or D. Everyone commits before the answer.'], 'Poll', 'Vote first');
        s.addNotes(`ANSWER: ${'ABCD'[p.a]}. ${p.o[p.a]}\nWHY: ${p.why}` + (j === 0 ? '\n\n' + notesFor(sl) : ''));
      }
    },
    async NEWS(sl) {
      const c = D.NEWS_ITEM, cats = D.NEWS_CATS;
      for (const key of [false, true]) {
        const s = newSlide(); eyebrow(s, key ? 'Answer key' : 'Listening · news structure', 'L'); title(s, key ? sl.h + ': answer key' : sl.h);
        if (!key) {
          c.sents.forEach(([t], j) => {
            const y = 1.95 + j * 0.62;
            card(s, 0.75, y, 11.8, 0.52, j % 2 ? 'FFFFFF' : 'F6F7FB', 'Sentence ' + (j + 1), false, j % 2 ? LINE : null);
            badge(s, 0.9, y + 0.06, 0.4, 'L', j + 1, 'Number ' + (j + 1), 12);
            T(s, t, { x: 1.5, y, w: 10.9, h: 0.52, fontSize: 15 });
          });
          await tipsBox(s, 0.75, 5.8, 11.8, 0.9, ['Headline = what happened · Details = who, where, when, how much · Impact = results and reactions'], 'Sort', 'Sort each sentence');
        } else {
          const ks = ['h', 'd', 'i'];
          ks.forEach((k, j) => {
            const x = 0.75 + j * 4.0, kk = ['O', 'L', 'W'][j];
            card(s, x, 1.95, 3.8, 4.75, SK[kk][0], cats[k][0]);
            T(s, cats[k][0], { x: x + 0.3, y: 2.15, w: 3.2, h: 0.5, fontSize: 19, bold: true, color: SK[kk][1] });
            T(s, cats[k][1], { x: x + 0.3, y: 2.65, w: 3.2, h: 0.4, fontSize: 12, color: MUTED });
            const items = c.sents.filter(x2 => x2[1] === k).map(x2 => x2[0]);
            T(s, items.map((t, q) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: q < items.length - 1, paraSpaceAfter: 8 } })), { x: x + 0.3, y: 3.2, w: 3.2, h: 3.3, fontSize: 13.5, valign: 'top' });
          });
        }
        s.addNotes(key ? 'Read the report in this order: headline, details, impact. Ask: "Which sentence would be a question about reactions?"' : notesFor(sl, 'Students call out Headline, Details or Impact for each sentence. The next slide is the answer key.'));
      }
    },
    async TRACK(sl) {
      const c = D.TRACK_CLIP;
      for (const key of [false, true]) {
        const s = newSlide(); eyebrow(s, key ? 'Answer key' : 'Listening · discussion', 'S'); title(s, key ? sl.h + ': answer key' : sl.h);
        T(s, c.context, { x: 0.75, y: 1.7, w: 11.8, h: 0.4, fontSize: 13, color: MUTED });
        const x0 = 0.75, nameW = 1.2, stW = 11.8 - nameW * 3;
        T(s, 'Who says this?', { x: x0 + 0.2, y: 2.2, w: stW - 0.2, h: 0.4, fontSize: 12, bold: true, color: MUTED });
        c.names.forEach((n, j) => T(s, n, { x: x0 + stW + j * nameW, y: 2.2, w: nameW, h: 0.4, fontSize: 14, bold: true, color: SK[['L', 'R', 'S'][j]][1], align: 'center' }));
        c.rows.forEach(([st, a], r) => {
          const y = 2.65 + r * 0.56;
          card(s, x0, y, 11.8, 0.48, r % 2 ? 'FFFFFF' : 'F6F7FB', 'Row ' + (r + 1), false, r % 2 ? LINE : null);
          T(s, st, { x: x0 + 0.2, y, w: stW - 0.3, h: 0.48, fontSize: 14.5 });
          c.names.forEach((n, j) => {
            const cx = x0 + stW + j * nameW + nameW / 2 - 0.19, on = key && a === j;
            s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: y + 0.05, w: 0.38, h: 0.38, rectRadius: 0.08, fill: { color: on ? 'DDF3E4' : 'FFFFFF' }, line: { color: on ? '17803D' : 'C9CFDD', width: 1 }, objectName: `Box ${r + 1}-${n}` });
            if (on) T(s, '✓', { x: cx, y: y + 0.05, w: 0.38, h: 0.38, fontSize: 14, bold: true, color: '17803D', align: 'center' });
          });
        });
        s.addNotes(key ? 'Play the discussion again with the transcript. Ask: "Who changed their mind, and why?"' : notesFor(sl, 'READ THE DISCUSSION ONCE (change your voice for each speaker):\n' + c.lines.map(([w, t]) => `${w}: ${t}`).join('\n')));
      }
    },
    async ATTITUDE(sl) {
      const c = D.ATT_TEXT;
      for (const key of [false, true]) {
        const s = newSlide(); eyebrow(s, key ? 'Answer key' : 'Listening · attitude', 'R'); title(s, key ? sl.h + ': answer key' : sl.h);
        card(s, 0.75, 1.95, 8.0, 4.7, 'F6F7FB', 'Report');
        T(s, c.context, { x: 1.05, y: 2.1, w: 7.4, h: 0.4, fontSize: 12, color: MUTED });
        const runs = c.segs.map(([t, k]) => ({ text: t, options: k ? (key ? { highlight: ATT_MARK[k][0], color: ATT_MARK[k][1], bold: true } : { underline: { style: 'dotted' }, bold: true }) : {} }));
        T(s, runs, { x: 1.05, y: 2.55, w: 7.4, h: 3.9, fontSize: 18, valign: 'top', lineSpacingMultiple: 1.3 });
        Object.entries(D.ATT_PENS).forEach(([k, n], j) => {
          const y = 1.95 + j * 0.8;
          card(s, 9.05, y, 3.5, 0.65, ATT_MARK[k][0], n);
          T(s, n, { x: 9.3, y, w: 3.1, h: 0.65, fontSize: 15, bold: true, color: ATT_MARK[k][1] });
        });
        card(s, 9.05, 5.2, 3.5, 1.45, 'FFFFFF', 'Questions', true, LINE);
        T(s, c.qs.map(([q, a], j) => ({ text: key ? `${q} → ${a}` : q, options: { breakLine: j < c.qs.length - 1, paraSpaceAfter: 4 } })), { x: 9.25, y: 5.28, w: 3.15, h: 1.3, fontSize: 11, valign: 'top' });
        s.addNotes(key ? 'Count the positive and negative phrases together, then ask the two questions.' : notesFor(sl, 'Students say the group for each underlined phrase: positive, negative, hedge or contrast. The next slide is the answer key.'));
      }
    }
  };

  // ---------- build ----------
  section('Welcome');
  const last = D.SLIDES.length - 1;
  for (let i = 0; i < D.SLIDES.length; i++) {
    const sl = D.SLIDES[i];
    if (i === 3) { sec = 'Teach'; section(sec); }
    if (i === last) { sec = 'Wrap-up'; section(sec); }
    const master = i === last ? 'CLOSING' : 'CONTENT';
    if (sl.layout === 'title') await titleSlide(sl);
    else if (i === 1 && sl.list) outcomesSlide(sl);
    else if (sl.widget) { if (!W[sl.widget]) throw new Error('No deck layout for ' + sl.widget); await W[sl.widget](sl); }
    else if (i === last && sl.list) await closingSlide(sl);
    else if (sl.cards) await cardsSlide(sl, i, master);
    else if (sl.list) await listSlide(sl, i, master);
    else throw new Error('Unknown slide ' + i);
  }
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log('written', OUT);
})().catch(e => { console.error(e); process.exit(1); });
