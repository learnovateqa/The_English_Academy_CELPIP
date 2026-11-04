// Session decks · CELPIP Prep Course · Elevate with Ms. Maria · Learnovate Qatar
// One generator for every session. The content of each deck lives in tools/decks/sNN.js,
// so every session follows the same 2-hour system and the same visual design.
//
// Build:  NODE_PATH=<folder with pptxgenjs, react-icons, react, react-dom, sharp>/node_modules \
//         APPLY_THEME=<pptx skill>/scripts/apply_theme.js node tools/session-deck.js 2 decks/session-02.pptx
const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const sharp = require('sharp');
const Lu = require('react-icons/lu');
const { applyTheme } = require(process.env.APPLY_THEME || './apply_theme.js');

const N = +process.argv[2];
const OUT = process.argv[3] || `session-${String(N).padStart(2, '0')}.pptx`;
const C = require(path.join(__dirname, 'decks', `s${String(N).padStart(2, '0')}.js`));

const THEME = { name: 'Elevate Pastel', headFontFace: 'Poppins', bodyFontFace: 'Poppins',
  colors: { dk1: '14213D', lt1: 'FFFFFF', dk2: '5B6782', lt2: 'F6F7FB', accent1: '1A5FCC', accent2: '6A2BC4', accent3: 'C93D1A',
    accent4: '007F7F', accent5: '008F55', accent6: 'E0A800', hlink: '1A5FCC', folHlink: '6A2BC4' } };
const SK = { L: ['EAF2FF', '1A5FCC'], R: ['F3ECFE', '6A2BC4'], W: ['FFEFEA', 'C93D1A'], S: ['E5F7F6', '007F7F'], O: ['E7F7EF', '008F55'], G: ['FFF7DB', '8A6A00'] };
const ROT = ['L', 'R', 'W', 'S', 'O'];
const ICON = { L: 'LuHeadphones', R: 'LuBookOpen', W: 'LuPenLine', S: 'LuMic', O: 'LuCompass' };
const SKILLNAME = { L: 'Listening', R: 'Reading', W: 'Writing', S: 'Speaking', O: 'Orientation' };
const NAVY = '14213D', MUTED = '5B6782', LINE = 'E6E9F2';
// highlight pens: [fill, ink, label, meaning]
const PENS = {
  d: ['FDE2E2', '9B1C1C', 'Distractor', 'Said first, then changed'], c: ['DDF3E4', '17603A', 'Answer cue', 'The final detail'], s: ['FFF1BF', '7A5A00', 'Signal word', 'Warns of a change'],
  p: ['DDF3E4', '17603A', 'Positive', 'Approves'], n: ['FDE2E2', '9B1C1C', 'Negative', 'Criticises'], h: ['FFF1BF', '7A5A00', 'Hedge', 'Not sure'], x: ['E3EEFF', '1A4FA8', 'Contrast', 'Changes direction'],
  k: ['E3EEFF', '1A4FA8', 'Claim', 'The writer’s point'], e: ['DDF3E4', '17603A', 'Evidence', 'Facts that support it'], q: ['FFF1BF', '7A5A00', 'Concession', 'Admits the other side'],
  m: ['FFF1BF', '7A5A00', 'Key word', 'Match this in the text'],
  g: ['DDF3E4', '17603A', 'Strong', 'Keep this'], f: ['FDE2E2', '9B1C1C', 'Fix', 'Error or weak choice'], t: ['FFF1BF', '7A5A00', 'Tone', 'Too casual or too stiff'] };

async function icon(name, color, size = 256) {
  const Cmp = Lu[name] || Lu.LuSparkles;
  const svg = renderToStaticMarkup(React.createElement(Cmp, { color: '#' + color, size: String(size), strokeWidth: 1.75 }));
  return 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
}
async function gradient(w, h, stops) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">${stops.map(([o, c]) => `<stop offset="${o}" stop-color="#${c}"/>`).join('')}</linearGradient>
    <radialGradient id="r" cx="0.85" cy="0.15" r="0.6"><stop offset="0" stop-color="#ffffff" stop-opacity="0.55"/><stop offset="1" stop-color="#ffffff" stop-opacity="0"/></radialGradient></defs>
    <rect width="100%" height="100%" fill="url(#g)"/><rect width="100%" height="100%" fill="url(#r)"/></svg>`;
  return 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
}
const shadow = () => ({ type: 'outer', color: '14213D', opacity: 0.08, blur: 14, offset: 3, angle: 90 });
const nobreak = t => String(t).replace(/(\S)–(\S)/g, '$1⁠–⁠$2');
// "[[phrase|d]]" marks a highlighted phrase
const parseMarks = t => { const out = []; String(t).split(/(\[\[[^\]]+\]\])/).forEach(p => { const m = p.match(/^\[\[(.+)\|(\w)\]\]$/); if (m) out.push([m[1], m[2]]); else if (p) out.push([p, '']); }); return out; };
const plain = t => parseMarks(t).map(x => x[0]).join('');

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.author = 'Ms. Maria · Learnovate Qatar';
  pres.title = `Session ${N} · ${C.title}`;
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const TC = pres.SchemeColor;
  const titleBg = await gradient(1920, 1080, [[0, 'F3ECFE'], [0.5, 'FFEFEA'], [1, 'EAF2FF']]);
  const closeBg = await gradient(1920, 1080, [[0, 'EAF2FF'], [1, 'E7F7EF']]);
  const dividerBg = await gradient(1920, 1080, [[0, SK[C.k][0]], [1, 'FFFFFF']]);
  pres.defineSlideMaster({ title: 'TITLE', background: { data: titleBg },
    objects: [
      { placeholder: { options: { name: 'title', type: 'title', x: 0.75, y: 2.3, w: 7.6, h: 2.3, fontSize: 42, bold: true, color: TC.text1, valign: 'top', align: 'left', margin: 0 }, text: '' } },
      { placeholder: { options: { name: 'body', type: 'body', x: 0.75, y: 4.75, w: 7.6, h: 0.6, fontSize: 18, color: TC.text2, align: 'left', margin: 0 }, text: '' } },
      { text: { text: 'ELEVATE WITH MS. MARIA  ·  LEARNOVATE QATAR', options: { x: 0.75, y: 6.6, w: 8, h: 0.35, fontSize: 10, bold: true, color: TC.text2, charSpacing: 3, margin: 0 } } }] });
  const contentObjects = () => [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.75, y: 0.88, w: 11.8, h: 0.8, fontSize: 32, bold: true, color: TC.text1, valign: 'top', align: 'left', margin: 0 }, text: '' } },
    { text: { text: `CELPIP Prep Course  ·  Session ${N}`, options: { x: 0.75, y: 6.95, w: 6, h: 0.3, fontSize: 9, color: TC.text2, margin: 0 } } }];
  const num = () => ({ x: 12.08, y: 6.95, w: 0.5, h: 0.3, fontSize: 9, color: '5B6782', align: 'right' });
  pres.defineSlideMaster({ title: 'CONTENT', background: { color: 'FFFFFF' }, objects: contentObjects(), slideNumber: num() });
  pres.defineSlideMaster({ title: 'CLOSING', background: { data: closeBg }, objects: contentObjects(), slideNumber: num() });
  pres.defineSlideMaster({ title: 'DIVIDER', background: { data: dividerBg }, objects: [{ text: { text: `CELPIP Prep Course  ·  Session ${N}`, options: { x: 0.75, y: 6.95, w: 6, h: 0.3, fontSize: 9, color: TC.text2, margin: 0 } } }], slideNumber: num() });

  // ---------- drawing helpers ----------
  const T = (s, text, o) => s.addText(text, { color: NAVY, margin: 0, isTextBox: true, valign: 'middle', ...o });
  const title = (s, t) => s.addText(nobreak(t), { placeholder: 'title', ...(String(t).length > 46 ? { fontSize: 26 } : {}) });
  const eyebrow = (s, text, k) => T(s, String(text).toUpperCase(), { x: 0.75, y: 0.5, w: 11, h: 0.32, fontSize: 11, bold: true, color: SK[k][1], charSpacing: 3, valign: 'top' });
  const card = (s, x, y, w, h, fill, name, sh, line) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.2, fill: { color: fill }, line: { color: line || fill, width: line ? 1 : 0.75 }, objectName: name, ...(sh ? { shadow: shadow() } : {}) });
  const chip = async (s, x, y, d, k, ic, name, onWhite) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: d, h: d, rectRadius: d * 0.3, fill: { color: onWhite ? SK[k][0] : 'FFFFFF' }, line: { color: onWhite ? SK[k][0] : 'FFFFFF' }, objectName: name + ' chip' });
    s.addImage({ data: await icon(ic, SK[k][1]), x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52, altText: name });
  };
  const oval = (s, x, y, d, k, label, name, solid = true) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: solid ? SK[k][1] : 'FFFFFF' }, line: { color: SK[k][1], width: solid ? 0 : 1.25 }, objectName: name });
    T(s, String(label), { x, y, w: d, h: d, align: 'center', fontSize: 12, bold: true, color: solid ? 'FFFFFF' : SK[k][1] });
  };
  const softPill = (s, x, y, w, h, k, label, name) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2, fill: { color: SK[k][0] }, line: { color: SK[k][0] }, objectName: name });
    T(s, String(label), { x, y, w, h, fontSize: 12, bold: true, color: SK[k][1], align: 'center' });
  };
  const rowLabel = (s, x, y, rowH, k, a, name) => {
    if (String(a).length <= 2) { oval(s, x, y + (rowH - 0.46) / 2, 0.46, k, a, name); return 0.46; }
    const w = Math.max(0.8, Math.min(2.0, 0.3 + String(a).length * 0.11));
    softPill(s, x, y + (rowH - 0.4) / 2, w, 0.4, k, a, name); return w;
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
  const ICON_RULES = [[/word|vocab/i, 'LuBookOpenCheck'], [/predict|context|preview/i, 'LuEye'], [/distractor|trap|mistake/i, 'LuAlertTriangle'], [/signal|change|mind/i, 'LuRefreshCw'],
    [/paraphrase/i, 'LuRepeat'], [/condition/i, 'LuFilter'], [/main idea|headline/i, 'LuNewspaper'], [/detail/i, 'LuListChecks'], [/impact/i, 'LuTrendingUp'],
    [/look back|score|result|data/i, 'LuBarChart3'], [/cause|scan|find|search/i, 'LuSearch'], [/fix/i, 'LuWrench'], [/goal|focus|target/i, 'LuTarget'], [/contrast/i, 'LuArrowLeftRight'],
    [/disagree|polite no|negative/i, 'LuThumbsDown'], [/partly/i, 'LuScale'], [/decision|^agree/i, 'LuThumbsUp'], [/reluctant/i, 'LuMeh'], [/surprise/i, 'LuZap'],
    [/relief|positive|attitude/i, 'LuSmile'], [/hedge/i, 'LuHelpCircle'], [/viewpoint|comment/i, 'LuMessageSquare'], [/imply/i, 'LuMessageCircle'], [/infer|idea/i, 'LuLightbulb'],
    [/tone/i, 'LuVolume2'], [/^what$/i, 'LuClipboardList'], [/^when$|time|minute|clock/i, 'LuClock'], [/bring|laptop/i, 'LuLaptop'], [/email|letter|reply/i, 'LuMail'],
    [/diagram|table|chart/i, 'LuTable'], [/paragraph/i, 'LuAlignLeft'], [/claim/i, 'LuFlag'], [/evidence/i, 'LuFileText'], [/concession/i, 'LuHandMetal'],
    [/plan|outline/i, 'LuListOrdered'], [/proofread|check/i, 'LuCheckCheck'], [/readab/i, 'LuGlasses'], [/content|coheren/i, 'LuLayers'], [/task|fulfil/i, 'LuBadgeCheck'],
    [/tone|formal|friend/i, 'LuUser'], [/book|test/i, 'LuAward'], [/ruler|clb/i, 'LuRuler'], [/start/i, 'LuPlay'], [/finish/i, 'LuFlag'], [/add|explain/i, 'LuPlus']];
  const iconFor = (t, k) => { const x = String(t).replace(/^\d+\s*·\s*/, ''); for (const [r, ic] of ICON_RULES) if (r.test(x)) return ic; return ICON[k] || 'LuSparkles'; };

  // ---------- timing + notes ----------
  let clock = 0; // minutes from the start of class
  const hm = m => `${Math.floor(m / 60)}:${String(Math.round(m % 60)).padStart(2, '0')}`;
  const notes = (sl, extra) => {
    const head = `⏱ ${sl.m || 1} min  ·  starts at ${hm(clock)}`;
    clock += sl.m || 1;
    return [head, sl.say ? 'SAY: ' + sl.say : '', sl.notes || '', extra || '', sl.check ? 'CHECK: ' + sl.check : ''].filter(Boolean).join('\n');
  };
  let sec = 'Welcome';
  const newSlide = (master = 'CONTENT') => pres.addSlide({ masterName: master, sectionTitle: sec });
  const eb = sl => sl.eb || (sl.k === 'O' ? 'Strategy' : SKILLNAME[C.k]);
  const kk = sl => sl.k || C.k;

  // ---------- layouts ----------
  const L = {};
  L.title = async sl => {
    const s = newSlide('TITLE');
    T(s, 'SESSION ' + String(N).padStart(2, '0'), { x: 0.75, y: 1.75, w: 4, h: 0.4, fontSize: 13, bold: true, color: SK[C.k][1], charSpacing: 4 });
    s.addText(nobreak(C.title), { placeholder: 'title' });
    s.addText('CELPIP Prep Course  ·  2-hour live class', { placeholder: 'body' });
    const four = [['L', 'LuHeadphones', 'Listening'], ['R', 'LuBookOpen', 'Reading'], ['W', 'LuPenLine', 'Writing'], ['S', 'LuMic', 'Speaking']];
    for (let i = 0; i < 4; i++) {
      const [k, ic, n] = four[i], x = 9.05 + (i % 2) * 1.85, y = 2.0 + Math.floor(i / 2) * 2.05, on = k === C.k;
      card(s, x, y, 1.65, 1.85, 'FFFFFF', n + ' tile', true, on ? SK[k][1] : null);
      await chip(s, x + 0.43, y + 0.3, 0.8, k, ic, n, true);
      T(s, n, { x, y: y + 1.2, w: 1.65, h: 0.4, fontSize: 12, bold: true, align: 'center', color: on ? SK[k][1] : NAVY });
    }
    s.addNotes(notes(sl, C.delivery ? 'DELIVERY TIPS FOR THIS CLASS\n' + C.delivery.map(x => '• ' + x).join('\n') : ''));
  };
  L.outcomes = async sl => {
    const s = newSlide();
    eyebrow(s, "Today's outcomes", 'R'); title(s, 'By the end of today you can…');
    sl.items.forEach((t, i) => {
      const y = 1.95 + i * 0.86, k = ['L', 'R', 'O', 'W', 'S'][i % 5];
      oval(s, 0.75, y + 0.12, 0.46, k, i + 1, 'Outcome ' + (i + 1));
      T(s, t, { x: 1.45, y, w: 6.6, h: 0.7, fontSize: 18 });
    });
    card(s, 8.6, 1.95, 4.0, 4.3, 'F6F7FB', 'Agenda');
    T(s, '2 HOURS', { x: 8.95, y: 2.2, w: 3.3, h: 0.3, fontSize: 11, bold: true, color: MUTED, charSpacing: 3 });
    C.agenda.forEach(([a, b], i) => {
      const y = 2.65 + i * 0.68;
      T(s, a, { x: 8.95, y, w: 0.8, h: 0.5, fontSize: 14, bold: true, color: '6A2BC4' });
      T(s, b, { x: 9.8, y, w: 2.6, h: 0.5, fontSize: 14 });
    });
    s.addNotes(notes(sl));
  };
  L.words = async sl => {
    const s = newSlide();
    eyebrow(s, 'Key words for today', 'O'); title(s, sl.h || `${sl.items.length} words you will hear today`);
    sl.items.forEach(([a, b], j) => {
      const k = ROT[j % 5], x = 0.75 + (j % 4) * 3.0, y = 1.95 + Math.floor(j / 4) * 2.3;
      card(s, x, y, 2.8, 2.1, SK[k][0], a);
      T(s, a, { x: x + 0.3, y: y + 0.25, w: 2.3, h: 0.5, fontSize: 19, bold: true, color: SK[k][1], valign: 'top' });
      T(s, b, { x: x + 0.3, y: y + 0.8, w: 2.3, h: 1.15, fontSize: 14, valign: 'top' });
    });
    s.addNotes(notes(sl));
  };
  L.divider = async sl => {
    sec = `Block ${sl.block} · ${sl.h}`.slice(0, 60); pres.addSection({ title: sec });
    const s = newSlide('DIVIDER'), k = kk(sl);
    T(s, `BLOCK ${sl.block}  ·  ${sl.time}`, { x: 0.75, y: 1.3, w: 8, h: 0.45, fontSize: 14, bold: true, color: SK[k][1], charSpacing: 4 });
    T(s, nobreak(sl.h), { x: 0.75, y: 1.8, w: 11.5, h: 1.3, fontSize: 40, bold: true, valign: 'top' });
    if (sl.sub) T(s, sl.sub, { x: 0.75, y: 3.1, w: 11, h: 0.6, fontSize: 18, color: MUTED, valign: 'top' });
    const steps = sl.steps || ['Teach', 'I do · model', 'We do · together'];
    const icons = ['LuPresentation', 'LuEye', 'LuUsers', 'LuCheckCheck'];
    for (let i = 0; i < steps.length; i++) {
      const w = (11.8 - 0.3 * (steps.length - 1)) / steps.length, x = 0.75 + i * (w + 0.3);
      card(s, x, 4.15, w, 1.9, 'FFFFFF', 'Step ' + (i + 1), true);
      await chip(s, x + 0.3, 4.45, 0.7, ROT[(i + 1) % 5], icons[i % 4], steps[i], true);
      T(s, String(i + 1).padStart(2, '0'), { x: x + w - 0.9, y: 4.5, w: 0.6, h: 0.4, fontSize: 12, bold: true, color: SK[ROT[(i + 1) % 5]][1], align: 'right' });
      T(s, steps[i], { x: x + 0.3, y: 5.3, w: w - 0.6, h: 0.6, fontSize: 17, bold: true });
    }
    s.addNotes(notes(sl));
  };
  L.cards = async sl => {
    const s = newSlide(sl.master); eyebrow(s, eb(sl), kk(sl)); title(s, sl.h);
    const cs = sl.items, n = cs.length, tips = sl.tips || [];
    const gap = 0.25, w = (11.8 - gap * (n - 1)) / n, h = tips.length ? 3.2 : 4.3;
    for (let j = 0; j < n; j++) {
      const [a, b, ic] = cs[j], k = ROT[(j + ROT.indexOf(kk(sl))) % 5], x = 0.75 + j * (w + gap);
      card(s, x, 1.95, w, h, SK[k][0], a);
      await chip(s, x + 0.3, 2.25, 0.75, k, ic || iconFor(a, k), a);
      T(s, String(j + 1).padStart(2, '0'), { x: x + w - 0.9, y: 2.3, w: 0.6, h: 0.4, fontSize: 12, bold: true, color: SK[k][1], align: 'right' });
      T(s, a, { x: x + 0.3, y: 3.15, w: w - 0.6, h: 0.6, fontSize: n > 3 ? 17 : 19, bold: true, valign: 'top' });
      T(s, b, { x: x + 0.3, y: 3.75, w: w - 0.6, h: h - 1.95, fontSize: n > 3 ? 14 : 15, valign: 'top' });
    }
    if (tips.length) await tipsBox(s, 0.75, 5.4, 11.8, 1.25, tips, 'Tips', sl.tipHead || 'Remember');
    s.addNotes(notes(sl));
  };
  L.rows = async sl => { // rows on the left, tips on the right
    const s = newSlide(sl.master); eyebrow(s, eb(sl), kk(sl)); title(s, sl.h);
    const R = sl.items, tips = sl.tips || [], rowH = Math.min(0.72, (4.75 - (R.length - 1) * 0.1) / R.length);
    R.forEach(([a, b], j) => {
      const y = 1.95 + j * (rowH + 0.1);
      card(s, 0.75, y, 7.3, rowH, j % 2 ? 'FFFFFF' : 'F6F7FB', 'Row ' + (j + 1), false, j % 2 ? LINE : null);
      const lw = rowLabel(s, 0.95, y, rowH, kk(sl), a, 'Label ' + (j + 1));
      T(s, b, { x: 1.2 + lw, y, w: 6.65 - lw, h: rowH, fontSize: 15 });
    });
    if (tips.length) await tipsBox(s, 8.35, 1.95, 4.15, Math.max(2.6, R.length * (rowH + 0.1) - 0.1), tips, 'Strategies', sl.tipHead || 'How to score higher');
    s.addNotes(notes(sl));
  };
  L.grid = async sl => { // two columns of rows
    const s = newSlide(sl.master); eyebrow(s, eb(sl), kk(sl)); title(s, sl.h);
    const R = sl.items, tips = sl.tips || [], rowsN = Math.ceil(R.length / 2), avail = tips.length ? 3.3 : 4.7;
    const rowH = Math.min(0.95, (avail - (rowsN - 1) * 0.15) / rowsN);
    R.forEach(([a, b], j) => {
      const col = j % 2, row = Math.floor(j / 2), x = 0.75 + col * 6.0, y = 1.95 + row * (rowH + 0.15), k = ROT[j % 5];
      card(s, x, y, 5.8, rowH, (row + col) % 2 ? 'FFFFFF' : 'F6F7FB', 'Item ' + (j + 1), false, (row + col) % 2 ? LINE : null);
      const lw = rowLabel(s, x + 0.25, y, rowH, k, a, 'Label ' + (j + 1));
      T(s, b, { x: x + 0.5 + lw, y, w: 5.1 - lw, h: rowH, fontSize: 15.5 });
    });
    if (tips.length) await tipsBox(s, 0.75, 5.45, 11.8, 1.2, tips, 'Tips', sl.tipHead || 'Remember');
    s.addNotes(notes(sl));
  };
  L.pairs = async sl => {
    const s = newSlide(); eyebrow(s, eb(sl), kk(sl)); title(s, sl.h);
    const P = sl.items, tips = sl.tips || [], avail = tips.length ? 3.3 : 4.7, rowH = Math.min(0.9, (avail - (P.length - 1) * 0.12) / P.length);
    const [la, lb] = sl.labels || ['', ''];
    for (let j = 0; j < P.length; j++) {
      const y = 1.95 + j * (rowH + 0.12);
      card(s, 0.75, y, 6.6, rowH, 'F6F7FB', 'Left ' + j);
      await chip(s, 0.95, y + (rowH - 0.5) / 2, 0.5, kk(sl), sl.iconA || 'LuMessageCircle', la || 'Left', true);
      T(s, [...(la ? [{ text: la.toUpperCase() + '   ', options: { bold: true, color: MUTED, fontSize: 10 } }] : []), { text: P[j][0] }], { x: 1.6, y, w: 5.6, h: rowH, fontSize: 15.5 });
      T(s, '→', { x: 7.4, y, w: 0.5, h: rowH, fontSize: 20, bold: true, color: MUTED, align: 'center' });
      card(s, 7.95, y, 4.6, rowH, SK.O[0], 'Right ' + j);
      await chip(s, 8.12, y + (rowH - 0.5) / 2, 0.5, 'O', sl.iconB || 'LuLightbulb', lb || 'Right');
      T(s, [...(lb ? [{ text: lb.toUpperCase() + '   ', options: { bold: true, color: SK.O[1], fontSize: 10 } }] : []), { text: P[j][1], options: { bold: true } }], { x: 8.75, y, w: 3.65, h: rowH, fontSize: 15.5 });
    }
    if (tips.length) await tipsBox(s, 0.75, 5.45, 11.8, 1.2, tips, 'Tips', sl.tipHead || 'Remember');
    s.addNotes(notes(sl));
  };
  L.steps = async sl => { // horizontal process with arrows
    const s = newSlide(); eyebrow(s, eb(sl), kk(sl)); title(s, sl.h);
    const S = sl.items, n = S.length, w = (11.8 - 0.25 * (n - 1)) / n, h = (sl.tips || []).length ? 3.35 : 4.3;
    for (let i = 0; i < n; i++) {
      const [a, b, ic] = S[i], k = ROT[(i + 4) % 5], x = 0.75 + i * (w + 0.25);
      card(s, x, 1.95, w, h, SK[k][0], a);
      await chip(s, x + 0.3, 2.25, 0.75, k, ic || iconFor(a, k), a);
      T(s, String(i + 1).padStart(2, '0'), { x: x + w - 0.9, y: 2.3, w: 0.6, h: 0.4, fontSize: 12, bold: true, color: SK[k][1], align: 'right' });
      T(s, a, { x: x + 0.3, y: 3.15, w: w - 0.5, h: 0.8, fontSize: 17, bold: true, valign: 'top' });
      T(s, b, { x: x + 0.3, y: 3.95, w: w - 0.5, h: h - 2.15, fontSize: 13.5, valign: 'top' });
      if (i < n - 1) T(s, '→', { x: x + w - 0.03, y: 3.35, w: 0.3, h: 0.5, fontSize: 18, bold: true, color: MUTED, align: 'center' });
    }
    if ((sl.tips || []).length) await tipsBox(s, 0.75, 5.5, 11.8, 1.15, sl.tips, 'Tips', sl.tipHead || 'Remember');
    s.addNotes(notes(sl));
  };
  // multiple choice: a "your turn" slide, plus an answer slide when reveal is set
  L.mcq = async sl => {
    for (const key of sl.reveal ? [false, true] : [false]) {
      const s = newSlide(); eyebrow(s, key ? 'Answer' : (sl.eb || 'Your turn'), kk(sl)); title(s, sl.h);
      let y = String(sl.h).length > 46 ? 2.25 : 1.95;
      if (sl.ctx) {
        const ch = sl.ctxH || (plain(sl.ctx).length > 140 ? 1.5 : 1.0);
        card(s, 0.75, y, 11.8, ch, SK[kk(sl)][0], 'Context');
        T(s, parseMarks(sl.ctx).map(([t, m]) => ({ text: t, options: m && key ? { bold: true, highlight: PENS[m][0], color: PENS[m][1] } : {} })), { x: 1.05, y, w: 11.2, h: ch, fontSize: 16.5, italic: true });
        y += ch + 0.25;
      }
      const rowsN = Math.ceil(sl.o.length / 2), oh = Math.min(0.9, (6.7 - y - 1.1 - 0.2 * (rowsN - 1) - 0.15) / rowsN);
      sl.o.forEach((o, k) => {
        const col = k % 2, row = Math.floor(k / 2), x = 0.75 + col * 6.0, yy = y + row * (oh + 0.2), right = key && k === sl.a;
        card(s, x, yy, 5.8, oh, right ? 'DDF3E4' : 'FFFFFF', 'Option ' + 'ABCD'[k], true, right ? '17803D' : LINE);
        oval(s, x + 0.25, yy + (oh - 0.48) / 2, 0.48, right ? 'O' : ROT[k], 'ABCD'[k], 'Letter ' + 'ABCD'[k]);
        T(s, o + (right ? '  ✓' : ''), { x: x + 0.95, y: yy, w: 4.7, h: oh, fontSize: 16.5, bold: right, color: key && !right ? MUTED : NAVY });
      });
      const ty = y + rowsN * oh + (rowsN - 1) * 0.2 + 0.15;
      if (key) await tipsBox(s, 0.75, ty, 11.8, Math.min(1.1, 6.7 - ty), [sl.why], 'Why', 'Why');
      else await tipsBox(s, 0.75, ty, 11.8, Math.min(1.1, 6.7 - ty), [sl.prompt || 'Hands up for A, B, C or D. Everyone commits before the answer.'], 'Vote', 'Vote first');
      s.addNotes(key ? `ANSWER: ${'ABCD'[sl.a]}. ${sl.o[sl.a]}\nWHY: ${sl.why}` : notes(sl, sl.reveal ? 'The next slide shows the answer.' : `ANSWER: ${'ABCD'[sl.a]}. ${sl.o[sl.a]}\nWHY: ${sl.why}`));
    }
  };
  // transcript or text with highlighted phrases: a task slide, then an answer-key slide
  L.marks = async sl => {
    const pens = sl.pens || ['d', 'c', 's'];
    for (const key of [false, true]) {
      const s = newSlide(); eyebrow(s, key ? 'Answer key' : (sl.eb || 'Your turn'), kk(sl)); title(s, key ? sl.h + ' · answer key' : sl.h);
      card(s, 0.75, 1.95, 8.0, 4.75, 'F6F7FB', 'Text');
      let y0 = 2.1;
      if (sl.ctx) { T(s, sl.ctx, { x: 1.05, y: 2.05, w: 7.4, h: 0.4, fontSize: 12, color: MUTED }); y0 = 2.5; }
      const runs = [];
      const lines = sl.lines || [[null, sl.text]];
      lines.forEach(([who, text], j) => {
        if (who) runs.push({ text: who + ': ', options: { bold: true } });
        const segs = parseMarks(text);
        segs.forEach(([t, m], si) => {
          const brk = si === segs.length - 1 && j < lines.length - 1;
          const o = m ? (key ? { highlight: PENS[m][0], color: PENS[m][1], bold: true } : { underline: { style: 'dotted' }, bold: !!sl.boldTask }) : {};
          runs.push({ text: t, options: { ...o, ...(brk ? { breakLine: true } : {}) } });
        });
      });
      T(s, runs, { x: 1.05, y: y0, w: 7.4, h: 6.55 - y0, fontSize: sl.size || (lines.length > 6 ? 14.5 : 16), valign: 'top', paraSpaceAfter: 6, lineSpacingMultiple: sl.lines ? 1.0 : 1.25 });
      const ph = Math.min(0.85, 2.9 / pens.length);
      pens.forEach((k, j) => {
        const y = 1.95 + j * (ph + 0.12);
        card(s, 9.05, y, 3.5, ph, PENS[k][0], PENS[k][2]);
        T(s, [{ text: PENS[k][2], options: { bold: true, color: PENS[k][1], breakLine: true } }, { text: PENS[k][3], options: { fontSize: 11, color: NAVY } }], { x: 9.3, y, w: 3.1, h: ph, fontSize: 14 });
      });
      const qy = 1.95 + pens.length * (ph + 0.12) + 0.05;
      if (sl.qs) {
        card(s, 9.05, qy, 3.5, 6.7 - qy, 'FFFFFF', 'Questions', true, LINE);
        T(s, sl.qs.map(([q, a], j) => ({ text: key ? `${q} → ${a}` : q, options: { breakLine: j < sl.qs.length - 1, paraSpaceAfter: 5 } })), { x: 9.25, y: qy + 0.1, w: 3.15, h: 6.5 - qy, fontSize: 11, valign: 'top' });
      }
      s.addNotes(key ? (sl.keyNotes || 'Reveal after the class has decided. Ask: "Which word gave it away?"') : notes(sl, sl.lines ? 'READ ALOUD ONCE:\n' + sl.lines.map(([w, t]) => `${w}: ${plain(t)}`).join('\n') : ''));
    }
  };
  // two-column notes (think aloud)
  L.notes2 = async sl => {
   for (const key of sl.reveal ? [false, true] : [true]) {
    const s = newSlide(); eyebrow(s, sl.reveal ? (key ? 'Model notes' : 'Your turn') : eb(sl), kk(sl)); title(s, sl.reveal && key ? sl.h + ' · model notes' : sl.h);
    T(s, sl.ctx, { x: 0.75, y: 1.7, w: 11.8, h: 0.4, fontSize: 13, color: MUTED });
    for (let ci = 0; ci < 2; ci++) {
      const x = 0.75 + ci * 6.0, k = ci ? 'R' : 'L';
      card(s, x, 2.2, 5.8, 3.0, SK[k][0], sl.cols[ci]);
      await chip(s, x + 0.3, 2.4, 0.55, k, 'LuUser', sl.cols[ci]);
      T(s, sl.cols[ci], { x: x + 1.05, y: 2.4, w: 4.5, h: 0.55, fontSize: 16, bold: true, color: SK[k][1] });
      const nt = key ? sl.lines.filter(l => l[2] === ci).map(l => l[3]) : ['…'];
      if (!key) { T(s, 'Your notes here', { x: x + 0.3, y: 3.1, w: 5.2, h: 0.5, fontSize: 14, italic: true, color: MUTED }); continue; }
      T(s, nt.map((t, j) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: j < nt.length - 1, paraSpaceAfter: 6 } })), { x: x + 0.3, y: 3.1, w: 5.2, h: 2.0, fontSize: 15, fontFace: 'Consolas', valign: 'top' });
    }
    await tipsBox(s, 0.75, 5.4, 11.8, 1.25, sl.tips || ['One column per speaker. Words, numbers, arrows. No full sentences.'], 'Tips', 'Remember');
    if (sl.reveal && key) { s.addNotes('Compare with the class. Ask: "What did you write for the change of plan?"' + (sl.qs ? '\nQUESTIONS:\n' + sl.qs.map(([q, a]) => `• ${q} → ${a}`).join('\n') : '')); continue; }
    s.addNotes(notes(sl, (sl.reveal ? 'Students take notes alone. READ THE CLIP ONCE at normal speed:\n' : 'READ ONE LINE AT A TIME, pause and think aloud after each line:\n') + sl.lines.map(l => `${l[0]}: ${l[1]}\n   → note: ${l[3]}${l[4] ? '   | think aloud: ' + l[4] : ''}`).join('\n') + (sl.qs ? '\n\nQUESTIONS:\n' + sl.qs.map(([q, a]) => `• ${q} → ${a}`).join('\n') : '')));
   }
  };
  // sort sentences into categories: a task slide and an answer-key slide
  L.sort = async sl => {
    for (const key of [false, true]) {
      const s = newSlide(); eyebrow(s, key ? 'Answer key' : (sl.eb || 'Your turn'), kk(sl)); title(s, key ? sl.h + ' · answer key' : sl.h);
      if (!key) {
        const rh = Math.min(0.6, (3.65 - (sl.items.length - 1) * 0.08) / sl.items.length);
        sl.items.forEach(([t], j) => {
          const y = 1.95 + j * (rh + 0.08);
          card(s, 0.75, y, 11.8, rh, j % 2 ? 'FFFFFF' : 'F6F7FB', 'Sentence ' + (j + 1), false, j % 2 ? LINE : null);
          oval(s, 0.9, y + (rh - 0.4) / 2, 0.4, kk(sl), j + 1, 'Number ' + (j + 1));
          T(s, t, { x: 1.5, y, w: 10.9, h: rh, fontSize: 14.5 });
        });
        await tipsBox(s, 0.75, 5.75, 11.8, 0.95, [sl.cats.map(c => `${c[1]} = ${c[2]}`).join('  ·  ')], 'Sort', 'Sort each one');
      } else {
        const n = sl.cats.length, w = (11.8 - 0.2 * (n - 1)) / n;
        sl.cats.forEach(([k, name, desc], j) => {
          const x = 0.75 + j * (w + 0.2), c = ['O', 'L', 'W', 'R'][j % 4];
          card(s, x, 1.95, w, 4.75, SK[c][0], name);
          T(s, name, { x: x + 0.3, y: 2.15, w: w - 0.6, h: 0.5, fontSize: 19, bold: true, color: SK[c][1] });
          T(s, desc, { x: x + 0.3, y: 2.65, w: w - 0.6, h: 0.4, fontSize: 12, color: MUTED });
          const its = sl.items.filter(x2 => x2[1] === k).map(x2 => x2[0]);
          T(s, its.map((t, q) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: q < its.length - 1, paraSpaceAfter: 8 } })), { x: x + 0.3, y: 3.2, w: w - 0.6, h: 3.3, fontSize: 13, valign: 'top' });
        });
      }
      s.addNotes(key ? (sl.keyNotes || 'Go through each item: "Why does it belong here?"') : notes(sl, 'The next slide is the answer key.'));
    }
  };
  // who said it: statements x speakers grid, with an answer key
  L.who = async sl => {
    for (const key of [false, true]) {
      const s = newSlide(); eyebrow(s, key ? 'Answer key' : (sl.eb || 'Your turn'), kk(sl)); title(s, key ? sl.h + ' · answer key' : sl.h);
      T(s, sl.ctx, { x: 0.75, y: 1.7, w: 11.8, h: 0.4, fontSize: 13, color: MUTED });
      const x0 = 0.75, nameW = 1.2, stW = 11.8 - nameW * sl.names.length;
      T(s, 'Who says this?', { x: x0 + 0.2, y: 2.2, w: stW - 0.2, h: 0.4, fontSize: 12, bold: true, color: MUTED });
      sl.names.forEach((n, j) => T(s, n, { x: x0 + stW + j * nameW, y: 2.2, w: nameW, h: 0.4, fontSize: 14, bold: true, color: SK[['L', 'R', 'S'][j]][1], align: 'center' }));
      const rh = Math.min(0.56, 4.0 / sl.rows.length);
      sl.rows.forEach(([st, a], r) => {
        const y = 2.65 + r * rh;
        card(s, x0, y, 11.8, rh - 0.08, r % 2 ? 'FFFFFF' : 'F6F7FB', 'Row ' + (r + 1), false, r % 2 ? LINE : null);
        T(s, st, { x: x0 + 0.2, y, w: stW - 0.3, h: rh - 0.08, fontSize: 14.5 });
        sl.names.forEach((n, j) => {
          const cx = x0 + stW + j * nameW + nameW / 2 - 0.19, on = key && a === j, by = y + (rh - 0.08 - 0.38) / 2;
          s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: by, w: 0.38, h: 0.38, rectRadius: 0.08, fill: { color: on ? 'DDF3E4' : 'FFFFFF' }, line: { color: on ? '17803D' : 'C9CFDD', width: 1 }, objectName: `Box ${r + 1}-${n}` });
          if (on) T(s, '✓', { x: cx, y: by, w: 0.38, h: 0.38, fontSize: 14, bold: true, color: '17803D', align: 'center' });
        });
      });
      s.addNotes(key ? 'Play or read the discussion again. Ask: "Who changed their mind, and why?"' : notes(sl, 'READ THE DISCUSSION ONCE (change your voice for each speaker):\n' + sl.lines.map(([w, t]) => `${w}: ${t}`).join('\n')));
    }
  };
  // reading passage + questions, then an answer key with proof lines
  L.passage = async sl => {
    for (const key of [false, true]) {
      const s = newSlide(); eyebrow(s, key ? 'Answer key' : (sl.eb || 'Your turn'), kk(sl)); title(s, key ? sl.h + ' · answer key' : sl.h);
      card(s, 0.75, 1.95, 6.9, 4.75, 'F6F7FB', 'Passage');
      T(s, sl.src, { x: 1.05, y: 2.05, w: 6.3, h: 0.4, fontSize: 12, bold: true, color: SK[kk(sl)][1] });
      const runs = [];
      sl.paras.forEach((p, j) => parseMarks(p).forEach(([t, m], si, arr) => runs.push({ text: t, options: { ...(m && key ? { highlight: PENS.m[0], bold: true } : {}), ...(si === arr.length - 1 && j < sl.paras.length - 1 ? { breakLine: true } : {}) } })));
      T(s, runs, { x: 1.05, y: 2.5, w: 6.3, h: 4.05, fontSize: sl.size || 15, valign: 'top', paraSpaceAfter: 10 });
      card(s, 7.9, 1.95, 4.65, 4.75, 'FFFFFF', 'Questions', true, LINE);
      const qs = sl.qs, qh = 4.55 / qs.length;
      qs.forEach(([q, opts, a, proof], j) => {
        const y = 2.05 + j * qh;
        oval(s, 8.1, y + 0.05, 0.36, kk(sl), j + 1, 'Q' + (j + 1));
        const parts = [{ text: q, options: { bold: true, breakLine: true } }];
        if (key) parts.push({ text: '✓ ' + (opts ? opts[a] : a), options: { color: '17803D', bold: true, breakLine: !!proof } }), proof && parts.push({ text: proof, options: { color: MUTED, fontSize: 10.5 } });
        else if (opts) parts.push({ text: opts.map((o, k) => `${'ABCD'[k]}) ${o}`).join('   '), options: { fontSize: 11 } });
        T(s, parts, { x: 8.6, y, w: 3.8, h: qh - 0.05, fontSize: 12, valign: 'top' });
      });
      s.addNotes(key ? (sl.keyNotes || 'For each answer, point to the proof line in the text. Ask: "Which words in the question match which words in the text?"') : notes(sl, 'ANSWERS: ' + qs.map(([q, o, a], j) => `${j + 1}) ${o ? o[a] : a}`).join('  ·  ')));
    }
  };
  L.table = async sl => {
    const s = newSlide(sl.master); eyebrow(s, eb(sl), kk(sl)); title(s, sl.h);
    const tips = sl.tips || [], head = sl.head, rows = sl.rows, cols = head.length;
    const firstW = sl.firstW || 4.6, otherW = (11.8 - firstW) / (cols - 1), avail = tips.length ? 3.15 : 4.4, rh = Math.min(0.62, avail / (rows.length + 1));
    head.forEach((h, c) => T(s, String(h).toUpperCase(), { x: 0.75 + (c ? firstW + (c - 1) * otherW : 0) + 0.2, y: 1.95, w: (c ? otherW : firstW) - 0.2, h: rh, fontSize: 10.5, bold: true, color: MUTED, align: c ? 'center' : 'left', charSpacing: 2 }));
    rows.forEach((r, i) => {
      const y = 1.95 + (i + 1) * rh;
      card(s, 0.75, y + 0.04, 11.8, rh - 0.08, i % 2 ? 'FFFFFF' : 'F6F7FB', 'Row ' + (i + 1), false, i % 2 ? LINE : null);
      r.forEach((v, c) => T(s, String(v), { x: 0.75 + (c ? firstW + (c - 1) * otherW : 0) + 0.2, y, w: (c ? otherW : firstW) - 0.2, h: rh, fontSize: c ? 14 : 14.5, bold: !!c && sl.boldVals, align: c ? 'center' : 'left', color: c && sl.colorVals ? SK[ROT[(c - 1) % 5]][1] : NAVY }));
    });
    if (tips.length) await tipsBox(s, 0.75, 5.45, 11.8, 1.2, tips, 'Tips', sl.tipHead || 'Remember');
    s.addNotes(notes(sl));
  };
  L.ruler = async sl => {
    const s = newSlide(); eyebrow(s, eb(sl), kk(sl)); title(s, sl.h);
    sl.items.forEach(([a, b], j) => {
      const x = 0.75 + j * 6.0, k = j ? 'R' : 'L';
      card(s, x, 1.95, 5.8, 1.7, SK[k][0], a);
      T(s, [{ text: a, options: { bold: true, fontSize: 20, color: SK[k][1], breakLine: true } }, { text: b, options: { fontSize: 14 } }], { x: x + 0.3, y: 1.95, w: 5.2, h: 1.7 });
    });
    T(s, 'CLB RULER', { x: 0.75, y: 3.95, w: 4, h: 0.35, fontSize: 11, bold: true, color: MUTED, charSpacing: 3 });
    for (let i = 1; i <= 12; i++) {
      const x = 0.75 + (i - 1) * 0.985, k = i <= 4 ? 'W' : i <= 6 ? 'R' : i <= 8 ? 'L' : 'O', on = sl.hi && sl.hi.includes(i);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.35, w: 0.9, h: 0.75, rectRadius: 0.12, fill: { color: on ? SK[k][1] : SK[k][0] }, line: { color: SK[k][1], width: on ? 0 : 0.75 }, objectName: 'CLB ' + i });
      T(s, String(i), { x, y: 4.35, w: 0.9, h: 0.75, fontSize: 18, bold: true, align: 'center', color: on ? 'FFFFFF' : SK[k][1] });
    }
    if (sl.tips) await tipsBox(s, 0.75, 5.45, 11.8, 1.2, sl.tips, 'Tips', sl.tipHead || 'Remember');
    s.addNotes(notes(sl));
  };
  // a model email/answer with labelled parts
  L.model = async sl => {
    const s = newSlide(); eyebrow(s, eb(sl), kk(sl)); title(s, sl.h);
    const tw = sl.task ? 4.1 : 0;
    if (sl.task) {
      card(s, 0.75, 1.95, 3.85, 4.75, SK[kk(sl)][0], 'Task');
      T(s, 'THE TASK', { x: 1.0, y: 2.1, w: 3.3, h: 0.35, fontSize: 11, bold: true, color: SK[kk(sl)][1], charSpacing: 3 });
      T(s, [{ text: sl.task, options: { breakLine: true } }, ...(sl.bullets || []).map((b, i) => ({ text: b, options: { bullet: { indent: 12 }, breakLine: i < sl.bullets.length - 1 } }))], { x: 1.0, y: 2.5, w: 3.4, h: 4.0, fontSize: 13, valign: 'top', paraSpaceAfter: 6 });
    }
    const P = sl.parts, ph = Math.min(1.25, (4.75 - (P.length - 1) * 0.1) / P.length);
    P.forEach(([lab, txt], j) => {
      const y = 1.95 + j * (ph + 0.1), k = ROT[j % 5], x = 0.75 + tw, w = 11.8 - tw;
      card(s, x, y, w, ph, j % 2 ? 'FFFFFF' : 'F6F7FB', 'Part ' + (j + 1), j % 2 === 1, j % 2 ? LINE : null);
      const lw = Math.max(1.3, Math.min(2.2, 0.3 + String(lab).length * 0.1));
      softPill(s, x + 0.2, y + 0.15, lw, 0.38, k, lab, 'Label ' + (j + 1));
      T(s, txt, { x: x + 0.25, y: y + 0.55, w: w - 0.45, h: ph - 0.6, fontSize: sl.size || 13, valign: 'top' });
    });
    s.addNotes(notes(sl));
  };
  // exit ticket: true / false statements, answers in the notes
  L.tf = async sl => {
    const s = newSlide(); eyebrow(s, sl.eb || 'Exit ticket', kk(sl)); title(s, sl.h || 'True or false?');
    const R = sl.items, rh = Math.min(0.72, (4.7 - (R.length - 1) * 0.1) / R.length);
    R.forEach(([st], j) => {
      const y = 1.95 + j * (rh + 0.1);
      card(s, 0.75, y, 11.8, rh, j % 2 ? 'FFFFFF' : 'F6F7FB', 'Statement ' + (j + 1), false, j % 2 ? LINE : null);
      oval(s, 0.95, y + (rh - 0.46) / 2, 0.46, ROT[j % 5], j + 1, 'Number ' + (j + 1));
      T(s, st, { x: 1.65, y, w: 8.4, h: rh, fontSize: 16 });
      softPill(s, 10.25, y + (rh - 0.4) / 2, 2.1, 0.4, 'G', 'True  /  False', 'TF ' + (j + 1));
    });
    s.addNotes(notes(sl, 'Students answer with fingers (1 = true, 2 = false).\nANSWERS: ' + R.map(([, a, w], j) => `${j + 1} ${a ? 'True' : 'False'}${w ? ' (' + w + ')' : ''}`).join(' · ')));
  };
  L.closing = async sl => {
    sec = 'Wrap-up'; pres.addSection({ title: sec });
    const s = newSlide('CLOSING');
    eyebrow(s, `Study Hub · Session ${N}`, 'O'); title(s, N < 12 ? `Before Session ${N + 1}` : 'Before your test');
    const todo = sl.items, rowH = Math.min(0.7, (4.6 - (todo.length - 1) * 0.12) / todo.length);
    todo.forEach((t, i) => {
      const y = 1.95 + i * (rowH + 0.12);
      card(s, 0.75, y, 7.4, rowH, 'FFFFFF', 'To-do ' + (i + 1), true);
      oval(s, 0.97, y + (rowH - 0.42) / 2, 0.42, 'O', '✓', 'Check ' + (i + 1));
      T(s, t, { x: 1.6, y, w: 6.4, h: rowH, fontSize: 16 });
    });
    card(s, 8.6, 1.95, 3.9, 4.15, 'FFFFFF', 'Next class', true);
    await chip(s, 8.95, 2.3, 0.85, 'O', 'LuFlag', 'Next class', true);
    T(s, [{ text: 'Your results open the next class.', options: { bold: true, fontSize: 18, breakLine: true } }, { text: sl.next || `Finish the Study Hub before Session ${N + 1} so we can work on your real mistakes.`, options: { fontSize: 14, color: MUTED } }],
      { x: 8.95, y: 3.4, w: 3.25, h: 2.4, valign: 'top', paraSpaceAfter: 10 });
    s.addNotes(notes(sl));
  };

  // ---------- build ----------
  pres.addSection({ title: sec });
  for (const sl of C.slides) {
    const fn = L[sl.t];
    if (!fn) throw new Error('Unknown slide type ' + sl.t);
    await fn(sl);
  }
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log('written', OUT, '·', pres.slides.length, 'slides ·', hm(clock), 'of teaching');
})().catch(e => { console.error(e); process.exit(1); });
