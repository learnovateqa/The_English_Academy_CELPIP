// Session 1 deck · CELPIP Prep Course · Elevate with Ms. Maria · Learnovate Qatar
// Build:  NODE_PATH=<folder with pptxgenjs, react-icons, react, react-dom, sharp>/node_modules \
//         APPLY_THEME=<pptx skill>/scripts/apply_theme.js node tools/session-01-deck.js decks/session-01.pptx
const pptxgen = require('pptxgenjs');
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const sharp = require('sharp');
const Lu = require('react-icons/lu');
const { applyTheme } = require(process.env.APPLY_THEME || './apply_theme.js');

const THEME = { name: 'Elevate Pastel', headFontFace: 'Poppins', bodyFontFace: 'Poppins',
  colors: { dk1: '14213D', lt1: 'FFFFFF', dk2: '5B6782', lt2: 'F6F7FB', accent1: '1A5FCC', accent2: '6A2BC4', accent3: 'C93D1A',
    accent4: '007F7F', accent5: '008F55', accent6: 'E0A800', hlink: '1A5FCC', folHlink: '6A2BC4' } };
// [soft fill, ink]
const SK = { L: ['EAF2FF', '1A5FCC'], R: ['F3ECFE', '6A2BC4'], W: ['FFEFEA', 'C93D1A'], S: ['E5F7F6', '007F7F'], O: ['E7F7EF', '008F55'], G: ['FFF7DB', '8A6A00'] };
const NAVY = '14213D', MUTED = '5B6782', LINE = 'E6E9F2';

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
  pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
  pres.author = 'Ms. Maria · Learnovate Qatar';
  pres.title = 'Session 1 · Decode the Test Engine and Set Your CLB Target';
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const titleBg = await gradient(1920, 1080, [[0, 'F3ECFE'], [0.5, 'FFEFEA'], [1, 'EAF2FF']]);
  const closeBg = await gradient(1920, 1080, [[0, 'EAF2FF'], [1, 'E7F7EF']]);

  pres.defineSlideMaster({ title: 'TITLE', background: { data: titleBg },
    objects: [
      { placeholder: { options: { name: 'title', type: 'title', x: 0.75, y: 2.3, w: 7.6, h: 2.3, fontSize: 42, bold: true, color: C.text1, valign: 'top', align: 'left', margin: 0 }, text: '' } },
      { placeholder: { options: { name: 'body', type: 'body', x: 0.75, y: 4.75, w: 7.6, h: 0.6, fontSize: 18, color: C.text2, align: 'left', margin: 0 }, text: '' } },
      { text: { text: 'ELEVATE WITH MS. MARIA  ·  LEARNOVATE QATAR', options: { x: 0.75, y: 6.6, w: 8, h: 0.35, fontSize: 10, bold: true, color: C.text2, charSpacing: 3, margin: 0 } } }
    ] });
  const contentObjects = () => [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.75, y: 0.88, w: 11.8, h: 0.8, fontSize: 32, bold: true, color: C.text1, valign: 'top', align: 'left', margin: 0 }, text: '' } },
    { text: { text: 'CELPIP Prep Course  ·  Session 1', options: { x: 0.75, y: 6.95, w: 6, h: 0.3, fontSize: 9, color: C.text2, margin: 0 } } }];
  pres.defineSlideMaster({ title: 'CONTENT', background: { color: 'FFFFFF' }, objects: contentObjects(),
    slideNumber: { x: 12.08, y: 6.95, w: 0.5, h: 0.3, fontSize: 9, color: '5B6782', align: 'right' } });
  pres.defineSlideMaster({ title: 'CLOSING', background: { data: closeBg }, objects: contentObjects(),
    slideNumber: { x: 12.08, y: 6.95, w: 0.5, h: 0.3, fontSize: 9, color: '5B6782', align: 'right' } });

  // ---------- helpers ----------
  const title = (s, t) => s.addText(t, { placeholder: 'title' });
  const eyebrow = (s, text, k) => s.addText(text.toUpperCase(), { x: 0.75, y: 0.5, w: 11, h: 0.32, fontSize: 11, bold: true, color: SK[k][1], charSpacing: 3, margin: 0, isTextBox: true });
  const card = (s, x, y, w, h, fill, name, sh) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.22, fill: { color: fill }, line: { color: fill }, objectName: name, ...(sh ? { shadow: shadow() } : {}) });
  const chip = async (s, x, y, d, k, ic, name, onWhite) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: d, h: d, rectRadius: d * 0.3, fill: { color: onWhite ? SK[k][0] : 'FFFFFF' }, line: { color: onWhite ? SK[k][0] : 'FFFFFF' }, objectName: name + ' chip' });
    s.addImage({ data: await icon(ic, SK[k][1]), x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52, altText: name });
  };
  const badge = (s, x, y, d, k, label, name, solid = true) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: solid ? SK[k][1] : 'FFFFFF' }, line: { color: SK[k][1], width: solid ? 0 : 1.25 }, objectName: name });
    s.addText(String(label), { x, y, w: d, h: d, align: 'center', valign: 'middle', fontSize: 12, bold: true, color: solid ? 'FFFFFF' : SK[k][1], margin: 0, isTextBox: true });
  };
  const tips = async (s, x, y, w, h, list, name, headline = 'Pro tips') => {
    card(s, x, y, w, h, 'FFFFFF', name, true);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.22, fill: { color: 'FFFFFF', transparency: 100 }, line: { color: LINE, width: 1 }, objectName: name + ' outline' });
    if (h < 1.6) {
      await chip(s, x + 0.3, y + (h - 0.55) / 2, 0.55, 'G', 'LuLightbulb', 'Tip', true);
      s.addText([{ text: headline + '   ', options: { bold: true } }, { text: list.join(' ') }], { x: x + 1.05, y: y + 0.1, w: w - 1.35, h: h - 0.2, fontSize: 14, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
      return;
    }
    await chip(s, x + 0.3, y + 0.3, 0.55, 'G', 'LuLightbulb', 'Tip', true);
    s.addText(headline, { x: x + 1.0, y: y + 0.3, w: w - 1.3, h: 0.55, fontSize: 15, bold: true, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
    s.addText(list.map((t, i) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < list.length - 1, paraSpaceAfter: 8 } })),
      { x: x + 0.3, y: y + 1.0, w: w - 0.6, h: h - 1.2, fontSize: 14, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  };
  // a clean list of rows (replaces tables): label left, value pill right
  const rows = (s, x, y, w, rowH, gap, items, k, name) => items.forEach(([label, val, sub], i) => {
    const yy = y + i * (rowH + gap);
    card(s, x, yy, w, rowH, i % 2 ? 'FFFFFF' : 'F6F7FB', name + ' row ' + (i + 1));
    s.addText(label, { x: x + 0.3, y: yy, w: w - 2.1, h: rowH, fontSize: 15, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
    if (sub) s.addText(sub, { x: x + w - 3.55, y: yy, w: 1.6, h: rowH, fontSize: 12, color: MUTED, valign: 'middle', align: 'right', margin: 0, isTextBox: true });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + w - 1.75, y: yy + (rowH - 0.4) / 2, w: 1.5, h: 0.4, rectRadius: 0.2, fill: { color: SK[k][0] }, line: { color: SK[k][0] }, objectName: name + ' value ' + (i + 1) });
    s.addText(val, { x: x + w - 1.75, y: yy + (rowH - 0.4) / 2, w: 1.5, h: 0.4, fontSize: 12, bold: true, color: SK[k][1], align: 'center', valign: 'middle', margin: 0, isTextBox: true });
  });

  // 1 · Title
  pres.addSection({ title: 'Welcome' });
  let s = pres.addSlide({ masterName: 'TITLE', sectionTitle: 'Welcome' });
  s.addText('SESSION 01', { x: 0.75, y: 1.75, w: 4, h: 0.4, fontSize: 13, bold: true, color: '6A2BC4', charSpacing: 4, margin: 0, isTextBox: true });
  s.addText('Decode the Test Engine and Set Your CLB Target', { placeholder: 'title' });
  s.addText('CELPIP Prep Course  ·  2-hour live class', { placeholder: 'body' });
  const four = [['L', 'LuHeadphones', 'Listening'], ['R', 'LuBookOpen', 'Reading'], ['W', 'LuPenLine', 'Writing'], ['S', 'LuMic', 'Speaking']];
  for (let i = 0; i < 4; i++) {
    const [k, ic, n] = four[i], x = 9.05 + (i % 2) * 1.85, y = 2.0 + Math.floor(i / 2) * 2.05;
    card(s, x, y, 1.65, 1.85, 'FFFFFF', n + ' tile', true);
    await chip(s, x + 0.43, y + 0.3, 0.8, k, ic, n, true);
    s.addText(n, { x, y: y + 1.2, w: 1.65, h: 0.4, fontSize: 12, bold: true, color: NAVY, align: 'center', margin: 0, isTextBox: true });
  }
  s.addNotes('Welcome everyone. Say: "Today you learn exactly how the test works, how it is scored, and what level you need. Every slide is something you will use on test day."\n\nDELIVERY TIPS FOR THIS CLASS\n• Pace: speak about 15% slower than normal conversation. Pause after every acronym (CLB, IRCC, TEER) and write it on the board.\n• Check understanding with questions, not "Do you understand?" Students often nod even when confused. Use the CHECK question in each slide\'s notes.\n• Point to the icons every time: headphones = Listening, book = Reading, pen = Writing, microphone = Speaking.');

  // 2 · Outcomes
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Welcome' });
  eyebrow(s, "Today's outcomes", 'R'); title(s, 'By the end of today you can…');
  const outs = ['Name every part of the test and its timing', 'Explain how your level is scored', 'Find the CLB level your program needs', 'Avoid the 10 mistakes that cost the most points', 'Set a target for each skill'];
  outs.forEach((t, i) => {
    const y = 1.95 + i * 0.86, k = ['L', 'R', 'O', 'W', 'S'][i];
    badge(s, 0.75, y + 0.12, 0.46, k, i + 1, 'Outcome ' + (i + 1));
    s.addText(t, { x: 1.45, y, w: 6.6, h: 0.7, fontSize: 18, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  });
  card(s, 8.6, 1.95, 4.0, 4.3, 'F6F7FB', 'Agenda');
  s.addText('2 HOURS', { x: 8.95, y: 2.2, w: 3.3, h: 0.3, fontSize: 11, bold: true, color: MUTED, charSpacing: 3, margin: 0, isTextBox: true });
  [['0:00', 'Map the test'], ['0:20', 'How scoring works'], ['0:45', 'The test screen'], ['1:15', 'Mistakes & error causes'], ['1:45', 'Set your targets']].forEach(([a, b], i) => {
    const y = 2.65 + i * 0.68;
    s.addText(a, { x: 8.95, y, w: 0.8, h: 0.5, fontSize: 14, bold: true, color: '6A2BC4', valign: 'middle', margin: 0, isTextBox: true });
    s.addText(b, { x: 9.8, y, w: 2.6, h: 0.5, fontSize: 14, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  });
  s.addNotes('Read the outcomes aloud. Return to this slide at the end and ask students to tick each one.\nCHECK: "Which outcome is most important for you?"');

  // 2b · Key words
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Welcome' });
  eyebrow(s, 'Key words for today', 'O'); title(s, '8 words you will hear today');
  const words = [['L', 'CLB', 'Canada’s ruler for English level, from 1 to 12.'], ['R', 'IRCC', 'Canada’s immigration department.'], ['W', 'Paraphrase', 'The same idea in different words.'], ['S', 'Distractor', 'A wrong answer that sounds right. Often said first, then changed.'],
    ['O', 'Inference', 'What someone means but does not say directly.'], ['L', 'Rater', 'A trained person who scores your writing and speaking.'], ['R', 'Task fulfillment', 'Doing everything the task asks.'], ['W', 'Prep time', 'Planning time before you speak or write.']];
  words.forEach(([k, w, d], i) => {
    const x = 0.75 + (i % 4) * 3.0, y = 1.95 + Math.floor(i / 4) * 2.3;
    card(s, x, y, 2.8, 2.1, SK[k][0], w);
    s.addText(w, { x: x + 0.3, y: y + 0.25, w: 2.3, h: 0.5, fontSize: 19, bold: true, color: SK[k][1], margin: 0, isTextBox: true });
    s.addText(d, { x: x + 0.3, y: y + 0.8, w: 2.3, h: 1.15, fontSize: 14, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  });
  s.addNotes('Say each word slowly and have students repeat it. Pause after CLB and IRCC.\nCHECK: "What is a distractor?" "Which word means the same idea in different words?"');

  // 3 · Right test
  pres.addSection({ title: 'The test' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The test' });
  eyebrow(s, 'Before you book', 'L'); title(s, 'Pick the right test');
  const tests = [['L', 'LuAward', 'CELPIP-General', 'All four skills', 'For permanent residence: Express Entry and other economic programs.'],
    ['S', 'LuMic', 'CELPIP-General LS', 'Listening + Speaking', 'For Canadian citizenship applications.'],
    ['G', 'LuSearch', 'Check first', 'Which test? Which levels?', 'Confirm with your program before you pay for a test.']];
  for (let i = 0; i < 3; i++) {
    const [k, ic, h, tag, b] = tests[i], x = 0.75 + i * 4.0;
    card(s, x, 1.95, 3.7, 4.3, SK[k][0], h);
    await chip(s, x + 0.35, 2.3, 0.85, k, ic, h);
    s.addText(h, { x: x + 0.35, y: 3.4, w: 3.0, h: 0.55, fontSize: 21, bold: true, color: NAVY, margin: 0, isTextBox: true });
    s.addText(tag, { x: x + 0.35, y: 3.95, w: 3.0, h: 0.4, fontSize: 13, bold: true, color: SK[k][1], margin: 0, isTextBox: true });
    s.addText(b, { x: x + 0.35, y: 4.5, w: 3.0, h: 1.5, fontSize: 15, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  }
  s.addNotes('Booking the wrong version costs money and time. Ask who is applying for PR and who for citizenship.\nCHECK: "You are applying for citizenship. Which test do you book?"');

  // 4 · Glance
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The test' });
  eyebrow(s, 'The test at a glance', 'O'); title(s, 'Four skills, about 3 hours, one sitting');
  const g = [['L', 'LuHeadphones', 'Listening', '47–55', 'min', '6 parts · 38 Qs'], ['R', 'LuBookOpen', 'Reading', '55–60', 'min', '4 parts · 38 Qs'], ['W', 'LuPenLine', 'Writing', '53–60', 'min', '2 tasks'], ['S', 'LuMic', 'Speaking', '15–20', 'min', '8 tasks']];
  for (let i = 0; i < 4; i++) {
    const [k, ic, n, big, unit, sub] = g[i], x = 0.75 + i * 3.0;
    card(s, x, 1.95, 2.75, 3.05, SK[k][0], n);
    await chip(s, x + 0.3, 2.25, 0.7, k, ic, n);
    s.addText(String(i + 1).padStart(2, '0'), { x: x + 1.9, y: 2.3, w: 0.6, h: 0.4, fontSize: 12, bold: true, color: SK[k][1], align: 'right', margin: 0, isTextBox: true });
    s.addText(n, { x: x + 0.3, y: 3.1, w: 2.2, h: 0.45, fontSize: 18, bold: true, color: NAVY, margin: 0, isTextBox: true });
    s.addText([{ text: big, options: { fontSize: 30, bold: true, color: SK[k][1] } }, { text: ' ' + unit, options: { fontSize: 14, color: MUTED } }], { x: x + 0.3, y: 3.55, w: 2.3, h: 0.7, margin: 0, valign: 'middle', isTextBox: true });
    s.addText(sub, { x: x + 0.3, y: 4.3, w: 2.3, h: 0.45, fontSize: 13, color: MUTED, margin: 0, isTextBox: true });
  }
  await tips(s, 0.75, 5.25, 11.75, 1.5, ['Order: Listening → Reading → Writing → Speaking. Train stamina with full timed practice; Speaking comes last.'], 'Glance tip', 'Test-day tip');
  s.addNotes('Point to each icon as you name the skill. Stress stamina. Confirm current details on the official CELPIP website before class.\nCHECK: "Which skill comes last?" "Which skill is the shortest?"');

  // 5–6 · Listening & Reading
  pres.addSection({ title: 'The four skills' });
  const skillSlide = async (k, ic, eb, ttl, items, tipList, notes, rowH = 0.52) => {
    const sl = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The four skills' });
    eyebrow(sl, eb, k); title(sl, ttl);
    rows(sl, 0.75, 1.95, 7.2, rowH, 0.1, items, k, ttl);
    await tips(sl, 8.35, 1.95, 4.15, 4.6, tipList, 'Strategies', 'How to score higher');
    sl.addNotes(notes);
  };
  await skillSlide('L', 'LuHeadphones', 'Listening · 6 parts · 38 questions', 'Hear it once. Answer with confidence.',
    [['Part 1 · Problem Solving', '8 Qs'], ['Part 2 · Daily Life Conversation', '5 Qs'], ['Part 3 · Listening for Information', '6 Qs'], ['Part 4 · Listening to a News Item', '5 Qs'], ['Part 5 · Listening to a Discussion (video)', '8 Qs'], ['Part 6 · Listening to Viewpoints', '6 Qs']],
    ['Audio plays once: read the context screen and predict first.', 'The first answer you hear is often corrected. Wait for the final decision.', 'Note who, what, when, and any change of plan.'],
    'Point to the headphones icon. Model one prediction from a context screen.\nCHECK: "How many times can you hear the audio?"');
  // 6b · Distractor example
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The four skills' });
  eyebrow(s, 'Listening · example', 'L'); title(s, 'Spot the corrected answer');
  card(s, 0.75, 1.95, 6.6, 3.0, 'F6F7FB', 'Transcript');
  await chip(s, 1.05, 2.25, 0.7, 'L', 'LuHeadphones', 'Audio', true);
  s.addText('You hear:', { x: 1.95, y: 2.25, w: 4, h: 0.7, fontSize: 15, bold: true, color: MUTED, valign: 'middle', margin: 0, isTextBox: true });
  s.addText([{ text: '“Let’s meet on ', options: {} }, { text: 'Tuesday', options: { bold: true, color: SK.W[1] } }, { text: '… oh wait, Tuesday is busy. ', options: {} }, { text: 'Let’s do ', options: {} }, { text: 'Wednesday', options: { bold: true, color: SK.O[1] } }, { text: '.”', options: {} }],
    { x: 1.05, y: 3.1, w: 6.0, h: 1.6, fontSize: 22, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  card(s, 7.65, 1.95, 4.85, 3.0, 'FFFFFF', 'Question', true);
  s.addText('When will they meet?', { x: 7.95, y: 2.2, w: 4.3, h: 0.5, fontSize: 17, bold: true, color: NAVY, margin: 0, isTextBox: true });
  [['W', 'A', 'Tuesday', 'Trap: said first, then changed'], ['O', 'B', 'Wednesday', 'Correct: the final decision']].forEach(([k, l, a, why], i) => {
    const y = 2.85 + i * 0.95;
    card(s, 7.95, y, 4.3, 0.8, SK[k][0], 'Option ' + l);
    badge(s, 8.12, y + 0.17, 0.46, k, l, 'Letter ' + l);
    s.addText([{ text: a, options: { bold: true, fontSize: 15, breakLine: true } }, { text: why, options: { fontSize: 12, color: MUTED } }], { x: 8.75, y, w: 3.4, h: 0.8, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  });
  await tips(s, 0.75, 5.2, 11.75, 1.55, ['Do not grab the first detail you hear. Listen for words that change the answer: but, oh wait, actually, instead, on second thought.'], 'Signal words', 'Listen for the change');
  s.addNotes('Read the line aloud with natural self-correction. Ask students to raise a hand when they hear the change.\nCHECK: "Why is Tuesday wrong?" "Which words told you the plan changed?"');

  await skillSlide('R', 'LuBookOpen', 'Reading · 4 parts · 38 questions', 'Scan fast. Match the paraphrase.',
    [['Part 1 · Reading Correspondence', '11 Qs'], ['Part 2 · Reading to Apply a Diagram', '8 Qs'], ['Part 3 · Reading for Information', '9 Qs'], ['Part 4 · Reading for Viewpoints', '10 Qs']],
    ['Read the questions first, then scan for the answer.', 'Answers are paraphrased: "postponed" = "moved to a later date".', 'About one minute per question. Never get stuck.'],
    'Point to the book icon. Explain question-first scanning.\nCHECK: "What do you read first: the passage or the questions?"', 0.72);

  // 8b · Paraphrase pairs
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The four skills' });
  eyebrow(s, 'Reading · example', 'R'); title(s, 'Paraphrase pairs to know');
  [['postponed', 'moved to a later date'], ['mandatory', 'you must'], ['complimentary', 'free'], ['refreshments', 'drinks and snacks'], ['seating is limited', 'not everyone gets a seat'], ['commute', 'travel to and from work']].forEach(([a, b], i) => {
    const col = i % 2, row = Math.floor(i / 2), x = 0.75 + col * 5.95, y = 1.95 + row * 1.15;
    card(s, x, y, 5.8, 0.95, i % 4 === 0 || i % 4 === 3 ? 'F6F7FB' : 'FFFFFF', 'Pair ' + (i + 1));
    s.addText('“' + a + '”', { x: x + 0.3, y, w: 2.3, h: 0.95, fontSize: 16, bold: true, color: SK.R[1], valign: 'middle', margin: 0, isTextBox: true });
    s.addText('=', { x: x + 2.6, y, w: 0.4, h: 0.95, fontSize: 18, bold: true, color: MUTED, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
    s.addText(b, { x: x + 3.05, y, w: 2.6, h: 0.95, fontSize: 16, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  });
  await tips(s, 0.75, 5.5, 11.75, 1.2, ['The question and the text rarely use the same words. Look for the same meaning, not the same word.'], 'Paraphrase tip', 'Same idea, new words');
  s.addNotes('Cover the right column and ask students to guess the paraphrase.\nCHECK: "If a notice says attendance is mandatory, do you have to go?"');

  // 7 · Writing
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The four skills' });
  eyebrow(s, 'Writing · 2 tasks', 'W'); title(s, 'Answer every part of the task');
  [['Task 1 · Email', '27', 'Cover all three bullets in 150–200 words.'], ['Task 2 · Survey', '26', 'Choose one option and defend it in 150–200 words.']].forEach(([h, m, b], i) => {
    const x = 0.75 + i * 3.95;
    card(s, x, 1.95, 3.7, 2.55, SK.W[0], h);
    s.addText(h, { x: x + 0.35, y: 2.2, w: 3.0, h: 0.45, fontSize: 17, bold: true, color: NAVY, margin: 0, isTextBox: true });
    s.addText([{ text: m, options: { fontSize: 34, bold: true, color: SK.W[1] } }, { text: ' min', options: { fontSize: 14, color: MUTED } }], { x: x + 0.35, y: 2.7, w: 3.0, h: 0.7, margin: 0, valign: 'middle', isTextBox: true });
    s.addText(b, { x: x + 0.35, y: 3.45, w: 3.0, h: 0.9, fontSize: 14, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  });
  card(s, 8.65, 1.95, 3.85, 2.55, SK.R[0], 'Raters score');
  s.addText('RATERS SCORE', { x: 9.0, y: 2.2, w: 3.2, h: 0.3, fontSize: 11, bold: true, color: SK.R[1], charSpacing: 3, margin: 0, isTextBox: true });
  ['Content & coherence', 'Vocabulary', 'Readability', 'Task fulfillment'].forEach((t, i) => {
    badge(s, 9.0, 2.65 + i * 0.44, 0.3, 'R', '✓', 'Criterion ' + (i + 1));
    s.addText(t, { x: 9.45, y: 2.6 + i * 0.44, w: 2.9, h: 0.4, fontSize: 14, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  });
  await tips(s, 0.75, 4.7, 11.75, 2.1, ['Plan 3 minutes, write 20, proofread 3.', 'Give each bullet its own paragraph so the rater can find it.', 'Match the tone: formal for a manager, friendly for a friend.'], 'Writing tips', 'How to score higher');
  s.addNotes('Point to the pen icon. A missed bullet lowers task fulfillment.\nCHECK: "How many minutes for the email?" "What happens if you skip one bullet?"');

  // 10b · Plan an email
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The four skills' });
  eyebrow(s, 'Writing · example', 'W'); title(s, 'Plan your email in 3 minutes');
  card(s, 0.75, 1.95, 4.2, 4.75, SK.W[0], 'Prompt');
  s.addText('THE TASK', { x: 1.05, y: 2.2, w: 3.6, h: 0.3, fontSize: 11, bold: true, color: SK.W[1], charSpacing: 3, margin: 0, isTextBox: true });
  s.addText([{ text: 'Your neighbour’s dog barks every night. Write to your neighbour:', options: { breakLine: true } }, { text: 'introduce yourself', options: { bullet: true, breakLine: true } }, { text: 'describe the problem', options: { bullet: true, breakLine: true } }, { text: 'suggest a solution', options: { bullet: true } }],
    { x: 1.05, y: 2.6, w: 3.6, h: 3.8, fontSize: 15, color: NAVY, valign: 'top', margin: 0, paraSpaceAfter: 6, isTextBox: true });
  [['1', 'Hello + who I am', 'Hi Sam, I’m Ana from unit 4B, next door.'], ['2', 'Problem + how it affects me', 'Your dog barks at night, so I cannot sleep before work.'], ['3', 'Solution + polite close', 'Would it be possible to keep him inside after 10 p.m.? Thank you for understanding.']].forEach(([n, h, ex], i) => {
    const y = 1.95 + i * 1.6;
    card(s, 5.2, y, 7.3, 1.45, 'FFFFFF', 'Paragraph ' + n, true);
    badge(s, 5.45, y + 0.25, 0.5, 'W', n, 'Paragraph number ' + n);
    s.addText([{ text: h, options: { bold: true, fontSize: 15, breakLine: true } }, { text: ex, options: { fontSize: 13.5, color: MUTED, italic: true } }], { x: 6.15, y: y + 0.12, w: 6.1, h: 1.2, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  });
  s.addNotes('One bullet = one paragraph. Useful phrases: "I’m writing to…", "I’m afraid that…", "Would it be possible to…?", "Thank you for understanding."\nCHECK: "How many paragraphs for three bullets?" "Is this email formal or friendly?"');

  // 8 · Speaking
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The four skills' });
  eyebrow(s, 'Speaking · 8 tasks · 15–20 minutes', 'S'); title(s, 'Fill the time. Stay organized.');
  const tasks = [['Giving Advice', '30 / 90 s'], ['Personal Experience', '30 / 60 s'], ['Describing a Scene', '30 / 60 s'], ['Making Predictions', '30 / 60 s'], ['Comparing & Persuading', '60+60 / 60 s'], ['Difficult Situation', '60 / 60 s'], ['Expressing Opinions', '30 / 90 s'], ['Unusual Situation', '30 / 60 s']];
  tasks.forEach(([t, v], i) => {
    const col = i < 4 ? 0 : 1, x = 0.75 + col * 3.75, y = 1.95 + (i % 4) * 1.12;
    card(s, x, y, 3.55, 0.95, i % 2 ? 'FFFFFF' : 'F6F7FB', 'Task ' + (i + 1));
    badge(s, x + 0.2, y + 0.24, 0.46, 'S', i + 1, 'Task number ' + (i + 1));
    s.addText([{ text: t, options: { bold: true, fontSize: 14, breakLine: true } }, { text: 'prep / speak  ' + v, options: { fontSize: 11, color: MUTED } }], { x: x + 0.8, y, w: 2.65, h: 0.95, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  });
  await tips(s, 8.35, 1.95, 4.15, 4.4, ['Raters score content & coherence, vocabulary, listenability and task fulfillment.', 'Plan 2–3 points in prep time, not full sentences.', 'Keep talking until time ends: add an example or a summary.'], 'Speaking tips', 'How to score higher');
  s.addNotes('Point to the microphone icon. Ask a volunteer to speak for 60 seconds about a favourite place. Task 5 has 60 seconds to choose, then 60 seconds to prepare.\nCHECK: "Which tasks are 90 seconds long?"');

  // 12b · Speaking phrases
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The four skills' });
  eyebrow(s, 'Speaking · useful phrases', 'S'); title(s, 'Phrases to fill your speaking time');
  [['L', 'Start', ['In my opinion, …', 'I would suggest …', 'The best option is …']], ['S', 'Add & explain', ['This is because …', 'For example, …', 'Another reason is …']], ['O', 'Finish', ['To sum up, …', 'That’s why I think …', 'So my advice is …']]].forEach(([k, h, list], i) => {
    const x = 0.75 + i * 4.0;
    card(s, x, 1.95, 3.7, 3.4, SK[k][0], h);
    s.addText(h.toUpperCase(), { x: x + 0.35, y: 2.2, w: 3.0, h: 0.35, fontSize: 12, bold: true, color: SK[k][1], charSpacing: 3, margin: 0, isTextBox: true });
    s.addText(list.map((t, j) => ({ text: t, options: { breakLine: j < list.length - 1 } })), { x: x + 0.35, y: 2.75, w: 3.0, h: 2.4, fontSize: 18, color: NAVY, valign: 'top', margin: 0, paraSpaceAfter: 14, isTextBox: true });
  });
  await tips(s, 0.75, 5.6, 11.75, 1.15, ['Survival rule: if you run out of ideas, keep talking until the clock stops. Summarize your main point or add a quick example.'], 'Survival rule', 'Never stop early');
  s.addNotes('Drill the phrases chorally, then ask volunteers to answer "Coffee or tea?" for 30 seconds using one phrase from each column.\nCHECK: "What do you say if you run out of ideas with 15 seconds left?"');

  // 13b · CELPIP vs CLB
  pres.addSection({ title: 'Scores and targets' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Scores and targets' });
  eyebrow(s, 'CELPIP and CLB', 'L'); title(s, 'The test and the ruler');
  card(s, 0.75, 1.95, 5.6, 2.6, SK.R[0], 'CELPIP');
  await chip(s, 1.05, 2.25, 0.75, 'R', 'LuPenLine', 'Test');
  s.addText([{ text: 'CELPIP is the test', options: { bold: true, fontSize: 22, breakLine: true } }, { text: 'The exam you sit on the computer.', options: { fontSize: 15 } }], { x: 2.0, y: 2.2, w: 4.1, h: 2.1, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  card(s, 6.6, 1.95, 5.9, 2.6, SK.L[0], 'CLB');
  await chip(s, 6.9, 2.25, 0.75, 'L', 'LuScale', 'Ruler');
  s.addText([{ text: 'CLB is the ruler', options: { bold: true, fontSize: 22, breakLine: true } }, { text: 'Canada uses the Canadian Language Benchmarks (1–12) to measure English. Immigration asks for CLB levels.', options: { fontSize: 15 } }], { x: 7.85, y: 2.2, w: 4.45, h: 2.1, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  s.addText('CLB RULER', { x: 0.75, y: 4.8, w: 3, h: 0.3, fontSize: 11, bold: true, color: MUTED, charSpacing: 3, margin: 0, isTextBox: true });
  for (let i = 1; i <= 12; i++) {
    const x = 0.75 + (i - 1) * 0.98, hl = i === 7;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 5.2, w: 0.88, h: 0.75, rectRadius: 0.12, fill: { color: hl ? SK.L[1] : (i < 7 ? 'F6F7FB' : SK.L[0]) }, line: { color: hl ? SK.L[1] : (i < 7 ? 'F6F7FB' : SK.L[0]) }, objectName: 'CLB ' + i });
    s.addText(String(i), { x, y: 5.2, w: 0.88, h: 0.75, fontSize: 18, bold: true, color: hl ? 'FFFFFF' : NAVY, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
  }
  s.addText('Your CELPIP result converts straight onto the ruler: CELPIP 7 = CLB 7.', { x: 0.75, y: 6.1, w: 11.75, h: 0.45, fontSize: 15, color: NAVY, margin: 0, isTextBox: true });
  s.addNotes('Analogy: CELPIP is like a bathroom scale; CLB is the kilogram. Different tests (scales), one unit Canada understands. Pause after "CLB" and write it on the board.\nCHECK: "Which one do you take: CELPIP or CLB?" "If you score CELPIP 8, what is your CLB?"');

  // 9 · Scoring
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Scores and targets' });
  eyebrow(s, 'How scoring works', 'R'); title(s, 'One level per skill. No penalty for guessing.');
  [['L', '4', 'levels', 'One CELPIP level each for Listening, Reading, Writing and Speaking.'], ['R', '7=7', 'CELPIP = CLB', 'A CELPIP level matches the same CLB level. Programs ask for CLB per skill.'], ['O', '0', 'penalty', 'A wrong answer costs nothing. Never leave a question blank.']].forEach(([k, big, lab, txt], i) => {
    const x = 0.75 + i * 4.0;
    card(s, x, 1.95, 3.7, 4.5, SK[k][0], 'Fact ' + (i + 1));
    s.addText(big, { x: x + 0.35, y: 2.25, w: 3.0, h: 1.5, fontSize: 72, bold: true, color: SK[k][1], margin: 0, valign: 'middle', isTextBox: true });
    s.addText(lab.toUpperCase(), { x: x + 0.35, y: 3.85, w: 3.0, h: 0.35, fontSize: 12, bold: true, color: NAVY, charSpacing: 3, margin: 0, isTextBox: true });
    s.addText(txt, { x: x + 0.35, y: 4.35, w: 3.0, h: 1.8, fontSize: 15, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  });
  s.addNotes('Listening and Reading are scored by the computer; Writing and Speaking by trained raters. Do not promise raw-score conversions; the official site publishes the score comparison.\nCHECK: "Do you get one score or four?" "What happens if you guess and get it wrong?"');

  // 15b · Immigration words
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Scores and targets' });
  eyebrow(s, 'Immigration words in plain English', 'O'); title(s, '3 words you will see on the next slide');
  [['L', 'LuMonitor', 'Express Entry', 'Canada’s online application system for permanent residence.', 'You create a profile and receive points. Higher English levels = more points.'],
   ['R', 'LuGraduationCap', 'TEER 0 / 1', 'Management and professional jobs.', 'e.g. restaurant manager, registered nurse, software engineer, accountant'],
   ['W', 'LuCompass', 'TEER 2 / 3', 'Skilled technical and trade jobs.', 'e.g. electrician, cook, dental assistant, heavy-equipment operator']].forEach(([k, ic, h, d, ex], i) => {
    const x = 0.75 + i * 4.0;
    card(s, x, 1.95, 3.7, 4.0, SK[k][0], h);
  });
  for (const [i, [k, ic, h, d, ex]] of [['L', 'LuMonitor', 'Express Entry', 'Canada’s online application system for permanent residence.', 'You create a profile and receive points. Higher English levels = more points.'],
   ['R', 'LuGraduationCap', 'TEER 0 / 1', 'Management and professional jobs.', 'e.g. restaurant manager, registered nurse, software engineer, accountant'],
   ['W', 'LuCompass', 'TEER 2 / 3', 'Skilled technical and trade jobs.', 'e.g. electrician, cook, dental assistant, heavy-equipment operator']].entries()) {
    const x = 0.75 + i * 4.0;
    await chip(s, x + 0.35, 2.25, 0.75, k, ic, h);
    s.addText(h, { x: x + 0.35, y: 3.2, w: 3.0, h: 0.5, fontSize: 21, bold: true, color: NAVY, margin: 0, isTextBox: true });
    s.addText([{ text: d, options: { bold: true, breakLine: true } }, { text: ex, options: { color: MUTED } }], { x: x + 0.35, y: 3.75, w: 3.0, h: 2.05, fontSize: 14, color: NAVY, valign: 'top', margin: 0, paraSpaceAfter: 8, isTextBox: true });
  }
  await tips(s, 0.75, 6.1, 11.75, 0.75, ['IRCC = Immigration, Refugees and Citizenship Canada, the department that checks your application.'], 'IRCC', 'One more');
  s.addNotes('Use workplace examples from your students’ own jobs: ask two students their job and decide together if it is TEER 0/1 or 2/3. Pause after each acronym.\nCHECK: "Is a cook TEER 1 or TEER 3?" "What is Express Entry in one sentence?"');

  // 10 · Requirements
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Scores and targets' });
  eyebrow(s, 'Minimum CLB by program', 'L'); title(s, 'What level do you need?');
  const head = ['Listening', 'Reading', 'Writing', 'Speaking'];
  head.forEach((h, j) => s.addText(h.toUpperCase(), { x: 6.6 + j * 1.5, y: 1.95, w: 1.4, h: 0.35, fontSize: 10, bold: true, color: MUTED, charSpacing: 2, align: 'center', margin: 0, isTextBox: true }));
  [['Express Entry · Federal Skilled Worker', [7, 7, 7, 7]], ['Canadian Experience Class · TEER 0 or 1', [7, 7, 7, 7]], ['Canadian Experience Class · TEER 2 or 3', [5, 5, 5, 5]], ['Federal Skilled Trades', [5, 4, 4, 5]], ['Citizenship (CELPIP-General LS)', [4, '–', '–', 4]]].forEach(([p, v], i) => {
    const y = 2.4 + i * 0.66;
    card(s, 0.75, y, 11.75, 0.56, i % 2 ? 'FFFFFF' : 'F6F7FB', 'Program ' + (i + 1));
    s.addText(p, { x: 1.05, y, w: 5.4, h: 0.56, fontSize: 15, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
    v.forEach((n, j) => s.addText(String(n), { x: 6.6 + j * 1.5, y, w: 1.4, h: 0.56, fontSize: 17, bold: true, color: n === '–' ? MUTED : SK.L[1], align: 'center', valign: 'middle', margin: 0, isTextBox: true }));
  });
  await tips(s, 0.75, 5.8, 11.75, 1.0, ['If immigration asks for CLB 7, aim for 8, so one small mistake on test day will not ruin your application. CLB 9+ also adds Express Entry points.'], 'Insurance tip', 'Your insurance policy');
  s.addNotes('Present the +1 rule as an insurance policy. Enter a sample profile (e.g. 8/7/6/7) in the website’s Requirements Checker. Remind students to confirm current rules on the IRCC website.\nCHECK: "Your program needs CLB 7. What level should you aim for, and why?"');

  // 11 · Test screen
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Scores and targets' });
  eyebrow(s, 'The test screen', 'S'); title(s, 'Know the rules before test day');
  const scr = [['L', 'LuHeadphones', 'Listening', 'Audio plays once. No pause, no replay. Questions appear after the audio.'], ['R', 'LuBookOpen', 'Reading', 'Like a moving train: once you click Next, you leave the station and cannot jump back.'], ['W', 'LuTimer', 'Writing', 'A timer runs for each task. Plan, write, then proofread.'], ['S', 'LuMic', 'Speaking', 'The microphone starts by itself after prep time. Silence is recorded too, so keep talking.']];
  for (let i = 0; i < 4; i++) {
    const [k, ic, n, t] = scr[i], x = 0.75 + (i % 2) * 5.95, y = 1.95 + Math.floor(i / 2) * 2.4;
    card(s, x, y, 5.8, 2.15, SK[k][0], n + ' rule');
    await chip(s, x + 0.35, y + 0.35, 0.75, k, ic, n);
    s.addText(n, { x: x + 1.35, y: y + 0.35, w: 4.2, h: 0.5, fontSize: 18, bold: true, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
    s.addText(t, { x: x + 1.35, y: y + 0.9, w: 4.2, h: 1.05, fontSize: 15, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  }
  s.addNotes('Demo the Test Screen Simulator on the website. Use the moving-train image for Reading. Reassure students about the microphone: it is automatic; their only job is to keep talking.\nCHECK: "What happens if you click Next in Reading?" "When does the microphone start?"');

  // 18b · Check your understanding
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Scores and targets' });
  eyebrow(s, 'Quick check', 'G'); title(s, 'True or false?');
  ['You get one overall CELPIP score.', 'A wrong answer loses you points.', 'CELPIP 8 = CLB 8.', 'You can replay the listening audio once.', 'The speaking microphone starts by itself.'].forEach((t, i) => {
    const y = 1.95 + i * 0.92;
    card(s, 0.75, y, 11.75, 0.78, i % 2 ? 'FFFFFF' : 'F6F7FB', 'Statement ' + (i + 1));
    badge(s, 0.97, y + 0.15, 0.48, 'R', i + 1, 'Statement number ' + (i + 1));
    s.addText(t, { x: 1.7, y, w: 7.8, h: 0.78, fontSize: 18, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 10.0, y: y + 0.17, w: 2.25, h: 0.44, rectRadius: 0.22, fill: { color: 'FFFFFF' }, line: { color: LINE, width: 1 }, objectName: 'Answer box ' + (i + 1) });
    s.addText('True  /  False', { x: 10.0, y: y + 0.17, w: 2.25, h: 0.44, fontSize: 12, bold: true, color: MUTED, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
  });
  s.addNotes('Students answer with fingers (1 = true, 2 = false) or on mini whiteboards.\nANSWERS: 1 False (one level per skill) · 2 False (no penalty) · 3 True · 4 False (plays once) · 5 True.');

  // 12 · Mistakes
  pres.addSection({ title: 'Mistakes' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Mistakes' });
  eyebrow(s, 'Avoid these', 'W'); title(s, '10 mistakes that cost points');
  ['Leaving a question blank', 'Missing a bullet in the email', 'Writing under 150 words', 'Wrong tone: too casual or too stiff', 'Picking the first answer you hear', 'Spending 3 minutes on one question', 'Stopping your answer early', 'Going off topic to "sound smart"', 'Memorized answers that miss the task', 'Skipping the proofreading pass'].forEach((t, i) => {
    const col = i < 5 ? 0 : 1, x = 0.75 + col * 5.95, y = 1.95 + (i % 5) * 0.96;
    card(s, x, y, 5.8, 0.8, i % 2 ? 'FFFFFF' : 'F6F7FB', 'Mistake ' + (i + 1));
    badge(s, x + 0.25, y + 0.17, 0.46, 'W', i + 1, 'Mistake number ' + (i + 1), false);
    s.addText(t, { x: x + 0.95, y, w: 4.7, h: 0.8, fontSize: 16, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  });
  s.addNotes('Normalize mistakes: everyone makes them. Ask which ones students have made before.\nCHECK: "Which mistake is easiest to fix today?"');

  // 13 · Error causes
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Mistakes' });
  eyebrow(s, 'Your detective notebook', 'R'); title(s, 'Find out why you lost the point');
  const causes = [['W', 'LuTarget', 'Distractor', 'You chose a detail that was mentioned, then corrected.'], ['L', 'LuBookOpen', 'Vocabulary', 'You missed a paraphrase or a word.'], ['S', 'LuClock', 'Timing', 'You ran out of time or rushed.'], ['R', 'LuBrain', 'Inference', 'You missed what was implied, not said.']];
  for (let i = 0; i < 4; i++) {
    const [k, ic, n, t] = causes[i], x = 0.75 + i * 3.0;
    card(s, x, 1.95, 2.75, 3.6, SK[k][0], n);
    await chip(s, x + 0.3, 2.25, 0.75, k, ic, n);
    s.addText(n, { x: x + 0.3, y: 3.2, w: 2.2, h: 0.5, fontSize: 19, bold: true, color: NAVY, margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.3, y: 3.75, w: 2.2, h: 1.6, fontSize: 14, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  }
  await tips(s, 0.75, 5.8, 11.75, 0.95, ['Your Error Log is a detective notebook, not a punishment. Tag each mistake with its cause; the clue tells you what to practise next.'], 'Error log tip', 'Be a detective');
  s.addNotes('Frame mistakes as normal and useful. Give one example of each cause from the diagnostic, then model tagging two errors live in the Error Log.\nCHECK: "You chose Tuesday, but the answer was Wednesday. Which cause is that?" (Distractor)');

  // 14 · Targets
  pres.addSection({ title: 'Wrap-up' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Wrap-up' });
  eyebrow(s, 'Do it now', 'O'); title(s, 'Set your targets in 3 minutes');
  [['O', 'Find the level your program needs'], ['L', 'Add +1 as a safety margin'], ['R', 'Set each skill on your dashboard'], ['W', 'Write down your weakest skill']].forEach(([k, t], i) => {
    const x = 0.75 + i * 3.0;
    card(s, x, 1.95, 2.75, 2.9, SK[k][0], 'Step ' + (i + 1));
    s.addText(String(i + 1).padStart(2, '0'), { x: x + 0.3, y: 2.2, w: 2, h: 0.9, fontSize: 40, bold: true, color: SK[k][1], margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.3, y: 3.15, w: 2.2, h: 1.5, fontSize: 17, bold: true, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  });
  await tips(s, 0.75, 5.15, 11.75, 1.15, ['Open the course website on your phone: My Dashboard → Your four skills → Target CLB.'], 'Action', 'On your phone');
  s.addNotes('Give 3 minutes for students to set targets on their phones. Walk around and check.\nCHECK: "What is your target for Speaking, and why?"');

  // 22b · Ask in English
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Wrap-up' });
  eyebrow(s, 'Classroom English', 'L'); title(s, 'Ask in English. Every question helps.');
  ['Could you repeat that, please?', 'Could you speak more slowly, please?', 'What does “___” mean?', 'How do you spell “___”?', 'Can you give an example?', 'Is this correct: “___”?'].forEach((t, i) => {
    const col = i % 2, row = Math.floor(i / 2), x = 0.75 + col * 5.95, y = 1.95 + row * 1.25, k = ['L', 'R', 'W', 'S', 'O', 'L'][i];
    card(s, x, y, 5.8, 1.05, SK[k][0], 'Phrase ' + (i + 1));
    s.addText('“', { x: x + 0.25, y, w: 0.5, h: 1.05, fontSize: 36, bold: true, color: SK[k][1], valign: 'middle', margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.8, y, w: 4.8, h: 1.05, fontSize: 18, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  });
  await tips(s, 0.75, 5.85, 11.75, 0.9, ['Asking questions in English is speaking practice too.'], 'Ask', 'Remember');
  s.addNotes('Practise each phrase chorally. Tell students you expect these questions in every class.\nCHECK: "How do you ask me to say something again?"');

  // 15 · Before Session 2
  s = pres.addSlide({ masterName: 'CLOSING', sectionTitle: 'Wrap-up' });
  eyebrow(s, 'Study Hub · Session 1', 'O'); title(s, 'Before Session 2');
  const todo = ['Listening and reading checks', 'Baseline email · 27 minutes', 'Two baseline speaking answers', 'Save your starting point', 'Tag every error in your Error Log'];
  for (let i = 0; i < todo.length; i++) {
    const y = 1.95 + i * 0.86;
    card(s, 0.75, y, 7.4, 0.7, 'FFFFFF', 'To-do ' + (i + 1), true);
    badge(s, 0.97, y + 0.14, 0.42, 'O', '✓', 'Check ' + (i + 1));
    s.addText(todo[i], { x: 1.6, y, w: 6.4, h: 0.7, fontSize: 16, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  }
  card(s, 8.6, 1.95, 3.9, 4.15, 'FFFFFF', 'Next class', true);
  await chip(s, 8.95, 2.3, 0.85, 'O', 'LuFlag', 'Next class', true);
  s.addText([{ text: 'Your results open the next class.', options: { bold: true, fontSize: 18, breakLine: true } }, { text: 'Finish the Study Hub before Session 2 so we can work on your real mistakes.', options: { fontSize: 14, color: MUTED } }],
    { x: 8.95, y: 3.4, w: 3.25, h: 2.4, color: NAVY, valign: 'top', margin: 0, paraSpaceAfter: 10, isTextBox: true });
  s.addNotes('Open the Study Hub tab for Session 1 and show the diagnostic. Return to slide 2 and tick the outcomes together.\nCHECK: "What do you need to finish before Session 2?"');

  const out = process.argv[2] || 'session-01.pptx';
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log('written', out);
})().catch(e => { console.error(e); process.exit(1); });
