// Session 1 deck · CELPIP Prep Course · Elevate with Ms. Maria · Learnovate Qatar
const pptxgen = require('pptxgenjs');
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const sharp = require('sharp');
const Lu = require('react-icons/lu');
const { applyTheme } = require(process.env.APPLY_THEME || './apply_theme.js');

const THEME = { name: 'Elevate Pastel', headFontFace: 'Poppins', bodyFontFace: 'Poppins',
  colors: { dk1: '14213D', lt1: 'FFFFFF', dk2: '4A5675', lt2: 'F4F6FA', accent1: '1A5FCC', accent2: '6A2BC4', accent3: 'C93D1A',
    accent4: '007F7F', accent5: '008F55', accent6: 'FFC300', hlink: '1A5FCC', folHlink: '6A2BC4' } };
// skill palette: [pastel fill, ink]
const SK = { L: ['E9F2FF', '1A5FCC'], R: ['F2EBFD', '6A2BC4'], W: ['FFEEEA', 'C93D1A'], S: ['E6F8F8', '007F7F'], O: ['E6F8F0', '008F55'], G: ['FFF6D6', '8A6A00'] };
const NAVY = '14213D', MUTED = '4A5675';

async function icon(name, color, size = 256) {
  const svg = renderToStaticMarkup(React.createElement(Lu[name], { color: '#' + color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return 'image/png;base64,' + png.toString('base64');
}

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
  pres.author = 'Ms. Maria · Learnovate Qatar';
  pres.title = 'Session 1 · Decode the Test Engine and Set Your CLB Target';
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  pres.defineSlideMaster({ title: 'TITLE', background: { color: 'F2EBFD' },
    objects: [
      { placeholder: { options: { name: 'title', type: 'title', x: 0.8, y: 2.2, w: 8.4, h: 2.0, fontSize: 44, bold: true, color: C.text1, valign: 'top' }, text: '' } },
      { placeholder: { options: { name: 'body', type: 'body', x: 0.8, y: 4.35, w: 8.4, h: 0.8, fontSize: 18, color: C.text2 }, text: '' } },
      { text: { text: 'ELEVATE WITH MS. MARIA  ·  LEARNOVATE QATAR', options: { x: 0.8, y: 1.45, w: 8.4, h: 0.4, fontSize: 12, bold: true, color: C.text1, charSpacing: 4, margin: 0 } } }
    ] });
  pres.defineSlideMaster({ title: 'CONTENT', background: { color: 'FFFFFF' }, margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.4, w: 12.1, h: 0.85, fontSize: 36, bold: true, color: C.text1, valign: 'middle', margin: 0 }, text: '' } },
      { text: { text: 'CELPIP Prep Course · Session 1', options: { x: 0.6, y: 6.95, w: 6, h: 0.3, fontSize: 10, color: C.text2, margin: 0 } } }
    ],
    slideNumber: { x: 12.2, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: '4A5675', align: 'right' } });
  pres.defineSlideMaster({ title: 'CLOSING', background: { color: 'E9F2FF' },
    objects: [
      { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.4, w: 12.1, h: 0.85, fontSize: 36, bold: true, color: C.text1, valign: 'middle', margin: 0 }, text: '' } },
      { text: { text: 'CELPIP Prep Course · Session 1', options: { x: 0.6, y: 6.95, w: 6, h: 0.3, fontSize: 10, color: C.text2, margin: 0 } } }
    ],
    slideNumber: { x: 12.2, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: '4A5675', align: 'right' } });

  // helpers
  const card = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.18, fill: { color: fill }, line: { color: fill }, objectName: name });
  const circleIcon = async (s, x, y, d, key, ic, name) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: 'FFFFFF' }, line: { color: SK[key][1], width: 1.5 }, objectName: name + ' circle' });
    s.addImage({ data: await icon(ic, SK[key][1]), x: x + d * 0.22, y: y + d * 0.22, w: d * 0.56, h: d * 0.56, altText: name });
  };
  const numDot = (s, x, y, d, key, label, name) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: SK[key][1] }, line: { color: SK[key][1] }, objectName: name });
    s.addText(String(label), { x, y, w: d, h: d, align: 'center', valign: 'middle', fontSize: 14, bold: true, color: 'FFFFFF', margin: 0, isTextBox: true });
  };
  const tipBox = (s, x, y, w, h, tips, name) => {
    card(s, x, y, w, h, SK.G[0], name);
    s.addText(tips.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < tips.length - 1, paraSpaceAfter: 6 } })),
      { x: x + 0.25, y: y + 0.15, w: w - 0.5, h: h - 0.3, fontSize: 14, color: NAVY, valign: 'top', isTextBox: true });
  };
  const title = (s, t) => s.addText(t, { placeholder: 'title' });

  // 1 · Title
  pres.addSection({ title: 'Welcome' });
  let s = pres.addSlide({ masterName: 'TITLE', sectionTitle: 'Welcome' });
  s.addText('Decode the Test Engine and Set Your CLB Target', { placeholder: 'title' });
  s.addText('Session 1 · CELPIP Prep Course', { placeholder: 'body' });
  const keys = ['L', 'R', 'W', 'S']; const ics = ['LuHeadphones', 'LuBookOpen', 'LuPenLine', 'LuMic'];
  for (let i = 0; i < 4; i++) {
    const x = 9.7 + (i % 2) * 1.55, y = 2.2 + Math.floor(i / 2) * 1.55;
    card(s, x, y, 1.35, 1.35, SK[keys[i]][0], 'Skill tile ' + keys[i]);
    s.addImage({ data: await icon(ics[i], SK[keys[i]][1]), x: x + 0.37, y: y + 0.37, w: 0.6, h: 0.6, altText: ['Listening', 'Reading', 'Writing', 'Speaking'][i] });
  }
  s.addNotes('Welcome everyone. Say: "Today you learn exactly how the test works, how it is scored, and what level you need. Every slide is something you will use on test day."');

  // 2 · Outcomes
  pres.addSection({ title: 'The test' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Welcome' });
  title(s, 'By the end of today you can…');
  const outs = ['Name every part of the test and its timing', 'Explain how your level is scored', 'Find the CLB level your program needs', 'Avoid the 10 mistakes that cost the most points', 'Set a target for each skill'];
  outs.forEach((t, i) => {
    const y = 1.55 + i * 0.98, k = ['L', 'R', 'O', 'W', 'S'][i];
    card(s, 0.6, y, 7.6, 0.8, SK[k][0], 'Outcome ' + (i + 1));
    numDot(s, 0.8, y + 0.15, 0.5, k, i + 1, 'Outcome number ' + (i + 1));
    s.addText(t, { x: 1.5, y, w: 6.6, h: 0.8, fontSize: 18, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  });
  card(s, 8.7, 1.55, 4.0, 4.72, SK.G[0], 'Agenda card');
  s.addText('Today · 2 hours', { x: 8.95, y: 1.7, w: 3.5, h: 0.5, fontSize: 18, bold: true, color: NAVY, margin: 0, isTextBox: true });
  const ag = [['0:00', 'Map the test'], ['0:20', 'How scoring works'], ['0:45', 'The test screen'], ['1:15', 'Mistakes and error causes'], ['1:45', 'Set your targets']];
  s.addText(ag.map(([a, b], i) => ([{ text: a + '  ', options: { bold: true, color: '8A6A00' } }, { text: b, options: { breakLine: i < ag.length - 1 } }])).flat(),
    { x: 8.95, y: 2.3, w: 3.5, h: 3.7, fontSize: 16, color: NAVY, valign: 'top', paraSpaceAfter: 14, isTextBox: true });
  s.addNotes('Read the outcomes aloud. Return to this slide at the end and ask students to tick each one.');

  // 3 · Pick the right test
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The test' });
  title(s, 'Pick the right test');
  const tests = [['L', 'LuAward', 'CELPIP-General', 'All four skills. For permanent residence: Express Entry and other economic programs.'],
    ['S', 'LuMic', 'CELPIP-General LS', 'Listening and Speaking only. For Canadian citizenship.'],
    ['G', 'LuSearch', 'Before you book', 'Check which test and which levels your program asks for.']];
  for (let i = 0; i < 3; i++) {
    const [k, ic, h, b] = tests[i], x = 0.6 + i * 4.1;
    card(s, x, 1.6, 3.85, 4.4, SK[k][0], 'Test card ' + (i + 1));
    await circleIcon(s, x + 0.3, 1.9, 0.9, k === 'G' ? 'O' : k, ic, h);
    s.addText(h, { x: x + 0.3, y: 3.0, w: 3.25, h: 0.6, fontSize: 22, bold: true, color: NAVY, margin: 0, isTextBox: true });
    s.addText(b, { x: x + 0.3, y: 3.65, w: 3.25, h: 2.1, fontSize: 16, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  }
  s.addNotes('Booking the wrong version costs money and time. Ask who is applying for PR and who for citizenship.');

  // 4 · Test at a glance
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The test' });
  title(s, 'The test at a glance');
  const glance = [['L', 'LuHeadphones', 'Listening', '47–55 min', '6 parts · 38 questions'], ['R', 'LuBookOpen', 'Reading', '55–60 min', '4 parts · 38 questions'],
    ['W', 'LuPenLine', 'Writing', '53–60 min', '2 tasks'], ['S', 'LuMic', 'Speaking', '15–20 min', '8 tasks']];
  for (let i = 0; i < 4; i++) {
    const [k, ic, n, t, c] = glance[i], x = 0.6 + i * 3.08;
    card(s, x, 1.55, 2.85, 3.0, SK[k][0], n + ' card');
    await circleIcon(s, x + 0.25, 1.8, 0.8, k, ic, n);
    s.addText(n, { x: x + 0.25, y: 2.75, w: 2.4, h: 0.5, fontSize: 22, bold: true, color: SK[k][1], margin: 0, isTextBox: true });
    s.addText([{ text: t, options: { bold: true, breakLine: true } }, { text: c }], { x: x + 0.25, y: 3.3, w: 2.4, h: 1.0, fontSize: 16, color: NAVY, margin: 0, valign: 'top', isTextBox: true });
    if (i < 3) s.addText('→', { x: x + 2.82, y: 2.7, w: 0.3, h: 0.5, fontSize: 20, bold: true, color: MUTED, align: 'center', margin: 0, isTextBox: true });
  }
  tipBox(s, 0.6, 4.85, 12.1, 1.35, ['Order: Listening → Reading → Writing → Speaking, about 3 hours in one sitting.', 'Train your stamina: do full timed practice before test day. Speaking comes last, when you are tired.'], 'Tips');
  s.addNotes('Walk through each skill. Stress stamina: Speaking comes last. Confirm current details on the official CELPIP website before class.');

  // 5 · Listening
  pres.addSection({ title: 'The four skills' });
  const partSlide = async (k, ic, ttl, rows, head, tips, notes) => {
    const sl = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The four skills' });
    title(sl, ttl);
    const tb = [head.map(h => ({ text: h, options: { bold: true, color: SK[k][1], fill: { color: SK[k][0] } } }))].concat(rows.map(r => r.map(c => ({ text: c }))));
    sl.addTable(tb, { x: 0.6, y: 1.55, w: 7.4, colW: head.length === 3 ? [4.0, 1.6, 1.8] : [5.6, 1.8], fontSize: 15, color: NAVY, border: { type: 'solid', pt: 0.75, color: 'DFE4EE' }, rowH: 0.48, margin: [4, 8, 4, 8] });
    await circleIcon(sl, 8.6, 1.55, 0.8, k, ic, ttl);
    sl.addText('How to score higher', { x: 9.55, y: 1.6, w: 3.2, h: 0.7, fontSize: 18, bold: true, color: NAVY, margin: 0, valign: 'middle', isTextBox: true });
    tipBox(sl, 8.6, 2.55, 4.1, 3.9, tips, 'Strategies');
    sl.addNotes(notes);
  };
  await partSlide('L', 'LuHeadphones', 'Listening · 6 parts, 38 questions',
    [['Part 1 · Problem Solving', '8'], ['Part 2 · Daily Life Conversation', '5'], ['Part 3 · Listening for Information', '6'], ['Part 4 · Listening to a News Item', '5'], ['Part 5 · Listening to a Discussion (video)', '8'], ['Part 6 · Listening to Viewpoints', '6']],
    ['Part', 'Questions'],
    ['Audio plays once: read the context screen and predict before it starts.', 'The first answer you hear is often corrected later. Wait for the final decision.', 'Take short notes: who, what, when, change of plan.'],
    'Model one prediction from a context screen. Then give a quick example of a corrected detail (Thursday → tomorrow).');
  await partSlide('R', 'LuBookOpen', 'Reading · 4 parts, 38 questions',
    [['Part 1 · Reading Correspondence', '11'], ['Part 2 · Reading to Apply a Diagram', '8'], ['Part 3 · Reading for Information', '9'], ['Part 4 · Reading for Viewpoints', '10']],
    ['Part', 'Questions'],
    ['Read the questions first, then scan for the answer.', 'Answers are paraphrased: "postponed" = "moved to a later date".', 'About one minute per question. Never get stuck.'],
    'Show one paraphrase pair. Ask who reads every word first; explain why question-first scanning saves time.');

  // 7 · Writing
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'The four skills' });
  title(s, 'Writing · 2 tasks, what raters score');
  const wt = [['Task 1 · Email', '27 minutes', '150–200 words · cover all 3 bullets'], ['Task 2 · Survey', '26 minutes', '150–200 words · choose one option and defend it']];
  for (let i = 0; i < 2; i++) {
    const x = 0.6 + i * 4.05;
    card(s, x, 1.55, 3.8, 2.4, SK.W[0], wt[i][0]);
    s.addText(wt[i][0], { x: x + 0.3, y: 1.75, w: 3.2, h: 0.5, fontSize: 20, bold: true, color: SK.W[1], margin: 0, isTextBox: true });
    s.addText([{ text: wt[i][1], options: { bold: true, breakLine: true } }, { text: wt[i][2] }], { x: x + 0.3, y: 2.35, w: 3.2, h: 1.4, fontSize: 16, color: NAVY, margin: 0, valign: 'top', isTextBox: true });
  }
  card(s, 8.75, 1.55, 3.95, 2.4, SK.R[0], 'Raters score');
  s.addText('Raters score', { x: 9.05, y: 1.75, w: 3.4, h: 0.5, fontSize: 20, bold: true, color: SK.R[1], margin: 0, isTextBox: true });
  s.addText(['Content & coherence', 'Vocabulary', 'Readability', 'Task fulfillment'].map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } })),
    { x: 9.05, y: 2.3, w: 3.4, h: 1.55, fontSize: 16, color: NAVY, valign: 'top', isTextBox: true });
  tipBox(s, 0.6, 4.3, 12.1, 1.9, ['Plan 3 minutes, write 20, proofread 3.', 'Give each bullet its own paragraph so the rater can find it.', 'Match the tone: formal for a manager, friendly for a friend.'], 'Writing strategies');
  s.addNotes('Ask: "What happens if you skip a bullet?" Task fulfillment drops. Show a one-line plan per bullet.');

  // 8 · Speaking
  await partSlide('S', 'LuMic', 'Speaking · 8 tasks, what raters score',
    [['1 · Giving Advice', '30 s', '90 s'], ['2 · Personal Experience', '30 s', '60 s'], ['3 · Describing a Scene', '30 s', '60 s'], ['4 · Making Predictions', '30 s', '60 s'], ['5 · Comparing and Persuading', '60 + 60 s', '60 s'], ['6 · Dealing with a Difficult Situation', '60 s', '60 s'], ['7 · Expressing Opinions', '30 s', '90 s'], ['8 · Describing an Unusual Situation', '30 s', '60 s']],
    ['Task', 'Prep', 'Speak'],
    ['Raters score content & coherence, vocabulary, listenability and task fulfillment.', 'Use prep time to plan 2–3 points, not full sentences.', 'Keep talking until time ends: add an example or a summary.'],
    'Ask a volunteer to speak for 60 seconds about a favourite place, so everyone feels how long 60 seconds is. Task 5 has 60 seconds to choose, then 60 seconds to prepare.');

  // 9 · Scoring
  pres.addSection({ title: 'Scores and targets' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Scores and targets' });
  title(s, 'How scoring works');
  const sc = [['L', '4', 'levels', 'You receive a CELPIP level for each skill: Listening, Reading, Writing and Speaking.'], ['R', '7 = 7', 'CELPIP = CLB', 'A CELPIP level lines up with the same CLB level. Programs ask for CLB per skill.'], ['O', '0', 'penalty', 'Wrong answers cost nothing. Never leave a question blank.']];
  for (let i = 0; i < 3; i++) {
    const [k, big, lab, txt] = sc[i], x = 0.6 + i * 4.1;
    card(s, x, 1.6, 3.85, 4.5, SK[k][0], 'Scoring fact ' + (i + 1));
    s.addText(big, { x: x + 0.3, y: 1.85, w: 3.25, h: 1.2, fontSize: 60, bold: true, color: SK[k][1], margin: 0, isTextBox: true });
    s.addText(lab, { x: x + 0.3, y: 3.05, w: 3.25, h: 0.5, fontSize: 16, bold: true, color: NAVY, margin: 0, isTextBox: true });
    s.addText(txt, { x: x + 0.3, y: 3.65, w: 3.25, h: 2.2, fontSize: 16, color: NAVY, valign: 'top', margin: 0, isTextBox: true });
  }
  s.addNotes('Listening and Reading are scored by the computer; Writing and Speaking by trained raters. Do not promise raw-score conversions; the official site publishes the score comparison.');

  // 10 · What level do you need
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Scores and targets' });
  title(s, 'What level do you need?');
  const req = [['Program', 'Listening', 'Reading', 'Writing', 'Speaking'], ['Express Entry · Federal Skilled Worker', '7', '7', '7', '7'], ['Canadian Experience Class · TEER 0 or 1', '7', '7', '7', '7'], ['Canadian Experience Class · TEER 2 or 3', '5', '5', '5', '5'], ['Federal Skilled Trades', '5', '4', '4', '5'], ['Canadian citizenship (CELPIP-General LS)', '4', '–', '–', '4']];
  s.addTable(req.map((r, i) => r.map((c, j) => ({ text: c, options: i === 0 ? { bold: true, color: SK.L[1], fill: { color: SK.L[0] }, align: j ? 'center' : 'left' } : { align: j ? 'center' : 'left', bold: j > 0 } }))),
    { x: 0.6, y: 1.55, w: 12.1, colW: [5.3, 1.7, 1.7, 1.7, 1.7], fontSize: 16, color: NAVY, border: { type: 'solid', pt: 0.75, color: 'DFE4EE' }, rowH: 0.55, margin: [4, 10, 4, 10] });
  tipBox(s, 0.6, 5.15, 12.1, 1.3, ['Aim one level above the minimum: it protects you and higher levels (especially CLB 9+) add Express Entry points.', 'Rules change. Always confirm current requirements on the official IRCC website.'], 'Target tips');
  s.addNotes('Enter a sample profile (e.g. 8/7/6/7) in the website’s Requirements Checker and show which programs it meets.');

  // 11 · Test screen
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Scores and targets' });
  title(s, 'What the test screen does');
  const scr = [['L', 'LuHeadphones', 'Listening', 'Audio plays once. No pause, no replay. Questions appear after the audio.'], ['R', 'LuBookOpen', 'Reading', 'Change answers until you press Next. After Next, you cannot go back.'], ['W', 'LuTimer', 'Writing', 'A timer runs for each task. Plan, write and leave time to proofread.'], ['S', 'LuMic', 'Speaking', 'Recording starts by itself when preparation ends. Silence is recorded too.']];
  for (let i = 0; i < 4; i++) {
    const [k, ic, n, t] = scr[i], x = 0.6 + (i % 2) * 6.1, y = 1.55 + Math.floor(i / 2) * 2.45;
    card(s, x, y, 5.95, 2.2, SK[k][0], n + ' screen');
    await circleIcon(s, x + 0.3, y + 0.3, 0.8, k, ic, n);
    s.addText(n, { x: x + 1.3, y: y + 0.3, w: 4.4, h: 0.6, fontSize: 20, bold: true, color: SK[k][1], margin: 0, valign: 'middle', isTextBox: true });
    s.addText(t, { x: x + 1.3, y: y + 0.95, w: 4.4, h: 1.1, fontSize: 16, color: NAVY, margin: 0, valign: 'top', isTextBox: true });
  }
  s.addNotes('Demo the Test Screen Simulator on the website. Key messages: no replay, Next is final, speaking records automatically.');

  // 12 · 10 mistakes
  pres.addSection({ title: 'Mistakes' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Mistakes' });
  title(s, '10 mistakes that cost points');
  const mis = ['Leaving a question blank', 'Missing a bullet in the email', 'Writing under 150 words', 'Wrong tone: too casual or too stiff', 'Picking an answer you only heard first', 'Spending 3 minutes on one question', 'Stopping your speaking answer early', 'Going off topic to "sound smart"', 'Memorized answers that do not fit the task', 'Skipping the proofreading pass'];
  mis.forEach((t, i) => {
    const col = i < 5 ? 0 : 1, row = i % 5, x = 0.6 + col * 6.1, y = 1.5 + row * 1.0, k = ['W', 'R', 'L', 'S', 'O'][row];
    card(s, x, y, 5.95, 0.82, SK[k][0], 'Mistake ' + (i + 1));
    numDot(s, x + 0.2, y + 0.16, 0.5, k, i + 1, 'Mistake number ' + (i + 1));
    s.addText(t, { x: x + 0.9, y, w: 4.9, h: 0.82, fontSize: 17, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  });
  s.addNotes('Ask students which of these they have made before. These become the checklist for every Study Hub task.');

  // 13 · Four causes
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Mistakes' });
  title(s, 'The four causes of errors');
  const causes = [['W', 'LuTarget', 'Distractor', 'You chose a detail that was mentioned but later corrected.'], ['L', 'LuBookOpen', 'Vocabulary', 'You missed a paraphrase or did not know a word.'], ['S', 'LuClock', 'Timing', 'You ran out of time or rushed.'], ['R', 'LuBrain', 'Inference', 'You missed what was implied, not said.']];
  for (let i = 0; i < 4; i++) {
    const [k, ic, n, t] = causes[i], x = 0.6 + (i % 2) * 6.1, y = 1.55 + Math.floor(i / 2) * 2.2;
    card(s, x, y, 5.95, 1.95, SK[k][0], n);
    await circleIcon(s, x + 0.3, y + 0.3, 0.8, k, ic, n);
    s.addText(n, { x: x + 1.3, y: y + 0.3, w: 4.4, h: 0.6, fontSize: 22, bold: true, color: SK[k][1], margin: 0, valign: 'middle', isTextBox: true });
    s.addText(t, { x: x + 1.3, y: y + 0.95, w: 4.4, h: 0.85, fontSize: 16, color: NAVY, margin: 0, valign: 'top', isTextBox: true });
  }
  tipBox(s, 0.6, 6.0, 12.1, 0.75, ['Tag every mistake in your Error Log. Your top cause tells you what to practise.'], 'Error log tip');
  s.addNotes('Give one example of each from the diagnostic. Model tagging two errors live in the Error Log.');

  // 14 · Set targets
  pres.addSection({ title: 'Wrap-up' });
  s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Wrap-up' });
  title(s, 'Set your targets now');
  const st = [['O', 'Find the level your program needs'], ['L', 'Add +1 as a safety margin if you can'], ['R', 'Set each skill on your dashboard'], ['W', 'Write down your weakest skill']];
  for (let i = 0; i < 4; i++) {
    const [k, t] = st[i], x = 0.6 + i * 3.08;
    card(s, x, 1.8, 2.85, 2.9, SK[k][0], 'Step ' + (i + 1));
    numDot(s, x + 0.3, 2.1, 0.7, k, i + 1, 'Step number ' + (i + 1));
    s.addText(t, { x: x + 0.3, y: 3.0, w: 2.3, h: 1.5, fontSize: 18, bold: true, color: NAVY, margin: 0, valign: 'top', isTextBox: true });
    if (i < 3) s.addText('→', { x: x + 2.82, y: 2.95, w: 0.3, h: 0.5, fontSize: 20, bold: true, color: MUTED, align: 'center', margin: 0, isTextBox: true });
  }
  tipBox(s, 0.6, 5.0, 12.1, 1.0, ['Open the course website on your phone now: My Dashboard → Your four skills → Target CLB.'], 'Action');
  s.addNotes('Give 3 minutes for students to set targets on their phones. Walk around and check.');

  // 15 · Before Session 2
  s = pres.addSlide({ masterName: 'CLOSING', sectionTitle: 'Wrap-up' });
  title(s, 'Before Session 2');
  const todo = ['Listening and reading checks', 'Baseline email (27 minutes)', 'Two baseline speaking answers', 'Save your starting point', 'Tag every error in your Error Log'];
  for (let i = 0; i < todo.length; i++) {
    const y = 1.55 + i * 0.95;
    card(s, 0.6, y, 7.6, 0.78, 'FFFFFF', 'To-do ' + (i + 1));
    s.addImage({ data: await icon('LuListChecks', SK.O[1]), x: 0.8, y: y + 0.17, w: 0.44, h: 0.44, altText: 'Check' });
    s.addText(todo[i], { x: 1.45, y, w: 6.6, h: 0.78, fontSize: 18, color: NAVY, valign: 'middle', margin: 0, isTextBox: true });
  }
  card(s, 8.7, 1.55, 4.0, 4.58, 'FFFFFF', 'Study Hub card');
  await circleIcon(s, 8.95, 1.8, 0.9, 'O', 'LuFlag', 'Study Hub');
  s.addText([{ text: 'Study Hub · Session 1', options: { bold: true, fontSize: 20, breakLine: true } }, { text: 'Your results open the next class, so finish before Session 2.', options: { fontSize: 16 } }],
    { x: 8.95, y: 2.9, w: 3.5, h: 2.9, color: NAVY, valign: 'top', margin: 0, paraSpaceAfter: 10, isTextBox: true });
  s.addNotes('Open the Study Hub tab for Session 1 and show the diagnostic. Return to slide 2 and tick the outcomes together.');

  const out = process.argv[2] || 'session-01.pptx';
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log('written', out);
})().catch(e => { console.error(e); process.exit(1); });
