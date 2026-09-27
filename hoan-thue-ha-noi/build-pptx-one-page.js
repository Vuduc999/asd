const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');

const C = {
  green: '86BC25', mid: '3C8F2A', deep: '046A38', dark: '1C3D26', pale: 'F1F6E4',
  black: '000000', ink: '222222', light: 'E6E6E6', white: 'FFFFFF',
  g4: 'BBBCBC', g7: '97999B', g10: '63666A',
  red: 'DA291C', orange: 'ED8B00',
};
const F = 'Arial';

async function icon(Comp, color = '#FFFFFF') {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
pres.title = 'VAT refund at HaNoi Tax Office: scenario model (one-page)';

function title(s, text, sub) {
  s.addText(text, { x: 0.5, y: 0.35, w: 12.3, h: 0.9, fontFace: F, fontSize: 24, bold: true, color: C.black, valign: 'top', margin: 0, isTextBox: true });
  if (sub) s.addText(sub, { x: 0.5, y: 1.22, w: 12.3, h: 0.35, fontFace: F, fontSize: 13, color: C.g10, valign: 'top', margin: 0, isTextBox: true });
}
function footer(s, n) {
  s.addText('VAT refund at Hanoi City Tax  |  Scenario model  |  Discussion document', { x: 0.5, y: 7.05, w: 9, h: 0.3, fontFace: F, fontSize: 9, color: C.g7, margin: 0, isTextBox: true });
  s.addText(String(n), { x: 12.33, y: 7.05, w: 0.5, h: 0.3, fontFace: F, fontSize: 9, color: C.g7, align: 'right', margin: 0, isTextBox: true });
}

// kinds: step (deep green fill), fav (pale green), mid (white, orange tag), adv (white, red tag), neutral, note (dashed)
function box(s, x, y, w, h, kind, eyebrow, head, body, tag) {
  const st = {
    step: { fill: C.deep, line: C.deep, eb: 'CDE6A7', hd: C.white, bd: 'E4F0D0' },
    start: { fill: C.black, line: C.black, eb: C.green, hd: C.white, bd: 'D0D0CE' },
    fav: { fill: C.pale, line: C.green, eb: C.mid, hd: C.black, bd: C.g10 },
    mid: { fill: C.white, line: C.g4, eb: C.g10, hd: C.black, bd: C.g10 },
    adv: { fill: C.white, line: C.g4, eb: C.g10, hd: C.black, bd: C.g10 },
    note: { fill: C.white, line: C.g7, eb: C.g10, hd: C.ink, bd: C.g10, dash: 'dash' },
  }[kind];
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: st.fill }, line: { color: st.line, width: 1.25, dashType: st.dash || 'solid' } });
  const runs = [];
  if (eyebrow) runs.push({ text: eyebrow.toUpperCase(), options: { fontSize: 9, bold: true, color: st.eb, charSpacing: 1, breakLine: true } });
  runs.push({ text: head, options: { fontSize: 12.5, bold: true, color: st.hd, breakLine: !!body } });
  if (body) runs.push({ text: body, options: { fontSize: 10.5, color: st.bd } });
  const tagW = tag ? 1.25 : 0;
  s.addText(runs, { x: x + 0.12, y: y + 0.06, w: w - 0.24, h: h - 0.12, fontFace: F, valign: 'middle', margin: 0, paraSpaceAfter: 2, isTextBox: true });
  if (tag) {
    const col = { fav: C.mid, mid: C.orange, adv: C.red, likely: C.deep }[tag.k];
    s.addText(tag.t.toUpperCase(), { x: x + w - tagW - 0.08, y: y + 0.08, w: tagW, h: 0.22, fontFace: F, fontSize: 7.5, bold: true, color: C.white, fill: { color: col }, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
  }
}
function seg(s, x1, y1, x2, y2, arrow, dash) {
  const o = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1) || 0.0001, h: Math.abs(y2 - y1) || 0.0001, line: { color: C.g7, width: 1.5, dashType: dash || 'solid' } };
  if (x2 < x1 || y2 < y1) { o.flipH = x2 < x1; o.flipV = y2 < y1; }
  if (arrow) o.line.endArrowType = 'triangle';
  s.addShape(pres.shapes.LINE, o);
}
// polyline with arrow at end
function path(s, pts, dash) {
  for (let i = 0; i < pts.length - 1; i++) seg(s, pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1], i === pts.length - 2, dash);
}
function dot(s, x, y, d, fill, line) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: line || fill, width: 1.5 } });
}
function numDot(s, x, y, d, n, fill = C.green, color = C.black) {
  dot(s, x, y, d, fill);
  s.addText(String(n), { x, y, w: d, h: d, fontFace: F, fontSize: d > 0.5 ? 16 : 12, bold: true, color, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
}
function sbox(s, x, y, w, h, kind, eb, hd, bd, tag) {
  const st = {
    step: { fill: C.deep, line: C.deep, eb: 'CDE6A7', hd: C.white, bd: 'E4F0D0' },
    start: { fill: C.black, line: C.black, eb: C.green, hd: C.white, bd: 'D0D0CE' },
    fav: { fill: C.pale, line: C.green, eb: C.mid, hd: C.black, bd: C.g10 },
    mid: { fill: C.white, line: C.g4, eb: C.g10, hd: C.black, bd: C.g10 },
    adv: { fill: C.white, line: C.g4, eb: C.g10, hd: C.black, bd: C.g10 },
    note: { fill: C.white, line: C.g7, eb: C.g10, hd: C.ink, bd: C.g10, dash: 'dash' },
  }[kind];
  const big = kind === 'step' || kind === 'start';
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: st.fill }, line: { color: st.line, width: 1.25, dashType: st.dash || 'solid' } });
  const runs = [{ text: eb.toUpperCase(), options: { fontSize: 8, bold: true, color: st.eb, charSpacing: 1, breakLine: true, paraSpaceAfter: 4 } },
    { text: hd, options: { fontSize: big ? 12 : 10.5, bold: true, color: st.hd, breakLine: !!bd } }];
  if (bd) runs.push({ text: bd, options: { fontSize: 9, color: st.bd } });
  s.addText(runs, { x: x + 0.1, y: y + 0.07, w: w - 0.2, h: h - 0.12, fontFace: F, valign: big ? 'middle' : 'top', margin: 0, paraSpaceAfter: 1, isTextBox: true });
  if (tag) {
    const col = { fav: C.mid, mid: C.orange, adv: C.red, likely: C.deep }[tag.k];
    s.addText(tag.t.toUpperCase(), { x: x + w - 1.0, y: y + 0.06, w: 0.92, h: 0.17, fontFace: F, fontSize: 6.5, bold: true, color: C.white, fill: { color: col }, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
  }
}
function legendY(s, y) {
  const items = [['step', 'YMVN action'], ['fav', 'Favourable outcome'], ['mid', 'Intermediate outcome'], ['adv', 'Negative outcome']];
  let x = 0.5;
  items.forEach(([k, t]) => {
    const fill = k === 'step' ? C.deep : k === 'fav' ? C.pale : C.white;
    const line = k === 'step' ? C.deep : k === 'fav' ? C.green : C.g4;
    s.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.05, w: 0.28, h: 0.18, fill: { color: fill }, line: { color: line, width: 1 } });
    if (k === 'mid' || k === 'adv') s.addShape(pres.shapes.RECTANGLE, { x: x + 0.16, y: y + 0.05, w: 0.12, h: 0.07, fill: { color: k === 'mid' ? C.orange : C.red }, line: { color: k === 'mid' ? C.orange : C.red, width: 0.5 } });
    s.addText(t, { x: x + 0.36, y, w: 2.0, h: 0.28, fontFace: F, fontSize: 10, color: C.g10, valign: 'middle', margin: 0, isTextBox: true });
    x += 2.35;
  });
}
function pageNo(s, n) { s.addText(String(n), { x: 12.33, y: 7.05, w: 0.5, h: 0.3, fontFace: F, fontSize: 9, color: C.g7, align: 'right', margin: 0, isTextBox: true }); }

const FAV = { k: 'fav', t: 'Favourable' }, MID = { k: 'mid', t: 'Intermediate' }, NEG = { k: 'adv', t: 'Negative' };

(async () => {
  // Slide 1: whole process on one page
  let s = pres.addSlide(); s.background = { color: C.white };
  s.addText('Three escalation steps, with a way out of the dispute at every stage', { x: 0.5, y: 0.3, w: 12.3, h: 0.55, fontFace: F, fontSize: 24, bold: true, color: C.black, valign: 'top', margin: 0, isTextBox: true });
  s.addText('Moving right means escalating. Boxes above the line close the case or lead to a decision; boxes below the line push the case to the next step.', { x: 0.5, y: 0.88, w: 12.3, h: 0.35, fontFace: F, fontSize: 12, color: C.g10, valign: 'top', margin: 0, isTextBox: true });

  const CX = [0.5, 3.75, 7.0, 10.25], CW = 2.55, IN = 0.32, IW = CW - IN;
  const RA = 1.42, RB = 2.47, RC = 4.72, RD = 5.77, RH = 0.97, SY = 3.57, SH = 0.97, SC = SY + SH / 2;
  const mid = y => y + RH / 2;

  // spine
  sbox(s, CX[0], SY, CW, SH, 'step', 'Step 1', 'Discussion with HaNoi Tax Office', 'Leadership level, informal');
  sbox(s, CX[1], SY, CW, SH, 'step', 'Step 2', 'Formal letter to HaNoi Tax Office', 'Copy DT; consider MoF, Japan Embassy, JCCI');
  sbox(s, CX[2], SY, CW, SH, 'step', 'Step 3', 'Petition to the Ministry of Finance', 'Only after HaNoi Tax Office has acted in writing');
  sbox(s, CX[3], SY, CW, SH, 'start', 'Result', 'Final unified position', 'MoF, DT and HaNoi Tax Office');
  [[0, 1], [1, 2]].forEach(([a, b]) => {
    path(s, [[CX[a] + CW, SC], [CX[b], SC]]);
    s.addText('NEGATIVE', { x: CX[a] + CW, y: SC - 0.3, w: CX[b] - CX[a] - CW, h: 0.22, fontFace: F, fontSize: 7.5, bold: true, color: C.red, align: 'center', margin: 0, isTextBox: true });
  });

  function up(i, rows) { const bx = CX[i] + 0.16; seg(s, bx, SY, bx, Math.min(...rows.map(mid))); rows.forEach(r => path(s, [[bx, mid(r)], [CX[i] + IN, mid(r)]])); }
  function down(i, ys, dash) { const bx = CX[i] + 0.16; seg(s, bx, SY + SH, bx, Math.max(...ys), false, dash); ys.forEach(y => path(s, [[bx, y], [CX[i] + IN, y]], dash)); }

  // Step 1 column
  sbox(s, CX[0] + IN, RA, IW, RH, 'fav', 'Scenario 1', 'Refund approved', 'Guidance clear to HaNoi Tax Office; case closed', FAV);
  sbox(s, CX[0] + IN, RB, IW, RH, 'mid', 'Scenario 2.1', 'Carry-forward; YMVN accepts', 'Case closed (YMVN to decide in advance)', MID);
  sbox(s, CX[0] + IN, RC, IW, RH, 'adv', 'Scenario 3', 'Refund rejected', 'No refund and no carry-forward', NEG);
  sbox(s, CX[0] + IN, RD, IW, RH, 'adv', 'Scenario 2.2', 'Carry-forward only; YMVN still requests a refund', 'Treated as Scenario 3', NEG);
  up(0, [RA, RB]); down(0, [mid(RC), mid(RD)]);

  // Step 2 column
  sbox(s, CX[1] + IN, RA, IW, RH, 'fav', 'Response A', 'HaNoi Tax Office approves the refund', 'Refund processed; case closed', FAV);
  sbox(s, CX[1] + IN, RB, IW, RH, 'mid', 'Response B', 'Joint meeting (DT chairs)', 'Agreement leads to a refund; if still negative, go to Step 3', MID);
  sbox(s, CX[1] + IN, RC, IW, RH, 'adv', 'Response C', 'Unilateral inspection minutes', 'YMVN records disagreement; asks to hold pending MoF guidance', NEG);
  sbox(s, CX[1] + IN, RD, IW, RH, 'adv', 'Response D', 'Rejection maintained', 'Written reply or no response; basis for escalation', NEG);
  up(1, [RA, RB]); down(1, [mid(RC), mid(RD)]);

  // Step 3 column
  sbox(s, CX[2] + IN, RA, IW, RH, 'mid', 'MoF response', 'Directs DT to convene an inter-agency meeting', null, { k: 'likely', t: 'Most likely' });
  sbox(s, CX[2] + IN, RB, IW, RH, 'mid', 'MoF response', 'Replies directly in writing');
  sbox(s, CX[2] + IN, RC, IW, RH, 'note', 'Contingency', 'MoF refers the file back to DT or HaNoi Tax Office', 'YMVN returns to discussions with HaNoi Tax Office');
  up(2, [RA, RB]); down(2, [mid(RC)], 'dash');
  const bx = CX[3] - 0.35;
  seg(s, CX[2] + CW, mid(RA), bx, mid(RA)); seg(s, CX[2] + CW, mid(RB), bx, mid(RB)); seg(s, bx, mid(RA), bx, SC);
  path(s, [[bx, SC], [CX[3], SC]]);

  // Final column
  sbox(s, CX[3] + IN, RB, IW, RH, 'fav', 'Approved', 'Refund processed', null, FAV);
  sbox(s, CX[3] + IN, RC, IW, RD + RH - RC, 'adv', 'Rejected', 'YMVN considers further recourse', 'Administrative complaint: low chance of success, as MoF and DT have already aligned.\nAdministrative lawsuit before the court.\nThe decision rests with YMVN.', NEG);
  up(3, [RB]); down(3, [RC + (RD + RH - RC) / 2]);

  legendY(s, 6.9); pageNo(s, 1);
  s.addNotes('YMVN decision needed in advance: is carry-forward acceptable? Under Scenario 2 the amount is not refunded in cash; it is offset against future tax payable, and it is not known how long full utilisation would take. The answer decides whether the file stops at 2.1 or moves on to Step 2.');

  // Slide 2: letter (as reviewed)
  s = pres.addSlide(); s.background = { color: C.white };
  s.addText('The Step 2 letter makes three points and signals the escalation path', { x: 0.5, y: 0.35, w: 12.3, h: 0.9, fontFace: F, fontSize: 24, bold: true, color: C.black, valign: 'top', margin: 0, isTextBox: true });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.6, w: 3.6, h: 5.1, fill: { color: C.dark }, line: { color: C.dark } });
  const meta = [['To', 'HaNoi Tax Office'], ['Copy', 'Department of Taxation, the authority that issued the guidance'], ['Consider copying', 'Ministry of Finance, the Japan Embassy, JCCI'], ['Timing', 'Immediately after Step 1, if HaNoi Tax Office’s position is unfavourable']];
  const mr = [];
  meta.forEach(([k, v], i) => { mr.push({ text: k.toUpperCase(), options: { fontSize: 9, bold: true, color: C.green, charSpacing: 1, breakLine: true } }); mr.push({ text: v, options: { fontSize: 14, color: C.white, breakLine: i < meta.length - 1 } }); });
  s.addText(mr, { x: 0.8, y: 1.85, w: 3.0, h: 4.6, fontFace: F, valign: 'top', margin: 0, paraSpaceAfter: 10, isTextBox: true });
  const pts = [
    ['Request the refund', 'YMVN reads the DT’s guidance as confirming eligibility for the refund, and asks HaNoi Tax Office to process it.'],
    ['Propose a joint meeting if the guidance is unclear', 'If HaNoi Tax Office considers the guidance insufficiently clear, it is asked to report to the DT and arrange a meeting of the parties (with the MoF, the Japan Embassy and JCCI where appropriate) to reach a unified view.'],
    ['State the intention to escalate', 'If HaNoi Tax Office still does not agree, YMVN will petition the Ministry of Finance and copy the relevant parties.'],
  ];
  pts.forEach(([h, b], i) => {
    const y = 1.7 + i * 1.7;
    numDot(s, 4.6, y, 0.6, i + 1);
    s.addText([{ text: h, options: { fontSize: 18, bold: true, color: C.black, breakLine: true } }, { text: b, options: { fontSize: 14, color: C.ink } }],
      { x: 5.4, y: y - 0.05, w: 7.4, h: 1.45, fontFace: F, valign: 'top', margin: 0, paraSpaceAfter: 4, isTextBox: true });
  });
  pageNo(s, 2);

  await pres.writeFile({ fileName: 'hanoi-vat-refund-one-page.pptx' });
  console.log('written');
})();
