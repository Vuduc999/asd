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
pres.title = 'VAT refund at HaNoi Tax Office: scenario model (merged)';

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
  // Slide 1: merged one-page model
  let s = pres.addSlide(); s.background = { color: C.white };
  s.addText('Every step either closes the case or escalates it to the next level', { x: 0.5, y: 0.3, w: 12.3, h: 0.55, fontFace: F, fontSize: 24, bold: true, color: C.black, valign: 'top', margin: 0, isTextBox: true });
  s.addText('Arrows up lead to a closed case. Arrows down show a negative outcome, which moves the case to the next step.', { x: 0.5, y: 0.88, w: 12.3, h: 0.35, fontFace: F, fontSize: 12, color: C.g10, valign: 'top', margin: 0, isTextBox: true });

  const CX = [0.5, 3.75, 7.0, 10.25], CW = 2.55;
  const LY = 1.45, LH = 0.8, SY = 3.4, SH = 1.05, SC = SY + SH / 2, TY = 4.85;
  const lbl = (x, y, w, h, runs) => s.addText(runs, { x, y, w, h, fontFace: F, fontSize: 8.5, color: C.ink, valign: 'middle', margin: 0, isTextBox: true });

  // top lane: case closed
  sbox(s, 0.5, LY, 1.95, LH, 'mid', 'Closed', 'Carry-forward accepted', null, MID);
  sbox(s, 2.65, LY, 12.83 - 2.65, LH, 'fav', 'Case closed', 'Refund processed', 'Can be reached at Step 1, at Step 2, or through the final unified position', FAV);

  // spine
  sbox(s, CX[0], SY, CW, SH, 'step', 'Step 1', 'Discussion with HaNoi Tax Office', 'Leadership level, informal');
  sbox(s, CX[1], SY, CW, SH, 'step', 'Step 2', 'Formal letter to HaNoi Tax Office', 'Copy DT; consider MoF, Japan Embassy, JCCI');
  sbox(s, CX[2], SY, CW, SH, 'step', 'Step 3', 'Petition to the Ministry of Finance', 'MoF convenes an inter-agency meeting (most likely) or replies in writing');
  sbox(s, CX[3], SY, CW, SH, 'start', 'Result', 'Final unified position', 'MoF, DT and HaNoi Tax Office');
  path(s, [[CX[2] + CW, SC], [CX[3], SC]]);

  // exits (up)
  const G = C.mid;
  path(s, [[1.3, SY], [1.3, LY + LH]]);
  lbl(1.4, 2.35, 1.35, 0.9, [{ text: 'Scenario 2.1: ', options: { bold: true, color: G } }, { text: 'carry-forward only; YMVN accepts' }]);
  path(s, [[2.85, SY], [2.85, LY + LH]]);
  lbl(2.95, 2.35, 0.95, 0.9, [{ text: 'Scenario 1: ', options: { bold: true, color: G } }, { text: 'approved' }]);
  path(s, [[4.2, SY], [4.2, LY + LH]]);
  lbl(4.3, 2.35, 2.3, 0.9, [{ text: 'Response A: ', options: { bold: true, color: G } }, { text: 'HaNoi Tax Office approves', options: { breakLine: true } }, { text: 'Response B: ', options: { bold: true, color: G } }, { text: 'joint meeting (DT chairs) agrees' }]);
  path(s, [[10.7, SY], [10.7, LY + LH]]);
  lbl(10.8, 2.35, 1.9, 0.9, [{ text: 'Final position: ', options: { bold: true, color: G } }, { text: 'refund approved' }]);

  // negative outcomes (down) and escalation
  const TH = [1.3, 1.5, 1.1, 1.4];
  sbox(s, CX[0], TY, CW, TH[0], 'adv', 'Scenarios 3, 2.2', 'No refund', 'HaNoi Tax Office rejects the refund, or allows carry-forward only and YMVN still wants a refund', NEG);
  sbox(s, CX[1], TY, CW, TH[1], 'adv', 'Responses B–D', 'Rejection maintained', 'Written rejection or no response, joint meeting still negative, or unilateral inspection minutes. YMVN records disagreement in any minutes and asks to hold the file pending MoF guidance.', NEG);
  sbox(s, CX[2], TY, CW, TH[2], 'note', 'Contingency', 'MoF refers the file back', 'YMVN returns to discussions with HaNoi Tax Office');
  sbox(s, CX[3], TY, CW, TH[3], 'adv', 'Rejected', 'YMVN considers further recourse', 'Administrative complaint: low chance of success, as MoF and DT have already aligned.\nAdministrative lawsuit before the court.\nThe decision rests with YMVN.', NEG);
  [0, 1, 3].forEach(i => path(s, [[CX[i] + CW / 2, SY + SH], [CX[i] + CW / 2, TY]]));
  path(s, [[CX[2] + CW / 2, SY + SH], [CX[2] + CW / 2, TY]], 'dash');
  [0, 1].forEach(i => {
    const x = CX[i] + CW + 0.35, y = TY + 0.45;
    path(s, [[CX[i] + CW, y], [x, y], [x, SC], [CX[i + 1], SC]]);
    s.addText('ESCALATE', { x: x - 0.33, y: SY + SH + 0.05, w: 0.66, h: 0.2, fontFace: F, fontSize: 7, bold: true, color: C.red, align: 'center', margin: 0, isTextBox: true, fill: { color: C.white } });
  });

  legendY(s, 6.75); pageNo(s, 1);
  s.addNotes('YMVN decision needed in advance: is carry-forward acceptable? Under Scenario 2 the amount is not refunded in cash; it is offset against future tax payable, and it is not known how long full utilisation would take. The answer decides whether the file closes (Scenario 2.1) or moves on to Step 2 (Scenario 2.2).');

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

  await pres.writeFile({ fileName: 'hanoi-vat-refund-merged.pptx' });
  console.log('written');
})();
