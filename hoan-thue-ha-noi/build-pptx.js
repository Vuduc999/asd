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
pres.title = 'VAT refund at Hanoi City Tax: scenario model';

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
function legend(s, y) {
  const items = [['step', 'Client action'], ['fav', 'Favourable outcome'], ['mid', 'Intermediate outcome'], ['adv', 'Adverse outcome']];
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

(async () => {
  const icons = {
    eye: await icon(fa.FaRegEye), stairs: await icon(fa.FaLayerGroup), file: await icon(fa.FaFileSignature), sliders: await icon(fa.FaSlidersH),
    check: await icon(fa.FaCheck, '#046A38'), user: await icon(fa.FaUserTie, '#000000'), team: await icon(fa.FaTasks, '#000000'),
  };

  // 1. Title
  let s = pres.addSlide();
  s.background = { color: C.black };
  s.addShape(pres.shapes.OVAL, { x: 8.9, y: 1.4, w: 4.7, h: 4.7, fill: { color: C.green }, line: { color: C.green } });
  s.addShape(pres.shapes.OVAL, { x: 10.55, y: 0.55, w: 1.2, h: 1.2, fill: { color: C.deep }, line: { color: C.deep } });
  s.addText('DISCUSSION DOCUMENT', { x: 0.7, y: 1.6, w: 7.5, h: 0.4, fontFace: F, fontSize: 12, bold: true, color: C.green, charSpacing: 2, margin: 0, isTextBox: true });
  s.addText('VAT refund at Hanoi City Tax', { x: 0.7, y: 2.1, w: 8, h: 1.1, fontFace: F, fontSize: 40, bold: true, color: C.white, margin: 0, isTextBox: true });
  s.addText('Scenario model and next actions following the Department of Taxation’s guidance letter', { x: 0.7, y: 3.25, w: 7.6, h: 1.0, fontFace: F, fontSize: 18, color: 'D0D0CE', valign: 'top', margin: 0, isTextBox: true });
  s.addText('September 2026', { x: 0.7, y: 6.3, w: 5, h: 0.4, fontFace: F, fontSize: 12, color: C.g7, margin: 0, isTextBox: true });

  // 2. Current status
  s = pres.addSlide(); s.background = { color: C.white };
  title(s, 'The refund file is on hold pending the Department of Taxation’s guidance', 'The parties have not yet agreed on how the relevant regulation should be interpreted.');
  const tl = [
    ['done', 'Guidance requested', 'Hanoi City Tax and the client each sent an official letter to the Department of Taxation (DoT) asking for guidance.'],
    ['now', 'DoT reviewing internally', 'The DoT is processing the request and preparing a guidance letter.'],
    ['next', 'DoT issues guidance', 'The trigger point for the scenario model in this document.'],
    ['next', 'Hanoi City Tax decides', 'Based on the guidance, Hanoi City Tax either approves or rejects the refund.'],
  ];
  const ty = 3.35, cw = 12.33 / 4;
  s.addShape(pres.shapes.LINE, { x: 0.5 + cw / 2, y: ty, w: cw * 3, h: 0, line: { color: C.g4, width: 2 } });
  s.addShape(pres.shapes.LINE, { x: 0.5 + cw / 2, y: ty, w: cw, h: 0, line: { color: C.green, width: 3 } });
  tl.forEach(([st, h, b], i) => {
    const cx = 0.5 + cw * i + cw / 2;
    if (st === 'now') {
      dot(s, cx - 0.32, ty - 0.32, 0.64, C.white, C.green);
      dot(s, cx - 0.2, ty - 0.2, 0.4, C.green);
      s.addText('WE ARE HERE', { x: cx - 0.9, y: ty - 0.95, w: 1.8, h: 0.3, fontFace: F, fontSize: 10, bold: true, color: C.deep, align: 'center', margin: 0, isTextBox: true, charSpacing: 1 });
    } else dot(s, cx - 0.15, ty - 0.15, 0.3, st === 'done' ? C.green : C.white, st === 'done' ? C.green : C.g7);
    s.addText(h, { x: cx - cw / 2 + 0.2, y: ty + 0.55, w: cw - 0.4, h: 0.45, fontFace: F, fontSize: 15, bold: true, color: C.black, align: 'center', margin: 0, isTextBox: true });
    s.addText(b, { x: cx - cw / 2 + 0.25, y: ty + 1.05, w: cw - 0.5, h: 1.2, fontFace: F, fontSize: 12, color: C.g10, align: 'center', valign: 'top', margin: 0, isTextBox: true });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.95, w: 12.33, h: 0.75, fill: { color: C.pale }, line: { color: C.pale } });
  s.addText([{ text: 'Why plan now: ', options: { bold: true, color: C.deep } }, { text: 'once the guidance is issued, Hanoi City Tax may act quickly. Agreeing the response to each scenario in advance lets the client react within days instead of weeks.', options: { color: C.ink } }],
    { x: 0.75, y: 5.95, w: 11.9, h: 0.75, fontFace: F, fontSize: 13, valign: 'middle', margin: 0, isTextBox: true });
  footer(s, 2);

  // 3. Overview
  s = pres.addSlide(); s.background = { color: C.white };
  title(s, 'Three escalation steps, each with a clear exit if the outcome is favourable', 'End-to-end view of the scenario model');
  const stages = [
    ['Trigger', 'DoT issues guidance letter', 'Obtain the letter on issuance and align on its interpretation with the client.', 'n/a', 'n/a'],
    ['Step 1', 'Leadership-level discussion with Hanoi City Tax', 'Informal meeting to learn Hanoi City Tax’s position before any minutes or decision are issued.', 'Refund approved, or client accepts carry-forward', 'Refund rejected, or carry-forward only and client still wants a refund'],
    ['Step 2', 'Formal letter to Hanoi City Tax', 'Copy to the DoT. Request the refund, propose a joint meeting, and state the intention to escalate.', 'Hanoi City Tax approves, or joint meeting agrees on refund', 'Rejection maintained, or unilateral inspection minutes'],
    ['Step 3', 'Petition to the Ministry of Finance', 'Address to MoF leadership; copy DoT, Hanoi City Tax, the Embassy and JCCI.', 'Tax authorities’ unified position approves the refund', 'Unified position rejects the refund'],
    ['Final', 'Client decides on further recourse', 'Administrative complaint (low chance of success) or administrative lawsuit.', 'n/a', 'n/a'],
  ];
  const sw = 2.25, sg = 0.27;
  stages.forEach(([eb, h, b, ok, esc], i) => {
    const x = 0.5 + i * (sw + sg);
    const fill = i === 0 ? C.black : i === 4 ? C.light : C.deep;
    const tc = i === 4 ? C.black : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.8, w: sw, h: 2.55, fill: { color: fill }, line: { color: fill } });
    if (i >= 1 && i <= 3) numDot(s, x + 0.15, 1.95, 0.42, i);
    s.addText([{ text: eb.toUpperCase(), options: { fontSize: 9, bold: true, color: i === 4 ? C.g10 : C.green, charSpacing: 1, breakLine: true } },
      { text: h, options: { fontSize: 13, bold: true, color: tc, breakLine: true } },
      { text: b, options: { fontSize: 10.5, color: i === 4 ? C.g10 : 'D8E8C8' } }],
      { x: x + 0.15, y: i >= 1 && i <= 3 ? 2.47 : 1.95, w: sw - 0.3, h: i >= 1 && i <= 3 ? 1.8 : 2.3, fontFace: F, valign: 'top', margin: 0, paraSpaceAfter: 4, isTextBox: true });
    if (i < 4) s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: x + sw + 0.01, y: 2.99, w: 0.25, h: 0.18, rotate: 90, fill: { color: C.g7 }, line: { color: C.g7 } });
    if (ok !== 'n/a') {
      s.addShape(pres.shapes.RECTANGLE, { x, y: 4.55, w: sw, h: 1.05, fill: { color: C.pale }, line: { color: C.green, width: 1 } });
      s.addText([{ text: 'EXIT IF', options: { fontSize: 8.5, bold: true, color: C.mid, charSpacing: 1, breakLine: true } }, { text: ok, options: { fontSize: 10.5, color: C.ink } }],
        { x: x + 0.12, y: 4.6, w: sw - 0.24, h: 0.95, fontFace: F, valign: 'top', margin: 0, isTextBox: true });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 5.75, w: sw, h: 1.05, fill: { color: C.white }, line: { color: C.g4, width: 1 } });
      s.addText([{ text: 'ESCALATE IF', options: { fontSize: 8.5, bold: true, color: C.red, charSpacing: 1, breakLine: true } }, { text: esc, options: { fontSize: 10.5, color: C.ink } }],
        { x: x + 0.12, y: 5.8, w: sw - 0.24, h: 0.95, fontFace: F, valign: 'top', margin: 0, isTextBox: true });
    }
  });
  // fix triangles: use chevron-like arrows
  footer(s, 3);

  // 4. Step 1
  s = pres.addSlide(); s.background = { color: C.white };
  title(s, 'Step 1: an early, informal discussion reveals which of three scenarios applies', 'Held at leadership level so the position we hear is reliable. No written submission is needed yet.');
  box(s, 0.5, 3.45, 2.6, 1.3, 'step', 'Step 1', 'Discussion with Hanoi City Tax', 'Leadership level, informal');
  box(s, 3.8, 1.75, 3.0, 1.0, 'fav', 'Scenario 1', 'Refund approved', 'Guidance is clear to Hanoi City Tax', { k: 'fav', t: 'Favourable' });
  box(s, 3.8, 3.5, 3.0, 1.2, 'mid', 'Scenario 2', 'No refund; carry-forward allowed', 'Amount credited against future tax periods', { k: 'mid', t: 'Intermediate' });
  box(s, 3.8, 5.45, 3.0, 1.0, 'adv', 'Scenario 3', 'Refund rejected', 'No refund and no carry-forward', { k: 'adv', t: 'Adverse' });
  path(s, [[3.1, 4.1], [3.45, 4.1]]);
  seg(s, 3.45, 2.25, 3.45, 5.95);
  path(s, [[3.45, 2.25], [3.8, 2.25]]); path(s, [[3.45, 4.1], [3.8, 4.1]]); path(s, [[3.45, 5.95], [3.8, 5.95]]);
  box(s, 7.55, 1.75, 2.75, 1.0, 'fav', 'Outcome', 'Refund processed; case closed');
  box(s, 7.55, 3.05, 2.75, 0.95, 'mid', 'Scenario 2.1', 'Client accepts carry-forward; case closed');
  box(s, 7.55, 4.2, 2.75, 0.95, 'adv', 'Scenario 2.2', 'Client still requests a refund; treated as Scenario 3');
  box(s, 7.55, 5.45, 2.75, 1.0, 'step', 'Next', 'Go to Step 2: formal letter');
  path(s, [[6.8, 2.25], [7.55, 2.25]]); path(s, [[6.8, 5.95], [7.55, 5.95]]);
  seg(s, 6.8, 4.1, 7.2, 4.1); seg(s, 7.2, 3.52, 7.2, 4.68);
  path(s, [[7.2, 3.52], [7.55, 3.52]]); path(s, [[7.2, 4.68], [7.55, 4.68]]);
  path(s, [[8.925, 5.15], [8.925, 5.45]]);
  s.addShape(pres.shapes.RECTANGLE, { x: 10.7, y: 1.75, w: 2.13, h: 4.7, fill: { color: C.light }, line: { color: C.light } });
  s.addText([{ text: 'CLIENT DECISION', options: { fontSize: 9, bold: true, color: C.deep, charSpacing: 1, breakLine: true } },
    { text: 'Is carry-forward acceptable?', options: { fontSize: 13, bold: true, color: C.black, breakLine: true } },
    { text: 'Under Scenario 2 the amount is not refunded in cash. It is offset against future tax payable, and it is not known how long full utilisation would take.', options: { fontSize: 10.5, color: C.ink, breakLine: true } },
    { text: 'Agreeing this in advance decides whether the file stops at 2.1 or moves on to Step 2.', options: { fontSize: 10.5, color: C.ink } }],
    { x: 10.85, y: 1.9, w: 1.83, h: 4.4, fontFace: F, valign: 'top', margin: 0, paraSpaceAfter: 8, isTextBox: true });
  legend(s, 6.65);
  footer(s, 4);

  // 5. Step 2
  s = pres.addSlide(); s.background = { color: C.white };
  title(s, 'Step 2: a formal letter creates a written record before Hanoi City Tax acts', 'Triggered by Scenario 3 or Scenario 2.2. Hanoi City Tax may respond in one of four ways.');
  box(s, 0.5, 2.75, 2.85, 2.5, 'step', 'Step 2', 'Formal letter to Hanoi City Tax', 'Copy to the DoT.\nConsider copying the MoF, the Embassy and JCCI.\nContent: see next slide.');
  const ry = [1.7, 2.95, 4.2, 5.45], rh = 1.0;
  const resp = [
    ['fav', 'Response A', 'Hanoi City Tax approves the refund', { k: 'fav', t: 'Favourable' }],
    ['mid', 'Response B', 'Joint meeting of the parties', 'Hanoi City Tax reports to the DoT, which chairs', { k: 'mid', t: 'Intermediate' }],
    ['adv', 'Response C', 'Unilateral inspection minutes', 'Hanoi City Tax is entitled to do this', { k: 'adv', t: 'Adverse' }],
    ['adv', 'Response D', 'Rejection maintained', 'In a written reply, or no response', { k: 'adv', t: 'Adverse' }],
  ];
  resp.forEach((r, i) => { const hasBody = typeof r[3] === 'string'; box(s, 4.05, ry[i], 3.1, rh, r[0], r[1], r[2], hasBody ? r[3] : null, hasBody ? r[4] : r[3]); });
  path(s, [[3.35, 4.0], [3.7, 4.0]]);
  seg(s, 3.7, ry[0] + rh / 2, 3.7, ry[3] + rh / 2);
  ry.forEach(y => path(s, [[3.7, y + rh / 2], [4.05, y + rh / 2]]));
  const outs = [
    ['fav', 'Outcome', 'Refund processed; case closed'],
    ['mid', 'Outcome', 'Agreement leads to a refund. If still adverse, go to Step 3'],
    ['adv', 'Client response', 'Record disagreement in the minutes; ask for the file to be held pending MoF guidance'],
    ['adv', 'Client response', 'Use the written position as the basis for escalation'],
  ];
  outs.forEach((o, i) => box(s, 7.6, ry[i], 3.0, rh, o[0], o[1], o[2]));
  ry.forEach(y => path(s, [[7.15, y + rh / 2], [7.6, y + rh / 2]]));
  box(s, 11.05, 2.95, 1.78, 3.5, 'step', 'Step 3', 'Petition to the Ministry of Finance');
  [1, 2, 3].forEach(i => path(s, [[10.6, ry[i] + rh / 2], [11.05, ry[i] + rh / 2]]));
  legend(s, 6.65);
  footer(s, 5);

  // 6. Step 3
  s = pres.addSlide(); s.background = { color: C.white };
  title(s, 'Step 3: escalation to the Ministry of Finance leads to a final, unified position', 'Sent directly to MoF leadership, not back to the DoT, and only once Hanoi City Tax has acted or responded in writing.');
  box(s, 0.5, 2.9, 2.75, 2.0, 'step', 'Step 3', 'Petition to the Ministry of Finance', 'To MoF leadership; copy DoT, Hanoi City Tax, the Embassy and JCCI');
  box(s, 3.85, 1.75, 3.1, 1.2, 'mid', 'MoF response', 'Directs the DoT to convene an inter-agency meeting', null, { k: 'likely', t: 'Most likely' });
  box(s, 3.85, 3.35, 3.1, 1.1, 'mid', 'MoF response', 'Replies directly in writing');
  box(s, 3.85, 4.85, 3.1, 1.3, 'note', 'Contingency', 'MoF refers the file back to the DoT or Hanoi City Tax', 'The client returns to discussions with Hanoi City Tax');
  path(s, [[3.25, 3.9], [3.55, 3.9]]);
  seg(s, 3.55, 2.35, 3.55, 5.5);
  path(s, [[3.55, 2.35], [3.85, 2.35]]); path(s, [[3.55, 3.9], [3.85, 3.9]]); path(s, [[3.55, 5.5], [3.85, 5.5]], 'dash');
  box(s, 7.55, 2.55, 2.55, 1.9, 'step', 'Result', 'Final unified position of the tax authorities', 'MoF, DoT and Hanoi City Tax');
  seg(s, 6.95, 2.35, 7.25, 2.35); seg(s, 6.95, 3.9, 7.25, 3.9); seg(s, 7.25, 2.35, 7.25, 3.9);
  path(s, [[7.25, 3.5], [7.55, 3.5]]);
  box(s, 10.7, 1.75, 2.13, 1.3, 'fav', 'Approved', 'Refund processed', null);
  box(s, 10.7, 3.45, 2.13, 2.7, 'adv', 'Rejected', 'Client considers further recourse', 'Administrative complaint: low chance of success, as the MoF and DoT have already aligned.\nAdministrative lawsuit before the court.\nThe decision rests with the client.');
  seg(s, 10.1, 3.5, 10.4, 3.5); seg(s, 10.4, 2.4, 10.4, 4.8);
  path(s, [[10.4, 2.4], [10.7, 2.4]]); path(s, [[10.4, 4.8], [10.7, 4.8]]);
  legend(s, 6.65);
  footer(s, 6);

  // 7. Letter
  s = pres.addSlide(); s.background = { color: C.white };
  title(s, 'The Step 2 letter makes three points and signals the escalation path');
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.6, w: 3.6, h: 5.1, fill: { color: C.dark }, line: { color: C.dark } });
  const meta = [['To', 'Hanoi City Tax'], ['Copy', 'Department of Taxation, the authority that issued the guidance'], ['Consider copying', 'Ministry of Finance, the Embassy, JCCI'], ['Timing', 'Immediately after Step 1, if Hanoi City Tax’s position is unfavourable']];
  const mr = [];
  meta.forEach(([k, v], i) => { mr.push({ text: k.toUpperCase(), options: { fontSize: 9, bold: true, color: C.green, charSpacing: 1, breakLine: true } }); mr.push({ text: v, options: { fontSize: 14, color: C.white, breakLine: i < meta.length - 1 } }); });
  s.addText(mr, { x: 0.8, y: 1.85, w: 3.0, h: 4.6, fontFace: F, valign: 'top', margin: 0, paraSpaceAfter: 10, isTextBox: true });
  const pts = [
    ['Request the refund', 'The client reads the DoT’s guidance as confirming eligibility for the refund, and asks Hanoi City Tax to process it.'],
    ['Propose a joint meeting if the guidance is unclear', 'If Hanoi City Tax considers the guidance insufficiently clear, it is asked to report to the DoT and arrange a meeting of the parties (with the MoF, the Embassy and JCCI where appropriate) to reach a unified view.'],
    ['State the intention to escalate', 'If Hanoi City Tax still does not agree, the client will petition the Ministry of Finance and copy the relevant parties.'],
  ];
  pts.forEach(([h, b], i) => {
    const y = 1.7 + i * 1.7;
    numDot(s, 4.6, y, 0.6, i + 1);
    s.addText([{ text: h, options: { fontSize: 18, bold: true, color: C.black, breakLine: true } }, { text: b, options: { fontSize: 14, color: C.ink } }],
      { x: 5.4, y: y - 0.05, w: 7.4, h: 1.45, fontFace: F, valign: 'top', margin: 0, paraSpaceAfter: 4, isTextBox: true });
  });
  footer(s, 7);

  // 8. Summary table
  s = pres.addSlide(); s.background = { color: C.white };
  title(s, 'Summary of scenarios and client actions');
  const hdr = ['Scenario', 'Hanoi City Tax position', 'Client action', 'Outcome'].map(t => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.dark }, fontSize: 11 } }));
  const oc = (t, k) => ({ text: t, options: { bold: true, color: k === 'fav' ? C.mid : k === 'mid' ? C.orange : C.red } });
  const rows = [
    ['1', 'Approves the refund in line with the DoT guidance', 'Complete the refund formalities', oc('Refund', 'fav')],
    ['2.1', 'No refund; carry-forward allowed', 'Accept carry-forward and close the file', oc('Carry-forward; closed', 'mid')],
    ['2.2', 'No refund; carry-forward allowed', 'Still request a refund (utilisation period uncertain); treat as Scenario 3', oc('Go to Step 2', 'adv')],
    ['3', 'Rejects the refund', 'Step 2 formal letter; if still adverse, Step 3 petition to the MoF', oc('Per final position', 'adv')],
    ['3, minutes', 'Issues inspection minutes unilaterally', 'Record disagreement; request the file be held pending MoF guidance', oc('Go to Step 3', 'adv')],
    ['Final', 'Unified position rejects the refund', 'Client considers an administrative complaint or lawsuit', oc('Client decision', 'adv')],
  ].map((r, i) => r.map((c, j) => {
    const base = { fontSize: 11, color: C.ink, fill: { color: i % 2 ? 'F4F4F4' : C.white } };
    if (typeof c === 'string') return { text: c, options: Object.assign(base, j === 0 ? { bold: true, color: C.black } : {}) };
    return { text: c.text, options: Object.assign(base, c.options) };
  }));
  s.addTable([hdr, ...rows], { x: 0.5, y: 1.55, w: 12.33, colW: [1.4, 3.4, 5.03, 2.5], fontFace: F, border: { type: 'solid', pt: 0.75, color: C.light }, valign: 'middle', margin: [0.08, 0.12, 0.08, 0.12], rowH: 0.62 });
  footer(s, 8);

  // 9. Principles
  s = pres.addSlide(); s.background = { color: C.white };
  title(s, 'Four principles guide how we manage the case');
  const pr = [
    [icons.eye, 'Learn the position early', 'Meet Hanoi City Tax at leadership level as soon as the guidance is issued. Waiting for minutes or a decision leaves too little time to respond.'],
    [icons.stairs, 'Escalate in sequence', 'Petition the MoF only after Hanoi City Tax has acted or replied in writing. Escalating earlier risks the MoF referring the file back, since the DoT has already given guidance.'],
    [icons.file, 'Keep a written record', 'The Step 2 letter is the basis for the client’s comments in any minutes and for escalation if Hanoi City Tax acts unilaterally.'],
    [icons.sliders, 'Manage what we can control', 'Hanoi City Tax is entitled to issue minutes and decisions unilaterally. The client cannot prevent this, but can manage the situation at every step.'],
  ];
  pr.forEach(([ic, h, b], i) => {
    const x = 0.5 + (i % 2) * 6.27, y = 1.75 + Math.floor(i / 2) * 2.55, w = 6.06, hh = 2.25;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: hh, fill: { color: i % 3 === 0 ? C.pale : 'F4F4F4' }, line: { color: i % 3 === 0 ? C.pale : 'F4F4F4' } });
    dot(s, x + 0.3, y + 0.3, 0.75, C.green);
    s.addImage({ data: ic, x: x + 0.5, y: y + 0.5, w: 0.35, h: 0.35 });
    s.addText([{ text: h, options: { fontSize: 18, bold: true, color: C.black, breakLine: true } }, { text: b, options: { fontSize: 14.5, color: C.ink } }],
      { x: x + 1.3, y: y + 0.3, w: w - 1.6, h: hh - 0.5, fontFace: F, valign: 'top', margin: 0, paraSpaceAfter: 6, isTextBox: true });
  });
  footer(s, 9);

  // 10. Next actions
  s = pres.addSlide(); s.background = { color: C.white };
  title(s, 'Next actions and decisions needed from the client');
  const cols = [
    [icons.team, 'Our team', [
      'Monitor the DoT’s guidance and obtain a copy as soon as it is issued',
      'Analyse the guidance and prepare the client’s arguments for Step 1',
      'Arrange the leadership-level meeting with Hanoi City Tax',
      'Pre-draft the Step 2 letter so it can be sent immediately if needed',
      'Compile all correspondence to date for a possible MoF petition',
    ]],
    [icons.user, 'Client decisions', [
      'Whether carry-forward (Scenario 2) is an acceptable outcome',
      'Whether to copy the MoF, the Embassy and JCCI on the Step 2 letter',
      'Appetite for an administrative complaint or lawsuit if the final position is adverse',
    ]],
  ];
  cols.forEach(([ic, h, items], i) => {
    const x = 0.5 + i * 6.27, w = 6.06;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.6, w, h: 5.1, fill: { color: i ? C.pale : 'F4F4F4' }, line: { color: i ? C.pale : 'F4F4F4' } });
    dot(s, x + 0.3, 1.85, 0.6, C.green);
    s.addImage({ data: ic, x: x + 0.45, y: 2.0, w: 0.3, h: 0.3 });
    s.addText(h, { x: x + 1.05, y: 1.85, w: w - 1.3, h: 0.6, fontFace: F, fontSize: 18, bold: true, color: C.black, valign: 'middle', margin: 0, isTextBox: true });
    s.addText(items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < items.length - 1 } })),
      { x: x + 0.35, y: 2.75, w: w - 0.7, h: 3.8, fontFace: F, fontSize: 15, color: C.ink, valign: 'top', margin: 0, paraSpaceAfter: 14, isTextBox: true });
  });
  footer(s, 10);

  await pres.writeFile({ fileName: 'hanoi-vat-refund-scenarios.pptx' });
  console.log('written');
})();
