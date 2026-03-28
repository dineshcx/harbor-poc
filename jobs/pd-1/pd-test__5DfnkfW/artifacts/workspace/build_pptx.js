const pptxgen = require('/usr/lib/node_modules/pptxgenjs');

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Internal Communications';
pres.title = 'Internal Communications Plan';

// ─── Color palette ───────────────────────────────────────────────────────────
const NAVY   = '1E2761';
const ICE    = 'CADCFC';
const WHITE  = 'FFFFFF';
const LIGHT_BG = 'F4F6FC';  // very light ice tint for content slides

// Slide dimensions: 10" × 5.625"
const W = 10;
const H = 5.625;

// ─── Helper: shadow factory (never reuse objects) ────────────────────────────
const makeShadow = () => ({ type: 'outer', blur: 8, offset: 3, angle: 135, color: '000000', opacity: 0.12 });
const makeShadowSm = () => ({ type: 'outer', blur: 4, offset: 2, angle: 135, color: '000000', opacity: 0.10 });

// ─── Helper: slide title (no underline, just whitespace) ─────────────────────
function addTitle(slide, text, { x = 0.5, y = 0.28, w = 9, dark = false } = {}) {
  slide.addText(text, {
    x, y, w, h: 0.65,
    fontFace: 'Georgia', fontSize: 32, bold: true,
    color: dark ? WHITE : NAVY,
    align: 'left', valign: 'middle', margin: 0,
  });
}

// ─── Helper: navy accent bar on left of title area (replaces underline) ──────
function addAccentBar(slide, { y = 0.28, h = 0.65, dark = false } = {}) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y, w: 0.06, h,
    fill: { color: dark ? ICE : NAVY },
    line: { color: dark ? ICE : NAVY, width: 0 },
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 1: TITLE SLIDE
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: NAVY };

  // Large decorative circle top-right
  slide.addShape(pres.shapes.OVAL, {
    x: 7.2, y: -1.2, w: 4.5, h: 4.5,
    fill: { color: ICE, transparency: 82 },
    line: { color: ICE, width: 0 },
  });
  // Small decorative circle bottom-left
  slide.addShape(pres.shapes.OVAL, {
    x: -0.8, y: 3.8, w: 3.0, h: 3.0,
    fill: { color: ICE, transparency: 88 },
    line: { color: ICE, width: 0 },
  });

  // Horizontal divider strip
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 2.88, w: W, h: 0.04,
    fill: { color: ICE, transparency: 60 },
    line: { color: ICE, width: 0 },
  });

  // Eyebrow label
  slide.addText('LEADERSHIP REVIEW', {
    x: 0.65, y: 1.25, w: 8, h: 0.35,
    fontFace: 'Calibri', fontSize: 11, bold: true,
    color: ICE, charSpacing: 4, align: 'left', margin: 0,
  });

  // Main title
  slide.addText('Internal Communications Plan', {
    x: 0.65, y: 1.65, w: 8.5, h: 1.0,
    fontFace: 'Georgia', fontSize: 36, bold: true,
    color: WHITE, align: 'left', valign: 'middle', margin: 0,
  });

  // Subtitle
  slide.addText('Translating Leadership Strategy into Employee Action', {
    x: 0.65, y: 2.95, w: 8, h: 0.55,
    fontFace: 'Calibri', fontSize: 18,
    color: ICE, align: 'left', valign: 'middle', margin: 0,
  });

  // Date / confidential
  slide.addText('March 2026  |  Confidential', {
    x: 0.65, y: 4.85, w: 5, h: 0.35,
    fontFace: 'Calibri', fontSize: 12,
    color: WHITE, transparency: 50,
    align: 'left', valign: 'middle', margin: 0,
  });

  // Ice-blue accent bar left edge
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: H,
    fill: { color: ICE, transparency: 50 },
    line: { color: ICE, width: 0 },
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 2: EXECUTIVE SUMMARY
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  addAccentBar(slide);
  addTitle(slide, 'Plan Overview', { x: 0.5 });

  // Stat callout cards
  const cards = [
    { big: '$638B', sub: '2024 Revenue', detail: '+11% YoY', x: 0.5 },
    { big: '21%', sub: 'Global Employee', detail: 'Engagement (↓2pp)', x: 3.55 },
    { big: '50%', sub: 'Employees', detail: 'Considering Leaving', x: 6.6 },
  ];
  cards.forEach(({ big, sub, detail, x }) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.1, w: 2.85, h: 2.1,
      fill: { color: NAVY },
      line: { color: NAVY, width: 0 },
      shadow: makeShadow(),
    });
    // Accent top strip
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.1, w: 2.85, h: 0.12,
      fill: { color: ICE },
      line: { color: ICE, width: 0 },
    });
    slide.addText(big, {
      x, y: 1.22, w: 2.85, h: 0.9,
      fontFace: 'Georgia', fontSize: 48, bold: true,
      color: WHITE, align: 'center', valign: 'middle', margin: 0,
    });
    slide.addText(sub, {
      x, y: 2.1, w: 2.85, h: 0.35,
      fontFace: 'Calibri', fontSize: 13, bold: true,
      color: ICE, align: 'center', valign: 'middle', margin: 0,
    });
    slide.addText(detail, {
      x, y: 2.42, w: 2.85, h: 0.38,
      fontFace: 'Calibri', fontSize: 12,
      color: WHITE, align: 'center', valign: 'middle', margin: 0,
    });
  });

  // Body text
  slide.addText(
    'This plan bridges leadership\'s strategic messaging with employee realities, addressing engagement gaps and translating priorities into employee action.',
    {
      x: 0.5, y: 3.4, w: 9, h: 0.85,
      fontFace: 'Calibri', fontSize: 15,
      color: NAVY, align: 'left', valign: 'top', margin: 0,
    }
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 3: LEADERSHIP PRIORITIES — two-column
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  addAccentBar(slide);
  addTitle(slide, 'What Leadership Is Signaling');

  // Column headers
  const colHeaderOpts = (x) => ({
    x, y: 1.05, w: 4.25, h: 0.4,
    fontFace: 'Georgia', fontSize: 16, bold: true,
    color: WHITE, align: 'left', valign: 'middle', margin: [0, 8, 0, 8],
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.05, w: 4.25, h: 0.4, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  slide.addText('Strategic Priorities', colHeaderOpts(0.4));

  slide.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.05, w: 4.25, h: 0.4, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  slide.addText('Investment Commitments', colHeaderOpts(5.15));

  const leftItems = [
    'AI as a defining investment (1,000+ GenAI applications)',
    'Speed and competitive urgency across all business units',
    'Startup culture: ownership, lean teams, accountability',
  ];
  const rightItems = [
    '$2.2B in frontline wages — largest investment ever',
    '34% improvement in safety incident rates year-over-year',
    '250,000+ Career Choice participants supported',
  ];

  leftItems.forEach((text, i) => {
    const y = 1.6 + i * 0.95;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 4.25, h: 0.82,
      fill: { color: WHITE }, line: { color: ICE, width: 1 },
      shadow: makeShadowSm(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 0.06, h: 0.82,
      fill: { color: ICE }, line: { color: ICE, width: 0 },
    });
    slide.addText([{ text: `${i + 1}.  `, options: { bold: true, color: NAVY } }, { text }], {
      x: 0.52, y: y + 0.04, w: 4.05, h: 0.74,
      fontFace: 'Calibri', fontSize: 13, color: NAVY,
      align: 'left', valign: 'middle', margin: 0,
    });
  });

  rightItems.forEach((text, i) => {
    const y = 1.6 + i * 0.95;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.15, y, w: 4.25, h: 0.82,
      fill: { color: NAVY }, line: { color: NAVY, width: 0 },
      shadow: makeShadowSm(),
    });
    slide.addText([{ text: `${i + 1}.  `, options: { bold: true, color: ICE } }, { text }], {
      x: 5.27, y: y + 0.04, w: 4.05, h: 0.74,
      fontFace: 'Calibri', fontSize: 13, color: WHITE,
      align: 'left', valign: 'middle', margin: 0,
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 4: SENTIMENT LANDSCAPE — data cards
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  addAccentBar(slide);
  addTitle(slide, 'What Employees Are Experiencing');

  const metrics = [
    { stat: '23% → 21%', label: 'Overall Engagement', note: 'Equals COVID-year decline' },
    { stat: '30% → 27%', label: 'Manager Engagement', note: 'Critical drop in a key layer' },
    { stat: '42%', label: 'Managers: Daily Stress', note: 'Reported high stress levels' },
    { stat: '−5pp', label: 'Young Manager Drop', note: 'Engagement decline, under-30s' },
    { stat: '−7pp', label: 'Female Manager Drop', note: 'Steepest engagement decline' },
    { stat: '70%', label: 'Team Eng. Driven By', note: 'Manager influence on teams' },
  ];

  // 2 rows × 3 cols
  metrics.forEach(({ stat, label, note }, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.4 + col * 3.15;
    const y = 1.12 + row * 1.88;

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.95, h: 1.65,
      fill: { color: WHITE }, line: { color: ICE, width: 1.5 },
      shadow: makeShadow(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.95, h: 0.08,
      fill: { color: i < 3 ? NAVY : ICE }, line: { color: NAVY, width: 0 },
    });
    slide.addText(stat, {
      x, y: y + 0.08, w: 2.95, h: 0.7,
      fontFace: 'Georgia', fontSize: 28, bold: true,
      color: NAVY, align: 'center', valign: 'middle', margin: 0,
    });
    slide.addText(label, {
      x, y: y + 0.78, w: 2.95, h: 0.42,
      fontFace: 'Calibri', fontSize: 12, bold: true,
      color: NAVY, align: 'center', valign: 'top', margin: 0,
    });
    slide.addText(note, {
      x, y: y + 1.17, w: 2.95, h: 0.38,
      fontFace: 'Calibri', fontSize: 11, italic: true,
      color: '6B7799', align: 'center', valign: 'top', margin: 0,
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 5: GAP ANALYSIS — two-column comparison
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  addAccentBar(slide);
  addTitle(slide, 'Where Messaging Meets Reality');

  const pairs = [
    ['\"Operate like a startup\"',        '\"We\'re cutting headcount\"'],
    ['\"Lean teams, scrappiness\"',        '\"Do more with less\"'],
    ['\"Speed and urgency\"',              '\"No room for mistakes\"'],
    ['\"1,000+ AI applications\"',         '\"AI is replacing my job\"'],
    ['\"Flatter organizations\"',          '\"Managers don\'t matter\"'],
    ['\"In-person collaboration\"',        '\"Flexibility is gone\"'],
  ];

  // Column headers
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.05, w: 4.25, h: 0.38, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  slide.addText('Leadership Says', {
    x: 0.4, y: 1.05, w: 4.25, h: 0.38,
    fontFace: 'Georgia', fontSize: 15, bold: true,
    color: WHITE, align: 'center', valign: 'middle', margin: 0,
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 5.35, y: 1.05, w: 4.25, h: 0.38, fill: { color: 'C0392B' }, line: { color: 'C0392B', width: 0 } });
  slide.addText('Employees May Hear', {
    x: 5.35, y: 1.05, w: 4.25, h: 0.38,
    fontFace: 'Georgia', fontSize: 15, bold: true,
    color: WHITE, align: 'center', valign: 'middle', margin: 0,
  });

  // Arrow divider
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 4.73, y: 1.05, w: 0.54, h: 0.38,
    fill: { color: 'E8E8E8' }, line: { color: 'E8E8E8', width: 0 },
  });
  slide.addText('→', {
    x: 4.73, y: 1.05, w: 0.54, h: 0.38,
    fontFace: 'Calibri', fontSize: 20, bold: true,
    color: NAVY, align: 'center', valign: 'middle', margin: 0,
  });

  pairs.forEach(([left, right], i) => {
    const y = 1.55 + i * 0.64;
    const bg = i % 2 === 0 ? WHITE : 'EEF2FC';

    slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 4.25, h: 0.55, fill: { color: bg }, line: { color: ICE, width: 0.75 } });
    slide.addText(left, {
      x: 0.55, y, w: 4.05, h: 0.55,
      fontFace: 'Calibri', fontSize: 13, italic: true,
      color: NAVY, align: 'left', valign: 'middle', margin: 0,
    });

    slide.addShape(pres.shapes.RECTANGLE, { x: 5.35, y, w: 4.25, h: 0.55, fill: { color: bg }, line: { color: 'FCDCDC', width: 0.75 } });
    slide.addText(right, {
      x: 5.5, y, w: 4.05, h: 0.55,
      fontFace: 'Calibri', fontSize: 13, italic: true,
      color: 'C0392B', align: 'left', valign: 'middle', margin: 0,
    });

    // Arrow connector
    slide.addShape(pres.shapes.RECTANGLE, { x: 4.73, y, w: 0.54, h: 0.55, fill: { color: bg }, line: { color: 'E8E8E8', width: 0.75 } });
    slide.addText('→', {
      x: 4.73, y, w: 0.54, h: 0.55,
      fontFace: 'Calibri', fontSize: 16,
      color: '888888', align: 'center', valign: 'middle', margin: 0,
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 6: CORE NARRATIVE — card layout
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  addAccentBar(slide);
  addTitle(slide, 'Our Employee-Facing Narrative');

  const items = [
    { icon: '01', header: 'Lead with Gratitude and Results', body: 'Acknowledge effort, recognize hard work, and lead with the progress made — before asking for more.' },
    { icon: '02', header: 'Connect Investment to People', body: 'Wages, safety improvements, and development programs are evidence of commitment — lead with them.' },
    { icon: '03', header: 'Frame Expectations Within Support', body: 'Set clear expectations, but pair each ask with the tools, resources, and backing employees need.' },
    { icon: '04', header: 'Address AI with Transparency', body: 'Position AI as a tool that amplifies human work — not a replacement. Share real use cases.' },
    { icon: '05', header: 'Commit to Ongoing Dialogue', body: 'Communication is not a one-way broadcast. Create channels for questions, concerns, and feedback.' },
  ];

  items.forEach(({ icon, header, body }, i) => {
    const y = 1.1 + i * 0.86;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 9.2, h: 0.78,
      fill: { color: i % 2 === 0 ? WHITE : 'EEF2FC' },
      line: { color: ICE, width: 1 },
      shadow: makeShadowSm(),
    });
    // Number badge
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 0.52, h: 0.78,
      fill: { color: NAVY }, line: { color: NAVY, width: 0 },
    });
    slide.addText(icon, {
      x: 0.4, y, w: 0.52, h: 0.78,
      fontFace: 'Georgia', fontSize: 16, bold: true,
      color: ICE, align: 'center', valign: 'middle', margin: 0,
    });
    slide.addText(header, {
      x: 1.0, y: y + 0.06, w: 3.2, h: 0.34,
      fontFace: 'Calibri', fontSize: 14, bold: true,
      color: NAVY, align: 'left', valign: 'top', margin: 0,
    });
    slide.addText(body, {
      x: 1.0, y: y + 0.38, w: 8.5, h: 0.36,
      fontFace: 'Calibri', fontSize: 12,
      color: '444C6E', align: 'left', valign: 'top', margin: 0,
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 7: AUDIENCE STRATEGY — three cards
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  addAccentBar(slide);
  addTitle(slide, 'Audience Segmentation');

  const audiences = [
    {
      label: 'Corporate',
      icon: '🏢',
      color: NAVY,
      points: [
        'Strategic vision and company direction',
        'Career growth within a flatter organization',
        'Wins against bureaucracy and speed barriers',
      ],
    },
    {
      label: 'Frontline',
      icon: '👷',
      color: '1A5276',
      points: [
        'Wage investment: $2.2B — largest ever',
        'Safety progress: 34% fewer incidents',
        'AI as assistance, not replacement',
      ],
    },
    {
      label: 'Managers',
      icon: '👥',
      color: '2874A6',
      points: [
        'Your role is critical to transformation',
        'Development investment prioritizes you',
        'Equipped with tools to lead through change',
      ],
    },
  ];

  audiences.forEach(({ label, points, color }, i) => {
    const x = 0.4 + i * 3.1;
    const cardH = 4.0;

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.1, w: 2.9, h: cardH,
      fill: { color: WHITE }, line: { color: ICE, width: 1.5 },
      shadow: makeShadow(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.1, w: 2.9, h: 0.62,
      fill: { color }, line: { color, width: 0 },
    });
    slide.addText(label, {
      x, y: 1.1, w: 2.9, h: 0.62,
      fontFace: 'Georgia', fontSize: 18, bold: true,
      color: WHITE, align: 'center', valign: 'middle', margin: 0,
    });
    points.forEach((pt, j) => {
      const py = 1.85 + j * 1.05;
      // Small diamond bullet
      slide.addShape(pres.shapes.OVAL, {
        x: x + 0.18, y: py + 0.05, w: 0.14, h: 0.14,
        fill: { color }, line: { color, width: 0 },
      });
      slide.addText(pt, {
        x: x + 0.38, y: py, w: 2.42, h: 0.92,
        fontFace: 'Calibri', fontSize: 13,
        color: NAVY, align: 'left', valign: 'top', margin: 0,
      });
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 8: CHANNEL AND TIMELINE
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  addAccentBar(slide);
  addTitle(slide, 'Two-Week Rollout Plan');

  // Week labels and activities
  const weeks = [
    {
      label: 'Week 1: Launch',
      color: NAVY,
      activities: [
        'Executive email to all employees',
        'Intranet hub go-live with full resources',
        'Manager briefing packets distributed',
        'Team discussions: structured conversation guide',
      ],
    },
    {
      label: 'Week 2: Deepen',
      color: '2874A6',
      activities: [
        'Follow-up email: highlight FAQ and themes',
        'Team meetings round 2: address questions',
        'AI spotlight: real tools, real stories',
        'FAQ and intranet updated with live feedback',
      ],
    },
  ];

  weeks.forEach(({ label, color, activities }, wi) => {
    const x = 0.4 + wi * 4.75;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.08, w: 4.55, h: 4.2,
      fill: { color: WHITE }, line: { color: ICE, width: 1 },
      shadow: makeShadow(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.08, w: 4.55, h: 0.5,
      fill: { color }, line: { color, width: 0 },
    });
    slide.addText(label, {
      x, y: 1.08, w: 4.55, h: 0.5,
      fontFace: 'Georgia', fontSize: 16, bold: true,
      color: WHITE, align: 'center', valign: 'middle', margin: 0,
    });

    activities.forEach((act, ai) => {
      const ay = 1.72 + ai * 0.82;
      // Day indicator
      slide.addShape(pres.shapes.RECTANGLE, {
        x: x + 0.22, y: ay, w: 0.38, h: 0.38,
        fill: { color }, line: { color, width: 0 },
      });
      slide.addText(`D${ai * 2 + (wi * 1) + 1}`, {
        x: x + 0.22, y: ay, w: 0.38, h: 0.38,
        fontFace: 'Calibri', fontSize: 10, bold: true,
        color: WHITE, align: 'center', valign: 'middle', margin: 0,
      });
      slide.addText(act, {
        x: x + 0.68, y: ay, w: 3.72, h: 0.38,
        fontFace: 'Calibri', fontSize: 13,
        color: NAVY, align: 'left', valign: 'middle', margin: 0,
      });
    });
  });

  // Channel summary bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.2, w: W, h: 0.425,
    fill: { color: ICE }, line: { color: ICE, width: 0 },
  });
  slide.addText('Channels: Executive Email  •  Intranet / News Hub  •  Manager-Led Team Meetings  •  FAQ Updates', {
    x: 0.4, y: 5.2, w: 9.2, h: 0.425,
    fontFace: 'Calibri', fontSize: 12, bold: true,
    color: NAVY, align: 'center', valign: 'middle', margin: 0,
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 9: MESSAGING APPROACHES — side-by-side
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  addAccentBar(slide);
  addTitle(slide, 'Two Approaches to Sensitive Themes');

  const approaches = [
    {
      label: 'Approach A',
      sub: 'Direct & Performance-Driven',
      color: '7D3C98',
      points: [
        'Leads with competitive urgency and clear performance bar',
        'Sets expectations upfront — accountability is explicit',
        'Highly motivating for top performers and senior leaders',
        'Risk: may alienate an already-exhausted workforce',
        'Risk: increases stress during a period of low engagement',
      ],
    },
    {
      label: 'Approach B',
      sub: 'Supportive & People-Focused',
      color: '1A7431',
      points: [
        'Leads with investment, recognition, and empathy',
        'Frames expectations within visible support structures',
        'Builds trust — essential during engagement decline',
        'Connects performance to purpose, not just pressure',
        'Appropriate for broad workforce audiences',
      ],
    },
  ];

  approaches.forEach(({ label, sub, color, points }, i) => {
    const x = 0.4 + i * 4.75;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.08, w: 4.55, h: 4.2,
      fill: { color: WHITE }, line: { color: ICE, width: 1.5 },
      shadow: makeShadow(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.08, w: 4.55, h: 0.55,
      fill: { color }, line: { color, width: 0 },
    });
    slide.addText(label, {
      x, y: 1.08, w: 4.55, h: 0.3,
      fontFace: 'Georgia', fontSize: 16, bold: true,
      color: WHITE, align: 'center', valign: 'middle', margin: 0,
    });
    slide.addText(sub, {
      x, y: 1.35, w: 4.55, h: 0.28,
      fontFace: 'Calibri', fontSize: 12, italic: true,
      color: WHITE, align: 'center', valign: 'middle', margin: 0,
    });
    points.forEach((pt, pi) => {
      const py = 1.75 + pi * 0.7;
      const isRisk = pt.startsWith('Risk:');
      slide.addShape(pres.shapes.OVAL, {
        x: x + 0.18, y: py + 0.08, w: 0.13, h: 0.13,
        fill: { color: isRisk ? 'C0392B' : color }, line: { color: isRisk ? 'C0392B' : color, width: 0 },
      });
      slide.addText(pt, {
        x: x + 0.38, y: py, w: 4.02, h: 0.62,
        fontFace: 'Calibri', fontSize: 13,
        color: isRisk ? 'C0392B' : NAVY,
        italic: isRisk, align: 'left', valign: 'top', margin: 0,
      });
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 10: EVALUATION MATRIX — table
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  addAccentBar(slide);
  addTitle(slide, 'Evaluation Matrix');

  const criteria = [
    ['Clarity of Message',        'High',     'Moderate'],
    ['Employee Trust',            'At Risk',  'High'],
    ['Manager Confidence',        'Moderate', 'High'],
    ['Alignment with Engagement Data', 'Low', 'High'],
    ['Risk of Misinterpretation', 'High',     'Low'],
  ];

  const ratingColor = (val) => {
    if (['High'].includes(val)) return { bg: '1A7431', fg: WHITE };
    if (['Moderate'].includes(val)) return { bg: 'F39C12', fg: WHITE };
    if (['Low', 'At Risk'].includes(val)) return { bg: 'C0392B', fg: WHITE };
    return { bg: 'CCCCCC', fg: NAVY };
  };

  // Header row
  const headerY = 1.1;
  const colXs = [0.4, 4.3, 7.0];
  const colWs = [3.75, 2.55, 2.55];
  const rowH = 0.68;

  ['Criteria', 'Approach A', 'Approach B'].forEach((hdr, ci) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: colXs[ci], y: headerY, w: colWs[ci], h: 0.5,
      fill: { color: NAVY }, line: { color: NAVY, width: 0 },
    });
    slide.addText(hdr, {
      x: colXs[ci], y: headerY, w: colWs[ci], h: 0.5,
      fontFace: 'Georgia', fontSize: 14, bold: true,
      color: WHITE, align: 'center', valign: 'middle', margin: 0,
    });
  });

  criteria.forEach(([crit, a, b], ri) => {
    const y = headerY + 0.5 + ri * rowH;
    const bg = ri % 2 === 0 ? WHITE : 'EEF2FC';

    // Criteria cell
    slide.addShape(pres.shapes.RECTANGLE, {
      x: colXs[0], y, w: colWs[0], h: rowH,
      fill: { color: bg }, line: { color: ICE, width: 0.75 },
    });
    slide.addText(crit, {
      x: colXs[0] + 0.15, y, w: colWs[0] - 0.15, h: rowH,
      fontFace: 'Calibri', fontSize: 14, bold: true,
      color: NAVY, align: 'left', valign: 'middle', margin: 0,
    });

    // Approach A cell
    const { bg: aBg, fg: aFg } = ratingColor(a);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: colXs[1], y, w: colWs[1], h: rowH,
      fill: { color: aBg }, line: { color: ICE, width: 0.75 },
    });
    slide.addText(a, {
      x: colXs[1], y, w: colWs[1], h: rowH,
      fontFace: 'Calibri', fontSize: 14, bold: true,
      color: aFg, align: 'center', valign: 'middle', margin: 0,
    });

    // Approach B cell
    const { bg: bBg, fg: bFg } = ratingColor(b);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: colXs[2], y, w: colWs[2], h: rowH,
      fill: { color: bBg }, line: { color: ICE, width: 0.75 },
    });
    slide.addText(b, {
      x: colXs[2], y, w: colWs[2], h: rowH,
      fontFace: 'Calibri', fontSize: 14, bold: true,
      color: bFg, align: 'center', valign: 'middle', margin: 0,
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 11: RECOMMENDATION — dark
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: NAVY };

  // Decorative circle — pushed far right so it doesn't overlap content
  slide.addShape(pres.shapes.OVAL, {
    x: 7.8, y: -1.2, w: 3.8, h: 3.8,
    fill: { color: ICE, transparency: 90 },
    line: { color: ICE, width: 0 },
  });

  addAccentBar(slide, { dark: true });

  slide.addText('Our Recommendation', {
    x: 0.5, y: 0.28, w: 9, h: 0.65,
    fontFace: 'Georgia', fontSize: 32, bold: true,
    color: WHITE, align: 'left', valign: 'middle', margin: 0,
  });

  // Recommendation badge
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 9.2, h: 0.72,
    fill: { color: ICE, transparency: 15 },
    line: { color: ICE, width: 1 },
  });
  slide.addText('Approach B: Supportive and People-Focused', {
    x: 0.4, y: 1.1, w: 9.2, h: 0.72,
    fontFace: 'Georgia', fontSize: 22, bold: true,
    color: ICE, align: 'center', valign: 'middle', margin: 0,
  });

  const rationale = [
    'Engagement is declining — leading with performance pressure accelerates attrition',
    'Manager burnout is at a critical inflection point — they need visible support to deliver',
    'Trust must be rebuilt before performance demands can land effectively',
  ];
  rationale.forEach((pt, i) => {
    const y = 2.08 + i * 0.82;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 0.48, h: 0.62,
      fill: { color: ICE, transparency: 20 }, line: { color: ICE, width: 0 },
    });
    slide.addText(`${i + 1}`, {
      x: 0.4, y, w: 0.48, h: 0.62,
      fontFace: 'Georgia', fontSize: 22, bold: true,
      color: ICE, align: 'center', valign: 'middle', margin: 0,
    });
    slide.addText(pt, {
      x: 0.95, y, w: 8.65, h: 0.62,
      fontFace: 'Calibri', fontSize: 15,
      color: WHITE, align: 'left', valign: 'middle', margin: 0,
    });
  });

  slide.addText('Note: Approach A directness is incorporated selectively for senior and corporate audiences.', {
    x: 0.4, y: 4.98, w: 9.2, h: 0.35,
    fontFace: 'Calibri', fontSize: 11, italic: true,
    color: ICE, align: 'left', valign: 'middle', margin: 0,
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 12: MEASUREMENT FRAMEWORK
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  addAccentBar(slide);
  addTitle(slide, 'How We\'ll Measure Success');

  const metrics = [
    { label: 'Email Open Rates', detail: 'Tracked by audience segment — corporate vs. frontline', target: '' },
    { label: 'Manager Meeting Completion', detail: 'All team-level discussions facilitated within 2 weeks', target: 'Target: 90%+' },
    { label: 'Message Comprehension', detail: 'Pulse survey: do employees understand key priorities?', target: 'Target: 75%+' },
    { label: 'Trust in Leadership', detail: 'Sentiment shift tracked via pulse survey', target: 'Target: +5pp' },
    { label: 'Leading Indicators', detail: 'Intranet search terms, HR inquiry volume and topics', target: '' },
    { label: 'Feedback Loops', detail: 'Pulse surveys, manager forms, Ask Leadership channel', target: '' },
  ];

  metrics.forEach(({ label, detail, target }, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 4.75;
    const y = 1.1 + row * 1.42;

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.55, h: 1.28,
      fill: { color: WHITE }, line: { color: ICE, width: 1 },
      shadow: makeShadowSm(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.06, h: 1.28,
      fill: { color: NAVY }, line: { color: NAVY, width: 0 },
    });
    slide.addText(label, {
      x: x + 0.14, y: y + 0.08, w: 4.3, h: 0.35,
      fontFace: 'Calibri', fontSize: 14, bold: true,
      color: NAVY, align: 'left', valign: 'top', margin: 0,
    });
    slide.addText(detail, {
      x: x + 0.14, y: y + 0.42, w: 4.3, h: 0.48,
      fontFace: 'Calibri', fontSize: 12,
      color: '444C6E', align: 'left', valign: 'top', margin: 0,
    });
    if (target) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: x + 0.14, y: y + 0.9, w: 1.4, h: 0.28,
        fill: { color: NAVY }, line: { color: NAVY, width: 0 },
      });
      slide.addText(target, {
        x: x + 0.14, y: y + 0.9, w: 1.4, h: 0.28,
        fontFace: 'Calibri', fontSize: 11, bold: true,
        color: WHITE, align: 'center', valign: 'middle', margin: 0,
      });
    }
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 13: POST-LAUNCH REVISION
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  addAccentBar(slide);
  addTitle(slide, 'Adapting Based on Feedback');

  // Two-column: What We Heard / What We Changed
  const cols = [
    {
      header: 'What We Heard',
      color: '8E44AD',
      items: [
        'Confusion around "startup culture" — employees feared job losses',
        'Significant AI anxiety — concern about role displacement',
        'Frontline disconnect — messaging felt corporate and distant',
      ],
    },
    {
      header: 'What We Changed',
      color: '1A7431',
      items: [
        'Led with investment and recognition before expectations',
        'Added explicit job security clarity to AI messaging',
        'Strengthened manager tools with frontline-specific talking points',
      ],
    },
  ];

  cols.forEach(({ header, color, items }, ci) => {
    const x = 0.4 + ci * 4.75;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.08, w: 4.55, h: 3.7,
      fill: { color: WHITE }, line: { color: ICE, width: 1 },
      shadow: makeShadow(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.08, w: 4.55, h: 0.48,
      fill: { color }, line: { color, width: 0 },
    });
    slide.addText(header, {
      x, y: 1.08, w: 4.55, h: 0.48,
      fontFace: 'Georgia', fontSize: 16, bold: true,
      color: WHITE, align: 'center', valign: 'middle', margin: 0,
    });
    items.forEach((item, ii) => {
      const iy = 1.7 + ii * 1.02;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: x + 0.18, y: iy + 0.06, w: 0.14, h: 0.14,
        fill: { color }, line: { color, width: 0 },
      });
      slide.addText(item, {
        x: x + 0.4, y: iy, w: 4.0, h: 0.92,
        fontFace: 'Calibri', fontSize: 13,
        color: NAVY, align: 'left', valign: 'top', margin: 0,
      });
    });
  });

  // Key principle banner
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 4.98, w: 9.2, h: 0.42,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 },
  });
  slide.addText('"Listen fast, respond visibly, adjust transparently."', {
    x: 0.4, y: 4.98, w: 9.2, h: 0.42,
    fontFace: 'Georgia', fontSize: 15, italic: true, bold: true,
    color: ICE, align: 'center', valign: 'middle', margin: 0,
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 14: CLOSING / NEXT STEPS — dark
// ═══════════════════════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: NAVY };

  // Decorative elements — kept to corners, clear of content
  slide.addShape(pres.shapes.OVAL, {
    x: -1.5, y: 4.2, w: 3.2, h: 3.2,
    fill: { color: ICE, transparency: 90 },
    line: { color: ICE, width: 0 },
  });
  slide.addShape(pres.shapes.OVAL, {
    x: 8.2, y: -0.8, w: 3.0, h: 3.0,
    fill: { color: ICE, transparency: 90 },
    line: { color: ICE, width: 0 },
  });

  addAccentBar(slide, { dark: true });

  slide.addText('Next Steps', {
    x: 0.5, y: 0.28, w: 9, h: 0.65,
    fontFace: 'Georgia', fontSize: 32, bold: true,
    color: WHITE, align: 'left', valign: 'middle', margin: 0,
  });

  const steps = [
    { num: '1', text: 'Leadership approval of the final messaging approach and narrative framework' },
    { num: '2', text: 'Manager briefing materials distributed and briefing sessions scheduled by [date]' },
    { num: '3', text: 'Week 1 launch executed with full monitoring and measurement activated' },
  ];

  steps.forEach(({ num, text }, i) => {
    const y = 1.18 + i * 1.05;
    slide.addShape(pres.shapes.OVAL, {
      x: 0.4, y, w: 0.62, h: 0.62,
      fill: { color: ICE }, line: { color: ICE, width: 0 },
    });
    slide.addText(num, {
      x: 0.4, y, w: 0.62, h: 0.62,
      fontFace: 'Georgia', fontSize: 22, bold: true,
      color: NAVY, align: 'center', valign: 'middle', margin: 0,
    });
    slide.addText(text, {
      x: 1.12, y: y + 0.04, w: 8.48, h: 0.62,
      fontFace: 'Calibri', fontSize: 16,
      color: WHITE, align: 'left', valign: 'middle', margin: 0,
    });
  });

  // Closing quote — solid darker strip for strong contrast
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.6, w: W, h: 1.025,
    fill: { color: '131A4A' },
    line: { color: '131A4A', width: 0 },
  });
  slide.addText('"Strong communication builds the trust that makes transformation possible."', {
    x: 0.5, y: 4.6, w: 9, h: 1.025,
    fontFace: 'Georgia', fontSize: 17, italic: true, bold: true,
    color: ICE, align: 'center', valign: 'middle', margin: 0,
  });
}

// ─── Write file ───────────────────────────────────────────────────────────────
pres.writeFile({ fileName: '/workspace/outputs/Comms_Plan_Leadership_Summary.pptx' })
  .then(() => console.log('✓ Saved: /workspace/outputs/Comms_Plan_Leadership_Summary.pptx'))
  .catch(err => { console.error('Error:', err); process.exit(1); });
