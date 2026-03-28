"use strict";

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, TableOfContents, HeadingLevel,
  BorderStyle, WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak,
  ExternalHyperlink,
} = require("docx");
const fs = require("fs");
const path = require("path");

// ─── Color Palette ───────────────────────────────────────────────────────────
const NAVY   = "1B3A5C";
const DGRAY  = "333333";
const LGRAY  = "F2F4F7";
const MGRAY  = "D9DEE6";
const WHITE  = "FFFFFF";
const ACCENT = "E8EDF3";

// ─── Page / Content Geometry (DXA) ──────────────────────────────────────────
const PAGE_W = 12240;   // 8.5 in
const PAGE_H = 15840;   // 11 in
const MARGIN  = 1080;   // 0.75 in  (top/bottom)
const HMARGIN = 1260;   // 0.875 in (left/right)
const CONTENT_W = PAGE_W - HMARGIN * 2;   // 9720

// ─── Borders helpers ─────────────────────────────────────────────────────────
const mkBorder = (color = "CCCCCC", sz = 4) => ({
  style: BorderStyle.SINGLE, size: sz, color,
});
const allBorders = (color, sz) => {
  const b = mkBorder(color, sz);
  return { top: b, bottom: b, left: b, right: b };
};
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noAllBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ─── Cell margin standard ───────────────────────────────────────────────────
const CELL_MARGINS = { top: 100, bottom: 100, left: 150, right: 150 };

// ─── Typography helpers ──────────────────────────────────────────────────────
function run(text, opts = {}) {
  return new TextRun({ text, font: "Arial", color: opts.color || DGRAY, ...opts });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: "Arial", bold: true, color: NAVY, size: 36 })],
    spacing: { before: 320, after: 160 },
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial", bold: true, color: NAVY, size: 28 })],
    spacing: { before: 240, after: 120 },
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Arial", bold: true, color: NAVY, size: 24 })],
    spacing: { before: 200, after: 100 },
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    children: [run(text, { color: DGRAY, size: 22, ...opts })],
    spacing: { before: 80, after: 80 },
  });
}

function spacer(pt = 80) {
  return new Paragraph({ children: [run("")], spacing: { before: pt, after: pt } });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// Bullet using numbering reference
function bullet(text, level = 0, refName = "bullets") {
  return new Paragraph({
    numbering: { reference: refName, level },
    children: [run(text, { size: 22, color: DGRAY })],
    spacing: { before: 60, after: 60 },
  });
}

// Numbered item
function numbered(text, refName = "numbers") {
  return new Paragraph({
    numbering: { reference: refName, level: 0 },
    children: [run(text, { size: 22, color: DGRAY })],
    spacing: { before: 60, after: 60 },
  });
}

// ─── Table helpers ───────────────────────────────────────────────────────────
function headerCell(text, width, bgColor = NAVY) {
  return new TableCell({
    borders: allBorders(NAVY, 6),
    width: { size: width, type: WidthType.DXA },
    shading: { fill: bgColor, type: ShadingType.CLEAR },
    margins: CELL_MARGINS,
    verticalAlign: VerticalAlign.CENTER,
    children: [
      new Paragraph({
        children: [new TextRun({ text, font: "Arial", bold: true, color: WHITE, size: 20 })],
        alignment: AlignmentType.LEFT,
      }),
    ],
  });
}

function dataCell(children, width, bgColor = WHITE) {
  const paragraphs = Array.isArray(children) ? children : [
    new Paragraph({
      children: [run(children, { size: 20 })],
    }),
  ];
  return new TableCell({
    borders: allBorders(MGRAY, 4),
    width: { size: width, type: WidthType.DXA },
    shading: { fill: bgColor, type: ShadingType.CLEAR },
    margins: CELL_MARGINS,
    verticalAlign: VerticalAlign.TOP,
    children: paragraphs,
  });
}

function dataCellAlt(children, width) {
  return dataCell(children, width, ACCENT);
}

// Make a simple table with header row
function makeTable(colWidths, headerTexts, rows) {
  const totalW = colWidths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    tableHeader: true,
    children: headerTexts.map((t, i) => headerCell(t, colWidths[i])),
  });

  const dataRows = rows.map((row, ri) => {
    const bg = ri % 2 === 0 ? WHITE : ACCENT;
    return new TableRow({
      children: row.map((cell, ci) => {
        if (typeof cell === "string") {
          return dataCell(cell, colWidths[ci], bg);
        } else {
          // cell is array of paragraphs
          return new TableCell({
            borders: allBorders(MGRAY, 4),
            width: { size: colWidths[ci], type: WidthType.DXA },
            shading: { fill: bg, type: ShadingType.CLEAR },
            margins: CELL_MARGINS,
            verticalAlign: VerticalAlign.TOP,
            children: cell,
          });
        }
      }),
    });
  });

  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows],
  });
}

// ─── Cover Page ──────────────────────────────────────────────────────────────
function buildCoverPage() {
  return [
    spacer(2000),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: "Internal Communications Plan:",
          font: "Arial", bold: true, color: NAVY, size: 56,
        }),
      ],
      spacing: { before: 0, after: 120 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: "Translating Leadership Strategy into Employee Action",
          font: "Arial", bold: true, color: NAVY, size: 52,
        }),
      ],
      spacing: { before: 0, after: 400 },
    }),
    // Decorative rule
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: noAllBorders,
              width: { size: CONTENT_W, type: WidthType.DXA },
              shading: { fill: NAVY, type: ShadingType.CLEAR },
              margins: { top: 4, bottom: 4, left: 0, right: 0 },
              children: [new Paragraph({ children: [run("")] })],
            }),
          ],
        }),
      ],
    }),
    spacer(200),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: "Strategic Priorities, Employee Engagement, and Organizational Alignment",
          font: "Arial", italics: true, color: "4A6FA5", size: 28,
        }),
      ],
      spacing: { before: 0, after: 600 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [run("March 2026", { bold: true, size: 26, color: DGRAY })],
      spacing: { before: 0, after: 200 },
    }),
    spacer(100),
    // Confidential banner table
    new Table({
      width: { size: 6000, type: WidthType.DXA },
      columnWidths: [6000],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: allBorders(NAVY, 8),
              width: { size: 6000, type: WidthType.DXA },
              shading: { fill: LGRAY, type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "CONFIDENTIAL — Internal Use Only",
                      font: "Arial", bold: true, color: NAVY, size: 22,
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    pageBreak(),
  ];
}

// ─── Table of Contents ───────────────────────────────────────────────────────
function buildTOC() {
  return [
    heading1("Table of Contents"),
    new TableOfContents("Table of Contents", {
      hyperlink: true,
      headingStyleRange: "1-3",
    }),
    pageBreak(),
  ];
}

// ─── Section 1: Executive Summary ────────────────────────────────────────────
function buildSection1() {
  return [
    heading1("Section 1: Executive Summary"),
    body(
      "This communications plan bridges the gap between Amazon's externally focused leadership messaging—as " +
      "articulated in the 2024 Letter to Shareholders and the 2025 Proxy Statement—and the day-to-day " +
      "experience of employees across the organization. Leadership has signaled ambitious priorities: " +
      "aggressive AI investment, operational speed and efficiency, a startup-like ownership culture, " +
      "flatter organizational structures, return-to-office collaboration, and sustained cost discipline."
    ),
    body(
      "At the same time, broader workforce sentiment data from the Gallup 2025 State of the Global Workplace " +
      "report reveals falling employee engagement (from 23% to 21% globally), declining manager wellbeing, " +
      "rising stress, and a growing disconnect between executive expectations and frontline realities."
    ),
    body(
      "This plan identifies friction points where leadership messaging may land poorly if delivered without " +
      "context, and builds a structured communications rollout that translates strategic intent into clear, " +
      "audience-appropriate narratives for corporate employees, frontline/operations employees, and people " +
      "managers. It includes dual messaging approaches for sensitive topics, a detailed FAQ, executive " +
      "communications, manager talking points, and revised materials reflecting post-launch feedback."
    ),
    pageBreak(),
  ];
}

// ─── Section 2: Analysis of Leadership Messaging ─────────────────────────────
function buildSection2() {
  return [
    heading1("Section 2: Analysis of Leadership Messaging"),

    heading2("2.1 Key Strategic Priorities from the 2024 Shareholder Letter"),
    body(
      "Andy Jassy's 2024 letter centers on Amazon's \"Why Culture\"—a philosophy of constant questioning, " +
      "customer obsession, and invention. The key priorities communicated are:"
    ),
    numbered("AI as a Defining Investment: Generative AI will reinvent every customer experience. Amazon is " +
      "investing aggressively—over 1,000 GenAI applications in development, new Trainium2 chips, Amazon Nova models, " +
      "Amazon Bedrock and SageMaker expansion. AI revenue is growing at triple-digit YoY percentages.", "n1"),
    numbered("Speed and Urgency: \"There are closing windows all around us.\" Speed is framed as a leadership " +
      "decision, not a trade-off against quality. The company must move faster in every business.", "n1"),
    numbered("Operating Like the World's Largest Startup: This means customer-obsession first, disproportionate " +
      "need for builders, ownership mindset, scrappiness and lean teams, willingness to take risks, and " +
      "delivering compelling results.", "n1"),
    numbered("Flatter Organizations and Ownership: Increasing the ratio of individual contributors to managers. " +
      "Empowering owners of two-way door decisions to move rapidly. Builders should feel fully accountable.", "n1"),
    numbered("Eliminating Bureaucracy: Jassy solicited bureaucracy examples from employees, received nearly " +
      "1,000 emails, and has driven over 375 changes. The company is committed to rooting out red tape.", "n1"),
    numbered("In-Person Collaboration: \"We've found that this process is far more effective in person than remote.\" " +
      "With AI reinventing every experience, \"there has never been a more important time to optimize to invent well.\"", "n1"),
    numbered("Scrappiness and Cost Discipline: \"Our best leaders get the most done with the least number of " +
      "resources.\" Historical team bloat has been addressed.", "n1"),

    spacer(120),
    heading2("2.2 Key Themes from the 2025 Proxy Statement"),
    body("The Proxy Statement reinforces governance, compensation, and operational themes:"),
    numbered("Compensation Aligned to Long-Term Performance: Executive compensation is anchored in time-vested " +
      "RSUs. No equity awards granted to the CEO in 2024.", "n2"),
    numbered("Safety Investments and Progress: Over $2 billion invested in safety since 2019. Recordable " +
      "Incident Rate improved 34% over five years. Plans to allocate hundreds of millions more in 2025.", "n2"),
    numbered("Workforce Investment: Largest-ever annual investment in U.S. hourly employee wages—more than " +
      "$2.2 billion in 2024. Average hourly pay for frontline operations employees exceeds $22/hour with " +
      "total compensation over $29/hour.", "n2"),
    numbered("Sustainability Commitments: Climate Pledge goal of net-zero carbon by 2040. Named world's largest " +
      "corporate purchaser of renewable energy for fifth straight year.", "n2"),
    numbered("Community and Social Impact: Over $1 trillion contributed to U.S. economy since 2010. Housing " +
      "Equity Fund, Career Choice program (250,000+ participants), disaster relief operations.", "n2"),
    numbered("AI Governance and Responsible Development: Board provides active oversight of AI risks. " +
      "Commitment to responsible AI development with privacy and security as core dimensions.", "n2"),

    spacer(120),
    heading2("2.3 Tone and Expectations"),
    body(
      "Leadership messaging is ambitious, fast-paced, and unapologetically high-expectation. Key tone markers " +
      "include: urgency (\"closing windows\"), ownership (\"what would I do if this was my own money\"), lean " +
      "operations (\"scrappy\"), competitive intensity, and a bias toward action over deliberation. The messaging " +
      "celebrates builders and inventors while setting clear expectations that results—not tenure, not charisma, " +
      "not managing up—are what matter."
    ),
    pageBreak(),
  ];
}

// ─── Section 3: Employee Sentiment Analysis ──────────────────────────────────
function buildSection3() {
  return [
    heading1("Section 3: Employee Sentiment Analysis"),

    heading2("3.1 Key Findings from the Gallup 2025 Report"),
    body(
      "The Gallup State of the Global Workplace 2025 report, based on the world's largest ongoing study of " +
      "the employee experience, reveals a workforce at a potential breaking point:"
    ),
    bullet("Global engagement fell from 23% to 21%—a decline equal to the COVID-19 lockdown year."),
    bullet("Manager engagement dropped sharply from 30% to 27%. Young managers (under 35) fell by 5 percentage " +
      "points; female managers fell by 7 points."),
    bullet("70% of team engagement is attributable to the manager—making manager disengagement an existential " +
      "business risk."),
    bullet("Employee life evaluations (wellbeing) fell to 33% thriving globally. Managers experienced the " +
      "steepest declines."),
    bullet("40% of employees experienced daily stress. Managers report higher stress (42%) than individual " +
      "contributors (39%)."),
    bullet("50% of employees globally are watching for or actively seeking new jobs."),
    bullet("Less than half (44%) of managers worldwide have received management training."),
    bullet("The U.S./Canada region, while still above global averages in engagement (31%), has seen historic " +
      "declines in wellbeing."),

    spacer(120),
    heading2("3.2 Friction Points and Gaps"),
    body(
      "When leadership messaging is mapped against sentiment data, several friction points emerge:"
    ),
    numbered("\"Operate Like a Startup\" vs. Manager Burnout: Leadership calls for urgency, lean teams, and " +
      "scrappiness. But managers—the group being asked to execute this—are experiencing their sharpest " +
      "engagement and wellbeing declines in years. Without support, startup intensity risks accelerating burnout.", "n3"),
    numbered("Flatter Organizations vs. Manager Identity: Increasing IC-to-manager ratios signals efficiency " +
      "but can be interpreted as devaluing people management. Managers may feel their roles are being diminished, " +
      "exactly when the data shows they need more support, not less.", "n3"),
    numbered("Speed and Bureaucracy Elimination vs. Psychological Safety: Employees may hear \"move faster and " +
      "do more with less\" as a performance threat rather than an empowerment message. The line between " +
      "eliminating bureaucracy and eliminating support structures is thin in perception.", "n3"),
    numbered("In-Person Collaboration Mandate vs. Employee Flexibility Preferences: The Gallup data shows " +
      "hybrid (23%) and exclusively remote (23%) workers report higher engagement than on-site remote-capable " +
      "workers (19%). Messaging that frames in-person work as non-negotiable may conflict with employee preferences.", "n3"),
    numbered("High-Expectation Culture vs. Recognition Gaps: Employees whose best managers \"were willing to " +
      "help,\" \"wanted to see me shine,\" and \"never let anything fall between the cracks\" describe exactly " +
      "the supportive behaviors that high-pressure cultures can erode. Leadership messaging emphasizes output and " +
      "results; employees crave recognition and development.", "n3"),
    numbered("AI Transformation vs. Job Security Anxiety: Over 1,000 GenAI applications being built across " +
      "Amazon. While framed as customer-experience innovation, employees may read this as role displacement—" +
      "especially when paired with \"lean teams\" and \"scrappiness\" language.", "n3"),

    pageBreak(),
  ];
}

// ─── Section 4: Core Employee-Facing Narrative ───────────────────────────────
function buildSection4() {
  return [
    heading1("Section 4: Core Employee-Facing Narrative"),

    heading2("The Narrative"),
    // Callout box
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: {
                top: mkBorder(NAVY, 12),
                bottom: mkBorder(NAVY, 12),
                left: mkBorder(NAVY, 36),
                right: mkBorder(MGRAY, 4),
              },
              width: { size: CONTENT_W, type: WidthType.DXA },
              shading: { fill: LGRAY, type: ShadingType.CLEAR },
              margins: { top: 160, bottom: 160, left: 240, right: 240 },
              children: [
                new Paragraph({
                  children: [run(
                    "Amazon achieved strong results in 2024—$638B in revenue, 11% growth, record delivery speeds, " +
                    "and major AI milestones. These results reflect the talent and commitment of teams across every " +
                    "part of the organization.",
                    { size: 22, color: DGRAY }
                  )],
                  spacing: { before: 80, after: 120 },
                }),
                new Paragraph({
                  children: [run(
                    "Looking ahead, we are investing aggressively in areas that will define the next decade: AI, " +
                    "faster delivery, healthcare, satellite broadband, and more. This requires us to stay fast, " +
                    "inventive, and customer-obsessed—the qualities that got us here.",
                    { size: 22, color: DGRAY }
                  )],
                  spacing: { before: 80, after: 120 },
                }),
                new Paragraph({
                  children: [run(
                    "We also know that strong results come from supported teams. That's why we're investing more " +
                    "than ever in frontline wages, safety, career development, and manager capability. We're " +
                    "eliminating bureaucracy that slows you down—not support that helps you succeed.",
                    { size: 22, color: DGRAY }
                  )],
                  spacing: { before: 80, after: 120 },
                }),
                new Paragraph({
                  children: [run(
                    "What does this mean for you? It means your work matters more than ever. It means we're " +
                    "committed to giving you the tools, clarity, and support to do your best work. And it means " +
                    "that as our business evolves with AI and new technologies, we'll be transparent about what's " +
                    "changing, why, and how we're investing in our people through every transition.",
                    { size: 22, color: DGRAY }
                  )],
                  spacing: { before: 80, after: 120 },
                }),
                new Paragraph({
                  children: [run(
                    "We operate like the world's largest startup because we believe that ownership, speed, and " +
                    "curiosity produce the best outcomes—for customers and for the people who build Amazon. This " +
                    "is a place for builders who want to make an impact, and we're working to make sure every " +
                    "builder here has what they need to thrive.",
                    { size: 22, color: DGRAY }
                  )],
                  spacing: { before: 80, after: 0 },
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    pageBreak(),
  ];
}

// ─── Section 5: Communications Plan ──────────────────────────────────────────
function buildSection5() {
  // 5.1 Audience table — 4 cols
  const audColWidths = [1600, 2100, 4200, 1820];
  const audRows = [
    [
      "Corporate Employees",
      "HQ and tech roles, corporate functions, hybrid/in-office",
      [
        new Paragraph({ children: [run("1. Strategic vision: AI investment and innovation priorities.", { size: 20 })], spacing: { before: 40, after: 40 } }),
        new Paragraph({ children: [run("2. What \"operate like a startup\" means for your team.", { size: 20 })], spacing: { before: 40, after: 40 } }),
        new Paragraph({ children: [run("3. Bureaucracy elimination: how your feedback drove 375+ changes.", { size: 20 })], spacing: { before: 40, after: 40 } }),
        new Paragraph({ children: [run("4. Career growth in a flatter organization.", { size: 20 })], spacing: { before: 40, after: 40 } }),
      ],
      "Intellectually engaging, transparent, forward-looking",
    ],
    [
      "Frontline / Operations Employees",
      "Fulfillment, logistics, delivery, hourly roles",
      [
        new Paragraph({ children: [run("1. Largest-ever wage investment—$2.2B in 2024.", { size: 20 })], spacing: { before: 40, after: 40 } }),
        new Paragraph({ children: [run("2. Safety progress: 34% improvement in incident rates over 5 years.", { size: 20 })], spacing: { before: 40, after: 40 } }),
        new Paragraph({ children: [run("3. Career Choice and development pathways.", { size: 20 })], spacing: { before: 40, after: 40 } }),
        new Paragraph({ children: [run("4. What AI means for operations (assistance, not replacement).", { size: 20 })], spacing: { before: 40, after: 40 } }),
      ],
      "Direct, concrete, benefit-focused",
    ],
    [
      "People Managers",
      "All people managers across corporate and operations",
      [
        new Paragraph({ children: [run("1. Your role is critical—70% of team engagement depends on you.", { size: 20 })], spacing: { before: 40, after: 40 } }),
        new Paragraph({ children: [run("2. New manager development investments and coaching resources.", { size: 20 })], spacing: { before: 40, after: 40 } }),
        new Paragraph({ children: [run("3. How to talk to your teams about change, AI, and expectations.", { size: 20 })], spacing: { before: 40, after: 40 } }),
        new Paragraph({ children: [run("4. Support structures: what's staying, what's changing, and what we're adding.", { size: 20 })], spacing: { before: 40, after: 40 } }),
      ],
      "Supportive, equipping, honest",
    ],
  ];

  // 5.2 Channel table — 4 cols
  const chanColWidths = [2400, 3200, 2200, 1920];
  const chanRows = [
    [
      "Email",
      "Executive announcements, milestone updates",
      "All employees",
      "Week 1 launch, Week 2 follow-up",
    ],
    [
      "Intranet / News Hub",
      "Detailed content: FAQs, resources, leader videos, narrative articles",
      "All employees",
      "Updated throughout rollout",
    ],
    [
      "Manager-Led Team Meetings",
      "Contextualized discussion, Q&A, local translation of key messages",
      "All employees (via managers)",
      "Week 1 and Week 2 sessions",
    ],
  ];

  // 5.3 Timeline table — 3 cols
  const timeColWidths = [1800, 2600, 5320];
  const timeRows = [
    [
      "Day 1 (Monday)\nWeek 1",
      "Launch & Contextualize",
      "Executive email from CEO to all employees. Intranet hub goes live with narrative, FAQ, and " +
      "supporting resources. Managers receive talking points and briefing guide.",
    ],
    [
      "Days 2–3",
      "Manager Briefings",
      "Manager briefing sessions (virtual/in-person). Managers encouraged to schedule team discussions " +
      "by end of Week 1.",
    ],
    [
      "Days 3–4",
      "Audience-Specific Content",
      "Intranet feature: \"What This Means for You\"—audience-specific articles for corporate, frontline, " +
      "and managers.",
    ],
    [
      "Day 5",
      "FAQ Refresh",
      "FAQ update based on early questions. Manager discussion guide supplement distributed.",
    ],
    [
      "Days 6–7\nWeek 2",
      "Deepen & Reinforce",
      "Follow-up email from senior leadership addressing early themes and questions. Intranet stories " +
      "featuring team-level examples of the strategy in action.",
    ],
    [
      "Days 8–9",
      "Team Meetings (Round 2)",
      "Manager-led team meetings (second round). Focus on local implications, team-specific questions.",
    ],
    [
      "Day 10",
      "AI Spotlight & Q&A",
      "Intranet: AI and Innovation spotlight—what it means for different roles. Employee Q&A feature " +
      "or \"Ask Leadership\" digest published.",
    ],
  ];

  return [
    heading1("Section 5: Communications Plan"),

    heading2("5.1 Audience Segmentation and Key Messages"),
    spacer(80),
    makeTable(audColWidths, ["Audience", "Profile", "Key Messages", "Tone"], audRows),

    spacer(200),
    heading2("5.2 Channel Strategy"),
    spacer(80),
    makeTable(chanColWidths, ["Channel", "Primary Use", "Audience", "Frequency"], chanRows),

    spacer(200),
    heading2("5.3 Two-Week Rollout Timeline"),
    spacer(80),
    makeTable(timeColWidths, ["Timing", "Phase", "Activities"], timeRows),

    pageBreak(),
  ];
}

// ─── Section 6: Executive Announcement Email ─────────────────────────────────
function buildSection6() {
  const emailColWidths = [CONTENT_W];
  const emailContent = [
    new Paragraph({
      children: [new TextRun({ text: "From: ", font: "Arial", bold: true, size: 20, color: DGRAY }),
        new TextRun({ text: "Andy Jassy, President and CEO", font: "Arial", size: 20, color: DGRAY })],
      spacing: { before: 80, after: 40 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "To: ", font: "Arial", bold: true, size: 20, color: DGRAY }),
        new TextRun({ text: "All Amazon Employees", font: "Arial", size: 20, color: DGRAY })],
      spacing: { before: 0, after: 40 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Subject: ", font: "Arial", bold: true, size: 20, color: DGRAY }),
        new TextRun({ text: "Where We're Headed — And What It Means for You", font: "Arial", size: 20, color: DGRAY })],
      spacing: { before: 0, after: 120 },
    }),
    // Divider
    new Table({
      width: { size: CONTENT_W - 600, type: WidthType.DXA },
      columnWidths: [CONTENT_W - 600],
      rows: [new TableRow({ children: [new TableCell({
        borders: noAllBorders, width: { size: CONTENT_W - 600, type: WidthType.DXA },
        shading: { fill: MGRAY, type: ShadingType.CLEAR },
        margins: { top: 2, bottom: 2, left: 0, right: 0 },
        children: [new Paragraph({ children: [run("")] })],
      })]})],
    }),
    new Paragraph({ children: [run("")], spacing: { before: 80, after: 80 } }),
    new Paragraph({
      children: [run("Team,", { bold: true, size: 20 })],
      spacing: { before: 0, after: 120 },
    }),
    new Paragraph({
      children: [run(
        "2024 was a strong year for Amazon. Revenue grew 11% to $638 billion, we shipped at record speeds for " +
        "the second consecutive year, AWS grew 19%, and we launched AI capabilities that are already reshaping how " +
        "we serve customers. These results reflect your work—across fulfillment centers, corporate offices, studios, " +
        "data centers, and every part of this company.",
        { size: 20 }
      )],
      spacing: { before: 0, after: 120 },
    }),
    new Paragraph({
      children: [run("I want to be direct about what comes next.", { size: 20 })],
      spacing: { before: 0, after: 120 },
    }),
    new Paragraph({
      children: [run(
        "We're investing aggressively in AI because we believe it will reinvent every customer experience we " +
        "know. We have over 1,000 GenAI applications in development, new custom silicon chips, and frontier models " +
        "that are already changing how customers shop, build, and create. This isn't a side initiative—it's central " +
        "to everything we do.",
        { size: 20 }
      )],
      spacing: { before: 0, after: 120 },
    }),
    new Paragraph({
      children: [run(
        "At the same time, we're recommitting to the operating principles that make Amazon effective: speed, " +
        "ownership, customer obsession, and lean teams that move fast and build things that matter. I've said we " +
        "want to operate like the world's largest startup, and I mean it. That means empowering builders, reducing " +
        "bureaucracy, and making sure decisions happen close to where the work gets done.",
        { size: 20 }
      )],
      spacing: { before: 0, after: 120 },
    }),
    new Paragraph({
      children: [run(
        "I also know that urgency without support is just pressure. That's why in 2024, we made our largest-ever " +
        "investment in frontline wages—over $2.2 billion. We continued to improve safety outcomes, with our " +
        "recordable incident rate improving 34% over five years. And we expanded Career Choice and development " +
        "programs that have now served over 250,000 employees.",
        { size: 20 }
      )],
      spacing: { before: 0, after: 120 },
    }),
    new Paragraph({
      children: [run(
        "For managers: your role has never been more important. You are the connective tissue between strategy " +
        "and execution. We're investing in new coaching and development resources specifically for you, because we " +
        "know that when managers are supported, their teams perform.",
        { size: 20 }
      )],
      spacing: { before: 0, after: 120 },
    }),
    new Paragraph({
      children: [run(
        "This is a moment of transformation. AI, customer expectations, and competitive dynamics are all " +
        "accelerating. We'll continue to share what this means for different parts of the business in the coming " +
        "weeks. You'll find a detailed FAQ and resources on [Intranet Hub].",
        { size: 20 }
      )],
      spacing: { before: 0, after: 120 },
    }),
    new Paragraph({
      children: [run(
        "Thank you for what you do every day. Amazon is built by people who ask \"why not?\"—and I'm grateful " +
        "to work alongside you.",
        { size: 20 }
      )],
      spacing: { before: 0, after: 120 },
    }),
    new Paragraph({
      children: [run("Andy", { bold: true, size: 20 })],
      spacing: { before: 0, after: 0 },
    }),
  ];

  return [
    heading1("Section 6: Executive Announcement Email"),
    spacer(80),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: emailColWidths,
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: {
                top: mkBorder(NAVY, 8),
                bottom: mkBorder(NAVY, 8),
                left: mkBorder(NAVY, 8),
                right: mkBorder(NAVY, 8),
              },
              width: { size: CONTENT_W, type: WidthType.DXA },
              shading: { fill: "FAFBFC", type: ShadingType.CLEAR },
              margins: { top: 160, bottom: 160, left: 300, right: 300 },
              children: emailContent,
            }),
          ],
        }),
      ],
    }),
    pageBreak(),
  ];
}

// ─── Section 7: Manager Talking Points ───────────────────────────────────────
function buildSection7() {
  const qaColWidths = [3500, 6220];
  const qaRows = [
    [
      "\"Are layoffs coming?\"",
      "There are no announcements about reductions. The company is growing and investing. If anything changes, " +
      "leadership has committed to transparency.",
    ],
    [
      "\"Does 'lean teams' mean my team is getting cut?\"",
      "Lean means efficient, not understaffed. The goal is removing bureaucracy, not removing support.",
    ],
    [
      "\"What does AI mean for my job?\"",
      "AI is creating tools to make your work more effective. New skills and roles will emerge. The company is " +
      "investing in development and Career Choice programs to support transitions.",
    ],
    [
      "\"Why do we have to be in the office?\"",
      "In-person collaboration drives better invention outcomes, especially during this period of AI transformation. " +
      "The company believes the energy and speed of in-person work produces better results for customers and teams.",
    ],
  ];

  return [
    heading1("Section 7: Manager Talking Points"),

    new Paragraph({
      children: [new TextRun({
        text: "Introduction for Managers",
        font: "Arial", bold: true, color: NAVY, size: 24, italics: true,
      })],
      spacing: { before: 80, after: 80 },
    }),
    body(
      "These talking points are designed to help you lead a team discussion about the company's strategic " +
      "direction and what it means for your team. You don't need to cover everything—focus on what's most " +
      "relevant to your group. Be honest about what you know and what's still being worked out."
    ),
    spacer(100),

    heading2("Key Points to Cover"),

    new Paragraph({
      children: [run("1.  Business Results and Context", { bold: true, size: 22, color: NAVY })],
      spacing: { before: 120, after: 60 },
    }),
    body(
      "Amazon grew 11% in revenue in 2024 to $638B. AWS grew 19%. We shipped at record speeds. These results " +
      "create the foundation for continued investment in our people and our future."
    ),

    new Paragraph({
      children: [run("2.  Strategic Priorities", { bold: true, size: 22, color: NAVY })],
      spacing: { before: 120, after: 60 },
    }),
    body(
      "We're investing heavily in AI, faster delivery, healthcare, and new technologies. We're also focused on " +
      "operating with more speed, less bureaucracy, and stronger ownership at every level."
    ),

    new Paragraph({
      children: [run("3.  What \"Operate Like a Startup\" Means", { bold: true, size: 22, color: NAVY })],
      spacing: { before: 120, after: 60 },
    }),
    body(
      "This means each of us takes ownership of outcomes. It means we question whether processes add value. It " +
      "means we move fast—not carelessly, but with urgency and purpose. It does not mean doing more with nothing. " +
      "It means being focused and intentional about what matters most."
    ),

    new Paragraph({
      children: [run("4.  Investing in Our People", { bold: true, size: 22, color: NAVY })],
      spacing: { before: 120, after: 60 },
    }),
    bullet("Largest-ever frontline wage investment ($2.2B)."),
    bullet("Safety incident rates improved 34% over five years."),
    bullet("Career Choice has served 250,000+ employees."),
    bullet("New manager development resources are being rolled out."),

    new Paragraph({
      children: [run("5.  AI and Your Team", { bold: true, size: 22, color: NAVY })],
      spacing: { before: 120, after: 60 },
    }),
    body(
      "AI is creating new tools and capabilities, not replacing the people who use them. Over 1,000 GenAI " +
      "applications are being developed to improve customer experiences. Many will create new roles and skills " +
      "requirements. We'll be transparent as these evolve."
    ),

    new Paragraph({
      children: [run("6.  What Managers Should Emphasize", { bold: true, size: 22, color: NAVY })],
      spacing: { before: 120, after: 60 },
    }),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: {
                top: mkBorder(NAVY, 12),
                bottom: mkBorder(NAVY, 12),
                left: mkBorder(NAVY, 36),
                right: mkBorder(MGRAY, 4),
              },
              width: { size: CONTENT_W, type: WidthType.DXA },
              shading: { fill: LGRAY, type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 240, right: 240 },
              children: [
                new Paragraph({
                  children: [run(
                    "\"Your work is valued, and the company is investing in you—in your pay, your safety, your " +
                    "development, and the tools you need to succeed. The pace is going to be fast, and that's " +
                    "intentional. But I'm here to support you, and leadership is committed to making sure we have " +
                    "what we need.\"",
                    { size: 22, italics: true, color: NAVY }
                  )],
                  spacing: { before: 0, after: 0 },
                }),
              ],
            }),
          ],
        }),
      ],
    }),

    spacer(160),
    heading2("Anticipated Questions and Suggested Responses"),
    spacer(80),
    makeTable(qaColWidths, ["Employee Question", "Suggested Response"], qaRows),

    pageBreak(),
  ];
}

// ─── Section 8: Employee FAQ ──────────────────────────────────────────────────
function buildSection8() {
  const faqs = [
    {
      q: "Q1. How did Amazon perform financially in 2024?",
      a: "Amazon achieved $638 billion in revenue in 2024—an 11% increase year-over-year. AWS grew 19%, we " +
         "delivered at record speeds for the second consecutive year, and we made major AI milestones across " +
         "the company. These results reflect the work of teams across every part of the organization.",
    },
    {
      q: "Q2. Will AI replace my job?",
      a: "No. AI is being developed to make your work more effective, not to eliminate roles. Over 1,000 GenAI " +
         "applications are being built to improve customer experiences—and many will create new roles and skill " +
         "opportunities. Where roles do evolve, we will be transparent about the changes and invest in helping " +
         "employees make successful transitions through Career Choice and development programs.",
    },
    {
      q: "Q3. What does it mean to \"operate like the world's largest startup\"?",
      a: "It means taking ownership of outcomes, questioning whether every process adds real value, making " +
         "decisions close to where the work gets done, and moving with urgency and purpose. It does not mean " +
         "working without support or doing the job of three people. It means being focused and intentional about " +
         "what matters most for customers.",
    },
    {
      q: "Q4. What is bureaucracy elimination, and how does it affect me?",
      a: "Leadership solicited examples of bureaucratic processes that slow people down. Nearly 1,000 examples " +
         "were submitted, and over 375 changes have been driven as a result. The goal is removing processes that " +
         "don't add value—not removing support, resources, or decision-making structures that help you do your job.",
    },
    {
      q: "Q5. What does a flatter organization mean for my career?",
      a: "Flatter organizations are designed to put more decision-making authority in the hands of individual " +
         "contributors and closer to where the work happens. This creates more opportunities for individual impact " +
         "and ownership. Career growth in this model is tied to scope and impact, not just title progression.",
    },
    {
      q: "Q6. Are layoffs coming?",
      a: "There are no announcements about reductions. The company is growing and continuing to invest. If " +
         "circumstances change, leadership is committed to communicating transparently with employees as early " +
         "as possible.",
    },
    {
      q: "Q7. What is happening with frontline wages?",
      a: "In 2024, Amazon made its largest-ever annual investment in U.S. hourly employee wages—more than $2.2 " +
         "billion. Average hourly pay for frontline operations employees now exceeds $22/hour, with total " +
         "compensation exceeding $29/hour including benefits.",
    },
    {
      q: "Q8. What safety improvements have been made?",
      a: "Over $2 billion has been invested in safety since 2019. Our Recordable Incident Rate has improved 34% " +
         "over the past five years. We have plans to allocate hundreds of millions more in 2025 to continue this " +
         "trajectory.",
    },
    {
      q: "Q9. What development opportunities are available to me?",
      a: "Career Choice is available to eligible employees and has now served over 250,000 participants. The " +
         "program pre-pays 95% of tuition for courses in high-demand fields. Additional development resources, " +
         "including manager coaching tools and skills programs, are being expanded across the company.",
    },
    {
      q: "Q10. Why is Amazon requiring employees to work in the office?",
      a: "Leadership believes in-person collaboration produces better invention outcomes—especially during this " +
         "period of rapid AI and business transformation. The company has found that the energy and speed of " +
         "in-person work is particularly important right now for building the future of Amazon.",
    },
    {
      q: "Q11. How does manager support look different going forward?",
      a: "We know that 70% of team engagement is driven by the manager experience. That's why we're investing " +
         "specifically in manager development—new coaching resources, training programs, and support tools. " +
         "Less than half of managers globally have received formal management training, and we intend to close " +
         "that gap.",
    },
    {
      q: "Q12. What does high performance look like in this environment?",
      a: "High performance starts with strong support. We're investing in manager coaching, clearer processes, " +
         "and development programs so that every team has the tools to succeed. Performance is evaluated on " +
         "customer impact and outcome ownership—not tenure or managing up.",
    },
    {
      q: "Q13. What is Amazon doing on sustainability?",
      a: "Amazon's Climate Pledge commits to net-zero carbon by 2040. We have been named the world's largest " +
         "corporate purchaser of renewable energy for the fifth straight year. Sustainability is integrated into " +
         "our long-term business strategy, not a side initiative.",
    },
    {
      q: "Q14. How does executive compensation connect to employee performance?",
      a: "Executive compensation is anchored primarily in time-vested RSUs tied to long-term company performance. " +
         "No equity awards were granted to the CEO in 2024. Executive pay is designed to align leadership " +
         "incentives with the long-term interests of employees, customers, and shareholders.",
    },
    {
      q: "Q15. How can I provide feedback or ask questions?",
      a: "You can submit questions and feedback through the Intranet Hub's dedicated feedback channel. Managers " +
         "will also be holding team sessions specifically for Q&A. Leadership reviews themes from feedback channels " +
         "regularly, and key questions will be addressed in future FAQ updates and communications.",
    },
    {
      q: "Q16. What does AI governance look like—who is responsible?",
      a: "The Board of Directors provides active oversight of AI risks. The company is committed to responsible " +
         "AI development with privacy and security as core dimensions. AI investments are subject to governance " +
         "reviews and are designed to meet both legal and ethical standards.",
    },
    {
      q: "Q17. How is Amazon contributing to the broader community?",
      a: "Amazon has contributed over $1 trillion to the U.S. economy since 2010. The Housing Equity Fund " +
         "supports affordable housing. Career Choice programs serve employees in every part of the country. " +
         "Disaster relief operations have been activated in numerous crisis situations.",
    },
    {
      q: "Q18. What happens if my role is impacted by AI or operational changes?",
      a: "If your role is impacted, we are committed to communicating early and investing in your transition. " +
         "That means access to Career Choice tuition support, upskilling resources, and internal mobility " +
         "pathways. We will not make major operational changes without giving employees advance notice and " +
         "transition support.",
    },
  ];

  const faqRows = faqs.map((faq, i) => {
    const bg = i % 2 === 0 ? WHITE : ACCENT;
    return new TableRow({
      children: [
        new TableCell({
          borders: allBorders(MGRAY, 4),
          width: { size: 3200, type: WidthType.DXA },
          shading: { fill: bg, type: ShadingType.CLEAR },
          margins: CELL_MARGINS,
          children: [new Paragraph({ children: [run(faq.q, { bold: true, size: 20, color: NAVY })] })],
        }),
        new TableCell({
          borders: allBorders(MGRAY, 4),
          width: { size: 6520, type: WidthType.DXA },
          shading: { fill: bg, type: ShadingType.CLEAR },
          margins: CELL_MARGINS,
          children: [new Paragraph({ children: [run(faq.a, { size: 20 })] })],
        }),
      ],
    });
  });

  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell("Question", 3200),
      headerCell("Answer", 6520),
    ],
  });

  return [
    heading1("Section 8: Employee FAQ"),
    spacer(80),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [3200, 6520],
      rows: [headerRow, ...faqRows],
    }),
    pageBreak(),
  ];
}

// ─── Section 9: Dual Messaging Approaches ────────────────────────────────────
function buildSection9() {
  const evalColWidths = [3200, 3260, 3260];
  const evalRows = [
    [
      "Clarity of Expectations",
      "High — employees know exactly what's expected",
      "Moderate — expectations are present but may feel secondary to support language",
    ],
    [
      "Employee Trust",
      "Risk — may feel transactional, especially for frontline employees experiencing burnout",
      "High — builds trust by leading with investment and empathy",
    ],
    [
      "Manager Confidence",
      "Moderate — managers may feel unsupported in delivering tough messages",
      "High — managers have language and framing that connects pressure to support",
    ],
    [
      "Alignment with Sentiment Data",
      "Low — directly conflicts with engagement decline data; risks amplifying burnout",
      "High — addresses the gaps identified in the Gallup report",
    ],
    [
      "Risk of Misinterpretation",
      "High — \"lean\" and \"speed\" can be heard as \"do more with less and be grateful\"",
      "Low — framing reduces likelihood of negative interpretation",
    ],
    [
      "Leadership Credibility",
      "High within performance-oriented segments; lower with exhausted workforce",
      "High — balanced approach preserves credibility across audience segments",
    ],
  ];

  return [
    heading1("Section 9: Dual Messaging Approaches"),

    heading2("9.1 Theme: Performance Expectations and Efficiency"),

    heading3("Approach A: Direct and Performance-Driven"),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: { top: mkBorder(MGRAY, 4), bottom: mkBorder(MGRAY, 4), left: mkBorder("CC4444", 24), right: mkBorder(MGRAY, 4) },
              width: { size: CONTENT_W, type: WidthType.DXA },
              shading: { fill: "FFF8F8", type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 240, right: 240 },
              children: [
                new Paragraph({ children: [new TextRun({ text: "Core Message:", font: "Arial", bold: true, size: 20, color: DGRAY })], spacing: { before: 0, after: 60 } }),
                new Paragraph({ children: [run("\"Amazon's competitive advantage depends on speed, ownership, and lean execution. We're raising the bar—not because we're dissatisfied, but because the opportunity ahead demands it. Builders who thrive here take ownership, move fast, and deliver results.\"", { size: 20, italics: true })], spacing: { before: 0, after: 120 } }),
                new Paragraph({ children: [new TextRun({ text: "Application in Executive Email:", font: "Arial", bold: true, size: 20, color: DGRAY })], spacing: { before: 60, after: 60 } }),
                new Paragraph({ children: [run("Emphasizes competitive urgency, \"closing windows,\" and the expectation that every employee operates with a startup mindset. References AI investment scale to underscore the magnitude of the moment.", { size: 20 })], spacing: { before: 0, after: 100 } }),
                new Paragraph({ children: [new TextRun({ text: "Application in Manager Talking Points:", font: "Arial", bold: true, size: 20, color: DGRAY })], spacing: { before: 60, after: 60 } }),
                new Paragraph({ children: [run("Frames expectations clearly: output and customer impact are what matter most. Encourages managers to set high standards and hold teams accountable.", { size: 20 })], spacing: { before: 0, after: 100 } }),
                new Paragraph({ children: [new TextRun({ text: "Application in FAQ:", font: "Arial", bold: true, size: 20, color: DGRAY })], spacing: { before: 60, after: 60 } }),
                new Paragraph({ children: [run("Answers performance questions directly: \"We expect every team to operate with urgency and ownership. This means faster decisions, leaner processes, and a focus on customer outcomes.\"", { size: 20 })], spacing: { before: 0, after: 0 } }),
              ],
            }),
          ],
        }),
      ],
    }),

    spacer(160),
    heading3("Approach B: Supportive and People-Focused"),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: { top: mkBorder(MGRAY, 4), bottom: mkBorder(MGRAY, 4), left: mkBorder("2E8B57", 24), right: mkBorder(MGRAY, 4) },
              width: { size: CONTENT_W, type: WidthType.DXA },
              shading: { fill: "F0FFF4", type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 240, right: 240 },
              children: [
                new Paragraph({ children: [new TextRun({ text: "Core Message:", font: "Arial", bold: true, size: 20, color: DGRAY })], spacing: { before: 0, after: 60 } }),
                new Paragraph({ children: [run("\"We're asking more of our teams because we believe in what we can build together. But asking more also means investing more—in your development, your tools, your managers, and your wellbeing. We succeed when our people succeed.\"", { size: 20, italics: true })], spacing: { before: 0, after: 120 } }),
                new Paragraph({ children: [new TextRun({ text: "Application in Executive Email:", font: "Arial", bold: true, size: 20, color: DGRAY })], spacing: { before: 60, after: 60 } }),
                new Paragraph({ children: [run("Leads with gratitude and investment. Frames speed and efficiency as outcomes of better support, not just higher expectations. Highlights wage investments, safety improvements, and career development before discussing operational goals.", { size: 20 })], spacing: { before: 0, after: 100 } }),
                new Paragraph({ children: [new TextRun({ text: "Application in Manager Talking Points:", font: "Arial", bold: true, size: 20, color: DGRAY })], spacing: { before: 60, after: 60 } }),
                new Paragraph({ children: [run("Equips managers to lead with empathy and context. Acknowledges that change is hard. Provides specific language for connecting company goals to team and individual support.", { size: 20 })], spacing: { before: 0, after: 100 } }),
                new Paragraph({ children: [new TextRun({ text: "Application in FAQ:", font: "Arial", bold: true, size: 20, color: DGRAY })], spacing: { before: 60, after: 60 } }),
                new Paragraph({ children: [run("Answers performance questions with context: \"High performance starts with strong support. We're investing in manager coaching, development programs, and clearer processes so that every team has what they need to succeed.\"", { size: 20 })], spacing: { before: 0, after: 0 } }),
              ],
            }),
          ],
        }),
      ],
    }),

    spacer(200),
    heading2("9.2 Evaluation"),
    spacer(80),
    makeTable(
      evalColWidths,
      ["Criteria", "Approach A (Direct)", "Approach B (Supportive)"],
      evalRows
    ),

    spacer(200),
    heading2("9.3 Recommendation"),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: allBorders(NAVY, 6),
              width: { size: CONTENT_W, type: WidthType.DXA },
              shading: { fill: LGRAY, type: ShadingType.CLEAR },
              margins: { top: 160, bottom: 160, left: 240, right: 240 },
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "Recommended Approach: ", font: "Arial", bold: true, size: 22, color: NAVY }), run("Approach B (Supportive and People-Focused) as the primary messaging framework, with selective use of Approach A's directness for corporate and senior audiences where competitive context is motivating rather than threatening.", { size: 22, color: DGRAY })],
                  spacing: { before: 0, after: 120 },
                }),
                new Paragraph({ children: [run("The rationale is straightforward: employee engagement is declining, manager burnout is at a critical point, and 50% of the global workforce is actively looking for new roles. Leading with performance pressure in this environment risks accelerating attrition and disengagement. Leading with support and investment—while still clearly communicating expectations—builds the trust necessary for employees to embrace the pace and ambition leadership is setting.", { size: 22 })], spacing: { before: 0, after: 120 } }),
                new Paragraph({ children: [run("In practice, this means: the executive email and frontline materials use Approach B as the primary frame. Manager talking points use Approach B but include Approach A framing for teams where competitive urgency is motivating. The FAQ blends both approaches, with answers that acknowledge the \"why\" (competitive urgency) and the \"how\" (support and investment).", { size: 22 })], spacing: { before: 0, after: 0 } }),
              ],
            }),
          ],
        }),
      ],
    }),

    pageBreak(),
  ];
}

// ─── Section 10: Revised Materials ───────────────────────────────────────────
function buildSection10() {
  const revisedEmailContent = [
    new Paragraph({
      children: [new TextRun({ text: "From: ", font: "Arial", bold: true, size: 20, color: DGRAY }), new TextRun({ text: "Andy Jassy, President and CEO", font: "Arial", size: 20, color: DGRAY })],
      spacing: { before: 80, after: 40 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "To: ", font: "Arial", bold: true, size: 20, color: DGRAY }), new TextRun({ text: "All Amazon Employees", font: "Arial", size: 20, color: DGRAY })],
      spacing: { before: 0, after: 40 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Subject: ", font: "Arial", bold: true, size: 20, color: DGRAY }), new TextRun({ text: "A Follow-Up on Our Direction — And What I Want to Clarify", font: "Arial", size: 20, color: DGRAY })],
      spacing: { before: 0, after: 120 },
    }),
    new Table({
      width: { size: CONTENT_W - 600, type: WidthType.DXA },
      columnWidths: [CONTENT_W - 600],
      rows: [new TableRow({ children: [new TableCell({ borders: noAllBorders, width: { size: CONTENT_W - 600, type: WidthType.DXA }, shading: { fill: MGRAY, type: ShadingType.CLEAR }, margins: { top: 2, bottom: 2, left: 0, right: 0 }, children: [new Paragraph({ children: [run("")] })] })] })],
    }),
    new Paragraph({ children: [run("")], spacing: { before: 80, after: 80 } }),
    new Paragraph({ children: [run("Team,", { bold: true, size: 20 })], spacing: { before: 0, after: 120 } }),
    new Paragraph({ children: [run("Two weeks ago, I shared a message about where Amazon is headed and what it means for our teams. Since then, I've heard from many of you—through your managers, through feedback channels, and directly. I appreciate the honesty, and I want to address what I'm hearing.", { size: 20 })], spacing: { before: 0, after: 120 } }),
    new Paragraph({ children: [run("First, let me be clear about something: when I talk about operating like a startup, I am not signaling layoffs or reductions. I'm talking about a mindset—one where we question whether every process adds value, where decisions happen close to the work, and where everyone feels ownership over outcomes. That's different from doing more with less. It's about doing the right things with clarity and purpose.", { size: 20 })], spacing: { before: 0, after: 120 } }),
    new Paragraph({ children: [run("Second, on AI: we are building AI to make your work better, not to replace you. Yes, roles will evolve—they always have at Amazon. But we're investing in Career Choice, manager coaching, and skills development precisely because we want everyone here to grow with the company. When AI changes what a role looks like, we'll be transparent about it and invest in helping you make that transition.", { size: 20 })], spacing: { before: 0, after: 120 } }),
    new Paragraph({ children: [run("Third, I want to acknowledge that speed without support is just pressure. I'm committed to making sure that our drive for urgency is matched by investment in the people doing the work. That means continuing our record investment in wages, safety, and development. It also means equipping managers with the training and resources they need—because we know that when managers are supported, their teams are too.", { size: 20 })], spacing: { before: 0, after: 120 } }),
    new Paragraph({ children: [run("Here's what you can expect in the coming weeks: your managers will be holding follow-up conversations focused on what these priorities mean specifically for your team. We're also adding new resources to the intranet FAQ based on your questions. And I'll continue sharing updates as our plans evolve.", { size: 20 })], spacing: { before: 0, after: 120 } }),
    new Paragraph({ children: [run("Amazon is at its best when we're honest with each other. Thank you for your candor, and thank you for your commitment to building something that matters.", { size: 20 })], spacing: { before: 0, after: 120 } }),
    new Paragraph({ children: [run("Andy", { bold: true, size: 20 })], spacing: { before: 0, after: 0 } }),
  ];

  return [
    heading1("Section 10: Revised Materials (Post-Feedback)"),

    heading2("10.1 Context: What Feedback Indicated"),
    body(
      "After the initial rollout, employee feedback indicated confusion and concern around performance " +
      "expectations. Specifically:"
    ),
    bullet("Employees interpreted \"operate like a startup\" and \"lean teams\" as signals of impending " +
      "headcount reductions."),
    bullet("Managers reported difficulty explaining the connection between \"speed and urgency\" and the " +
      "company's investment in employee wellbeing."),
    bullet("Frontline employees felt the messaging was \"corporate talk\" that didn't address their " +
      "day-to-day experience."),
    bullet("Questions about AI and job security increased significantly in manager Q&A sessions."),

    spacer(160),
    heading2("10.2 What Changed and Why"),

    new Paragraph({
      children: [run("1.  Core Narrative", { bold: true, size: 22, color: NAVY })],
      spacing: { before: 120, after: 60 },
    }),
    body(
      "Revised to lead explicitly with investment and support before addressing expectations. Added specific, " +
      "concrete examples of what \"startup culture\" means in practice (and what it does not mean). Included a " +
      "direct statement addressing job security concerns related to AI."
    ),

    new Paragraph({
      children: [run("2.  Executive Email", { bold: true, size: 22, color: NAVY })],
      spacing: { before: 120, after: 60 },
    }),
    body(
      "Added a new opening section acknowledging employee feedback and addressing confusion head-on. Reordered " +
      "content to place people investment before operational expectations. Added specific manager support " +
      "commitments. Included a commitment to follow-up communication cadence."
    ),

    new Paragraph({
      children: [run("3.  Manager Talking Points", { bold: true, size: 22, color: NAVY })],
      spacing: { before: 120, after: 60 },
    }),
    body(
      "Added a new section: \"Addressing Misperceptions.\" Included specific language for the most common " +
      "misinterpretations. Added guidance for managers on how to translate \"startup culture\" to frontline " +
      "contexts. Strengthened the AI transition messaging with concrete commitments."
    ),

    spacer(160),
    heading2("10.3 Revised Executive Email"),
    spacer(80),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: allBorders(NAVY, 8),
              width: { size: CONTENT_W, type: WidthType.DXA },
              shading: { fill: "FAFBFC", type: ShadingType.CLEAR },
              margins: { top: 160, bottom: 160, left: 300, right: 300 },
              children: revisedEmailContent,
            }),
          ],
        }),
      ],
    }),

    spacer(200),
    heading2("10.4 Revised Manager Talking Points"),

    heading3("Addressing Misperceptions (New Section)"),
    body(
      "After the initial rollout, several misperceptions became common themes in team Q&A sessions. Use " +
      "the following guidance to address them directly and confidently."
    ),

    spacer(80),
    makeTable(
      [3200, 6520],
      ["Misperception", "How to Address It"],
      [
        [
          "\"Lean teams means we're getting laid off.\"",
          "Lean means focused and efficient, not understaffed. The company is growing its workforce in areas " +
          "of strategic investment. If that changes, leadership has committed to communicating early.",
        ],
        [
          "\"Startup culture means doing three jobs for one salary.\"",
          "It means ownership over outcomes, not overwork. It means removing bureaucracy so you can focus " +
          "on work that matters—not doing more of everything.",
        ],
        [
          "\"AI is going to take over my role.\"",
          "AI is a tool, not a replacement. Roles will evolve, as they always have at Amazon. We're investing " +
          "specifically in Career Choice and skills programs to help everyone grow with the technology.",
        ],
        [
          "\"The company doesn't care about our wellbeing.\"",
          "The $2.2 billion wage investment, 34% safety improvement, and 250,000+ Career Choice participants " +
          "are direct evidence of investment in our people. Acknowledging that the pace is fast is not " +
          "incompatible with caring about your experience.",
        ],
      ]
    ),

    spacer(200),
    heading3("Strengthened AI Transition Messaging"),
    body(
      "When discussing AI with your team, use the following concrete commitments:"
    ),
    bullet("AI applications are being built to enhance your work, create new tools, and improve customer " +
      "experiences—not to eliminate positions."),
    bullet("Where roles evolve due to AI, employees will receive advance notice and access to transition " +
      "resources, including Career Choice tuition support and internal mobility programs."),
    bullet("New AI-related roles and skill requirements will be communicated transparently as they are " +
      "defined. Watch the Intranet Hub for updates."),
    bullet("If you have specific questions about AI and your role, bring them to me. I'll get answers if " +
      "I don't have them."),

    pageBreak(),
  ];
}

// ─── Section 11: Supporting Assets Summary ───────────────────────────────────
function buildSection11() {
  const assetColWidths = [3000, 2800, 2160, 1760];
  const assetRows = [
    ["Executive Announcement Email (Initial)", "Corporate Communications", "Complete", "Section 6"],
    ["Revised Executive Email (Post-Feedback)", "Corporate Communications", "Complete", "Section 10.3"],
    ["Manager Talking Points (Initial)", "People & Culture", "Complete", "Section 7"],
    ["Revised Manager Talking Points", "People & Culture", "Complete", "Section 10.4"],
    ["Employee FAQ (18 Q&A Pairs)", "Corporate Communications", "Complete", "Section 8"],
    ["Audience Segmentation Matrix", "Communications Strategy", "Complete", "Section 5.1"],
    ["Channel Strategy", "Communications Strategy", "Complete", "Section 5.2"],
    ["Two-Week Rollout Timeline", "Project Management", "Complete", "Section 5.3"],
    ["Dual Messaging Framework", "Communications Strategy", "Complete", "Section 9"],
    ["Dual Messaging Evaluation Matrix", "Communications Strategy", "Complete", "Section 9.2"],
    ["Intranet Hub Content Brief", "Digital Communications", "Pending — Launch", "—"],
    ["Audience-Specific Articles (x3)", "Corporate Communications", "Pending — Week 1", "—"],
    ["Manager Briefing Deck", "People & Culture", "Pending — Day 1", "—"],
    ["AI Spotlight Feature", "Corporate Communications", "Pending — Week 2", "—"],
    ["\"Ask Leadership\" Q&A Digest", "Executive Office", "Pending — Day 10", "—"],
  ];

  return [
    heading1("Section 11: Supporting Assets Summary"),
    body(
      "The following table summarizes all deliverables included in this communications plan, along with " +
      "ownership, current status, and document reference."
    ),
    spacer(80),
    makeTable(
      assetColWidths,
      ["Deliverable", "Owner", "Status", "Reference"],
      assetRows
    ),
    spacer(160),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: allBorders(NAVY, 6),
              width: { size: CONTENT_W, type: WidthType.DXA },
              shading: { fill: LGRAY, type: ShadingType.CLEAR },
              margins: { top: 140, bottom: 140, left: 240, right: 240 },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun({ text: "CONFIDENTIAL — Internal Use Only", font: "Arial", bold: true, size: 20, color: NAVY })],
                  spacing: { before: 0, after: 60 },
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [run("This document is intended solely for internal distribution. Do not forward, reproduce, or share outside of authorized recipients.", { size: 18, color: "666666" })],
                  spacing: { before: 0, after: 0 },
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ];
}

// ─── Assemble Document ───────────────────────────────────────────────────────
async function buildDocument() {
  const doc = new Document({
    numbering: {
      config: [
        {
          reference: "bullets",
          levels: [{
            level: 0, format: LevelFormat.BULLET, text: "\u2022",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          }],
        },
        // Independent numbered lists for each section
        ...[
          "numbers", "n1", "n2", "n3",
        ].map(ref => ({
          reference: ref,
          levels: [{
            level: 0, format: LevelFormat.DECIMAL, text: "%1.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          }],
        })),
      ],
    },
    styles: {
      default: {
        document: { run: { font: "Arial", size: 22, color: DGRAY } },
      },
      paragraphStyles: [
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 36, bold: true, font: "Arial", color: NAVY },
          paragraph: {
            spacing: { before: 320, after: 160 },
            outlineLevel: 0,
          },
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 28, bold: true, font: "Arial", color: NAVY },
          paragraph: {
            spacing: { before: 240, after: 120 },
            outlineLevel: 1,
          },
        },
        {
          id: "Heading3",
          name: "Heading 3",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 24, bold: true, font: "Arial", color: NAVY },
          paragraph: {
            spacing: { before: 200, after: 100 },
            outlineLevel: 2,
          },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: PAGE_W, height: PAGE_H },
            margin: { top: MARGIN, right: HMARGIN, bottom: MARGIN, left: HMARGIN },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "Internal Communications Plan — CONFIDENTIAL", font: "Arial", size: 16, color: "888888" }),
                ],
                alignment: AlignmentType.RIGHT,
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "March 2026  |  Internal Use Only  |  Page ", font: "Arial", size: 16, color: "888888" }),
                  new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "888888" }),
                  new TextRun({ text: " of ", font: "Arial", size: 16, color: "888888" }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 16, color: "888888" }),
                ],
                alignment: AlignmentType.CENTER,
              }),
            ],
          }),
        },
        children: [
          ...buildCoverPage(),
          ...buildTOC(),
          ...buildSection1(),
          ...buildSection2(),
          ...buildSection3(),
          ...buildSection4(),
          ...buildSection5(),
          ...buildSection6(),
          ...buildSection7(),
          ...buildSection8(),
          ...buildSection9(),
          ...buildSection10(),
          ...buildSection11(),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  const outPath = "/workspace/outputs/Internal_Comms_Plan.docx";
  fs.writeFileSync(outPath, buffer);
  console.log("Document written to:", outPath);
  console.log("File size:", buffer.length, "bytes");
}

buildDocument().catch(err => { console.error(err); process.exit(1); });
