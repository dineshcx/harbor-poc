const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Student Marks Distribution";

// Colors
const PRIMARY = "1E2761";
const SECONDARY = "CADCFC";
const ACCENT = "408EC6";
const WHITE = "FFFFFF";
const DARK_TEXT = "1E293B";

// ========= SLIDE 1: Title =========
let slide1 = pres.addSlide();
slide1.background = { color: PRIMARY };
slide1.addText("Student Marks\nDistribution Report", {
  x: 0.8, y: 1.2, w: 8.4, h: 2.5,
  fontSize: 42, fontFace: "Georgia", color: WHITE,
  bold: true, align: "left", lineSpacingMultiple: 1.2
});
slide1.addText("Academic Performance Overview", {
  x: 0.8, y: 3.8, w: 6, h: 0.6,
  fontSize: 18, fontFace: "Calibri", color: SECONDARY, italic: true
});
slide1.addShape(pres.shapes.RECTANGLE, {
  x: 0.8, y: 4.6, w: 2.5, h: 0.05, fill: { color: ACCENT }
});

// ========= SLIDE 2: Grade Distribution Bar Chart =========
let slide2 = pres.addSlide();
slide2.background = { color: "F8FAFC" };
slide2.addText("Grade Distribution", {
  x: 0.5, y: 0.3, w: 9, h: 0.7,
  fontSize: 32, fontFace: "Georgia", color: PRIMARY, bold: true, margin: 0
});
slide2.addText("Number of students per grade category", {
  x: 0.5, y: 0.95, w: 9, h: 0.4,
  fontSize: 14, fontFace: "Calibri", color: "64748B", margin: 0
});

slide2.addChart(pres.charts.BAR, [{
  name: "Students",
  labels: ["A (90-100)", "B (80-89)", "C (70-79)", "D (60-69)", "F (<60)"],
  values: [3, 3, 2, 2, 0]
}], {
  x: 0.5, y: 1.5, w: 9, h: 3.8, barDir: "col",
  chartColors: ["1E6F5C", "289672", "29BB89", "E6DD3B", "E84545"],
  chartArea: { fill: { color: WHITE }, roundedCorners: true },
  catAxisLabelColor: "64748B", catAxisLabelFontSize: 12,
  valAxisLabelColor: "64748B", valAxisLabelFontSize: 10,
  valGridLine: { color: "E2E8F0", size: 0.5 },
  catGridLine: { style: "none" },
  showValue: true, dataLabelPosition: "outEnd", dataLabelColor: DARK_TEXT,
  showLegend: false
});

// ========= SLIDE 3: Score Distribution (individual scores) =========
let slide3 = pres.addSlide();
slide3.background = { color: "F8FAFC" };
slide3.addText("Individual Student Scores", {
  x: 0.5, y: 0.3, w: 9, h: 0.7,
  fontSize: 32, fontFace: "Georgia", color: PRIMARY, bold: true, margin: 0
});
slide3.addText("Score out of 100 for each student", {
  x: 0.5, y: 0.95, w: 9, h: 0.4,
  fontSize: 14, fontFace: "Calibri", color: "64748B", margin: 0
});

slide3.addChart(pres.charts.BAR, [{
  name: "Score",
  labels: ["Alice J.", "Bob S.", "Charlie B.", "Diana R.", "Ethan H.", "Fiona A.", "George L.", "Hannah M.", "Ivan D.", "Julia R."],
  values: [92, 78, 85, 67, 91, 73, 88, 95, 62, 81]
}], {
  x: 0.5, y: 1.5, w: 9, h: 3.8, barDir: "col",
  chartColors: ["408EC6"],
  chartArea: { fill: { color: WHITE }, roundedCorners: true },
  catAxisLabelColor: "64748B", catAxisLabelFontSize: 9,
  valAxisLabelColor: "64748B", valAxisLabelFontSize: 10,
  valGridLine: { color: "E2E8F0", size: 0.5 },
  catGridLine: { style: "none" },
  showValue: true, dataLabelPosition: "outEnd", dataLabelColor: DARK_TEXT, dataLabelFontSize: 10,
  showLegend: false
});

// ========= SLIDE 4: Summary Stats =========
let slide4 = pres.addSlide();
slide4.background = { color: PRIMARY };
slide4.addText("Key Statistics", {
  x: 0.5, y: 0.3, w: 9, h: 0.7,
  fontSize: 32, fontFace: "Georgia", color: WHITE, bold: true, margin: 0
});

const stats = [
  { label: "Highest Score", value: "95", sub: "Hannah Montana" },
  { label: "Lowest Score", value: "62", sub: "Ivan Drago" },
  { label: "Average Score", value: "81.2", sub: "Class Mean" },
  { label: "Total Students", value: "10", sub: "Enrolled" }
];
const cardW = 2.0, gap = 0.35, totalW = cardW * 4 + gap * 3;
const startX = (10 - totalW) / 2;

stats.forEach((s, i) => {
  const cx = startX + i * (cardW + gap);
  slide4.addShape(pres.shapes.RECTANGLE, {
    x: cx, y: 1.5, w: cardW, h: 3.0, fill: { color: "283A6B" },
    shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.2 }
  });
  slide4.addText(s.value, {
    x: cx, y: 1.8, w: cardW, h: 1.2,
    fontSize: 48, fontFace: "Georgia", color: SECONDARY, bold: true, align: "center", valign: "middle"
  });
  slide4.addText(s.label, {
    x: cx, y: 3.0, w: cardW, h: 0.5,
    fontSize: 14, fontFace: "Calibri", color: WHITE, bold: true, align: "center"
  });
  slide4.addText(s.sub, {
    x: cx, y: 3.5, w: cardW, h: 0.4,
    fontSize: 11, fontFace: "Calibri", color: "8899BB", align: "center"
  });
});

pres.writeFile({ fileName: "/workspace/marks_distribution.pptx" })
  .then(() => console.log("PPTX created."))
  .catch(err => console.error(err));
