const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Student Marks Distribution";

// Color palette - Ocean Gradient
const PRIMARY = "065A82";
const SECONDARY = "1C7293";
const ACCENT = "21295C";
const LIGHT = "E8F4F8";
const WHITE = "FFFFFF";
const DARK_TEXT = "1E293B";
const MUTED = "64748B";

// Student data
const students = [
  { name: "Alice Johnson", score: 92 },
  { name: "Bob Smith", score: 78 },
  { name: "Charlie Brown", score: 85 },
  { name: "Diana Ross", score: 67 },
  { name: "Ethan Hunt", score: 91 },
  { name: "Fiona Apple", score: 73 },
  { name: "George Lucas", score: 88 },
  { name: "Hannah Montana", score: 95 },
  { name: "Ivan Drago", score: 62 },
  { name: "Julia Roberts", score: 81 }
];

// Grade distribution
const gradeCount = { A: 0, B: 0, C: 0, D: 0, F: 0 };
students.forEach(s => {
  if (s.score >= 90) gradeCount.A++;
  else if (s.score >= 80) gradeCount.B++;
  else if (s.score >= 70) gradeCount.C++;
  else if (s.score >= 60) gradeCount.D++;
  else gradeCount.F++;
});

// ========== SLIDE 1: TITLE ==========
let slide1 = pres.addSlide();
slide1.background = { color: ACCENT };

// Large decorative circle
slide1.addShape(pres.shapes.OVAL, {
  x: 6.5, y: -1.5, w: 5.5, h: 5.5,
  fill: { color: PRIMARY, transparency: 40 }
});
slide1.addShape(pres.shapes.OVAL, {
  x: 7.5, y: 2.5, w: 4, h: 4,
  fill: { color: SECONDARY, transparency: 50 }
});

slide1.addText("Student Marks\nDistribution Report", {
  x: 0.8, y: 1.2, w: 6, h: 2.5,
  fontSize: 38, fontFace: "Georgia", color: WHITE, bold: true,
  lineSpacingMultiple: 1.2
});

slide1.addText("Performance Analysis & Grade Breakdown", {
  x: 0.8, y: 3.7, w: 6, h: 0.6,
  fontSize: 16, fontFace: "Calibri", color: "94B8D0", italic: true
});

const avg = (students.reduce((s, st) => s + st.score, 0) / students.length).toFixed(1);
slide1.addText([
  { text: `${students.length}`, options: { fontSize: 28, bold: true, color: WHITE } },
  { text: " Students  |  ", options: { fontSize: 14, color: "94B8D0" } },
  { text: `${avg}`, options: { fontSize: 28, bold: true, color: WHITE } },
  { text: " Avg Score", options: { fontSize: 14, color: "94B8D0" } }
], { x: 0.8, y: 4.5, w: 6, h: 0.7 });


// ========== SLIDE 2: Individual Scores Bar Chart ==========
let slide2 = pres.addSlide();
slide2.background = { color: "F8FAFB" };

slide2.addText("Individual Student Scores", {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, fontFace: "Georgia", color: ACCENT, bold: true, margin: 0
});
slide2.addText("Score out of 100 for each student", {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 13, fontFace: "Calibri", color: MUTED, margin: 0
});

// Color each bar based on grade
const barColors = students.map(s => {
  if (s.score >= 90) return "0D9488";
  if (s.score >= 80) return "2563EB";
  if (s.score >= 70) return "D97706";
  if (s.score >= 60) return "DC2626";
  return "6B7280";
});

slide2.addChart(pres.charts.BAR, [{
  name: "Score",
  labels: students.map(s => s.name.split(" ")[0]),
  values: students.map(s => s.score)
}], {
  x: 0.3, y: 1.4, w: 9.4, h: 3.8, barDir: "col",
  chartColors: barColors,
  chartArea: { fill: { color: WHITE }, roundedCorners: true },
  catAxisLabelColor: MUTED, catAxisLabelFontSize: 10,
  valAxisLabelColor: MUTED, valAxisLabelFontSize: 9,
  valGridLine: { color: "E2E8F0", size: 0.5 },
  catGridLine: { style: "none" },
  showValue: true, dataLabelPosition: "outEnd", dataLabelColor: DARK_TEXT, dataLabelFontSize: 10,
  showLegend: false,
  valAxisMinVal: 0, valAxisMaxVal: 100
});

// Legend for colors
const legendItems = [
  { label: "A (90-100)", color: "0D9488" },
  { label: "B (80-89)", color: "2563EB" },
  { label: "C (70-79)", color: "D97706" },
  { label: "D (60-69)", color: "DC2626" }
];
legendItems.forEach((item, i) => {
  const lx = 0.5 + i * 2.3;
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: lx, y: 5.15, w: 0.25, h: 0.25, fill: { color: item.color }
  });
  slide2.addText(item.label, {
    x: lx + 0.32, y: 5.15, w: 1.8, h: 0.25,
    fontSize: 10, fontFace: "Calibri", color: MUTED, margin: 0
  });
});


// ========== SLIDE 3: Grade Distribution Pie Chart ==========
let slide3 = pres.addSlide();
slide3.background = { color: "F8FAFB" };

slide3.addText("Grade Distribution", {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, fontFace: "Georgia", color: ACCENT, bold: true, margin: 0
});
slide3.addText("Percentage of students in each grade band", {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 13, fontFace: "Calibri", color: MUTED, margin: 0
});

// Pie chart on the left
const gradeLabels = Object.keys(gradeCount);
const gradeValues = Object.values(gradeCount);

slide3.addChart(pres.charts.PIE, [{
  name: "Grades",
  labels: gradeLabels.map((g, i) => `Grade ${g}`),
  values: gradeValues
}], {
  x: 0.3, y: 1.3, w: 5, h: 4,
  chartColors: ["0D9488", "2563EB", "D97706", "DC2626", "6B7280"],
  showPercent: true,
  showLegend: true, legendPos: "b", legendFontSize: 11, legendColor: DARK_TEXT,
  dataLabelColor: WHITE, dataLabelFontSize: 14,
  chartArea: { fill: { color: WHITE }, roundedCorners: true }
});

// Stat cards on the right
const grades = ["A", "B", "C", "D", "F"];
const gradeColors = ["0D9488", "2563EB", "D97706", "DC2626", "6B7280"];
const gradeRanges = ["90–100", "80–89", "70–79", "60–69", "Below 60"];

grades.forEach((g, i) => {
  const cy = 1.4 + i * 0.78;
  slide3.addShape(pres.shapes.RECTANGLE, {
    x: 5.8, y: cy, w: 3.8, h: 0.65,
    fill: { color: WHITE },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 1, angle: 135, opacity: 0.08 }
  });
  // Color accent bar
  slide3.addShape(pres.shapes.RECTANGLE, {
    x: 5.8, y: cy, w: 0.08, h: 0.65,
    fill: { color: gradeColors[i] }
  });
  slide3.addText(`Grade ${g}`, {
    x: 6.05, y: cy, w: 1.4, h: 0.65,
    fontSize: 14, fontFace: "Calibri", color: DARK_TEXT, bold: true, valign: "middle", margin: 0
  });
  slide3.addText(gradeRanges[i], {
    x: 7.3, y: cy, w: 1.0, h: 0.65,
    fontSize: 11, fontFace: "Calibri", color: MUTED, valign: "middle", margin: 0
  });
  slide3.addText(`${gradeCount[g]}`, {
    x: 8.5, y: cy, w: 0.9, h: 0.65,
    fontSize: 22, fontFace: "Georgia", color: gradeColors[i], bold: true,
    align: "center", valign: "middle", margin: 0
  });
});


// Save
pres.writeFile({ fileName: "/workspace/marks_distribution.pptx" }).then(() => {
  console.log("PPTX created successfully");
});
