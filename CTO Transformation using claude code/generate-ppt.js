const PptxGenJS = require("pptxgenjs");

const pptx = new PptxGenJS();

// ═══════════════════════════════════════════════════
// DESIGN SYSTEM — Global Wealth Management
// ═══════════════════════════════════════════════════
const C = {
  trustBlue:    "003B70",
  crimson:      "D11242",
  slate:        "0F1632",
  slateMid:     "142040",
  cream:        "FAF8F2",
  white:        "FFFFFF",
  text:         "1A1F36",
  textLight:    "5A6178",
  green:        "0D9B5C",
  greenLight:   "EDF7F0",
  red:          "D11242",
  redLight:     "FDF0F3",
  lightBlue:    "7AB3E0",
  midBlue:      "004D99",
  amber:        "D4930B",
};

pptx.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pptx.author = "Chief AI Officer";
pptx.company = "Office of the Chief AI Officer";
pptx.subject = "Technology Transformation Strategy";
pptx.title = "The Five-Person Bank";

// ─── MASTER SLIDES ───
pptx.defineSlideMaster({
  title: "DARK",
  background: { color: C.slate },
});
pptx.defineSlideMaster({
  title: "BLUE",
  background: { color: C.trustBlue },
});
pptx.defineSlideMaster({
  title: "CREAM",
  background: { color: C.cream },
});
pptx.defineSlideMaster({
  title: "WHITE",
  background: { color: C.white },
});

// ─── HELPERS ───
function addFooter(slide, textColor = "FFFFFF") {
  slide.addText("Office of the Chief AI Officer  |  Technology Transformation Strategy  |  March 2026", {
    x: 0.5, y: 7.0, w: 12.33, h: 0.3,
    fontSize: 8, color: textColor, fontFace: "Inter", align: "left", transparency: 60,
  });
}

function horizonLine(slide, y, color = C.trustBlue) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 1.5, y: y, w: 10.33, h: 0.01,
    fill: { color: color, transparency: 80 },
  });
}

function sectionEyebrow(slide, text, x, y, color = C.crimson) {
  slide.addText(text.toUpperCase(), {
    x: x, y: y, w: 11, h: 0.3,
    fontSize: 10, fontFace: "Inter", bold: true, color: color, letterSpacing: 4, align: "left",
  });
}

function sectionTitle(slide, text, x, y, opts = {}) {
  slide.addText(text, {
    x: x, y: y, w: opts.w || 11, h: opts.h || 0.8,
    fontSize: opts.fontSize || 32, fontFace: "Inter", bold: true,
    color: opts.color || C.white, align: opts.align || "left",
    lineSpacing: opts.lineSpacing || 38,
  });
}

function bodyText(slide, text, x, y, w, opts = {}) {
  slide.addText(text, {
    x: x, y: y, w: w, h: opts.h || 1.5,
    fontSize: opts.fontSize || 13, fontFace: "Inter", color: opts.color || C.textLight,
    align: opts.align || "left", lineSpacing: opts.lineSpacing || 22, valign: "top",
    ...(opts.bold && { bold: true }),
  });
}

function statBox(slide, value, label, x, y, w, opts = {}) {
  const bgColor = opts.bg || C.white;
  const topBorder = opts.topColor || null;
  slide.addShape(pptx.ShapeType.roundRect, {
    x: x, y: y, w: w, h: opts.boxH || 1.1,
    fill: { color: bgColor }, rectRadius: 0.1,
    shadow: { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.06 },
    ...(topBorder && { line: { color: topBorder, width: 0 } }),
  });
  if (topBorder) {
    slide.addShape(pptx.ShapeType.rect, {
      x: x, y: y, w: w, h: 0.05, fill: { color: topBorder },
    });
  }
  slide.addText(value, {
    x: x, y: y + 0.15, w: w, h: 0.5,
    fontSize: opts.valueFontSize || 26, fontFace: "Inter", bold: true,
    color: opts.valueColor || C.slate, align: "center",
  });
  slide.addText(label.toUpperCase(), {
    x: x, y: y + 0.65, w: w, h: 0.3,
    fontSize: 8, fontFace: "Inter", color: opts.labelColor || C.textLight,
    align: "center", letterSpacing: 1.5,
  });
}

function darkStatBox(slide, value, label, x, y, w) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: x, y: y, w: w, h: 1.2,
    fill: { color: C.white, transparency: 94 },
    line: { color: C.white, width: 0.5, transparency: 88 },
    rectRadius: 0.1,
  });
  slide.addText(value, {
    x: x, y: y + 0.15, w: w, h: 0.5,
    fontSize: 30, fontFace: "Inter", bold: true, color: C.crimson, align: "center",
  });
  slide.addText(label.toUpperCase(), {
    x: x, y: y + 0.7, w: w, h: 0.35,
    fontSize: 8, fontFace: "Inter", color: C.lightBlue,
    align: "center", letterSpacing: 1.5, lineSpacing: 13,
  });
}

function comparisonBar(slide, label, oldVal, newVal, oldPct, newPct, x, y, barW) {
  const barH = 0.25;
  slide.addText(label, { x: x, y: y, w: 1.5, h: barH * 2 + 0.1, fontSize: 10, fontFace: "Inter", bold: true, color: C.white, align: "right", valign: "middle" });
  // Old bar
  slide.addShape(pptx.ShapeType.roundRect, { x: x + 1.65, y: y, w: barW, h: barH, fill: { color: C.white, transparency: 94 }, rectRadius: 0.04 });
  slide.addShape(pptx.ShapeType.roundRect, { x: x + 1.65, y: y, w: barW * oldPct, h: barH, fill: { color: C.red, transparency: 40 }, rectRadius: 0.04 });
  slide.addText(oldVal, { x: x + 1.65, y: y, w: barW * oldPct, h: barH, fontSize: 9, fontFace: "Inter", bold: true, color: C.white, align: "right", margin: [0, 8, 0, 0] });
  // New bar
  slide.addShape(pptx.ShapeType.roundRect, { x: x + 1.65, y: y + barH + 0.06, w: barW, h: barH, fill: { color: C.white, transparency: 94 }, rectRadius: 0.04 });
  slide.addShape(pptx.ShapeType.roundRect, { x: x + 1.65, y: y + barH + 0.06, w: Math.max(barW * newPct, 0.3), h: barH, fill: { color: C.green }, rectRadius: 0.04 });
  slide.addText(newVal, { x: x + 1.65, y: y + barH + 0.06, w: Math.max(barW * newPct, 0.3) + 0.8, h: barH, fontSize: 9, fontFace: "Inter", bold: true, color: C.white, align: "right", margin: [0, 8, 0, 0] });
}


// ═══════════════════════════════════════════════════
// SLIDE 1 — TITLE
// ═══════════════════════════════════════════════════
let slide = pptx.addSlide({ masterName: "DARK" });

// Subtle orb
slide.addShape(pptx.ShapeType.ellipse, {
  x: 8, y: -2, w: 8, h: 8,
  fill: { color: C.trustBlue, transparency: 80 },
});

slide.addText("BOARD STRATEGY PAPER", {
  x: 1, y: 1.2, w: 6, h: 0.3,
  fontSize: 10, fontFace: "Inter", bold: true, color: C.lightBlue, letterSpacing: 4,
});
slide.addText("What if your entire\ntechnology organisation\nwas", {
  x: 1, y: 1.8, w: 7, h: 2.2,
  fontSize: 42, fontFace: "Inter", bold: true, color: C.white, lineSpacing: 50,
});
slide.addText("5", {
  x: 1, y: 3.9, w: 2, h: 1.2,
  fontSize: 96, fontFace: "Inter", bold: true, color: C.crimson,
});
slide.addText("people and a file?", {
  x: 2.7, y: 4.15, w: 6, h: 0.8,
  fontSize: 42, fontFace: "Inter", bold: true, color: C.lightBlue,
});

// Equation row
const eqY = 5.6;
const roles = ["bank.md", "CPO", "Engineer", "AI Eng", "Data Arch", "CCO"];
const roleLabels = ["Enterprise DNA", "Product Vision", "Full-Stack", "Agentic & MCP", "Data & AI", "Compliance"];
roles.forEach((role, i) => {
  const xPos = 1 + i * 1.9;
  slide.addShape(pptx.ShapeType.roundRect, {
    x: xPos, y: eqY, w: 1.6, h: 0.9,
    fill: { color: C.white, transparency: 94 },
    line: { color: C.white, width: 0.5, transparency: 85 },
    rectRadius: 0.08,
  });
  slide.addText(role, {
    x: xPos, y: eqY + 0.1, w: 1.6, h: 0.4,
    fontSize: 13, fontFace: "Inter", bold: true, color: C.white, align: "center",
  });
  slide.addText(roleLabels[i].toUpperCase(), {
    x: xPos, y: eqY + 0.5, w: 1.6, h: 0.3,
    fontSize: 7, fontFace: "Inter", color: C.lightBlue, align: "center", letterSpacing: 1.5,
  });
  if (i < roles.length - 1) {
    slide.addText("+", {
      x: xPos + 1.55, y: eqY + 0.15, w: 0.4, h: 0.5,
      fontSize: 18, fontFace: "Inter", color: C.textLight, align: "center", transparency: 50,
    });
  }
});

addFooter(slide);


// ═══════════════════════════════════════════════════
// SLIDE 2 — THE BRUTAL TRUTH
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "CREAM" });
sectionEyebrow(slide, "The Brutal Truth", 1, 0.5);
sectionTitle(slide, "Your Technology Organisation Was Designed\nfor a World That No Longer Exists", 1, 0.9, { color: C.slate, fontSize: 28, h: 1, lineSpacing: 34 });
bodyText(slide, "You employ 2,000 technologists. You spend $813M a year. And a single application still takes 44 weeks to reach production.", 1, 1.9, 8, { color: C.text, fontSize: 14 });

// Pyramid representation
const pyramidData = [
  { role: "CTO", count: "1", w: 2 },
  { role: "VPs", count: "10", w: 4 },
  { role: "Managers", count: "62", w: 6 },
  { role: "Engineers & BAs", count: "1,040", w: 9 },
  { role: "Compliance, PMs, Contractors", count: "555", w: 11 },
];
pyramidData.forEach((row, i) => {
  const rw = row.w;
  const rx = (13.33 - rw) / 2;
  const ry = 2.7 + i * 0.55;
  slide.addShape(pptx.ShapeType.roundRect, {
    x: rx, y: ry, w: rw, h: 0.45,
    fill: { color: C.crimson, transparency: 10 + i * 5 },
    rectRadius: 0.06,
  });
  slide.addText(`${row.role}  —  ${row.count}`, {
    x: rx, y: ry, w: rw, h: 0.45,
    fontSize: 11, fontFace: "Inter", bold: true, color: C.white, align: "center",
  });
});

slide.addText("2,008 PEOPLE  •  $813M / YEAR", {
  x: 3, y: 5.6, w: 7.33, h: 0.4,
  fontSize: 16, fontFace: "Inter", bold: true, color: C.crimson, align: "center",
});

// Pain stats
const painStats = [
  { val: "62%", label: "Wait Time", desc: "Only 38% of effort adds value" },
  { val: "14", label: "Team Handoffs", desc: "3-5 day queue per handoff" },
  { val: "$4.2M", label: "Per Application", desc: "100 apps = $420M/year" },
  { val: "25%", label: "Time Coding", desc: "75% meetings, docs, process" },
];
painStats.forEach((s, i) => {
  statBox(slide, s.val, s.label, 1 + i * 2.9, 6.1, 2.6, { topColor: C.crimson, valueColor: C.crimson, boxH: 1.2 });
  slide.addText(s.desc, { x: 1 + i * 2.9, y: 7.0, w: 2.6, h: 0.3, fontSize: 8, fontFace: "Inter", color: C.textLight, align: "center" });
});
addFooter(slide, C.textLight);


// ═══════════════════════════════════════════════════
// SLIDE 3 — CURRENT SDLC: 8,500 HOURS
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "WHITE" });
sectionEyebrow(slide, "Current State", 1, 0.5);
sectionTitle(slide, "One Application: 8,500 Hours Across 44 Weeks", 1, 0.9, { color: C.slate, fontSize: 26, h: 0.6 });
bodyText(slide, "14 handoffs  •  6 governance gates  •  23 approval steps  •  34% rework rate", 1, 1.5, 10, { fontSize: 11, h: 0.3, color: C.textLight });

// Waterfall bars
const phases = [
  { name: "Requirements", hours: "1,200 hrs", weeks: "8 wks", start: 0, width: 0.18 },
  { name: "Architecture", hours: "800 hrs", weeks: "5 wks", start: 0.15, width: 0.12 },
  { name: "Development", hours: "2,800 hrs", weeks: "14 wks", start: 0.25, width: 0.32 },
  { name: "Testing", hours: "1,800 hrs", weeks: "9 wks", start: 0.54, width: 0.20 },
  { name: "Release", hours: "1,100 hrs", weeks: "5 wks", start: 0.72, width: 0.14 },
];
const barLeft = 2.5;
const barFullW = 9.5;
phases.forEach((p, i) => {
  const ry = 2.1 + i * 0.65;
  slide.addText(p.name, { x: 0.5, y: ry, w: 1.8, h: 0.4, fontSize: 10, fontFace: "Inter", bold: true, color: C.text, align: "right" });
  // Track
  slide.addShape(pptx.ShapeType.roundRect, { x: barLeft, y: ry + 0.05, w: barFullW, h: 0.32, fill: { color: "F5F5F3" }, rectRadius: 0.04 });
  // Bar
  slide.addShape(pptx.ShapeType.roundRect, { x: barLeft + barFullW * p.start, y: ry + 0.05, w: barFullW * p.width, h: 0.32, fill: { color: C.crimson }, rectRadius: 0.04 });
  slide.addText(`${p.hours} • ${p.weeks}`, { x: barLeft + barFullW * p.start, y: ry + 0.05, w: barFullW * p.width, h: 0.32, fontSize: 8, fontFace: "Inter", bold: true, color: C.white, align: "center" });
});
// Governance bar
const govY = 2.1 + 5 * 0.65;
slide.addText("Governance", { x: 0.5, y: govY, w: 1.8, h: 0.4, fontSize: 10, fontFace: "Inter", bold: true, color: C.text, align: "right" });
slide.addShape(pptx.ShapeType.roundRect, { x: barLeft, y: govY + 0.05, w: barFullW * 0.9, h: 0.32, fill: { color: C.crimson, transparency: 85 }, line: { color: C.crimson, width: 0.75, dashType: "dash" }, rectRadius: 0.04 });
slide.addText("800 hrs — continuous overhead", { x: barLeft, y: govY + 0.05, w: barFullW * 0.9, h: 0.32, fontSize: 8, fontFace: "Inter", bold: true, color: C.crimson, align: "center" });

// Total bar
slide.addShape(pptx.ShapeType.roundRect, { x: 1, y: 5.8, w: 11.33, h: 0.7, fill: { color: C.redLight }, line: { color: C.crimson, width: 1.5 }, rectRadius: 0.1 });
slide.addText("TOTAL DELIVERY EFFORT PER APPLICATION", { x: 1.3, y: 5.85, w: 6, h: 0.3, fontSize: 10, fontFace: "Inter", bold: true, color: C.crimson, letterSpacing: 2 });
slide.addText("62% wait time  •  34% rework rate  •  $4.2M blended cost", { x: 1.3, y: 6.15, w: 6, h: 0.25, fontSize: 9, fontFace: "Inter", color: C.textLight });
slide.addText("8,500 hrs", { x: 8, y: 5.85, w: 4, h: 0.6, fontSize: 32, fontFace: "Inter", bold: true, color: C.crimson, align: "right", margin: [0, 15, 0, 0] });
addFooter(slide, C.textLight);


// ═══════════════════════════════════════════════════
// SLIDE 4 — THE PROVOCATION
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "WHITE" });
sectionEyebrow(slide, "The Provocation", 1, 0.5);
sectionTitle(slide, "What If 95% of This Effort\nWas Never Necessary?", 1, 0.9, { color: C.slate, fontSize: 30, h: 1, lineSpacing: 36 });

bodyText(slide, "Strip away the coordination, the handoffs, the documentation theatre, the environment provisioning, the regression testing, the manual code review, the compliance evidence gathering.\n\nWhat remains is judgement. Product judgement. Technical judgement. Data judgement. Risk judgement.\n\nEverything else is execution that Claude Code can perform autonomously, at machine speed, with perfect consistency.", 1, 2.0, 5.5, { color: C.text, fontSize: 12, h: 3.5, lineSpacing: 20 });

// Right side — execution vs judgement list
const execItems = ["Writing code", "Writing tests", "Code review", "Deployment scripting", "Documentation", "Environment setup", "Compliance evidence", "Security scanning", "Regression testing", "Incident triage"];
execItems.forEach((item, i) => {
  slide.addText(item, { x: 7.5, y: 2.0 + i * 0.35, w: 3.5, h: 0.3, fontSize: 11, fontFace: "Inter", color: C.textLight });
  slide.addText("EXECUTION", { x: 10.5, y: 2.0 + i * 0.35, w: 1.5, h: 0.3, fontSize: 8, fontFace: "Inter", bold: true, color: C.crimson, align: "right" });
});
horizonLine(slide, 5.55, C.trustBlue);
const humanItems = ["Architecture & judgement", "Product & risk decisions"];
humanItems.forEach((item, i) => {
  slide.addText(item, { x: 7.5, y: 5.7 + i * 0.35, w: 3.5, h: 0.3, fontSize: 11, fontFace: "Inter", bold: true, color: C.slate });
  slide.addText("HUMAN", { x: 10.5, y: 5.7 + i * 0.35, w: 1.5, h: 0.3, fontSize: 8, fontFace: "Inter", bold: true, color: C.green, align: "right" });
});

// Big stat
slide.addText("95%", { x: 1, y: 5.5, w: 3, h: 1.2, fontSize: 64, fontFace: "Inter", bold: true, color: C.crimson });
slide.addText("of technology effort is execution,\nnot judgement", { x: 1, y: 6.4, w: 5, h: 0.6, fontSize: 12, fontFace: "Inter", color: C.textLight, lineSpacing: 18 });
addFooter(slide, C.textLight);


// ═══════════════════════════════════════════════════
// SLIDE 5 — BANK.MD
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "CREAM" });
sectionEyebrow(slide, "The Enterprise DNA", 1, 0.5);
sectionTitle(slide, "bank.md: One File to Rule Them All", 1, 0.9, { color: C.slate, fontSize: 28, h: 0.6 });
bodyText(slide, "Instead of 120 GRC analysts, 18 architects, and 340 QA engineers enforcing standards manually, a single markdown file becomes the immutable constitution of your technology estate.", 1, 1.5, 8, { fontSize: 12, h: 0.6, color: C.text });

// Code block
slide.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 2.3, w: 7, h: 4.5, fill: { color: "0D1117" }, rectRadius: 0.12 });
// Title bar
slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.3, w: 7, h: 0.35, fill: { color: "161B22" } });
slide.addShape(pptx.ShapeType.ellipse, { x: 1.0, y: 2.38, w: 0.15, h: 0.15, fill: { color: "FF5F57" } });
slide.addShape(pptx.ShapeType.ellipse, { x: 1.2, y: 2.38, w: 0.15, h: 0.15, fill: { color: "FFBD2E" } });
slide.addShape(pptx.ShapeType.ellipse, { x: 1.4, y: 2.38, w: 0.15, h: 0.15, fill: { color: "28CA41" } });
slide.addText("bank.md — Enterprise Archetype", { x: 1.7, y: 2.32, w: 4, h: 0.3, fontSize: 9, fontFace: "Courier New", color: "484F58" });

const codeLines = [
  { t: "# BANK.MD — Global Enterprise Archetype", c: C.lightBlue },
  { t: "", c: "484F58" },
  { t: "## 1. Architecture Standards", c: "5A9FD4" },
  { t: "default_pattern: event-driven-microservices", c: "C9D1D9" },
  { t: "api_standard: OpenAPI 3.1, contract-first", c: "C9D1D9" },
  { t: "max_service_size: 500 lines business logic", c: "C9D1D9" },
  { t: "", c: "484F58" },
  { t: "## 2. Security & Compliance", c: "5A9FD4" },
  { t: "RULE: No secrets in code. All via Vault.", c: "C678DD" },
  { t: "RULE: PII encrypted AES-256 at rest, TLS 1.3", c: "C678DD" },
  { t: "RULE: OWASP Top 10 scan on every commit", c: "C678DD" },
  { t: "RULE: 4-eyes principle — no self-approval", c: "C678DD" },
  { t: "", c: "484F58" },
  { t: "## 3. Testing (shift-left, no QA phase)", c: "5A9FD4" },
  { t: "MANDATORY: Unit coverage ≥ 90%", c: "E5C07B" },
  { t: "MANDATORY: Integration tests per API contract", c: "E5C07B" },
  { t: "test_generation: AT CODE CREATION, not after", c: "C9D1D9" },
  { t: "", c: "484F58" },
  { t: "## 4-6. Design, Deployment, Governance ...", c: "5A9FD4" },
];
codeLines.forEach((line, i) => {
  slide.addText(line.t, {
    x: 1.1, y: 2.75 + i * 0.2, w: 6.4, h: 0.2,
    fontSize: 8.5, fontFace: "Courier New", color: line.c,
  });
});

// Right side — what it replaces
const replaces = [
  { icon: "Architecture", replaces: "18 architects + ARB committees", desc: "Patterns auto-enforced. Deviations impossible." },
  { icon: "Security", replaces: "Manual security reviews + pen test queues", desc: "Controls real-time, not post-hoc." },
  { icon: "Testing", replaces: "340 QA engineers + 1,800 hrs/app", desc: "Tests generated with code. QA phase eliminated." },
  { icon: "Governance", replaces: "120 compliance analysts", desc: "Every action is an audit event. Zero manual effort." },
];
replaces.forEach((r, i) => {
  const ry = 2.3 + i * 1.15;
  slide.addShape(pptx.ShapeType.roundRect, { x: 8.2, y: ry, w: 4.5, h: 1.0, fill: { color: C.white }, rectRadius: 0.08, shadow: { type: "outer", blur: 4, offset: 1, color: "000000", opacity: 0.04 } });
  slide.addShape(pptx.ShapeType.rect, { x: 8.2, y: ry, w: 4.5, h: 0.04, fill: { color: C.crimson } });
  slide.addText(r.icon.toUpperCase(), { x: 8.4, y: ry + 0.1, w: 4.1, h: 0.25, fontSize: 9, fontFace: "Inter", bold: true, color: C.slate, letterSpacing: 1.5 });
  slide.addText(`Replaces: ${r.replaces}`, { x: 8.4, y: ry + 0.35, w: 4.1, h: 0.25, fontSize: 9, fontFace: "Inter", bold: true, color: C.crimson });
  slide.addText(r.desc, { x: 8.4, y: ry + 0.6, w: 4.1, h: 0.3, fontSize: 9, fontFace: "Inter", color: C.textLight });
});
addFooter(slide, C.textLight);


// ═══════════════════════════════════════════════════
// SLIDE 6 — THE FIVE
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "WHITE" });
sectionEyebrow(slide, "The Five", 1, 0.5);
sectionTitle(slide, "Five Humans. Infinite Capability.", 1, 0.9, { color: C.slate, fontSize: 30, h: 0.6 });

const fiveRoles = [
  { num: "01", title: "CPO", sub: "Chief Product Officer", owns: "Product vision, roadmap,\nstakeholder alignment,\nfeature trade-offs", claude: "User stories, regulatory mapping,\ncompetitive analysis, PRD generation" },
  { num: "02", title: "Full-Stack\nEngineer", sub: "Master Builder", owns: "Architecture evolution,\ncode review & approval,\nedge-case engineering", claude: "All code generation, testing,\nCI/CD, IaC, deployment,\nincident remediation" },
  { num: "03", title: "AI\nEngineer", sub: "Agentic & MCP", owns: "MCP servers, agent roles,\nguardrails & permissions,\nprompt engineering", claude: "Executing agentic workflows,\nmulti-step orchestration,\nautonomous operations" },
  { num: "04", title: "Data\nArchitect", sub: "Data & AI", owns: "Data model, lineage,\nresidency rules, ML\ngovernance, quality", claude: "Schema gen, ETL pipelines,\nML training, dashboards,\ndata documentation" },
  { num: "05", title: "CCO", sub: "Chief Compliance Officer", owns: "Regulatory policy rules,\nTier-1 approvals, risk\nframework, audit coord.", claude: "Evidence generation, SOX\nmonitoring, change detection,\ncontrol testing" },
];
fiveRoles.forEach((r, i) => {
  const cx = 0.5 + i * 2.5;
  // Header
  slide.addShape(pptx.ShapeType.roundRect, { x: cx, y: 1.7, w: 2.3, h: 5.0, fill: { color: C.white }, rectRadius: 0.1, shadow: { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.06 } });
  slide.addShape(pptx.ShapeType.rect, { x: cx, y: 1.7, w: 2.3, h: 1.1, fill: { color: C.slate } });
  // Round top corners manually
  slide.addShape(pptx.ShapeType.roundRect, { x: cx, y: 1.7, w: 2.3, h: 1.2, fill: { color: C.slate }, rectRadius: 0.1 });
  slide.addShape(pptx.ShapeType.rect, { x: cx, y: 2.3, w: 2.3, h: 0.6, fill: { color: C.slate } });
  slide.addText(r.num, { x: cx, y: 1.75, w: 2.3, h: 0.45, fontSize: 28, fontFace: "Inter", bold: true, color: C.crimson, align: "center" });
  slide.addText(r.title, { x: cx, y: 2.15, w: 2.3, h: 0.55, fontSize: 11, fontFace: "Inter", bold: true, color: C.white, align: "center", lineSpacing: 14 });
  // Owns
  slide.addText("OWNS", { x: cx + 0.15, y: 2.95, w: 2, h: 0.2, fontSize: 8, fontFace: "Inter", bold: true, color: C.crimson, letterSpacing: 2 });
  slide.addText(r.owns, { x: cx + 0.15, y: 3.15, w: 2, h: 1.2, fontSize: 8.5, fontFace: "Inter", color: C.text, lineSpacing: 13, valign: "top" });
  // Claude handles
  slide.addShape(pptx.ShapeType.roundRect, { x: cx + 0.1, y: 4.5, w: 2.1, h: 1.2, fill: { color: C.redLight }, rectRadius: 0.06 });
  slide.addText("CLAUDE HANDLES", { x: cx + 0.2, y: 4.55, w: 1.9, h: 0.2, fontSize: 7, fontFace: "Inter", bold: true, color: C.crimson, letterSpacing: 1.5 });
  slide.addText(r.claude, { x: cx + 0.2, y: 4.75, w: 1.9, h: 0.9, fontSize: 8, fontFace: "Inter", color: C.textLight, lineSpacing: 12, valign: "top" });
});

// Bottom comparison strip
slide.addShape(pptx.ShapeType.roundRect, { x: 1.5, y: 5.95, w: 4.5, h: 0.7, fill: { color: C.redLight }, line: { color: C.crimson, width: 1.5 }, rectRadius: 0.1 });
slide.addText("2,008", { x: 1.7, y: 5.95, w: 2, h: 0.7, fontSize: 28, fontFace: "Inter", bold: true, color: C.crimson });
slide.addText("PEOPLE TODAY\n$813M / YEAR", { x: 3.2, y: 5.95, w: 2.5, h: 0.7, fontSize: 9, fontFace: "Inter", color: C.crimson, lineSpacing: 14 });

slide.addText("→", { x: 6.2, y: 5.95, w: 0.8, h: 0.7, fontSize: 24, fontFace: "Inter", color: C.textLight, align: "center" });

slide.addShape(pptx.ShapeType.roundRect, { x: 7.2, y: 5.95, w: 4.5, h: 0.7, fill: { color: C.greenLight }, line: { color: C.green, width: 1.5 }, rectRadius: 0.1 });
slide.addText("5", { x: 7.4, y: 5.95, w: 1.5, h: 0.7, fontSize: 28, fontFace: "Inter", bold: true, color: C.green });
slide.addText("PEOPLE + CLAUDE CODE\n$9M / YEAR", { x: 8.5, y: 5.95, w: 3, h: 0.7, fontSize: 9, fontFace: "Inter", color: C.green, lineSpacing: 14 });
addFooter(slide, C.textLight);


// ═══════════════════════════════════════════════════
// SLIDE 7 — NEW SDLC: 72 HOURS
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "CREAM" });
sectionEyebrow(slide, "The New SDLC", 1, 0.5);
sectionTitle(slide, "From 8,500 Hours to 72 Hours", 1, 0.9, { color: C.slate, fontSize: 30, h: 0.6 });

const newSteps = [
  { num: "1", name: "Intent", time: "2h", who: "CPO + Engineer", desc: "Natural language\nrequirements" },
  { num: "2", name: "Generate", time: "8h", who: "Claude Code", desc: "Code, tests, IaC,\ndocs, compliance" },
  { num: "3", name: "Review", time: "6h", who: "Eng + Data + CCO", desc: "Human judgement\non Claude output" },
  { num: "4", name: "Iterate", time: "4h", who: "Claude (directed)", desc: "Feedback loops\nin minutes" },
  { num: "5", name: "Deploy", time: "2h", who: "Claude + CCO", desc: "Canary deploy,\nCCO approves" },
];
newSteps.forEach((s, i) => {
  const sx = 0.8 + i * 2.45;
  slide.addShape(pptx.ShapeType.roundRect, { x: sx, y: 1.7, w: 2.2, h: 2.8, fill: { color: C.white }, line: { color: C.green, width: 1 }, rectRadius: 0.1, shadow: { type: "outer", blur: 4, offset: 1, color: "000000", opacity: 0.04 } });
  // Number circle
  slide.addShape(pptx.ShapeType.ellipse, { x: sx + 0.85, y: 1.55, w: 0.4, h: 0.4, fill: { color: C.green } });
  slide.addText(s.num, { x: sx + 0.85, y: 1.55, w: 0.4, h: 0.4, fontSize: 14, fontFace: "Inter", bold: true, color: C.white, align: "center" });
  slide.addText(s.name.toUpperCase(), { x: sx, y: 2.1, w: 2.2, h: 0.3, fontSize: 10, fontFace: "Inter", bold: true, color: C.slate, align: "center", letterSpacing: 1.5 });
  slide.addText(s.time, { x: sx, y: 2.4, w: 2.2, h: 0.5, fontSize: 28, fontFace: "Inter", bold: true, color: C.green, align: "center" });
  slide.addText(s.who, { x: sx, y: 2.9, w: 2.2, h: 0.25, fontSize: 9, fontFace: "Inter", bold: true, color: C.crimson, align: "center" });
  slide.addText(s.desc, { x: sx + 0.15, y: 3.2, w: 1.9, h: 0.8, fontSize: 9, fontFace: "Inter", color: C.textLight, align: "center", lineSpacing: 14 });
  // Arrow
  if (i < newSteps.length - 1) {
    slide.addText("→", { x: sx + 2.1, y: 2.6, w: 0.5, h: 0.4, fontSize: 18, fontFace: "Inter", color: C.green, align: "center", transparency: 40 });
  }
});

// Total
slide.addShape(pptx.ShapeType.roundRect, { x: 1, y: 4.8, w: 11.33, h: 0.65, fill: { color: C.greenLight }, line: { color: C.green, width: 1.5 }, rectRadius: 0.1 });
slide.addText("TOTAL: INTENT TO PRODUCTION", { x: 1.3, y: 4.85, w: 5, h: 0.25, fontSize: 10, fontFace: "Inter", bold: true, color: C.green, letterSpacing: 2 });
slide.addText("Same application  •  3 calendar days  •  5 humans involved", { x: 1.3, y: 5.1, w: 5, h: 0.2, fontSize: 9, fontFace: "Inter", color: C.textLight });
slide.addText("~22 hrs human + ~50 hrs Claude", { x: 7, y: 4.85, w: 5, h: 0.55, fontSize: 20, fontFace: "Inter", bold: true, color: C.green, align: "right", margin: [0, 15, 0, 0] });

// Testing death callout
slide.addText('"Where did Testing go?" — It didn\'t go anywhere. It\'s embedded in Step 2.\nTests are generated with code, run in CI in seconds, and gate every merge. 1,800 hours → 0 separate hours.', {
  x: 1, y: 5.7, w: 11.33, h: 0.8,
  fontSize: 11, fontFace: "Inter", color: C.text, italic: true, align: "center", lineSpacing: 18,
});

// Comparison boxes
slide.addShape(pptx.ShapeType.roundRect, { x: 1, y: 6.6, w: 5.5, h: 0.7, fill: { color: C.redLight }, rectRadius: 0.08 });
slide.addText("TODAY:  32 people  •  8,500 hrs  •  44 weeks  •  $4.2M  •  18 defects", {
  x: 1.2, y: 6.65, w: 5.1, h: 0.6, fontSize: 10, fontFace: "Inter", bold: true, color: C.crimson, lineSpacing: 16,
});

slide.addShape(pptx.ShapeType.roundRect, { x: 6.8, y: 6.6, w: 5.5, h: 0.7, fill: { color: C.greenLight }, rectRadius: 0.08 });
slide.addText("FUTURE:  5 people  •  22 hrs  •  3 days  •  ~$12K  •  ~0 defects", {
  x: 7.0, y: 6.65, w: 5.1, h: 0.6, fontSize: 10, fontFace: "Inter", bold: true, color: C.green, lineSpacing: 16,
});
addFooter(slide, C.textLight);


// ═══════════════════════════════════════════════════
// SLIDE 8 — DEMO: WEALTH CRM IN 3 DAYS (Overview)
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "DARK" });
sectionEyebrow(slide, "Live Demonstration", 0.8, 0.5, C.lightBlue);
sectionTitle(slide, "A Wealth Advisory CRM\nBuilt in 3 Days by 5 People", 0.8, 0.9, { fontSize: 34, h: 1.2, lineSpacing: 42 });
bodyText(slide, "Client 360  •  AI recommendations  •  Advisor dashboard  •  MiFID II compliance  •  Client portal\n2,500 HNW clients  •  180 relationship managers  •  $127B AUM", 0.8, 2.2, 8, { color: C.lightBlue, fontSize: 13, h: 0.8 });

const demoActs = [
  { act: "1", title: "Foundation", desc: "Load bank.md as CLAUDE.md", time: "09:00", who: "Engineer" },
  { act: "2", title: "AI Nervous System", desc: "8 MCP servers, 6 agent roles, guardrails", time: "09:05", who: "AI Engineer" },
  { act: "3", title: "Requirements & UI", desc: "47 user stories, 68 React components", time: "10:30", who: "CPO" },
  { act: "4", title: "Database & Data", desc: "24 tables, 287K synthetic transactions", time: "13:30", who: "Claude" },
  { act: "5", title: "Deploy to AWS", desc: "Terraform, ECS, RDS, CloudFront", time: "15:30", who: "Engineer" },
  { act: "6", title: "Production Build", desc: "12 services, 3 ML models, 2,847 tests", time: "Day 2", who: "Data + Eng" },
  { act: "7", title: "CCO Compliance", desc: "18/18 controls pass, 142-pg evidence", time: "Day 3", who: "CCO" },
];
demoActs.forEach((a, i) => {
  const row = Math.floor(i / 4);
  const col = i % 4;
  const ax = 0.8 + col * 3.05;
  const ay = 3.3 + row * 1.8;
  slide.addShape(pptx.ShapeType.roundRect, { x: ax, y: ay, w: 2.8, h: 1.5, fill: { color: C.white, transparency: 94 }, line: { color: C.white, width: 0.5, transparency: 85 }, rectRadius: 0.1 });
  slide.addText(`ACT ${a.act}`, { x: ax + 0.15, y: ay + 0.1, w: 1, h: 0.25, fontSize: 8, fontFace: "Inter", bold: true, color: C.crimson, letterSpacing: 2 });
  slide.addText(a.time, { x: ax + 1.5, y: ay + 0.1, w: 1.1, h: 0.25, fontSize: 8, fontFace: "Inter", color: C.lightBlue, align: "right" });
  slide.addText(a.title, { x: ax + 0.15, y: ay + 0.4, w: 2.5, h: 0.3, fontSize: 13, fontFace: "Inter", bold: true, color: C.white });
  slide.addText(a.desc, { x: ax + 0.15, y: ay + 0.7, w: 2.5, h: 0.35, fontSize: 9, fontFace: "Inter", color: C.lightBlue, lineSpacing: 13 });
  slide.addText(a.who, { x: ax + 0.15, y: ay + 1.1, w: 2.5, h: 0.25, fontSize: 8, fontFace: "Inter", bold: true, color: C.lightBlue, transparency: 40 });
});
addFooter(slide);


// ═══════════════════════════════════════════════════
// SLIDE 9 — DEMO: AI ENGINEER + MCP
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "WHITE" });
sectionEyebrow(slide, "Act 2: AI Engineer Wires the Nervous System", 1, 0.4);
sectionTitle(slide, "MCP Servers, Agentic Workflows & Guardrails", 1, 0.8, { color: C.slate, fontSize: 24, h: 0.5 });

// MCP Servers
slide.addText("8 MCP SERVERS — 47 TOOLS", { x: 1, y: 1.5, w: 5, h: 0.3, fontSize: 10, fontFace: "Inter", bold: true, color: C.trustBlue, letterSpacing: 2 });
const mcpServers = [
  { name: "SAP FICO", tools: "get_client, get_balances", auth: "Read-only" },
  { name: "Bloomberg B-PIPE", tools: "get_quotes, stream_prices", auth: "Read-only" },
  { name: "Refinitiv KYC", tools: "screen_client, check_pep", auth: "Read-only" },
  { name: "AWS Infra", tools: "deploy_service, scale", auth: "Read-write" },
  { name: "PostgreSQL", tools: "query, migrate, seed", auth: "Read-write" },
  { name: "Auth Service", tools: "tokens, RBAC", auth: "Read-write" },
  { name: "Jira/Confluence", tools: "create_ticket, docs", auth: "Read-write" },
  { name: "Compliance Eng.", tools: "get_rules, validate", auth: "Read-only" },
];
mcpServers.forEach((s, i) => {
  const row = Math.floor(i / 2);
  const col = i % 2;
  const mx = 1 + col * 2.8;
  const my = 1.9 + row * 0.5;
  slide.addText(s.name, { x: mx, y: my, w: 1.3, h: 0.35, fontSize: 9, fontFace: "Inter", bold: true, color: C.slate });
  slide.addText(s.tools, { x: mx + 1.3, y: my, w: 1.6, h: 0.35, fontSize: 8, fontFace: "Courier New", color: C.textLight });
});

// Agent roles
slide.addText("6 AGENT ROLES", { x: 7, y: 1.5, w: 5, h: 0.3, fontSize: 10, fontFace: "Inter", bold: true, color: C.trustBlue, letterSpacing: 2 });
const agentRoles = [
  { name: "code-gen", perm: "write/edit code, read MCP" },
  { name: "test-gen", perm: "write tests, run suites" },
  { name: "reviewer", perm: "read-only code, flag issues" },
  { name: "deployer", perm: "trigger deploy (human OK)" },
  { name: "compliance", perm: "read-only, run audits" },
  { name: "incident-bot", perm: "read logs, draft RCA" },
];
agentRoles.forEach((a, i) => {
  const ay = 1.9 + i * 0.4;
  slide.addText(a.name, { x: 7, y: ay, w: 1.5, h: 0.3, fontSize: 9, fontFace: "Courier New", bold: true, color: C.slate });
  slide.addText("→  " + a.perm, { x: 8.5, y: ay, w: 3.5, h: 0.3, fontSize: 9, fontFace: "Inter", color: C.textLight });
});

// Hooks & Guardrails
slide.addText("3 AUTOMATED HOOKS", { x: 1, y: 4.3, w: 5, h: 0.3, fontSize: 10, fontFace: "Inter", bold: true, color: C.trustBlue, letterSpacing: 2 });
const hooks = [
  "pre-commit: SAST scan + CLAUDE.md policy check",
  "pre-merge: full test suite + compliance gate",
  "post-deploy: smoke test + canary health check",
];
hooks.forEach((h, i) => {
  slide.addText("✓  " + h, { x: 1, y: 4.65 + i * 0.35, w: 5.5, h: 0.3, fontSize: 10, fontFace: "Inter", color: C.green });
});

slide.addText("5 HARD GUARDRAILS", { x: 7, y: 4.3, w: 5, h: 0.3, fontSize: 10, fontFace: "Inter", bold: true, color: C.crimson, letterSpacing: 2 });
const guardrails = [
  "MAX_FILES_PER_SESSION: 50",
  "BLOCKED: direct production DB writes",
  "BLOCKED: secret access without CCO approval",
  "BLOCKED: deployment without 4-eyes sign-off",
  "REQUIRE_HUMAN: any Tier-1 system change",
];
guardrails.forEach((g, i) => {
  slide.addText("⛔  " + g, { x: 7, y: 4.65 + i * 0.35, w: 5.5, h: 0.3, fontSize: 10, fontFace: "Inter", color: C.crimson });
});

// Summary bar
slide.addShape(pptx.ShapeType.roundRect, { x: 1, y: 6.3, w: 11.33, h: 0.6, fill: { color: "7C3AED" }, rectRadius: 0.1 });
slide.addText("AI ENGINEER: NERVOUS SYSTEM LIVE  —  8 MCP servers (47 tools)  •  6 agent roles  •  3 hooks  •  5 guardrails  •  Full observability", {
  x: 1.3, y: 6.3, w: 10.73, h: 0.6, fontSize: 10, fontFace: "Inter", bold: true, color: C.white, align: "center",
});
addFooter(slide, C.textLight);


// ═══════════════════════════════════════════════════
// SLIDE 10 — DEMO: RESULTS
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "DARK" });
sectionEyebrow(slide, "Demo Results", 0.8, 0.5, C.lightBlue);
sectionTitle(slide, "Production-Grade Wealth Advisory CRM.\n3 Days. 5 People. Zero Defects.", 0.8, 0.9, { fontSize: 30, h: 1.1, lineSpacing: 38 });

const demoStats = [
  { val: "3", label: "Days" },
  { val: "5", label: "People" },
  { val: "18/18", label: "Controls\nPassed" },
  { val: "2,847", label: "Automated\nTests" },
  { val: "~$12K", label: "Total\nCost" },
];
demoStats.forEach((s, i) => {
  darkStatBox(slide, s.val, s.label, 0.8 + i * 2.45, 2.3, 2.2);
});

// What was built / what it replaces
slide.addText("WHAT WAS BUILT", { x: 0.8, y: 3.9, w: 5.5, h: 0.3, fontSize: 9, fontFace: "Inter", bold: true, color: C.lightBlue, letterSpacing: 2 });
const built = "68 UI components  •  24-table encrypted database  •  12 microservices\n8 data pipelines  •  3 ML models  •  Kafka event streaming\nOAuth 2.0 + RBAC  •  Full observability  •  DR tested  •  Canary deploy\n142-page compliance evidence package";
slide.addText(built, { x: 0.8, y: 4.25, w: 5.5, h: 1.5, fontSize: 10, fontFace: "Inter", color: C.lightBlue, lineSpacing: 17, transparency: 30 });

slide.addText("WHAT THIS REPLACES", { x: 7, y: 3.9, w: 5.5, h: 0.3, fontSize: 9, fontFace: "Inter", bold: true, color: C.lightBlue, letterSpacing: 2 });
const replaces2 = "32-person team  •  8,500 person-hours  •  44-week timeline\n$4.2M budget  •  14 handoffs  •  6 governance gates\n23 approval steps  •  1,800 hours of manual testing\n800 hours of compliance documentation";
slide.addText(replaces2, { x: 7, y: 4.25, w: 5.5, h: 1.5, fontSize: 10, fontFace: "Inter", color: C.lightBlue, lineSpacing: 17, transparency: 30 });

horizonLine(slide, 5.9, C.lightBlue);
slide.addText('"This is not a proof of concept. This is proof that the old model is obsolete."', {
  x: 1.5, y: 6.1, w: 10.33, h: 0.5,
  fontSize: 14, fontFace: "Inter", italic: true, color: C.white, align: "center", transparency: 30,
});
addFooter(slide);


// ═══════════════════════════════════════════════════
// SLIDE 11 — CONTROLS & GOVERNANCE
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "DARK" });
sectionEyebrow(slide, "Controls & Governance", 0.8, 0.5, C.lightBlue);
sectionTitle(slide, "More Control, Not Less", 0.8, 0.9, { fontSize: 32, h: 0.6 });
bodyText(slide, 'Your regulator\'s first question: "how do you maintain control with five people?"\nThe answer: better than you do with two thousand.', 0.8, 1.5, 8, { color: C.lightBlue, fontSize: 12, h: 0.7, lineSpacing: 20 });

// Before / After table
const ctrlRows = [
  { metric: "When controls run", old: "Post-hoc (gates)", new: "Real-time (embedded)" },
  { metric: "Coverage", old: "Sampling (5-10%)", new: "100% of every change" },
  { metric: "Evidence quality", old: "Manual, inconsistent", new: "Machine-generated, immutable" },
  { metric: "Time to detect", old: "Weeks to months", new: "Milliseconds (prevented)" },
  { metric: "Audit preparation", old: "6-8 weeks, $2M", new: "Instant (always ready)" },
  { metric: "Policy enforcement", old: "Advisory (honour system)", new: "Structural (cannot violate)" },
];
// Table header
slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.5, w: 11.73, h: 0.4, fill: { color: C.white, transparency: 92 } });
slide.addText("METRIC", { x: 0.95, y: 2.5, w: 3, h: 0.4, fontSize: 8, fontFace: "Inter", bold: true, color: C.lightBlue, letterSpacing: 2 });
slide.addText("TODAY", { x: 4.5, y: 2.5, w: 3.5, h: 0.4, fontSize: 8, fontFace: "Inter", bold: true, color: C.crimson, letterSpacing: 2, align: "center" });
slide.addText("WITH BANK.MD + CLAUDE", { x: 8.3, y: 2.5, w: 4, h: 0.4, fontSize: 8, fontFace: "Inter", bold: true, color: C.green, letterSpacing: 2, align: "center" });

ctrlRows.forEach((r, i) => {
  const ry = 2.95 + i * 0.45;
  if (i % 2 === 0) {
    slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: ry, w: 11.73, h: 0.42, fill: { color: C.white, transparency: 95 } });
  }
  slide.addText(r.metric, { x: 0.95, y: ry, w: 3.3, h: 0.42, fontSize: 10, fontFace: "Inter", bold: true, color: C.white });
  slide.addText(r.old, { x: 4.5, y: ry, w: 3.5, h: 0.42, fontSize: 10, fontFace: "Inter", color: C.crimson, align: "center" });
  slide.addText(r.new, { x: 8.3, y: ry, w: 4, h: 0.42, fontSize: 10, fontFace: "Inter", bold: true, color: C.green, align: "center" });
});

// Control cards
const ctrlCards = [
  { title: "Immutable Audit Trail", desc: "Every action logged with full provenance", badge: "SOX / Basel III" },
  { title: "Structural Enforcement", desc: "bank.md rules are hard constraints", badge: "Preventive" },
  { title: "4-Eyes with 5 People", desc: "Human approval on all production changes", badge: "Human-in-Loop" },
  { title: "100% Security Coverage", desc: "OWASP scan at generation time, not after", badge: "Shift-Left" },
  { title: "Data Sovereignty", desc: "Claude runs in your private cloud", badge: "GDPR" },
  { title: "Reg Change Detection", desc: "Auto-flag when new rules impact code", badge: "Future-Proof" },
];
ctrlCards.forEach((c, i) => {
  const row = Math.floor(i / 3);
  const col = i % 3;
  const cx = 0.8 + col * 4;
  const cy = 5.7 + row * 0.9;
  slide.addShape(pptx.ShapeType.roundRect, { x: cx, y: cy, w: 3.7, h: 0.75, fill: { color: C.white, transparency: 94 }, line: { color: C.white, width: 0.5, transparency: 88 }, rectRadius: 0.08 });
  slide.addText(c.title, { x: cx + 0.15, y: cy + 0.05, w: 2.4, h: 0.25, fontSize: 10, fontFace: "Inter", bold: true, color: C.white });
  slide.addText(c.desc, { x: cx + 0.15, y: cy + 0.3, w: 2.6, h: 0.25, fontSize: 8, fontFace: "Inter", color: C.lightBlue, transparency: 30 });
  slide.addText(c.badge.toUpperCase(), { x: cx + 2.8, y: cy + 0.05, w: 0.75, h: 0.25, fontSize: 6, fontFace: "Inter", bold: true, color: C.green, align: "right", letterSpacing: 1 });
});
addFooter(slide);


// ═══════════════════════════════════════════════════
// SLIDE 12 — ROI / FINANCIAL IMPACT
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "DARK" });
sectionEyebrow(slide, "Financial Impact", 0.8, 0.5, C.lightBlue);
sectionTitle(slide, "The Numbers That Will Define\nYour Legacy", 0.8, 0.9, { fontSize: 30, h: 1, lineSpacing: 38 });

// Hero stats
const roiStats = [
  { val: "$804M", label: "Annual Cost\nEliminated" },
  { val: "99%", label: "Tech Org Cost\nReduction" },
  { val: "350x", label: "Delivery Speed\nImprovement" },
  { val: "12 wk", label: "Full Transition\nTimeline" },
];
roiStats.forEach((s, i) => {
  darkStatBox(slide, s.val, s.label, 0.8 + i * 3.05, 2.1, 2.8);
});

// Comparison bars
comparisonBar(slide, "Tech Spend", "$813M", "$9M", 1.0, 0.015, 0.3, 3.7, 9.5);
comparisonBar(slide, "Time to Deploy", "44 weeks", "3 days", 1.0, 0.01, 0.3, 4.4, 9.5);
comparisonBar(slide, "Apps / Year", "~100", "~1,000+", 0.1, 1.0, 0.3, 5.1, 9.5);
comparisonBar(slide, "Defects / App", "~18", "~0", 0.9, 0.02, 0.3, 5.8, 9.5);

// Bottom
slide.addShape(pptx.ShapeType.roundRect, { x: 3, y: 6.5, w: 7.33, h: 0.5, fill: { color: C.green, transparency: 85 }, line: { color: C.green, width: 1 }, rectRadius: 0.08 });
slide.addText("$804M redirected to growth, M&A, customer experience, or returned to shareholders", {
  x: 3, y: 6.5, w: 7.33, h: 0.5, fontSize: 10, fontFace: "Inter", color: C.green, align: "center",
});
addFooter(slide);


// ═══════════════════════════════════════════════════
// SLIDE 13 — ROADMAP
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "WHITE" });
sectionEyebrow(slide, "Implementation", 1, 0.5);
sectionTitle(slide, "Twelve Weeks to Transformation", 1, 0.9, { color: C.slate, fontSize: 28, h: 0.5 });

const roadmap = [
  { phase: "1", time: "Weeks 1-2", title: "Write bank.md, Assemble The Five", items: "Appoint CPO, Engineer, AI Engineer, Data Architect, CCO\nAuthor bank.md v1.0  •  Deploy Claude Code  •  Select 3 pilot apps" },
  { phase: "2", time: "Weeks 3-4", title: "Deliver 3 Applications in Days", items: "Intent → Generate → Review → Iterate → Deploy cycle\nDemo to board and regulators  •  Refine bank.md" },
  { phase: "3", time: "Weeks 5-8", title: "Migrate Portfolio, Transition Workforce", items: "Systematic migration of existing applications\n15-20 apps migrated per week  •  Workforce transition programme" },
  { phase: "4", time: "Weeks 9-12", title: "The Five-Person Org Is Live", items: "Full portfolio under bank.md governance\nAnnual capacity: 1,000+ applications  •  Zero backlog  •  $9M/year" },
];
roadmap.forEach((r, i) => {
  const ry = 1.7 + i * 1.35;
  // Circle
  slide.addShape(pptx.ShapeType.ellipse, { x: 1, y: ry, w: 0.55, h: 0.55, fill: { color: C.crimson } });
  slide.addText(r.phase, { x: 1, y: ry, w: 0.55, h: 0.55, fontSize: 18, fontFace: "Inter", bold: true, color: C.white, align: "center" });
  // Connector line
  if (i < roadmap.length - 1) {
    slide.addShape(pptx.ShapeType.rect, { x: 1.255, y: ry + 0.55, w: 0.04, h: 0.8, fill: { color: C.textLight, transparency: 75 } });
  }
  // Content
  slide.addText(r.time.toUpperCase(), { x: 1.8, y: ry - 0.05, w: 3, h: 0.25, fontSize: 9, fontFace: "Inter", bold: true, color: C.crimson, letterSpacing: 2 });
  slide.addText(r.title, { x: 1.8, y: ry + 0.2, w: 5, h: 0.3, fontSize: 14, fontFace: "Inter", bold: true, color: C.slate });
  slide.addText(r.items, { x: 1.8, y: ry + 0.5, w: 10, h: 0.7, fontSize: 10, fontFace: "Inter", color: C.textLight, lineSpacing: 16 });
});
addFooter(slide, C.textLight);


// ═══════════════════════════════════════════════════
// SLIDE 14 — CASE STUDY: KLARNA
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "CREAM" });
sectionEyebrow(slide, "Real-World Evidence", 1, 0.4);
sectionTitle(slide, "This Is Not Theory. It Is Already Happening.", 1, 0.8, { color: C.slate, fontSize: 26, h: 0.5 });

// Klarna hero
slide.addShape(pptx.ShapeType.roundRect, { x: 1, y: 1.5, w: 7, h: 3.2, fill: { color: C.white }, rectRadius: 0.12, shadow: { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.06 } });
slide.addText("PRIMARY CASE STUDY", { x: 1.3, y: 1.6, w: 2, h: 0.25, fontSize: 8, fontFace: "Inter", bold: true, color: C.crimson, letterSpacing: 2 });
slide.addText("Klarna", { x: 1.3, y: 1.85, w: 6, h: 0.4, fontSize: 24, fontFace: "Inter", bold: true, color: C.slate });
slide.addText("$87B fintech  •  150M consumers  •  45 countries", { x: 1.3, y: 2.2, w: 6, h: 0.25, fontSize: 10, fontFace: "Inter", color: C.textLight });

bodyText(slide, "Klarna reduced from 5,000 to ~3,500 employees through natural attrition — not mass layoffs — with a target of 2,000. Revenue per employee increased 73%. They replaced Salesforce, Workday, and 12 marketing agencies with AI-built alternatives. Their AI assistant handled 2.3M conversations in month one — equivalent to 700 agents. Resolution time dropped from 11 to 2 minutes.\n\nIn July 2025, Klarna completed its IPO at $87B — the largest European tech IPO in a decade. The market explicitly valued the AI-native model.", 1.3, 2.5, 6.4, { fontSize: 11, h: 2.0, lineSpacing: 17, color: C.text });

// Right side — quote + metrics
slide.addShape(pptx.ShapeType.roundRect, { x: 8.3, y: 1.5, w: 4.3, h: 1.5, fill: { color: C.white }, rectRadius: 0.12, shadow: { type: "outer", blur: 4, offset: 1, color: "000000", opacity: 0.04 } });
slide.addShape(pptx.ShapeType.rect, { x: 8.3, y: 1.5, w: 0.06, h: 1.5, fill: { color: C.crimson } });
slide.addText('"AI can already do all of the jobs that we as humans do. We have basically stopped hiring. The AI is doing the job."', {
  x: 8.6, y: 1.55, w: 3.8, h: 1.0, fontSize: 11, fontFace: "Inter", italic: true, color: C.slate, lineSpacing: 17,
});
slide.addText("Sebastian Siemiatkowski, CEO — Bloomberg, 2024", {
  x: 8.6, y: 2.55, w: 3.8, h: 0.3, fontSize: 8, fontFace: "Inter", color: C.textLight,
});

// Klarna metrics
const klarnaMetrics = [
  { val: "5K→2K", label: "Headcount\nTarget" },
  { val: "+73%", label: "Rev / Employee" },
  { val: "700", label: "Agents\nReplaced" },
  { val: "$87B", label: "IPO\nValuation" },
];
klarnaMetrics.forEach((m, i) => {
  statBox(slide, m.val, m.label, 8.3 + i * 1.1, 3.2, 1.0, { topColor: C.crimson, valueColor: C.crimson, valueFontSize: 18, boxH: 0.95 });
});

// Other companies strip
slide.addText("CORROBORATING EVIDENCE", { x: 1, y: 5.0, w: 5, h: 0.25, fontSize: 9, fontFace: "Inter", bold: true, color: C.trustBlue, letterSpacing: 2 });

const otherCos = [
  { name: "Shopify", stat: "AI-First", desc: "Mandate: prove AI can't do it before hiring" },
  { name: "Bolt", stat: "95%", desc: "Of new code written by AI, humans review" },
  { name: "Cursor", stat: "$3.3M", desc: "Revenue per employee with ~60 people" },
  { name: "Duolingo", stat: "AI+Review", desc: "Replaced contractor workforce with AI" },
  { name: "GitHub Data", stat: "55%", desc: "Faster task completion across 77K orgs" },
  { name: "Mercado Libre", stat: "40%", desc: "Faster dev cycles with AI-native workflows" },
];
otherCos.forEach((co, i) => {
  const cx = 1 + i * 2.0;
  slide.addShape(pptx.ShapeType.roundRect, { x: cx, y: 5.35, w: 1.85, h: 1.3, fill: { color: C.white }, rectRadius: 0.08, shadow: { type: "outer", blur: 3, offset: 1, color: "000000", opacity: 0.04 } });
  slide.addText(co.name, { x: cx + 0.1, y: 5.4, w: 1.65, h: 0.25, fontSize: 10, fontFace: "Inter", bold: true, color: C.slate });
  slide.addText(co.stat, { x: cx + 0.1, y: 5.65, w: 1.65, h: 0.3, fontSize: 18, fontFace: "Inter", bold: true, color: C.crimson });
  slide.addText(co.desc, { x: cx + 0.1, y: 5.95, w: 1.65, h: 0.55, fontSize: 8, fontFace: "Inter", color: C.textLight, lineSpacing: 12 });
});

addFooter(slide, C.textLight);


// ═══════════════════════════════════════════════════
// SLIDE 15 — LESSONS FOR OUR BANK
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "WHITE" });
sectionEyebrow(slide, "Lessons for Our Bank", 1, 0.4);
sectionTitle(slide, "What the Evidence Tells Us", 1, 0.8, { color: C.slate, fontSize: 26, h: 0.5 });

const lessons = [
  { title: "1. Attrition, Not Layoffs", desc: "Klarna achieved 60% reduction through a hiring freeze and natural attrition — not mass layoffs. Our 12-week plan follows the same principle." },
  { title: "2. Commercial AI Platforms Win", desc: "None built their own models. They all adopted commercial AI (OpenAI, Copilot, Claude) and wired it in. Our bank.md + Claude Code approach follows proven practice." },
  { title: "3. The Market Rewards It", desc: "Klarna's $87B IPO proves investors see AI-native ops as structural advantage. For us: improved cost-to-income, higher ROE, re-rating opportunity." },
  { title: "4. Start with Execution", desc: "Every company started by replacing execution work while keeping strategic judgement human. Our model: 5 humans provide judgement, Claude handles execution." },
  { title: "5. Banking Has More to Gain", desc: "If a 5,000-person fintech gains 73% efficiency, a bank with 10x the process overhead, compliance burden, and coordination cost gains disproportionately more." },
  { title: "6. The Window Is Closing", desc: "Every competitor is reading the same data. The first-mover advantage in banking AI-native operations will be claimed within 12 months. We must move now." },
];
lessons.forEach((l, i) => {
  const row = Math.floor(i / 2);
  const col = i % 2;
  const lx = 1 + col * 5.8;
  const ly = 1.5 + row * 1.7;
  slide.addShape(pptx.ShapeType.roundRect, { x: lx, y: ly, w: 5.5, h: 1.45, fill: { color: C.white }, rectRadius: 0.1, shadow: { type: "outer", blur: 4, offset: 1, color: "000000", opacity: 0.04 } });
  slide.addShape(pptx.ShapeType.rect, { x: lx, y: ly, w: 0.06, h: 1.45, fill: { color: C.trustBlue } });
  slide.addText(l.title, { x: lx + 0.25, y: ly + 0.1, w: 5, h: 0.3, fontSize: 13, fontFace: "Inter", bold: true, color: C.slate });
  slide.addText(l.desc, { x: lx + 0.25, y: ly + 0.45, w: 5, h: 0.9, fontSize: 10, fontFace: "Inter", color: C.textLight, lineSpacing: 16, valign: "top" });
});

// Evidence summary bar
slide.addShape(pptx.ShapeType.roundRect, { x: 1, y: 6.6, w: 11.33, h: 0.55, fill: { color: C.slate }, rectRadius: 0.1 });
slide.addText("Klarna: 60% fewer people, 73% more revenue, $87B IPO  •  Bolt: 95% AI code  •  Cursor: $3.3M/employee  •  The model is validated.", {
  x: 1.3, y: 6.6, w: 10.73, h: 0.55, fontSize: 10, fontFace: "Inter", bold: true, color: C.white, align: "center",
});
addFooter(slide, C.textLight);


// ═══════════════════════════════════════════════════
// SLIDE 16 — RECOMMENDATION (closing)
// ═══════════════════════════════════════════════════
slide = pptx.addSlide({ masterName: "DARK" });

slide.addShape(pptx.ShapeType.ellipse, {
  x: 4, y: 0.5, w: 10, h: 10,
  fill: { color: C.trustBlue, transparency: 85 },
});

slide.addText("RECOMMENDATION", {
  x: 1.5, y: 1.5, w: 10.33, h: 0.3,
  fontSize: 10, fontFace: "Inter", bold: true, color: C.lightBlue, letterSpacing: 4, align: "center",
});

slide.addText("We have two thousand people\ndoing the work of five.", {
  x: 1.5, y: 2.2, w: 10.33, h: 1.5,
  fontSize: 36, fontFace: "Inter", bold: true, color: C.white, align: "center", lineSpacing: 46,
});

slide.addText("I am requesting board approval to initiate a 12-week proof-of-concept\nwith a dedicated team of five, beginning immediately.\n\nThe projected first-year savings alone will fund the transformation\nten times over. The question is no longer if this is possible.\nIt is how quickly we move.", {
  x: 2, y: 3.9, w: 9.33, h: 2,
  fontSize: 14, fontFace: "Inter", color: C.lightBlue, align: "center", lineSpacing: 24, transparency: 20,
});

horizonLine(slide, 6.0, C.lightBlue);

slide.addText("Presented by", { x: 2, y: 6.2, w: 3, h: 0.2, fontSize: 9, fontFace: "Inter", color: C.textLight, align: "center" });
slide.addText("Chief AI Officer", { x: 2, y: 6.4, w: 3, h: 0.25, fontSize: 12, fontFace: "Inter", bold: true, color: C.white, align: "center" });

slide.addText("To", { x: 5.5, y: 6.2, w: 3, h: 0.2, fontSize: 9, fontFace: "Inter", color: C.textLight, align: "center" });
slide.addText("Chief Executive Officer", { x: 5.5, y: 6.4, w: 3, h: 0.25, fontSize: 12, fontFace: "Inter", bold: true, color: C.white, align: "center" });

slide.addText("Date", { x: 9, y: 6.2, w: 2.5, h: 0.2, fontSize: 9, fontFace: "Inter", color: C.textLight, align: "center" });
slide.addText("March 2026", { x: 9, y: 6.4, w: 2.5, h: 0.25, fontSize: 12, fontFace: "Inter", bold: true, color: C.white, align: "center" });

addFooter(slide);


// ═══════════════════════════════════════════════════
// GENERATE
// ═══════════════════════════════════════════════════
const path = require("path");
const outputPath = path.resolve(__dirname, "CTO-Transformation-White-Paper.pptx");
pptx.writeFile({ fileName: outputPath }).then(() => {
  console.log(`PPTX generated: ${outputPath}`);
}).catch(err => {
  console.error("Error:", err);
});
