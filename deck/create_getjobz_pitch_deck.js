const fs = require("fs");
const path = require("path");

const pptxgen = require(path.join(process.env.APPDATA, "npm", "node_modules", "pptxgenjs"));

const root = path.resolve(__dirname, "..");
const primaryOutPath = path.join(__dirname, "Getjobz_Pitch_Deck.pptx");
const fallbackOutPath = path.join(__dirname, "Getjobz_Pitch_Deck_Updated.pptx");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Getjobz";
pptx.company = "Getjobz";
pptx.subject = "Getjobz AI Recruitment Agency pitch deck";
pptx.title = "Getjobz Pitch Deck";
pptx.lang = "en-US";
pptx.theme = {
  headFontFace: "Aptos Display",
  bodyFontFace: "Aptos",
  lang: "en-US"
};
pptx.margin = 0;

const C = {
  ink: "172033",
  muted: "5C6676",
  paper: "F5F6F1",
  panel: "FFFFFF",
  line: "DCE2E7",
  teal: "0F766E",
  tealDark: "0B514E",
  coral: "C95F45",
  gold: "D4A017",
  sage: "D9E7DC"
};

function svgData(name) {
  const svg = fs.readFileSync(path.join(root, "assets", name), "utf8");
  return `data:image/svg+xml;base64,${Buffer.from(svg).toString("base64")}`;
}

function addBackground(slide, dark = false) {
  slide.background = { color: dark ? C.ink : C.paper };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.12,
    fill: { color: dark ? C.gold : C.teal },
    line: { color: dark ? C.gold : C.teal }
  });
}

function addLogo(slide, dark = false) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: 0.38,
    w: 0.5,
    h: 0.5,
    rectRadius: 0.06,
    fill: { color: dark ? C.panel : C.tealDark },
    line: { color: dark ? C.panel : C.tealDark }
  });
  slide.addText("GJ", {
    x: 0.64,
    y: 0.51,
    w: 0.33,
    h: 0.16,
    fontFace: "Aptos",
    fontSize: 10,
    bold: true,
    color: dark ? C.tealDark : C.panel,
    margin: 0
  });
  slide.addText("Getjobz", {
    x: 1.15,
    y: 0.5,
    w: 1.2,
    h: 0.22,
    fontFace: "Aptos Display",
    fontSize: 16,
    bold: true,
    color: dark ? C.panel : C.ink,
    margin: 0
  });
}

function addFooter(slide, number, dark = false) {
  slide.addText(`Getjobz AI Recruitment Agency / ${String(number).padStart(2, "0")}`, {
    x: 0.62,
    y: 7.1,
    w: 4.2,
    h: 0.18,
    fontSize: 8,
    color: dark ? "D9E1E7" : C.muted,
    margin: 0
  });
}

function addTitle(slide, eyebrow, title, body, dark = false) {
  slide.addText(eyebrow.toUpperCase(), {
    x: 0.72,
    y: 1.05,
    w: 4.5,
    h: 0.22,
    fontSize: 9,
    bold: true,
    color: dark ? C.gold : C.coral,
    charSpace: 1.1,
    margin: 0
  });
  slide.addText(title, {
    x: 0.68,
    y: 1.36,
    w: 5.9,
    h: 1.0,
    fontFace: "Aptos Display",
    fontSize: 34,
    bold: true,
    color: dark ? C.panel : C.ink,
    breakLine: false,
    fit: "shrink",
    margin: 0
  });
  if (body) {
    slide.addText(body, {
      x: 0.72,
      y: 2.5,
      w: 5.25,
      h: 0.82,
      fontSize: 14,
      color: dark ? "D9E1E7" : C.muted,
      breakLine: false,
      fit: "shrink",
      margin: 0
    });
  }
}

function card(slide, x, y, w, h, title, body, accent = C.teal) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.05,
    fill: { color: C.panel },
    line: { color: C.line, width: 1 }
  });
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w,
    h: 0.08,
    fill: { color: accent },
    line: { color: accent }
  });
  slide.addText(title, {
    x: x + 0.2,
    y: y + 0.22,
    w: w - 0.4,
    h: 0.26,
    fontSize: 15,
    bold: true,
    color: C.ink,
    margin: 0
  });
  slide.addText(body, {
    x: x + 0.2,
    y: y + 0.62,
    w: w - 0.4,
    h: h - 0.78,
    fontSize: 10.5,
    color: C.muted,
    fit: "shrink",
    breakLine: false,
    margin: 0
  });
}

function pill(slide, x, y, text, color, width = 1.3) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w: width,
    h: 0.38,
    rectRadius: 0.12,
    fill: { color },
    line: { color }
  });
  slide.addText(text, {
    x: x + 0.08,
    y: y + 0.11,
    w: width - 0.16,
    h: 0.14,
    fontSize: 8.5,
    bold: true,
    color: color === C.sage ? C.tealDark : C.panel,
    align: "center",
    margin: 0
  });
}

function addBullets(slide, x, y, items, dark = false) {
  items.forEach((item, index) => {
    const top = y + index * 0.72;
    slide.addShape(pptx.ShapeType.ellipse, {
      x,
      y: top + 0.08,
      w: 0.16,
      h: 0.16,
      fill: { color: dark ? C.gold : C.teal },
      line: { color: dark ? C.gold : C.teal }
    });
    slide.addText(item, {
      x: x + 0.28,
      y: top,
      w: 4.95,
      h: 0.42,
      fontSize: 13,
      color: dark ? "E8EEF0" : C.ink,
      fit: "shrink",
      margin: 0
    });
  });
}

function personCard(slide, x, y, w, h, initials, name, role, body, color) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.06,
    fill: { color: C.panel },
    line: { color: C.line, width: 1 }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + 0.22,
    y: y + 0.24,
    w: 0.72,
    h: 0.72,
    fill: { color },
    line: { color }
  });
  slide.addText(initials, {
    x: x + 0.36,
    y: y + 0.47,
    w: 0.44,
    h: 0.18,
    fontSize: 10,
    bold: true,
    color: C.panel,
    align: "center",
    margin: 0
  });
  slide.addText(name, {
    x: x + 1.08,
    y: y + 0.27,
    w: w - 1.28,
    h: 0.24,
    fontSize: 14,
    bold: true,
    color: C.ink,
    margin: 0
  });
  slide.addText(role, {
    x: x + 1.08,
    y: y + 0.58,
    w: w - 1.28,
    h: 0.18,
    fontSize: 8.5,
    bold: true,
    color: C.tealDark,
    margin: 0
  });
  slide.addText(body, {
    x: x + 0.24,
    y: y + 1.08,
    w: w - 0.48,
    h: h - 1.28,
    fontSize: 9.2,
    color: C.muted,
    fit: "shrink",
    margin: 0
  });
}

function slide1() {
  const slide = pptx.addSlide();
  addBackground(slide, true);
  addLogo(slide, true);
  addTitle(
    slide,
    "Pitch deck",
    "Getjobz: AI recruitment agency for faster, higher-quality shortlists.",
    "We screen CVs, extract candidate evidence, rank applicants against job requirements, and help HR teams move the right people forward.",
    true
  );
  slide.addImage({
    data: svgData("hero-recruitment.svg"),
    x: 6.35,
    y: 1.08,
    w: 6.3,
    h: 4.02
  });
  pill(slide, 0.72, 3.64, "CV screening", C.teal, 1.45);
  pill(slide, 2.34, 3.64, "Candidate ranking", C.coral, 1.75);
  pill(slide, 4.27, 3.64, "Shortlist reports", C.gold, 1.6);
  slide.addText("Saving recruiters hours while keeping hiring decisions explainable and human-led.", {
    x: 0.72,
    y: 5.92,
    w: 7.0,
    h: 0.42,
    fontSize: 15,
    color: "D9E1E7",
    margin: 0
  });
  addFooter(slide, 1, true);
}

function slide2() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Problem", "Recruiting teams are overloaded by CV review.", "High-volume hiring creates repetitive manual work, inconsistent screening, and delayed shortlists.");
  card(slide, 6.55, 1.05, 2.0, 1.55, "CV overload", "Hundreds of applicants can arrive before recruiters have time to assess role fit.", C.coral);
  card(slide, 8.8, 1.05, 2.0, 1.55, "Hidden talent", "Strong candidates are missed when relevant evidence is buried in unstructured CVs.", C.gold);
  card(slide, 11.05, 1.05, 1.85, 1.55, "Slow decisions", "Hiring managers wait for shortlists while recruiters repeat the same screening steps.", C.teal);
  slide.addShape(pptx.ShapeType.line, { x: 1.0, y: 4.72, w: 10.9, h: 0, line: { color: C.line, width: 2 } });
  ["Apply", "Manual review", "Shortlist", "Interview"].forEach((label, i) => {
    const x = 1.0 + i * 3.55;
    slide.addShape(pptx.ShapeType.ellipse, {
      x,
      y: 4.4,
      w: 0.64,
      h: 0.64,
      fill: { color: i === 1 ? C.coral : C.panel },
      line: { color: i === 1 ? C.coral : C.teal, width: 2 }
    });
    slide.addText(label, { x: x - 0.25, y: 5.2, w: 1.3, h: 0.25, fontSize: 11, bold: true, color: C.ink, align: "center", margin: 0 });
  });
  slide.addText("The biggest bottleneck sits before the first interview.", {
    x: 3.7,
    y: 5.82,
    w: 6.1,
    h: 0.34,
    fontSize: 18,
    bold: true,
    color: C.tealDark,
    align: "center",
    margin: 0
  });
  addFooter(slide, 2);
}

function slide3() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Solution", "Getjobz automates the first screening pass.", "The platform creates a structured hiring scorecard, analyses CVs, ranks candidates, and produces shortlist notes recruiters can defend.");
  card(slide, 6.6, 1.15, 2.85, 1.55, "1. Job profile", "Capture must-have skills, seniority, experience level, and role context.", C.teal);
  card(slide, 9.85, 1.15, 2.85, 1.55, "2. CV intelligence", "Extract skills, tools, industries, achievements, and years of experience.", C.gold);
  card(slide, 6.6, 3.15, 2.85, 1.55, "3. Fit scoring", "Rank every profile against weighted requirements and gaps.", C.coral);
  card(slide, 9.85, 3.15, 2.85, 1.55, "4. Shortlist", "Generate manager-ready notes and interview focus areas.", C.tealDark);
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.8,
    y: 4.74,
    w: 5.0,
    h: 1.12,
    rectRadius: 0.08,
    fill: { color: C.ink },
    line: { color: C.ink }
  });
  slide.addText("Outcome: recruiters spend less time filtering and more time assessing the candidates most likely to succeed.", {
    x: 1.08,
    y: 5.06,
    w: 4.45,
    h: 0.44,
    fontSize: 14,
    color: C.panel,
    fit: "shrink",
    margin: 0
  });
  addFooter(slide, 3);
}

function slide4() {
  const slide = pptx.addSlide();
  addBackground(slide, true);
  addLogo(slide, true);
  addTitle(slide, "Product", "Recruiter dashboard for live candidate matching.", "Getjobz makes the shortlist visible: ranked profiles, skill coverage, match scores, gaps, and next-step recommendations.", true);
  slide.addImage({
    data: svgData("product-dashboard.svg"),
    x: 6.05,
    y: 0.94,
    w: 6.65,
    h: 4.28
  });
  addBullets(slide, 0.92, 3.7, [
    "Paste or upload CV text for quick role matching.",
    "Rank candidates against weighted requirements.",
    "Show matched skills and missing criteria for each candidate.",
    "Export shortlist notes for hiring manager review."
  ], true);
  addFooter(slide, 4, true);
}

function slide5() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Technology", "Screening engine built around evidence, not guesswork.", "The model pipeline converts unstructured CVs and job descriptions into comparable hiring signals.");
  const steps = [
    ["CV ingest", "Text, resume files, candidate profiles"],
    ["Extraction", "Skills, tools, industries, achievements"],
    ["Matching", "Weighted scorecard, keywords, seniority"],
    ["Recommendation", "Rank, gaps, interview focus"]
  ];
  steps.forEach(([title, body], i) => {
    const x = 0.85 + i * 3.08;
    card(slide, x, 3.05, 2.45, 1.55, title, body, [C.teal, C.gold, C.coral, C.tealDark][i]);
    if (i < steps.length - 1) {
      slide.addShape(pptx.ShapeType.chevron, {
        x: x + 2.55,
        y: 3.62,
        w: 0.32,
        h: 0.38,
        fill: { color: C.line },
        line: { color: C.line }
      });
    }
  });
  slide.addText("Designed for recruiter review, auditability, and configurable scorecards.", {
    x: 2.35,
    y: 5.55,
    w: 8.65,
    h: 0.36,
    fontSize: 18,
    bold: true,
    color: C.tealDark,
    align: "center",
    margin: 0
  });
  addFooter(slide, 5);
}

function slide6() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Customers", "Built for teams that hire often and move fast.", "Beachhead customers can adopt Getjobz for high-volume roles where screening quality and recruiter time matter.");
  card(slide, 6.6, 1.2, 2.9, 1.6, "Recruitment agencies", "Screen large applicant pools and send stronger shortlists to clients.", C.teal);
  card(slide, 9.8, 1.2, 2.9, 1.6, "SME hiring teams", "Add AI screening without buying a heavy enterprise HR suite.", C.gold);
  card(slide, 6.6, 3.2, 2.9, 1.6, "HR departments", "Standardize evaluation across roles, teams, and hiring managers.", C.coral);
  card(slide, 9.8, 3.2, 2.9, 1.6, "Startups", "Move faster from applications to interviews during growth hiring.", C.tealDark);
  addBullets(slide, 0.95, 3.65, [
    "Initial focus: data, sales, operations, customer success, and support roles.",
    "Entry point: one role, one shortlist, measurable time saved.",
    "Expansion: team scorecards, pipelines, analytics, and integrations."
  ]);
  addFooter(slide, 6);
}

function slide7() {
  const slide = pptx.addSlide();
  addBackground(slide, true);
  addLogo(slide, true);
  addTitle(slide, "Business model", "Software plus agency-grade screening services.", "Getjobz can sell recurring software access and done-with-you hiring support for teams that need immediate results.", true);
  card(slide, 0.85, 3.05, 2.85, 1.7, "Starter", "Low-cost monthly access for small teams running a few roles.", C.teal);
  card(slide, 3.98, 3.05, 2.85, 1.7, "Growth", "Team seats, shared pipelines, analytics, and shortlist exports.", C.gold);
  card(slide, 7.1, 3.05, 2.85, 1.7, "Agency service", "Per-role screening, ranking, and shortlist report delivery.", C.coral);
  card(slide, 10.22, 3.05, 2.35, 1.7, "Enterprise", "Integrations, governance, access controls, and custom scorecards.", C.tealDark);
  slide.addText("Revenue expands as customers add roles, users, candidate volume, and workflow integrations.", {
    x: 1.2,
    y: 5.52,
    w: 10.9,
    h: 0.36,
    fontSize: 16,
    bold: true,
    color: C.panel,
    align: "center",
    margin: 0
  });
  addFooter(slide, 7, true);
}

function slide8() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Go-to-market", "Win with pilots that prove time saved.", "Getjobz starts where the pain is visible: roles with high applicant volume and limited recruiter bandwidth.");
  addBullets(slide, 0.92, 3.05, [
    "Offer two-week screening pilots for one active role.",
    "Target recruitment agencies, HR managers, startup founders, and operations teams.",
    "Use shortlist quality, time saved, and interview conversion as proof points.",
    "Build repeatable templates for common roles before deeper ATS integrations."
  ]);
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 7.2,
    y: 1.25,
    w: 4.8,
    h: 4.9,
    rectRadius: 0.06,
    fill: { color: C.panel },
    line: { color: C.line }
  });
  ["Pilot", "Measure", "Convert", "Expand"].forEach((label, i) => {
    const y = 1.78 + i * 0.93;
    slide.addShape(pptx.ShapeType.ellipse, { x: 7.68, y, w: 0.42, h: 0.42, fill: { color: [C.teal, C.gold, C.coral, C.tealDark][i] }, line: { color: [C.teal, C.gold, C.coral, C.tealDark][i] } });
    slide.addText(label, { x: 8.35, y: y + 0.06, w: 2.2, h: 0.22, fontSize: 16, bold: true, color: C.ink, margin: 0 });
    if (i < 3) slide.addShape(pptx.ShapeType.line, { x: 7.89, y: y + 0.48, w: 0, h: 0.45, line: { color: C.line, width: 2 } });
  });
  addFooter(slide, 8);
}

function slide9() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Positioning", "A recruiting workflow product, not just a chatbot.", "Getjobz combines structured scorecards, CV intelligence, ranked shortlists, and responsible AI controls.");
  const rows = [
    ["Capability", "Manual review", "ATS only", "Generic AI", "Getjobz"],
    ["CV evidence extraction", "Slow", "Basic", "Prompted", "Structured"],
    ["Role-fit ranking", "Inconsistent", "Limited", "Manual setup", "Built in"],
    ["Shortlist reports", "Manual", "Limited", "Unstructured", "Ready"],
    ["Human review controls", "Yes", "Workflow", "Varies", "Core"]
  ];
  const x = 0.75;
  const y = 2.5;
  const colW = [2.45, 2.1, 2.0, 2.0, 2.0];
  rows.forEach((row, r) => {
    let left = x;
    row.forEach((cell, c) => {
      slide.addShape(pptx.ShapeType.rect, {
        x: left,
        y: y + r * 0.55,
        w: colW[c],
        h: 0.55,
        fill: { color: r === 0 ? C.ink : c === 4 ? "E6F3EF" : C.panel },
        line: { color: C.line, width: 0.6 }
      });
      slide.addText(cell, {
        x: left + 0.08,
        y: y + r * 0.55 + 0.16,
        w: colW[c] - 0.16,
        h: 0.16,
        fontSize: r === 0 ? 9.5 : 8.5,
        bold: r === 0 || c === 4,
        color: r === 0 ? C.panel : C.ink,
        fit: "shrink",
        margin: 0
      });
      left += colW[c];
    });
  });
  slide.addText("Differentiation: recruiter-ready workflow, transparent scoring, and shortlist outputs that save time immediately.", {
    x: 1.25,
    y: 6.0,
    w: 10.9,
    h: 0.32,
    fontSize: 15,
    bold: true,
    color: C.tealDark,
    align: "center",
    margin: 0
  });
  addFooter(slide, 9);
}

function slide10() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Responsible AI", "Explainable ranking with human control.", "Getjobz is designed to support better recruiter decisions, not replace them.");
  slide.addImage({ data: svgData("compliance-visual.svg"), x: 7.05, y: 1.22, w: 5.0, h: 3.45 });
  addBullets(slide, 0.92, 3.28, [
    "Scores show matched skills, missing criteria, and evidence notes.",
    "Recruiters can edit scorecards before screening and review every recommendation.",
    "Candidate data retention and access permissions can be configured.",
    "Final hiring decisions remain with the human hiring team."
  ]);
  addFooter(slide, 10);
}

function slide11() {
  const slide = pptx.addSlide();
  addBackground(slide, true);
  addLogo(slide, true);
  addTitle(slide, "Operating plan", "Use funding to build, validate, and sell repeatable screening workflows.", "The next phase should prove product demand with pilots and turn shortlist quality into repeat usage.", true);
  card(slide, 0.9, 3.05, 2.65, 1.55, "Product", "Improve CV parsing, scorecard setup, shortlist reports, and analytics.", C.teal);
  card(slide, 3.85, 3.05, 2.65, 1.55, "Security", "Add access controls, audit logs, data retention, and compliance workflows.", C.gold);
  card(slide, 6.8, 3.05, 2.65, 1.55, "Sales", "Run pilots with agencies, SMEs, and HR teams in target sectors.", C.coral);
  card(slide, 9.75, 3.05, 2.65, 1.55, "Operations", "Create role templates, service playbooks, and customer success support.", C.tealDark);
  slide.addText("Milestones: pilot customers, repeatable role templates, measurable screening time saved, and paid conversions.", {
    x: 1.05,
    y: 5.52,
    w: 11.2,
    h: 0.34,
    fontSize: 15,
    bold: true,
    color: C.panel,
    align: "center",
    margin: 0
  });
  addFooter(slide, 11, true);
}

function slide12() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "Decision team", "Human experts own the final shortlist decision.", "Getjobz uses AI to rank candidates, then a responsible review team validates evidence, fit, and client context before final recommendations are sent.");
  personCard(slide, 0.85, 3.0, 2.8, 1.72, "AO", "Adaora Okeke", "Head of Talent Decisions", "Approves final shortlists and calibrates role-fit recommendations with hiring managers.", C.teal);
  personCard(slide, 3.92, 3.0, 2.8, 1.72, "MT", "Michael Thompson", "People Operations Lead", "Reviews experience depth, candidate readiness, compensation fit, and interview process quality.", C.ink);
  personCard(slide, 6.99, 3.0, 2.8, 1.72, "NB", "Nadia Bello", "AI Screening Analyst", "Audits match scores, verifies extracted skills, and prepares interview focus notes.", C.coral);
  personCard(slide, 10.06, 3.0, 2.45, 1.72, "DE", "David Evans", "Client Hiring Partner", "Aligns recommendations with business goals, culture signals, and hiring timelines.", C.gold);
  slide.addText("Decision rule: AI recommends, the Getjobz team reviews, and the hiring manager makes the final employment decision.", {
    x: 1.1,
    y: 5.55,
    w: 11.0,
    h: 0.34,
    fontSize: 15,
    bold: true,
    color: C.tealDark,
    align: "center",
    margin: 0
  });
  addFooter(slide, 12);
}

function slide13() {
  const slide = pptx.addSlide();
  addBackground(slide);
  addLogo(slide);
  addTitle(slide, "The ask", "Partner with Getjobz to make shortlisting faster and smarter.", "We are looking for pilot customers, HR partners, and early investors who want a practical AI layer for recruiting teams.");
  addBullets(slide, 0.94, 3.1, [
    "Pilot Getjobz on one active hiring role.",
    "Share hiring manager feedback on shortlist quality.",
    "Partner on ATS integrations and recruiter workflows.",
    "Support the build-out of a trusted AI recruitment agency."
  ]);
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 7.05,
    y: 2.6,
    w: 5.0,
    h: 2.2,
    rectRadius: 0.08,
    fill: { color: C.ink },
    line: { color: C.ink }
  });
  slide.addText("Getjobz", {
    x: 7.55,
    y: 3.08,
    w: 4.0,
    h: 0.54,
    fontFace: "Aptos Display",
    fontSize: 32,
    bold: true,
    color: C.panel,
    align: "center",
    margin: 0
  });
  slide.addText("CV screening. Candidate ranking. Shortlist intelligence.", {
    x: 7.55,
    y: 3.78,
    w: 4.0,
    h: 0.32,
    fontSize: 13,
    color: "D9E1E7",
    align: "center",
    margin: 0
  });
  slide.addText("Contact form available on the Getjobz website: share role details, hiring volume, and shortlist timeline.", {
    x: 7.55,
    y: 4.22,
    w: 4.0,
    h: 0.42,
    fontSize: 10.5,
    color: "D9E1E7",
    align: "center",
    fit: "shrink",
    margin: 0
  });
  addFooter(slide, 13);
}

[
  slide1,
  slide2,
  slide3,
  slide4,
  slide5,
  slide6,
  slide7,
  slide8,
  slide9,
  slide10,
  slide11,
  slide12,
  slide13
].forEach((build) => build());

pptx
  .writeFile({ fileName: primaryOutPath })
  .then(() => {
    console.log(primaryOutPath);
  })
  .catch((error) => {
    if (error && error.code === "EBUSY") {
      return pptx.writeFile({ fileName: fallbackOutPath }).then(() => {
        console.log(fallbackOutPath);
      });
    }
    throw error;
  });
