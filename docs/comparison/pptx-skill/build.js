// Harbor Roast board deck built following Anthropic's pptx skill (pptxgenjs, structured deck).
// Same brief as docs/comparison/brief.md. Run: node build.js  (needs pptxgenjs, react-icons, react, react-dom, sharp)
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const SKILL = process.env.PPTX_SKILL_DIR;
const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));

const THEME = {
  name: "Harbor Roast",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "2B1E16", lt1: "FFFFFF", dk2: "4A2E1F", lt2: "F7F3EE",
    accent1: "4A2E1F", accent2: "C99A3E", accent3: "8A5A3B", accent4: "7A8C6E", accent5: "B5483B", accent6: "A89A8E",
    hlink: "8A5A3B", folHlink: "6E5A4E",
  },
};
const HEX = THEME.colors;

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";                       // 13.33 x 7.5
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "Harbor Roast: Northfield recommendation";
  const C = pres.SchemeColor;

  // ---- layouts
  pres.defineSlideMaster({
    title: "TITLE_DARK",
    background: { color: HEX.dk2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.3, w: 9, h: 1.6, fontFace: "+mj-lt", fontSize: 44, bold: true, color: C.background1, valign: "top" }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.1, w: 9, h: 1.0, fontSize: 20, color: C.accent2, valign: "top" }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "TITLE_ONLY",
    background: { color: "FFFFFF" },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 1.0, fontFace: "+mj-lt", fontSize: 32, bold: true, color: C.text2, valign: "middle" }, text: "" } },
    ],
    slideNumber: { x: 12.3, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: "8A7D72", align: "right" },
  });

  const content = (title) => {
    const s = pres.addSlide({ masterName: "TITLE_ONLY" });
    s.addText(title, { placeholder: "title" });
    return s;
  };
  const source = (s, t) => s.addText(t, { x: 0.6, y: 6.9, w: 11.4, h: 0.35, fontSize: 10, color: "8A7D72", italic: true, isTextBox: true, margin: 0 });
  const axisFont = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt" };

  // 1 ---- title
  let s = pres.addSlide({ masterName: "TITLE_DARK" });
  s.addText("Should Harbor Roast open in Northfield?", { placeholder: "title" });
  s.addText("Recommendation to the board  |  February 2027", { placeholder: "body" });
  s.addImage({ data: await icon(fa.FaCoffee, HEX.accent2), x: 10.6, y: 2.4, w: 1.8, h: 1.8, objectName: "Coffee icon" });

  // 2 ---- executive summary: stat callouts
  s = content("Recommendation: open a six-store pilot in Northfield");
  const stats = [["8%", "a year market growth in Northfield, vs 3% at home"], ["14%", "store margin, above our 11% average"],
                 ["2.7 yrs", "payback on each £350K fit-out"], ["£2.4M", "to open six stores, March–September 2027"]];
  stats.forEach(([big, lab], i) => {
    const x = 0.6 + i * 3.1;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.9, w: 2.85, h: 3.2, fill: { color: C.background2 }, rectRadius: 0.1, line: { color: C.background2 }, objectName: `Stat card ${i + 1}` });
    s.addText(big, { x: x + 0.2, y: 2.2, w: 2.45, h: 1.2, fontSize: 54, bold: true, color: i === 3 ? C.accent2 : C.text2, fontFace: "+mj-lt", isTextBox: true, margin: 0 });
    s.addText(lab, { x: x + 0.2, y: 3.5, w: 2.45, h: 1.3, fontSize: 16, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
  });
  s.addText("We recommend starting in the two university districts, which have high footfall and the least competition.",
    { x: 0.6, y: 5.6, w: 12.1, h: 0.6, fontSize: 18, color: C.text1, isTextBox: true, margin: 0 });

  // 3 ---- market growth: line chart
  s = content("Northfield's market is growing much faster than ours");
  const yrs = ["2020", "2021", "2022", "2023", "2024", "2025", "2026E", "2027E", "2028E", "2029E", "2030E"];
  s.addChart(pres.charts.LINE, [
    { name: "Northfield", labels: yrs, values: [120, 131, 143, 156, 170, 185, 200, 216, 233, 251, 270] },
    { name: "Home region", labels: yrs, values: [210, 218, 226, 233, 240, 247, 254, 260, 267, 273, 280] },
  ], { x: 0.6, y: 1.6, w: 8.2, h: 5.0, chartColors: [HEX.accent2, HEX.dk2], lineSize: 3, lineDataSymbol: "none",
       showTitle: true, title: "Specialty coffee spend, £M", titleFontSize: 14, titleColor: HEX.dk1,
       showLegend: true, legendPos: "b", legendFontSize: 12, catAxisLabelColor: "6E5A4E", valAxisLabelColor: "6E5A4E",
       catAxisLabelFontSize: 11, valAxisLabelFontSize: 11, valGridLine: { color: "E5DED6", size: 0.5 }, catGridLine: { style: "none" }, ...axisFont });
  s.addText([
    { text: "Northfield: ", options: { bold: true, color: C.accent3 } }, { text: "about 8% a year, reaching £270M by 2030", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Home region: ", options: { bold: true, color: C.accent3 } }, { text: "about 3% a year", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "By 2030 Northfield is nearly the size of our home market", options: { italic: true } },
  ], { x: 9.2, y: 2.0, w: 3.5, h: 4.0, fontSize: 16, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
  source(s, "Source: Regional Consumer Spend Survey 2025; 2026–2030 team forecast (fictional data)");

  // 4 ---- districts: scatter
  s = content("The university districts combine high footfall with few competitors");
  const pts = [["University Quarter", 48, 0.8], ["Riverside Campus", 41, 1.1], ["Old Town", 52, 3.8], ["Station", 45, 3.1],
               ["Retail Park", 30, 1.3], ["Eastgate", 22, 1.8], ["Westbrook", 18, 1.4], ["Market Hill", 35, 2.6]];
  s.addChart(pres.charts.SCATTER, [
    { name: "Footfall", values: pts.map((p) => p[1]) },
    { name: "Cafés per 10,000 residents", values: pts.map((p) => p[2]), labels: pts.map((p) => p[0]) },
  ], { x: 0.6, y: 1.6, w: 8.2, h: 5.0, chartColors: [HEX.accent2], lineSize: 0, lineDataSymbolSize: 12,
       showTitle: true, title: "Footfall (thousands/day) vs cafés per 10,000 residents", titleFontSize: 14, titleColor: HEX.dk1,
       showValAxisTitle: true, valAxisTitle: "Cafés per 10,000 residents", showCatAxisTitle: true, catAxisTitle: "Weekday footfall (thousands)",
       valAxisTitleFontSize: 12, catAxisTitleFontSize: 12, showLabel: true, dataLabelFormatScatter: "custom", dataLabelFontSize: 10,
       catAxisLabelColor: "6E5A4E", valAxisLabelColor: "6E5A4E", valGridLine: { color: "E5DED6", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, ...axisFont });
  const why = [[fa.FaWalking, "Footfall", "41–48 thousand people a day"], [fa.FaStore, "Competition", "Fewest specialty cafés per resident"], [fa.FaUserGraduate, "Customers", "Students match our core customer"]];
  for (let i = 0; i < why.length; i++) {
    const [Ic, h, t] = why[i]; const y = 2.0 + i * 1.45;
    s.addShape(pres.shapes.OVAL, { x: 9.2, y, w: 0.7, h: 0.7, fill: { color: C.accent1 }, objectName: `Icon circle ${i + 1}` });
    s.addImage({ data: await icon(Ic, "FFFFFF"), x: 9.37, y: y + 0.17, w: 0.36, h: 0.36, objectName: `Icon ${h}` });
    s.addText([{ text: h, options: { bold: true, breakLine: true, color: C.text2 } }, { text: t }], { x: 10.1, y: y - 0.05, w: 2.6, h: 1.0, fontSize: 15, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
  }
  source(s, "Source: council footfall counters 2025; café listings, January 2027 (fictional data)");

  // 5 ---- store economics: bar chart of the P&L
  s = content("Each store would earn a 14% margin, above our 11% average");
  s.addChart(pres.charts.BAR, [{ name: "£K per year", labels: ["Revenue", "Coffee & food", "Staff", "Rent", "Other costs", "Operating profit"], values: [1000, 310, 280, 170, 100, 140] }],
    { x: 0.6, y: 1.6, w: 7.6, h: 5.0, barDir: "bar", chartColors: [HEX.dk2], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelColor: HEX.dk1,
      showTitle: true, title: "Annual P&L for one store, steady state (£K)", titleFontSize: 14, titleColor: HEX.dk1, catAxisOrientation: "maxMin",
      catAxisLabelColor: "6E5A4E", valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, ...axisFont });
  [["14%", "operating margin"], ["17%", "rent as a share of revenue, 3 pts below home"]].forEach(([big, lab], i) => {
    const y = 2.1 + i * 2.1;
    s.addText(big, { x: 8.8, y, w: 3.9, h: 0.9, fontSize: 54, bold: true, color: C.accent2, fontFace: "+mj-lt", isTextBox: true, margin: 0 });
    s.addText(lab, { x: 8.8, y: y + 0.95, w: 3.9, h: 0.7, fontSize: 16, color: C.text1, isTextBox: true, margin: 0 });
  });
  source(s, "Year 2 onwards, per store. Source: home-region store P&Ls 2026; Northfield rent quotes (fictional data)");

  // 6 ---- payback
  s = content("Each store pays back its fit-out in under three years");
  s.addChart(pres.charts.BAR, [
    { name: "Fit-out cost", labels: ["Fit-out", "Year 1", "Year 2", "Year 3 (8 months)"], values: [350, 0, 0, 0] },
    { name: "Cumulative profit", labels: ["Fit-out", "Year 1", "Year 2", "Year 3 (8 months)"], values: [0, 112, 252, 350] },
  ], { x: 0.6, y: 1.6, w: 8.2, h: 5.0, barDir: "col", barGrouping: "stacked", chartColors: [HEX.accent3, HEX.accent2], showValue: true, dataLabelPosition: "ctr",
       dataLabelColor: "FFFFFF", dataLabelFontSize: 12, showTitle: true, title: "Fit-out vs cumulative profit, one store (£K)", titleFontSize: 14, titleColor: HEX.dk1,
       showLegend: true, legendPos: "b", legendFontSize: 12, catAxisLabelColor: "6E5A4E", valAxisLabelColor: "6E5A4E", valGridLine: { color: "E5DED6", size: 0.5 }, catGridLine: { style: "none" }, ...axisFont });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.2, y: 2.2, w: 3.5, h: 2.6, fill: { color: C.background2 }, rectRadius: 0.1, line: { color: C.background2 }, objectName: "Payback card" });
  s.addText([{ text: "2 years, 8 months", options: { fontSize: 32, bold: true, color: C.text2, breakLine: true, fontFace: "+mj-lt" } },
             { text: "to recover the £350K fit-out, assuming year 1 reaches 80% of steady-state profit", options: { fontSize: 15, color: C.text1 } }],
    { x: 9.45, y: 2.4, w: 3.0, h: 2.2, valign: "top", isTextBox: true, margin: 0 });
  source(s, "Source: team model; no discounting (fictional data)");

  // 7 ---- options: table
  s = content("A six-store pilot is the best of four options");
  const rag = { Good: ["DCEBD9", "2F6B2F"], Fair: ["F6E7C8", "8A6212"], Poor: ["F3D9D4", "9A3324"] };
  const cell = (t) => ({ text: t, options: { fill: { color: rag[t][0] }, color: rag[t][1], bold: true, align: "center" } });
  const head = ["Option", "Value", "Risk", "Speed", "Brand control", "Overall"].map((t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.dk2 }, align: t === "Option" ? "left" : "center" } }));
  const rows = [["Six-store pilot ★", "Fair", "Good", "Good", "Good", "Good"], ["20-store rollout", "Good", "Poor", "Fair", "Good", "Fair"],
                ["Franchise partner", "Fair", "Good", "Good", "Poor", "Fair"], ["Do nothing", "Poor", "Good", "Good", "Good", "Poor"]];
  s.addTable([head, ...rows.map((r) => [{ text: r[0], options: { bold: r[0].includes("★") } }, ...r.slice(1).map(cell)])],
    { x: 0.6, y: 1.9, w: 12.1, colW: [3.6, 1.7, 1.7, 1.7, 1.7, 1.7], rowH: 0.75, fontSize: 16, fontFace: "+mn-lt", color: C.text1, valign: "middle",
      border: { type: "solid", pt: 1, color: "FFFFFF" } });
  source(s, "Ratings against the board's criteria agreed in November 2026 (fictional)");

  // 8 ---- risks: icon rows
  s = content("The main risks can be managed");
  const risks = [[fa.FaFileContract, "Leases taken by competitors", "Medium", "Sign heads of terms for four sites by March; two backups identified"],
                 [fa.FaUserTie, "Can't hire experienced managers locally", "Medium", "Transfer two home-region managers for year 1; train local deputies"],
                 [fa.FaSun, "Student demand drops in summer", "High", "Summer menu and hours; target staff and visitors"],
                 [fa.FaTools, "Fit-out costs overrun", "Low", "Fixed-price contracts; 10% contingency inside the £2.4M"]];
  for (let i = 0; i < risks.length; i++) {
    const [Ic, r, lik, mit] = risks[i]; const y = 1.75 + i * 1.2;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.1, w: 0.7, h: 0.7, fill: { color: C.accent1 }, objectName: `Risk icon circle ${i + 1}` });
    s.addImage({ data: await icon(Ic, "FFFFFF"), x: 0.77, y: y + 0.27, w: 0.36, h: 0.36, objectName: `Risk icon ${i + 1}` });
    s.addText([{ text: r, options: { bold: true, breakLine: true, color: C.text2 } }, { text: `Likelihood: ${lik}`, options: { color: C.accent3, fontSize: 13 } }],
      { x: 1.5, y, w: 4.6, h: 0.95, fontSize: 16, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(mit, { x: 6.4, y, w: 6.3, h: 0.95, fontSize: 16, color: C.text1, valign: "middle", isTextBox: true, margin: 0 });
  }

  // 9 ---- plan: timeline
  s = content("Six stores would open between March and September 2027");
  const steps = [["Jan–Feb", "Sign leases"], ["March", "Stores 1–2 open"], ["June", "Stores 3–4 open"], ["September", "Stores 5–6 open"], ["December", "Board review"]];
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 3.6, w: 11.3, h: 0, line: { color: HEX.accent6, width: 3 }, objectName: "Timeline" });
  steps.forEach(([when, what], i) => {
    const cx = 1.0 + i * (11.3 / 4);
    s.addShape(pres.shapes.OVAL, { x: cx - 0.25, y: 3.35, w: 0.5, h: 0.5, fill: { color: i === 4 ? C.accent2 : C.accent1 }, objectName: `Milestone ${i + 1}` });
    s.addText(when, { x: cx - 1.2, y: 2.5, w: 2.4, h: 0.5, fontSize: 18, bold: true, color: C.text2, align: "center", isTextBox: true, margin: 0 });
    s.addText(what, { x: cx - 1.2, y: 4.1, w: 2.4, h: 0.6, fontSize: 16, color: C.text1, align: "center", isTextBox: true, margin: 0 });
  });

  // 10 ---- ask
  s = pres.addSlide({ masterName: "TITLE_DARK" });
  s.addText("We ask the board to approve £2.4M", { placeholder: "title" });
  s.addText("to open six Northfield stores between March and September 2027", { placeholder: "body" });

  const out = path.join(__dirname, "pptx-skill-deck.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log(out);
})();
