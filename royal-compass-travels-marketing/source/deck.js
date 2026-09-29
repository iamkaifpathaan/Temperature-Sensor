const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const NAVY = "16324F", TEAL = "0E6E68", BRASS = "B08A3E", INK = "1F2933", MUTED = "5F6B77",
      LINE = "DDE3E8", SOFT = "F2F5F7", TEALSOFT = "E3F0EE", WHITE = "FFFFFF";
const HF = "Cambria", BF = "Calibri";

async function icon(Comp, color, size = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
async function rose(color, opacity) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200" width="1200" height="1200" fill="none" stroke="#${color}" stroke-opacity="${opacity}" stroke-width="0.8">
  <circle cx="100" cy="100" r="96"/><circle cx="100" cy="100" r="78"/><circle cx="100" cy="100" r="10"/>
  ${Array.from({length:72},(_,i)=>{const a=i*5*Math.PI/180,r1=i%18===0?84:(i%2?92:89);return `<line x1="${100+r1*Math.sin(a)}" y1="${100-r1*Math.cos(a)}" x2="${100+96*Math.sin(a)}" y2="${100-96*Math.cos(a)}"/>`}).join("")}
  <path d="M100 8 L112 100 L100 192 L88 100 Z"/><path d="M8 100 L100 88 L192 100 L100 112 Z"/>
  <path d="M100 8 L100 192 M8 100 L192 100"/>
  <path d="M40 40 L104 96 L160 160 L96 104 Z"/><path d="M160 40 L104 104 L40 160 L96 96 Z"/></svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.title = "Royal Compass Travels — Month 1 Digital Marketing Plan";
  pres.company = "Royal Compass Travels";

  const I = {
    search: await icon(fa.FaSearch, WHITE), robot: await icon(fa.FaRegLightbulb, WHITE), qa: await icon(fa.FaRegQuestionCircle, WHITE),
    eye: await icon(fa.FaRegEye, TEAL), link: await icon(fa.FaRegEnvelope, TEAL),
    ig: await icon(fa.FaInstagram, TEAL), meta: await icon(fa.FaBullhorn, TEAL), web: await icon(fa.FaGlobeAsia, TEAL),
    lock: await icon(fa.FaLock, WHITE), user: await icon(fa.FaUser, WHITE), form: await icon(fa.FaRegEdit, WHITE), phone: await icon(fa.FaPhoneAlt, WHITE),
    check: await icon(fa.FaCheck, TEAL), doc: await icon(fa.FaRegFileAlt, TEAL), tool: await icon(fa.FaTools, TEAL), gsc: await icon(fa.FaChartLine, TEAL), key: await icon(fa.FaKey, TEAL),
  };
  const roseLight = await rose("FFFFFF", 0.16);
  const roseDark = await rose(NAVY, 0.07);

  const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: BF, color: INK, margin: 0, valign: "top" }, o));
  const title = (s, kicker, text) => {
    T(s, kicker.toUpperCase(), { x: 0.6, y: 0.42, w: 8, h: 0.25, fontSize: 10, color: TEAL, bold: true, charSpacing: 3 });
    T(s, text, { x: 0.6, y: 0.68, w: 8.8, h: 0.6, fontSize: 28, fontFace: HF, bold: true, color: NAVY });
  };
  const footer = (s, n) => {
    T(s, "Royal Compass Travels  ·  Month 1 Plan", { x: 0.6, y: 5.22, w: 5, h: 0.2, fontSize: 8.5, color: "8A96A3" });
    T(s, String(n), { x: 8.9, y: 5.22, w: 0.5, h: 0.2, fontSize: 8.5, color: "8A96A3", align: "right" });
  };
  const circle = (s, x, y, d, fill, img, pad) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
    const p = pad ?? d * 0.27;
    s.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  };

  // ---------- 1. Title ----------
  let s = pres.addSlide(); s.background = { color: NAVY };
  s.addImage({ data: roseLight, x: 5.6, y: 0.55, w: 4.9, h: 4.9 });
  T(s, "ROYAL COMPASS TRAVELS", { x: 0.7, y: 1.35, w: 6, h: 0.3, fontSize: 12, color: "C9B27A", bold: true, charSpacing: 4 });
  T(s, "Month 1 Digital\nMarketing Plan", { x: 0.7, y: 1.8, w: 6, h: 1.6, fontSize: 42, fontFace: HF, bold: true, color: WHITE });
  T(s, "Website search · Instagram · Meta Ads · Lead system", { x: 0.7, y: 3.55, w: 6, h: 0.35, fontSize: 14, color: "C7D3DF" });
  T(s, "royalcompasstravels.com", { x: 0.7, y: 4.75, w: 4, h: 0.25, fontSize: 10, color: "8FA6BC" });
  s.addNotes("Introduce the plan as a first-month foundation and testing plan. Keep it short: we'll walk through what we'll do, how we'll measure it, and what it costs.");

  // ---------- 2. Objective ----------
  s = pres.addSlide(); s.background = { color: WHITE };
  s.addImage({ data: roseDark, x: 6.2, y: -1.2, w: 5.2, h: 5.2 });
  title(s, "Objective", "What Month 1 is for");
  T(s, "Build stronger online visibility and establish an initial lead-generation system.",
    { x: 0.6, y: 1.55, w: 5.2, h: 1.3, fontSize: 22, fontFace: HF, color: INK, italic: true });
  const obj = [
    [I.eye, "Visibility", "A clearer, better-structured website and a consistent Instagram presence."],
    [I.link, "Enquiries", "A tested ad setup and a simple system that sends every website enquiry to your team."],
  ];
  obj.forEach(([ic, h, d], i) => {
    const y = 3.15 + i * 0.95;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.6, h: 0.6, fill: { color: TEALSOFT }, line: { color: TEALSOFT } });
    s.addImage({ data: ic, x: 0.76, y: y + 0.16, w: 0.28, h: 0.28 });
    T(s, h, { x: 1.4, y: y + 0.02, w: 4.5, h: 0.3, fontSize: 15, bold: true, color: NAVY });
    T(s, d, { x: 1.4, y: y + 0.32, w: 5.2, h: 0.5, fontSize: 12.5, color: MUTED });
  });
  footer(s, 2);
  s.addNotes("The first month is about foundations and learning, not scale. Two outcomes: visibility (website + Instagram) and an enquiry system (ads + lead area).");

  // ---------- 3. Website Search Optimization ----------
  s = pres.addSlide(); s.background = { color: WHITE };
  title(s, "Website search optimization", "Helping customers find and understand you");
  const cols3 = [
    [I.search, "SEO", "Search engines", ["Page titles & descriptions", "Clear headings and structure", "Relevant travel keywords", "Google Search Console setup"]],
    [I.robot, "GEO", "AI-based search tools", ["Clear business information", "Services & destinations explained", "Consistent details across pages", "Structured, easy-to-read content"]],
    [I.qa, "AEO", "Customer questions", ["Common traveller questions", "Short, direct answers", "Helpful FAQ sections", "Better informational pages"]],
  ];
  cols3.forEach(([ic, h, sub, items], i) => {
    const x = 0.6 + i * 3.0, y = 1.6, w = 2.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 3.0, fill: { color: SOFT }, line: { color: SOFT }, rectRadius: 0.08 });
    circle(s, x + 0.25, y + 0.28, 0.62, i === 1 ? TEAL : NAVY, ic);
    T(s, h, { x: x + 1.02, y: y + 0.3, w: 1.6, h: 0.35, fontSize: 20, fontFace: HF, bold: true, color: NAVY });
    T(s, sub, { x: x + 1.02, y: y + 0.64, w: 1.7, h: 0.25, fontSize: 11, color: MUTED });
    T(s, items.map((t, k) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: k < items.length - 1 } })),
      { x: x + 0.25, y: y + 1.2, w: w - 0.45, h: 1.95, fontSize: 12, color: INK, paraSpaceAfter: 6 });
  });
  footer(s, 3);
  s.addNotes("SEO: making the site readable and relevant for Google. GEO: making information clear for AI-based search tools like ChatGPT or Google's AI answers. AEO: answering real customer questions directly. They overlap — the same clean pages help all three. Search visibility builds over time; Month 1 sets the foundation.");

  // ---------- 4. Instagram ----------
  s = pres.addSlide(); s.background = { color: WHITE };
  title(s, "Instagram", "A consistent, travel-focused presence");
  const stats = [["1", "post every day"], ["30", "posts per month"], ["10", "reels per month"]];
  stats.forEach(([n, l], i) => {
    const x = 0.6 + i * 1.95;
    T(s, n, { x, y: 1.65, w: 1.8, h: 0.95, fontSize: 54, fontFace: HF, bold: true, color: i === 2 ? TEAL : NAVY });
    T(s, l, { x, y: 2.6, w: 1.8, h: 0.3, fontSize: 12.5, color: MUTED });
  });
  T(s, "Content themes", { x: 0.6, y: 3.3, w: 4, h: 0.3, fontSize: 13, bold: true, color: NAVY });
  const chips = ["Destination inspiration", "Travel tips", "Travel information", "Our services", "Packages & offers", "FAQs", "Trust-building"];
  let cx = 0.6, cy = 3.72;
  chips.forEach(c => {
    const w = 0.22 + c.length * 0.083;
    if (cx + w > 6.3) { cx = 0.6; cy += 0.48; }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: cy, w, h: 0.36, fill: { color: TEALSOFT }, line: { color: TEALSOFT }, rectRadius: 0.18 });
    T(s, c, { x: cx, y: cy, w, h: 0.36, fontSize: 11, color: TEAL, align: "center", valign: "middle" });
    cx += w + 0.12;
  });
  // right panel
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.75, y: 1.6, w: 2.65, h: 3.35, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.08 });
  circle(s, 7.0, 1.85, 0.55, WHITE, I.ig, 0.13);
  T(s, "How it supports the plan", { x: 7.0, y: 2.55, w: 2.2, h: 0.3, fontSize: 13, bold: true, color: WHITE });
  const sup = ["Builds trust when people check your profile after seeing an ad", "Shows which topics your audience responds to", "Gives us material to reuse in ads"];
  T(s, sup.map((t, k) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: k < sup.length - 1 } })),
    { x: 7.0, y: 2.95, w: 2.2, h: 1.9, fontSize: 11, color: "D5DEE7", paraSpaceAfter: 6 });
  footer(s, 4);
  s.addNotes("30 posts plus 10 reels in the month, all travel-focused. Content plan shared in advance for approval. Real trip photos and videos from your side make the content more credible. Growth on Instagram is gradual; we'll track which content people save, share and message about.");

  // ---------- 5. Meta Ads ----------
  s = pres.addSlide(); s.background = { color: WHITE };
  title(s, "Meta Ads", "A focused test on Facebook & Instagram");
  const steps = [["5", "creatives developed", "Different images, messages and angles"], ["Test", "all five", "Early spend spread evenly to compare"], ["Best 3", "keep running", "Weaker ads paused based on results"], ["Optimise", "through the month", "Audience and budget refined"]];
  steps.forEach(([big, small, d], i) => {
    const x = 0.6 + i * 2.25, y = 1.65, w = 2.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 1.75, fill: { color: i === 2 ? NAVY : SOFT }, line: { color: i === 2 ? NAVY : SOFT }, rectRadius: 0.08 });
    T(s, big, { x: x + 0.2, y: y + 0.2, w: w - 0.4, h: 0.55, fontSize: 26, fontFace: HF, bold: true, color: i === 2 ? WHITE : NAVY });
    T(s, small, { x: x + 0.2, y: y + 0.78, w: w - 0.4, h: 0.3, fontSize: 12, bold: true, color: i === 2 ? "C9B27A" : TEAL });
    T(s, d, { x: x + 0.2, y: y + 1.1, w: w - 0.4, h: 0.55, fontSize: 10.5, color: i === 2 ? "D5DEE7" : MUTED });
    if (i < 3) T(s, "›", { x: x + w, y: y + 0.6, w: 0.25, h: 0.5, fontSize: 24, color: BRASS, align: "center", valign: "middle" });
  });
  // bottom band
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 3.7, w: 8.8, h: 1.25, fill: { color: TEALSOFT }, line: { color: TEALSOFT }, rectRadius: 0.08 });
  T(s, "₹10,000", { x: 0.9, y: 3.88, w: 2.2, h: 0.55, fontSize: 28, fontFace: HF, bold: true, color: TEAL });
  T(s, "ad budget, paid by you directly to Meta", { x: 0.9, y: 4.43, w: 2.6, h: 0.4, fontSize: 10.5, color: INK });
  T(s, [
    { text: "Focus on relevant enquiries", options: { bold: true, color: NAVY, breakLine: true } },
    { text: "We judge ads by the quality and cost of enquiries — not likes or cheap clicks. Month 1 is a testing budget to learn what works for your audience before any decision to spend more.", options: { color: INK } },
  ], { x: 3.7, y: 3.9, w: 5.5, h: 0.95, fontSize: 11.5, paraSpaceAfter: 3 });
  footer(s, 5);
  s.addNotes("Five creatives are made and tested; the best three continue and are optimised. The ₹10,000 is spent from your own ad account, so you can see every rupee. We need from you: Meta Business access, priority destinations/packages, any offers, and quick approval of creatives. We don't predict a lead count before the test — the test tells us the real cost per enquiry.");

  // ---------- 6. Website & Lead System ----------
  s = pres.addSlide(); s.background = { color: WHITE };
  title(s, "Website & lead system", "Every website enquiry, straight to your team");
  const flow = [[I.user, "Visitor", "Finds you via search,\nInstagram or ads"], [I.form, "Enquiry form", "Name, phone,\ndestination, dates"], [I.lock, "Your admin area", "Protected; only you\nhold the password"], [I.phone, "Your team", "Calls, WhatsApp,\nquotes & bookings"]];
  flow.forEach(([ic, h, d], i) => {
    const x = 0.6 + i * 2.3;
    circle(s, x + 0.62, 1.6, 0.72, i === 2 ? TEAL : NAVY, ic);
    T(s, h, { x, y: 2.45, w: 1.96, h: 0.3, fontSize: 13, bold: true, color: NAVY, align: "center" });
    T(s, d, { x, y: 2.75, w: 1.96, h: 0.5, fontSize: 10.5, color: MUTED, align: "center" });
    if (i < 3) s.addShape(pres.shapes.LINE, { x: x + 1.5, y: 1.96, w: 1.1, h: 0, line: { color: BRASS, width: 1.25, endArrowType: "triangle", dashType: "dash" } });
  });
  const tiles = [[I.doc, "Legal pages", "Privacy, terms, cancellation & refund"], [I.gsc, "Search Console", "Connected & verified with Google"], [I.tool, "Website fixes", "Issues affecting experience or enquiries"], [I.key, "Client-controlled", "You own the admin login and leads"]];
  tiles.forEach(([ic, h, d], i) => {
    const x = 0.6 + i * 2.25, y = 3.55, w = 2.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 1.4, fill: { color: SOFT }, line: { color: SOFT }, rectRadius: 0.08 });
    s.addImage({ data: ic, x: x + 0.2, y: y + 0.2, w: 0.3, h: 0.3 });
    T(s, h, { x: x + 0.2, y: y + 0.58, w: w - 0.35, h: 0.28, fontSize: 12.5, bold: true, color: NAVY });
    T(s, d, { x: x + 0.2, y: y + 0.86, w: w - 0.35, h: 0.48, fontSize: 10.5, color: MUTED });
  });
  footer(s, 6);
  s.addNotes("Enquiries go to your own protected admin area — not to us. You set and keep the password. Your team handles follow-up, quotations and bookings; quick replies make a real difference. Legal pages are drafted from your actual policies for your review. This is fixes and setup, not a website redesign.");

  // ---------- 7. Month 1 Execution ----------
  s = pres.addSlide(); s.background = { color: WHITE };
  title(s, "Month 1 execution", "Four weeks, step by step");
  const weeks = [
    ["Week 1", "Set up", ["Website audit", "Access & setup", "Search Console", "Content & ad planning"]],
    ["Week 2", "Build", ["SEO improvements", "Website fixes & legal pages", "Admin lead area", "Ad creatives"]],
    ["Week 3", "Launch", ["Meta Ads test goes live", "Instagram continues", "Monitor early results", "Adjust"]],
    ["Week 4", "Review", ["Run best-performing ads", "Content continues", "Review the numbers", "Month-end report"]],
  ];
  s.addShape(pres.shapes.LINE, { x: 0.85, y: 1.78, w: 8.3, h: 0, line: { color: LINE, width: 1.5 } });
  weeks.forEach(([wk, h, items], i) => {
    const x = 0.6 + i * 2.25;
    s.addShape(pres.shapes.OVAL, { x: x + 0.12, y: 1.65, w: 0.26, h: 0.26, fill: { color: i === 3 ? TEAL : NAVY }, line: { color: WHITE, width: 2 } });
    T(s, wk.toUpperCase(), { x, y: 2.05, w: 2, h: 0.25, fontSize: 10, bold: true, color: TEAL, charSpacing: 2 });
    T(s, h, { x, y: 2.3, w: 2, h: 0.4, fontSize: 19, fontFace: HF, bold: true, color: NAVY });
    T(s, items.map((t, k) => ({ text: t, options: { bullet: { indent: 11 }, breakLine: k < items.length - 1 } })),
      { x, y: 2.8, w: 2.05, h: 1.45, fontSize: 11.5, color: INK, paraSpaceAfter: 4 });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 4.4, w: 8.8, h: 0.6, fill: { color: SOFT }, line: { color: SOFT }, rectRadius: 0.06 });
  T(s, [{ text: "To keep this on schedule, we'll need from you: ", options: { bold: true, color: NAVY } },
        { text: "account access, package and destination details, photos/videos, and timely approvals.", options: { color: INK } }],
    { x: 0.85, y: 4.4, w: 8.4, h: 0.6, fontSize: 11.5, valign: "middle" });
  footer(s, 7);
  s.addNotes("Instagram posting starts as soon as the content plan is approved. Week 1 depends on receiving access; any delay shifts the later steps. Ads may need a day or so for Meta's review. If the ad budget isn't fully used by day 30, the test can run a few extra days to finish it.");

  // ---------- 8. What we will measure ----------
  s = pres.addSlide(); s.background = { color: WHITE };
  title(s, "What we will measure", "Clear numbers, reported at month-end");
  const meas = [
    [I.web, "Website", ["Search impressions", "Clicks from Google", "Search queries", "Top-performing pages"]],
    [I.ig, "Instagram", ["Reach", "Engagement", "Profile visits", "Content performance"]],
    [I.meta, "Meta Ads", ["Spend", "Impressions & reach", "CTR & CPC", "Leads & cost per lead"]],
  ];
  meas.forEach(([ic, h, items], i) => {
    const x = 0.6 + i * 3.0, y = 1.6, w = 2.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 2.7, fill: { color: WHITE }, line: { color: LINE, width: 1 }, rectRadius: 0.08 });
    circle(s, x + 0.25, y + 0.25, 0.55, TEALSOFT, ic, 0.14);
    T(s, h, { x: x + 0.95, y: y + 0.33, w: 1.7, h: 0.4, fontSize: 17, fontFace: HF, bold: true, color: NAVY });
    items.forEach((t, k) => {
      const yy = y + 1.05 + k * 0.38;
      s.addImage({ data: I.check, x: x + 0.3, y: yy + 0.06, w: 0.16, h: 0.16 });
      T(s, t, { x: x + 0.6, y: yy, w: w - 0.8, h: 0.3, fontSize: 12, color: INK, bold: i === 2 && k === 3 });
    });
  });
  T(s, [{ text: "What matters most: ", options: { bold: true, color: TEAL } }, { text: "relevant enquiries and what each one costs. Reach and impressions are shown for context.", options: { color: INK } }],
    { x: 0.6, y: 4.52, w: 8.8, h: 0.4, fontSize: 12 });
  footer(s, 8);
  s.addNotes("A short mid-month update and a month-end report: work completed, real numbers from Google Search Console, Instagram and Meta, what worked, what didn't, and our recommendation. Please share feedback on lead quality — it's what lets us improve targeting.");

  // ---------- 9. Investment ----------
  s = pres.addSlide(); s.background = { color: WHITE };
  title(s, "Investment", "Month 1");
  const inv = [["₹20,000", "Our service fee", "SEO, GEO & AEO · Instagram · Ad creatives & management · Website fixes & lead system · Reporting", NAVY],
               ["₹10,000", "Meta Ads budget", "Paid by you directly to Meta from your own ad account. Not included in our fee.", TEAL]];
  inv.forEach(([n, l, d, c], i) => {
    const x = 0.6 + i * 3.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.6, w: 2.85, h: 2.3, fill: { color: SOFT }, line: { color: SOFT }, rectRadius: 0.08 });
    T(s, l.toUpperCase(), { x: x + 0.3, y: 1.85, w: 2.4, h: 0.25, fontSize: 10, bold: true, color: c, charSpacing: 2 });
    T(s, n, { x: x + 0.3, y: 2.15, w: 2.4, h: 0.7, fontSize: 34, fontFace: HF, bold: true, color: c });
    T(s, d, { x: x + 0.3, y: 2.95, w: 2.3, h: 1.1, fontSize: 11, color: MUTED });
  });
  T(s, "+", { x: 3.45, y: 2.55, w: 0.2, h: 0.4, fontSize: 20, color: BRASS, align: "center" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.75, y: 1.6, w: 2.65, h: 2.3, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.08 });
  T(s, "TOTAL INITIAL OUTLAY", { x: 7.05, y: 1.85, w: 2.2, h: 0.25, fontSize: 10, bold: true, color: "C9B27A", charSpacing: 2 });
  T(s, "₹30,000", { x: 7.05, y: 2.15, w: 2.2, h: 0.7, fontSize: 34, fontFace: HF, bold: true, color: WHITE });
  T(s, "₹20,000 to us + ₹10,000 to Meta", { x: 7.05, y: 2.95, w: 2.1, h: 0.6, fontSize: 11, color: "D5DEE7" });
  T(s, [{ text: "Please note: ", options: { bold: true, color: NAVY } }, { text: "the ₹10,000 ad spend is not included in the ₹20,000 service fee. Meta may apply taxes on ad spend as per its billing terms.", options: { color: INK } }],
    { x: 0.6, y: 4.2, w: 8.8, h: 0.5, fontSize: 11.5 });
  footer(s, 9);
  s.addNotes("The ad budget never passes through us — you add your own payment method in Meta Ads Manager and can see all spend there. Payment terms for the service fee will be confirmed in writing.");

  // ---------- 10. Month 1 Goal ----------
  s = pres.addSlide(); s.background = { color: NAVY };
  s.addImage({ data: roseLight, x: 7.5, y: -1.9, w: 4.2, h: 4.2 });
  T(s, "MONTH 1 GOAL", { x: 0.6, y: 0.5, w: 5, h: 0.25, fontSize: 10, bold: true, color: "C9B27A", charSpacing: 3 });
  T(s, "A solid foundation,\nand real data to decide the next step", { x: 0.6, y: 0.8, w: 7.2, h: 1.2, fontSize: 28, fontFace: HF, bold: true, color: WHITE });
  const goals = [["01", "Set up", "the foundation"], ["02", "Test", "audience & creatives"], ["03", "Generate", "initial enquiries"], ["04", "Decide", "whether continuing or scaling makes sense"]];
  goals.forEach(([n, h, d], i) => {
    const x = 0.6 + i * 2.25;
    T(s, n, { x, y: 2.45, w: 1, h: 0.3, fontSize: 12, bold: true, color: "C9B27A" });
    T(s, h, { x, y: 2.8, w: 2.05, h: 0.4, fontSize: 19, fontFace: HF, bold: true, color: WHITE });
    T(s, d, { x, y: 3.2, w: 1.95, h: 0.55, fontSize: 11.5, color: "C7D3DF" });
  });
  s.addShape(pres.shapes.LINE, { x: 0.6, y: 4.3, w: 8.8, h: 0, line: { color: "3A5572", width: 0.75 } });
  T(s, "Digital marketing results depend on audience, offer, creative, competition, budget and customer response. Specific rankings, traffic, leads or sales are not guaranteed.",
    { x: 0.6, y: 4.45, w: 8.8, h: 0.5, fontSize: 9.5, color: "9FB2C5", italic: true });
  s.addNotes("Close with the goal: foundation, test, initial enquiries, then a data-based decision together. Invite questions. After the call, send a written summary and the access/information checklist within 24 hours.");

  await pres.writeFile({ fileName: "out/Royal_Compass_Month1_Client_Presentation.pptx" });
  console.log("done");
})();
