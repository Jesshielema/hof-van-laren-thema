const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Hof van Laren — Webshop Voorstel";
pres.author = "Jesse Hielema";

// ─── COLORS ───────────────────────────────────────────────────────────────────
const C = {
  bg:    "F5EEE4",
  bg2:   "EFE8DE",
  red:   "900020",
  gold:  "C9A85C",
  dark:  "1C1410",
  muted: "8B7060",
  white: "FFFFFF",
  black: "000000",
};

// ─── FONTS ────────────────────────────────────────────────────────────────────
const F = {
  head: "Skeena",
  body: "Archivo",
};

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function makeShadow() {
  return { type: "outer", blur: 8, offset: 3, angle: 135, color: C.black, opacity: 0.10 };
}

// Small logo text in top-right corner (for slides 2–10)
function addLogoCorner(slide) {
  // Shield shape
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 8.8, y: 0.12, w: 0.22, h: 0.32,
    fill: { color: C.red }, line: { color: C.red, width: 0 },
  });
  // Logo text
  slide.addText("Hof van Laren", {
    x: 7.3, y: 0.14, w: 2.55, h: 0.28,
    fontSize: 9, fontFace: F.head, color: C.red,
    bold: true, align: "right", margin: 0,
  });
}

// Gold horizontal rule
function addGoldRule(slide, y) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.55, y, w: 8.9, h: 0.03,
    fill: { color: C.gold }, line: { color: C.gold, width: 0 },
  });
}

// Section label above title
function addSectionLabel(slide, text, y) {
  slide.addText(text.toUpperCase(), {
    x: 0.55, y, w: 9, h: 0.22,
    fontSize: 8, fontFace: F.body, color: C.gold,
    bold: true, charSpacing: 4, margin: 0,
  });
}

// Slide title
function addTitle(slide, text, y = 0.52) {
  slide.addText(text, {
    x: 0.55, y, w: 8.9, h: 0.7,
    fontSize: 30, fontFace: F.head, color: C.red,
    bold: true, margin: 0,
  });
}

// Gold checkmark bullet item
function goldItem(text) {
  return { text: `♦  ${text}`, options: { color: C.dark, fontSize: 11, fontFace: F.body, paraSpaceAfter: 4, breakLine: true } };
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — COVER
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  // Left crimson accent bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 5.625,
    fill: { color: C.red }, line: { color: C.red, width: 0 },
  });

  // Bottom gold line
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.38, w: 10, h: 0.06,
    fill: { color: C.gold }, line: { color: C.gold, width: 0 },
  });

  // Logo block — shield shape
  s.addShape(pres.shapes.RECTANGLE, {
    x: 1.8, y: 0.55, w: 0.55, h: 0.75,
    fill: { color: C.red }, line: { color: C.red, width: 0 },
    shadow: makeShadow(),
  });
  s.addText("van", {
    x: 1.8, y: 0.78, w: 0.55, h: 0.3,
    fontSize: 9, fontFace: F.body, color: C.white,
    bold: true, align: "center", margin: 0,
  });

  // Logo name
  s.addText("Hof van Laren", {
    x: 2.45, y: 0.65, w: 5, h: 0.55,
    fontSize: 28, fontFace: F.head, color: C.red,
    bold: true, margin: 0,
  });

  // Divider
  s.addShape(pres.shapes.RECTANGLE, {
    x: 1.8, y: 1.42, w: 3.5, h: 0.035,
    fill: { color: C.gold }, line: { color: C.gold, width: 0 },
  });

  // Main title
  s.addText("Een luxe digitale etalage", {
    x: 1.8, y: 1.6, w: 7.5, h: 1.4,
    fontSize: 42, fontFace: F.head, color: C.red,
    bold: true, margin: 0,
  });

  // Subtitle
  s.addText("Webshop voorstel voor Hof van Laren", {
    x: 1.8, y: 3.1, w: 6.5, h: 0.45,
    fontSize: 14, fontFace: F.body, color: C.muted,
    margin: 0,
  });

  // Bottom line
  s.addText("Opgesteld door Jesse Hielema  ·  Juni 2026  ·  Exclusief BTW", {
    x: 1.8, y: 5.0, w: 7, h: 0.28,
    fontSize: 9, fontFace: F.body, color: C.muted,
    margin: 0,
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — DE MARKT VRAAGT EROM
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addLogoCorner(s);
  addSectionLabel(s, "01 — Markt & kans", 0.18);
  addTitle(s, "De markt vraagt erom", 0.42);
  addGoldRule(s, 1.22);

  // 3 stat blocks
  const blocks = [
    { num: "+11,8%", label: "groei online\nbloemenmarkt per jaar", x: 0.55 },
    { num: "74%",    label: "betaalt meer voor\npremium online beleving", x: 3.72 },
    { num: "12.000", label: "Instagram volgers\nklaar om te converteren", x: 6.88 },
  ];

  blocks.forEach(b => {
    // Card bg
    s.addShape(pres.shapes.RECTANGLE, {
      x: b.x, y: 1.4, w: 2.85, h: 2.9,
      fill: { color: C.white }, line: { color: C.bg2, width: 1 },
      shadow: makeShadow(),
    });
    // Red accent top
    s.addShape(pres.shapes.RECTANGLE, {
      x: b.x, y: 1.4, w: 2.85, h: 0.07,
      fill: { color: C.red }, line: { color: C.red, width: 0 },
    });
    // Number
    s.addText(b.num, {
      x: b.x, y: 1.6, w: 2.85, h: 1.1,
      fontSize: 40, fontFace: F.head, color: C.red,
      bold: true, align: "center", margin: 0,
    });
    // Label
    s.addText(b.label, {
      x: b.x + 0.15, y: 2.75, w: 2.55, h: 0.8,
      fontSize: 11, fontFace: F.body, color: C.muted,
      align: "center", margin: 0,
    });
  });

  // Bottom quote
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.55, w: 10, h: 0.9,
    fill: { color: C.dark }, line: { color: C.dark, width: 0 },
  });
  s.addText("Het aantal bloemenwinkels daalt — 8,6% in 5 jaar. Wie nu online sterk staat, pakt dat marktaandeel.", {
    x: 0.55, y: 4.62, w: 8.9, h: 0.5,
    fontSize: 12, fontFace: F.body, color: C.bg,
    align: "center", margin: 0,
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — DE VISIE
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addLogoCorner(s);
  addSectionLabel(s, "02 — Visie", 0.18);
  addTitle(s, "Niet een webshop. Een digitale etalage.", 0.42);
  addGoldRule(s, 1.22);

  // LEFT column — NIET DIT
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.55, y: 1.4, w: 4.1, h: 2.85,
    fill: { color: C.bg2 }, line: { color: C.bg2, width: 0 },
  });
  s.addText("❌  Niet dit:", {
    x: 0.75, y: 1.55, w: 3.7, h: 0.35,
    fontSize: 13, fontFace: F.body, color: C.muted,
    bold: true, margin: 0,
  });
  s.addText([
    { text: "Snel een boeket bestellen", options: { breakLine: true, strikeThrough: true } },
    { text: "Generieke bloemenwebshop", options: { breakLine: true, strikeThrough: true } },
    { text: "Concurreren met Fleurop", options: { strikeThrough: true } },
  ], {
    x: 0.75, y: 2.05, w: 3.7, h: 1.8,
    fontSize: 13, fontFace: F.body, color: C.muted,
    margin: 0, paraSpaceAfter: 10,
  });

  // RIGHT column — WEL DIT
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.0, y: 1.4, w: 4.45, h: 2.85,
    fill: { color: C.red }, line: { color: C.red, width: 0 },
    shadow: makeShadow(),
  });
  s.addText("♦  Wel dit:", {
    x: 5.2, y: 1.55, w: 4.0, h: 0.35,
    fontSize: 13, fontFace: F.body, color: C.gold,
    bold: true, margin: 0,
  });
  s.addText([
    { text: "Potten, vazen & zijdebloemen verkopen", options: { breakLine: true } },
    { text: "Marchien en haar verhaal centraal", options: { breakLine: true } },
    { text: "Mensen naar de winkel trekken", options: { breakLine: false } },
  ], {
    x: 5.2, y: 2.05, w: 4.0, h: 1.8,
    fontSize: 13, fontFace: F.body, color: C.white,
    margin: 0, paraSpaceAfter: 10,
  });

  // Quote strip at bottom
  s.addText("“Mensen komen niet voor een boeket. Ze komen voor Marchien.”", {
    x: 0.55, y: 4.5, w: 8.9, h: 0.8,
    fontSize: 14, fontFace: F.head, color: C.red,
    italic: true, align: "center", margin: 0,
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — MVP PAKKET
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addLogoCorner(s);
  addSectionLabel(s, "03 — Pakket 1", 0.18);
  addTitle(s, "Het Fundament — MVP", 0.42);
  addGoldRule(s, 1.22);

  // Price block
  s.addText("€ 1.750", {
    x: 0.55, y: 1.3, w: 3.2, h: 0.85,
    fontSize: 48, fontFace: F.head, color: C.red,
    bold: true, margin: 0,
  });
  s.addText("MVP · Opgeleverd binnen 2 weken", {
    x: 0.55, y: 2.15, w: 3.5, h: 0.32,
    fontSize: 11, fontFace: F.body, color: C.gold,
    bold: true, margin: 0,
  });

  // Deliverables — col 1
  const col1 = [
    "Luxe Shopify thema op maat",
    "Homepage met alle kernonderdelen",
    "Winkelpagina (adres, kaart, uren)",
    "Pickup in Store Zuidlaren",
    "Domein koppeling",
  ];
  const col2 = [
    "iDEAL & betaalmethoden",
    "Eerste collectie producten ingevoerd",
    "1 seizoenstemplate (kerst)",
    "Instructievideo producten toevoegen",
  ];

  col1.forEach((item, i) => {
    s.addText(`♦  ${item}`, {
      x: 0.55, y: 2.6 + i * 0.38, w: 4.3, h: 0.32,
      fontSize: 11, fontFace: F.body, color: C.dark, margin: 0,
    });
  });
  col2.forEach((item, i) => {
    s.addText(`♦  ${item}`, {
      x: 5.05, y: 2.6 + i * 0.38, w: 4.45, h: 0.32,
      fontSize: 11, fontFace: F.body, color: C.dark, margin: 0,
    });
  });

  // Timeline bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.55, y: 4.95, w: 4.35, h: 0.5,
    fill: { color: C.dark }, line: { color: C.dark, width: 0 },
  });
  s.addText("Week 1 — Design & fundament", {
    x: 0.7, y: 4.97, w: 4.1, h: 0.36,
    fontSize: 11, fontFace: F.body, color: C.white, bold: true, margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 4.95, w: 4.35, h: 0.5,
    fill: { color: C.red }, line: { color: C.red, width: 0 },
  });
  s.addText("Week 2 — Producten & lancering", {
    x: 5.25, y: 4.97, w: 4.1, h: 0.36,
    fontSize: 11, fontFace: F.body, color: C.white, bold: true, margin: 0,
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 5 — GROEI PAKKET
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addLogoCorner(s);
  addSectionLabel(s, "04 — Pakket 2", 0.18);
  addTitle(s, "De Volledige Etalage", 0.42);
  addGoldRule(s, 1.22);

  // AANBEVOLEN badge
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.5, y: 0.4, w: 1.95, h: 0.35,
    fill: { color: C.red }, line: { color: C.red, width: 0 },
  });
  s.addText("⭐  AANBEVOLEN", {
    x: 7.5, y: 0.42, w: 1.95, h: 0.3,
    fontSize: 9, fontFace: F.body, color: C.white,
    bold: true, align: "center", margin: 0,
  });

  // Price
  s.addText("€ 3.500", {
    x: 0.55, y: 1.3, w: 3.5, h: 0.85,
    fontSize: 48, fontFace: F.head, color: C.red,
    bold: true, margin: 0,
  });
  s.addText("4–6 weken · Inclusief MVP Pakket", {
    x: 0.55, y: 2.15, w: 3.8, h: 0.32,
    fontSize: 11, fontFace: F.body, color: C.muted, margin: 0,
  });

  const left = [
    "Alles uit Pakket 1",
    "Storytelling modules (Marchien kiest, Uit Denemarken)",
    "Alle 7 seizoenspagina’s",
    "Productpagina’s met sfeerbeelden",
    "Filters: prijs, kleur, hoogte, materiaal",
  ];
  const right = [
    "“Boek adviesgesprek met Marchien” flow",
    "Instagram Shopping koppeling",
    "Nieuwsbrief koppeling",
    "Google Business optimalisatie",
    "Uitgebreide instructiesessie (2 uur)",
  ];

  left.forEach((item, i) => {
    s.addText(`♦  ${item}`, {
      x: 0.55, y: 2.6 + i * 0.42, w: 4.3, h: 0.35,
      fontSize: 10.5, fontFace: F.body, color: C.dark, margin: 0,
    });
  });
  right.forEach((item, i) => {
    s.addText(`♦  ${item}`, {
      x: 5.05, y: 2.6 + i * 0.42, w: 4.45, h: 0.35,
      fontSize: 10.5, fontFace: F.body, color: C.dark, margin: 0,
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 6 — PREMIUM PAKKET
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addLogoCorner(s);
  addSectionLabel(s, "05 — Pakket 3", 0.18);
  addTitle(s, "Het Complete Merk", 0.42);
  addGoldRule(s, 1.22);

  // Price
  s.addText("€ 5.500", {
    x: 0.55, y: 1.3, w: 3.5, h: 0.85,
    fontSize: 48, fontFace: F.head, color: C.red, bold: true, margin: 0,
  });
  s.addText("6–8 weken · Alles inbegrepen", {
    x: 0.55, y: 2.15, w: 3.5, h: 0.32,
    fontSize: 11, fontFace: F.body, color: C.muted, margin: 0,
  });

  const items = [
    "Alles uit Pakket 2",
    "Fotografie-briefing voor professionele shoot",
    "Instagram contentstrategie",
    "Wishlist / Favorieten functionaliteit",
    "Local Inventory voor Google Shopping",
    "Cadeaubon systeem",
    "SEO setup (titels, meta, structured data)",
    "3 maanden begeleiding (1x/maand)",
    "Prioriteit support",
  ];

  const half = Math.ceil(items.length / 2);
  items.slice(0, half).forEach((item, i) => {
    s.addText(`♦  ${item}`, {
      x: 0.55, y: 2.6 + i * 0.38, w: 4.3, h: 0.32,
      fontSize: 10.5, fontFace: F.body, color: C.dark, margin: 0,
    });
  });
  items.slice(half).forEach((item, i) => {
    s.addText(`♦  ${item}`, {
      x: 5.05, y: 2.6 + i * 0.38, w: 3.2, h: 0.32,
      fontSize: 10.5, fontFace: F.body, color: C.dark, margin: 0,
    });
  });

  // Dark panel right
  s.addShape(pres.shapes.RECTANGLE, {
    x: 8.4, y: 1.3, w: 1.45, h: 4.1,
    fill: { color: C.dark }, line: { color: C.dark, width: 0 },
  });
  s.addText("Inclusief\n3 maanden\nbegeleiding", {
    x: 8.4, y: 2.3, w: 1.45, h: 1.5,
    fontSize: 11, fontFace: F.body, color: C.gold,
    bold: true, align: "center", margin: 0,
  });
  s.addText("♦", {
    x: 8.4, y: 3.9, w: 1.45, h: 0.4,
    fontSize: 16, fontFace: F.body, color: C.gold,
    align: "center", margin: 0,
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 7 — MAANDELIJKS BEHEER
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addLogoCorner(s);
  addSectionLabel(s, "06 — Doorlopend beheer", 0.18);
  addTitle(s, "Na de lancering", 0.42);
  addGoldRule(s, 1.22);

  // CARD 1 — light
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.55, y: 1.4, w: 4.1, h: 3.95,
    fill: { color: C.bg2 }, line: { color: C.bg2, width: 0 },
    shadow: makeShadow(),
  });
  s.addText("Basisbeheer", {
    x: 0.75, y: 1.55, w: 3.7, h: 0.38,
    fontSize: 16, fontFace: F.head, color: C.dark, bold: true, margin: 0,
  });
  s.addText("€ 149", {
    x: 0.75, y: 1.95, w: 2, h: 0.55,
    fontSize: 32, fontFace: F.head, color: C.red, bold: true, margin: 0,
  });
  s.addText("per maand · excl. BTW", {
    x: 0.75, y: 2.55, w: 3.5, h: 0.28,
    fontSize: 10, fontFace: F.body, color: C.muted, margin: 0,
  });
  [
    "Seizoenstemplate wisselen (4x/jaar)",
    "Max. 5 nieuwe producten per maand",
    "Technische Shopify updates",
    "E-mail support binnen 48 uur",
  ].forEach((item, i) => {
    s.addText(`•  ${item}`, {
      x: 0.75, y: 3.0 + i * 0.36, w: 3.7, h: 0.3,
      fontSize: 10.5, fontFace: F.body, color: C.dark, margin: 0,
    });
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.55, y: 5.1, w: 4.1, h: 0.25,
    fill: { color: C.bg }, line: { color: C.bg, width: 0 },
  });
  s.addText("±2 uur per maand · daarboven €65/uur", {
    x: 0.75, y: 5.12, w: 3.5, h: 0.22,
    fontSize: 9, fontFace: F.body, color: C.muted, margin: 0,
  });

  // CARD 2 — dark
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.05, y: 1.4, w: 4.4, h: 3.95,
    fill: { color: C.dark }, line: { color: C.dark, width: 0 },
    shadow: makeShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.05, y: 1.4, w: 4.4, h: 0.07,
    fill: { color: C.gold }, line: { color: C.gold, width: 0 },
  });
  s.addText("Groei & Zichtbaarheid", {
    x: 5.25, y: 1.56, w: 4.0, h: 0.38,
    fontSize: 16, fontFace: F.head, color: C.white, bold: true, margin: 0,
  });
  s.addText("€ 295", {
    x: 5.25, y: 1.96, w: 2, h: 0.55,
    fontSize: 32, fontFace: F.head, color: C.gold, bold: true, margin: 0,
  });
  s.addText("per maand · excl. BTW", {
    x: 5.25, y: 2.56, w: 3.5, h: 0.28,
    fontSize: 10, fontFace: F.body, color: C.muted, margin: 0,
  });
  [
    "Alles uit Basisbeheer",
    "Onbeperkt producten toevoegen",
    "Alle seizoenspagina’s actueel",
    "Instagram koppelingen bijhouden",
    "Maandelijks voortgangsgesprek",
  ].forEach((item, i) => {
    s.addText(`♦  ${item}`, {
      x: 5.25, y: 3.0 + i * 0.36, w: 4.0, h: 0.3,
      fontSize: 10.5, fontFace: F.body, color: C.white, margin: 0,
    });
  });
  s.addText("±4 uur per maand · daarboven €65/uur", {
    x: 5.25, y: 5.12, w: 3.7, h: 0.22,
    fontSize: 9, fontFace: F.body, color: C.muted, margin: 0,
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 8 — LOSSE DIENSTEN
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addLogoCorner(s);
  addSectionLabel(s, "07 — Optioneel", 0.18);
  addTitle(s, "Losse diensten bij te boeken", 0.42);
  addGoldRule(s, 1.22);

  const services = [
    { price: "€ 250",  title: "Extra seizoenspagina",    desc: "Volledig uitgewerkte landingspagina voor één gelegenheid" },
    { price: "€ 195",  title: "Instagram Shopping",      desc: "Shopify koppelen, producten tagbaar maken in posts en stories" },
    { price: "€ 175",  title: "Fotografie-briefing",     desc: "Document voor fotograaf: stijl, licht en sfeer per product" },
    { price: "€ 350",  title: "SEO basissetup",          desc: "Titels, meta-descriptions en Google Search Console koppeling" },
    { price: "€ 125",  title: "Cadeaubon systeem",       desc: "Digitale cadeaubonnen die klanten online kunnen kopen" },
    { price: "€ 65/u", title: "Maatwerkuren",            desc: "Uurtarief voor alles buiten de pakketten. Min. 1 uur." },
  ];

  const cols = [0.55, 3.6, 6.65];
  const rows = [1.4, 3.1];

  services.forEach((svc, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = cols[col];
    const y = rows[row];

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.85, h: 1.55,
      fill: { color: C.white }, line: { color: C.bg2, width: 1 },
      shadow: makeShadow(),
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.06, h: 1.55,
      fill: { color: C.red }, line: { color: C.red, width: 0 },
    });
    s.addText(svc.price, {
      x: x + 0.15, y: y + 0.1, w: 2.6, h: 0.42,
      fontSize: 22, fontFace: F.head, color: C.red, bold: true, margin: 0,
    });
    s.addText(svc.title, {
      x: x + 0.15, y: y + 0.54, w: 2.6, h: 0.28,
      fontSize: 11, fontFace: F.body, color: C.dark, bold: true, margin: 0,
    });
    s.addText(svc.desc, {
      x: x + 0.15, y: y + 0.84, w: 2.6, h: 0.6,
      fontSize: 9, fontFace: F.body, color: C.muted, margin: 0,
    });
  });

  // Note at bottom
  s.addText("Alle prijzen zijn exclusief BTW · Losse diensten worden gefactureerd op basis van offerte of uurtarief", {
    x: 0.55, y: 4.9, w: 8.9, h: 0.35,
    fontSize: 9, fontFace: F.body, color: C.muted, align: "center", margin: 0,
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 9 — WERKWIJZE & VOORWAARDEN
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  addLogoCorner(s);
  addSectionLabel(s, "08 — Samenwerking", 0.18);
  addTitle(s, "Hoe we samenwerken", 0.42);
  addGoldRule(s, 1.22);

  const conditions = [
    { icon: "💳", title: "Betaling",           desc: "50% bij akkoord, 50% bij oplevering. Maandelijkse diensten vooraf per maand." },
    { icon: "📅", title: "Opzegtermijn",       desc: "Maandelijks opzegbaar, 1 maand van tevoren doorgeven per e-mail." },
    { icon: "📸", title: "Content aanlevering", desc: "Foto’s en teksten door de klant aangeleverd, tenzij anders afgesproken." },
    { icon: "🔄", title: "Revisierondes",       desc: "Elk pakket bevat 2 revisierondes. Extra revisies op uurbasis €65/uur." },
    { icon: "🏠", title: "Eigendom",            desc: "Na volledige betaling volledig eigendom van de klant, inclusief alle bestanden." },
    { icon: "💻", title: "Shopify abonnement",  desc: "±€32/maand voor rekening klant · Basic plan voldoet voor de start." },
  ];

  const cols2 = [0.55, 5.05];
  conditions.forEach((c, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = cols2[col];
    const y = 1.4 + row * 1.15;
    const w = 4.25;

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h: 1.0,
      fill: { color: C.white }, line: { color: C.bg2, width: 1 },
    });
    // Gold left border
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.055, h: 1.0,
      fill: { color: C.gold }, line: { color: C.gold, width: 0 },
    });
    s.addText(c.icon, {
      x: x + 0.15, y: y + 0.1, w: 0.45, h: 0.4,
      fontSize: 18, margin: 0,
    });
    s.addText(c.title, {
      x: x + 0.65, y: y + 0.1, w: w - 0.8, h: 0.3,
      fontSize: 11, fontFace: F.body, color: C.dark, bold: true, margin: 0,
    });
    s.addText(c.desc, {
      x: x + 0.65, y: y + 0.44, w: w - 0.8, h: 0.45,
      fontSize: 10, fontFace: F.body, color: C.muted, margin: 0,
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 10 — AFSLUITING
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };

  // Large crimson circle
  s.addShape(pres.shapes.OVAL, {
    x: 3.2, y: -1.5, w: 7, h: 7,
    fill: { color: C.red, transparency: 85 }, line: { color: C.red, width: 0 },
  });

  // Main title
  s.addText("Klaar om te beginnen?", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.4,
    fontSize: 44, fontFace: F.head, color: C.white,
    bold: true, align: "center", margin: 0,
  });

  // Gold line
  s.addShape(pres.shapes.RECTANGLE, {
    x: 3.8, y: 2.75, w: 2.4, h: 0.04,
    fill: { color: C.gold }, line: { color: C.gold, width: 0 },
  });

  // Sub
  s.addText("MVP live binnen 2 weken na akkoord", {
    x: 0.8, y: 2.9, w: 8.4, h: 0.4,
    fontSize: 13, fontFace: F.body, color: C.muted,
    align: "center", margin: 0,
  });

  // Contact
  s.addText("Jesse Hielema", {
    x: 0.8, y: 3.55, w: 8.4, h: 0.42,
    fontSize: 18, fontFace: F.head, color: C.gold,
    bold: true, align: "center", margin: 0,
  });
  s.addText("jessehielema@gmail.com  ·  ZZP Webdesign & Shopify", {
    x: 0.8, y: 4.0, w: 8.4, h: 0.35,
    fontSize: 12, fontFace: F.body, color: C.muted,
    align: "center", margin: 0,
  });

  // Logo
  s.addText("Hof van Laren", {
    x: 3.0, y: 4.55, w: 4, h: 0.4,
    fontSize: 13, fontFace: F.head, color: "F5EEE4",
    bold: true, align: "center", margin: 0, transparency: 40,
  });

  // Fine print
  s.addText("Voorstel geldig 30 dagen na datum  ·  Exclusief BTW", {
    x: 0.8, y: 5.2, w: 8.4, h: 0.25,
    fontSize: 8.5, fontFace: F.body, color: C.muted,
    align: "center", margin: 0,
  });
}

// ─── WRITE ────────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "Hof-van-Laren-Voorstel.pptx" })
  .then(() => console.log("✅ Presentatie opgeslagen: Hof-van-Laren-Voorstel.pptx"))
  .catch(err => console.error("❌ Fout:", err));
