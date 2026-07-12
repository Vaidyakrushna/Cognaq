const pptxgen = require("pptxgenjs");

let pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";

// ─── STRICT MINIMALIST PALETTE (GRAY, WHITE, #9D2706) ─────────────
const C = {
  brand:      "9D2706",   // Deep crimson red (used only for accents, text, icons)
  white:      "FFFFFF",   // Pure white for slide/card backgrounds
  bgLight:    "F5F5F7",   // Light gray/off-white for slide backgrounds
  textDark:   "2C3E50",   // Slate-charcoal for primary text (no pure black)
  textMedium: "5A6B7C",   // Medium slate-gray for descriptions
  textLight:  "8A9BA8",   // Light slate-gray for metadata/subtitles
  borderGray: "E2E8F0",   // Slate-border gray
  fillGray:   "EDF2F7",   // Light filled gray for header accents / tables
};

// ─── HELPERS ─────────────────────────────────────────────────────

// Section header slides (light gray background, minimal branding)
function sectionSlide(title, step = "") {
  let s = pptx.addSlide();
  s.background = { color: C.bgLight };
  
  // Left border strip in crimson
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.15, h:"100%", fill:{ color: C.brand } });
  
  if (step) {
    s.addText(step, { x:0.8, y:1.8, w:6, fontSize:72, bold:true, color: C.borderGray });
  }
  s.addText(title, { x:0.9, y:2.4, w:11, fontSize:32, bold:true, color: C.textDark, lineSpacingMultiple:1.3 });
  s.addShape(pptx.ShapeType.rect, { x:0.9, y:3.6, w:3.5, h:0.06, fill:{ color: C.brand } });
  
  s.addText("COGNAQ  |  3D & AR PLATFORM ROADMAP", { x:0.9, y:6.2, w:8, fontSize:10, color: C.textLight, bold:true, charSpacing:2 });
  return s;
}

// Standard content slide (light header, clean borders)
function contentSlide(title, subtitle = "") {
  let s = pptx.addSlide();
  s.background = { color: C.bgLight };
  
  // Left brand accent strip
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.1, h:"100%", fill:{ color: C.brand } });
  
  // Top header bar (pure white)
  s.addShape(pptx.ShapeType.rect, { x:0.1, y:0, w:"99%", h:0.9, fill:{ color: C.white } });
  s.addShape(pptx.ShapeType.rect, { x:0.1, y:0.88, w:"99%", h:0.03, fill:{ color: C.borderGray } });
  
  // Logo
  s.addText("COGNAQ", { x:0.35, y:0.2, w:2.5, h:0.5, fontSize:12, bold:true, color: C.brand, charSpacing:4 });
  // Title
  s.addText(title, { x:3.0, y:0.12, w:9.7, h:0.65, fontSize:22, bold:true, color: C.textDark, align:"right" });
  
  // Subtitle
  if (subtitle) {
    s.addText(subtitle, { x:0.35, y:0.95, w:"92%", fontSize:11, color: C.textMedium, italic:true });
  }
  
  // Footer
  s.addShape(pptx.ShapeType.rect, { x:0.1, y:6.9, w:"99%", h:0.1, fill:{ color: C.white } });
  s.addShape(pptx.ShapeType.rect, { x:0.1, y:6.88, w:"99%", h:0.02, fill:{ color: C.borderGray } });
  s.addText("PROJECT IMPLEMENTATION PLAN", { x:0.35, y:6.91, w:5, fontSize:8, bold:true, color: C.textLight, charSpacing:2 });
  
  return s;
}

// Card with thin border and white fill
function card(slide, x, y, w, h, bg = C.white, accentColor = C.borderGray) {
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill:{ color: bg }, line:{ color: accentColor, pt:1 } });
}

// Accent badge
function badge(slide, x, y, w, label, accent = false) {
  let bg = accent ? C.brand : C.fillGray;
  let fg = accent ? C.white : C.textDark;
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h:0.38, fill:{ color: bg }, line:{ color: C.borderGray, pt:0.5 } });
  slide.addText(label, { x, y:y+0.04, w, h:0.3, fontSize:9.5, bold:true, color: fg, align:"center" });
}

// Left accent bar
function accentBar(slide, x, y, h = 0.6) {
  slide.addShape(pptx.ShapeType.rect, { x, y, w:0.06, h, fill:{ color: C.brand } });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 1: COVER
// ═══════════════════════════════════════════════════════════════
{
  let s = pptx.addSlide();
  s.background = { color: C.bgLight };
  
  // Decorative layout boxes
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.2, h:"100%", fill:{ color: C.brand } });
  s.addShape(pptx.ShapeType.rect, { x:0.2, y:0, w:5.0, h:"100%", fill:{ color: C.white } });
  s.addShape(pptx.ShapeType.rect, { x:5.2, y:0, w:0.03, h:"100%", fill:{ color: C.borderGray } });
  
  // Logo & Title Left
  s.addText("🧊", { x:0.8, y:1.5, w:3.8, h:1.2, fontSize:64, align:"center" });
  s.addText("COGNAQ", { x:0.8, y:2.8, w:3.8, h:0.8, fontSize:36, bold:true, color: C.brand, align:"center", charSpacing:6 });
  s.addShape(pptx.ShapeType.rect, { x:1.2, y:3.65, w:3.0, h:0.04, fill:{ color: C.brand } });
  s.addText("3D CONFIGURATOR & AR PLATFORM", { x:0.8, y:3.8, w:3.8, fontSize:10.5, color: C.textDark, bold:true, align:"center", charSpacing:1.5 });
  
  // Right side titles
  s.addText("Product Roadmap & Business Plan", { x:5.8, y:1.6, w:7.2, fontSize:32, bold:true, color: C.textDark });
  s.addText("End-to-End Visualisation SaaS & 3D Services", { x:5.8, y:2.35, w:7.2, fontSize:16, color: C.brand, bold:true });
  s.addShape(pptx.ShapeType.rect, { x:5.8, y:2.85, w:7.0, h:0.04, fill:{ color: C.brand } });
  
  s.addText("A structured strategic blueprint detailing our solution architecture, target industry verticals, 3D modelling asset pipelines, competitive differentiators, and phased implementation timeline.", {
    x:5.8, y:3.1, w:7.2, fontSize:12.5, color: C.textMedium, lineSpacingMultiple:1.6
  });
  
  // Badges
  let tags = ["🏠 Room AR", "👟 Virtual Try-On", "🛒 Shopify Native App", "🎨 3D Asset Service"];
  tags.forEach((t, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    s.addShape(pptx.ShapeType.rect, { x:5.8+col*3.5, y:4.8+row*0.6, w:3.2, h:0.45, fill:{ color: C.white }, line:{ color: C.borderGray, pt:1 } });
    s.addText(t, { x:6.0+col*3.5, y:4.9, w:2.9, fontSize:11, color: C.textDark, bold:true });
  });
  
  s.addText("July 2026  |  Version 3.0  |  github.com/Vaidyakrushna/Cognaq", {
    x:5.8, y:6.35, w:7.2, fontSize:10, color: C.textLight, italic:true
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 2: TABLE OF CONTENTS
// ═══════════════════════════════════════════════════════════════
{
  let s = contentSlide("Table of Contents");
  let items = [
    ["01", "Executive Summary",        "The big picture and strategic vision"],
    ["02", "Problem Statement",        "The market challenges we address"],
    ["03", "Our Solution",             "What Cognaq delivers end-to-end"],
    ["04", "Target Industries",        "Market opportunity and verticals"],
    ["05", "AR Technology Deep Dive",  "Room AR and Virtual Try-On mechanics"],
    ["06", "Technology Stack",         "Full client, server, and AR architecture"],
    ["07", "Project Roadmap",          "Phased week-by-week implementation"],
    ["08", "3D Asset Pipeline",        "End-to-end model creation workflow"],
    ["09", "Competitor Analysis",      "How Cognaq compares with market options"],
    ["10", "Business Model",           "Subscription pricing and revenue streams"],
    ["11", "Technical Requirements",   "Core specs, scaling, and compliance"],
    ["12", "Success Metrics",          "KPIs and target launch deliverables"],
  ];
  items.forEach((item, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = 0.35 + col * 6.55;
    let y = 1.15 + row * 0.93;
    card(s, x, y, 6.2, 0.8, C.white);
    s.addShape(pptx.ShapeType.rect, { x, y, w:0.75, h:0.8, fill:{ color: C.fillGray } });
    s.addText(item[0], { x, y:y+0.22, w:0.75, fontSize:14, bold:true, color: C.brand, align:"center" });
    s.addText(item[1], { x:x+0.9, y:y+0.1, w:4.8, fontSize:13, bold:true, color: C.textDark });
    s.addText(item[2], { x:x+0.9, y:y+0.45, w:4.8, fontSize:11, color: C.textMedium });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 3: EXECUTIVE SUMMARY
// ═══════════════════════════════════════════════════════════════
{
  let s = contentSlide("Executive Summary", "What is Cognaq? — The platform at a glance.");
  accentBar(s, 0.35, 1.1, 1.2);
  s.addText("Cognaq is a next-generation 3D Product Configurator and Augmented Reality (AR) SaaS platform. We enable online brands to let their customers interact with, customise, and virtually try on products directly from their mobile or desktop web browser, driving sales conversion while reducing returns.", {
    x:0.55, y:1.05, w:12.4, h:1.3, fontSize:13.5, color: C.textDark, italic:true, lineSpacingMultiple:1.5
  });
  
  let pillars = [
    { icon:"🧊", title:"Interactive 3D Configurator", desc:"Real-time customisation of colors, parts, and materials directly in-browser. Zero plugins required." },
    { icon:"🪄", title:"Augmented Reality (AR)",      desc:"Instant 'View in Space' placement and Virtual Try-On powered by WebXR and on-device AI." },
    { icon:"🛒", title:"Shopify Integration",         desc:"Native Shopify app installation. Syncs dynamic custom variants to cart and checkout instantly." },
    { icon:"🎨", title:"End-to-End Asset Service",    desc:"Full service model: we partner with experienced 3D artists to model and digitise client catalogues." },
    { icon:"📊", title:"Merchant SaaS Dashboard",     desc:"Self-serve admin panel where merchants upload models, configure options, and see analytics." },
    { icon:"💰", title:"Dynamic Pricing Engine",      desc:"Material and accessory upgrades dynamically recalculate total price in the shopper's view." },
  ];
  pillars.forEach((p, i) => {
    let col = i % 3;
    let row = Math.floor(i / 3);
    let x = 0.35 + col * 4.25;
    let y = 2.65 + row * 2.05;
    card(s, x, y, 3.95, 1.85, C.white);
    s.addShape(pptx.ShapeType.rect, { x, y, w:3.95, h:0.55, fill:{ color: C.fillGray } });
    s.addText(p.icon + "  " + p.title, { x:x+0.15, y:y+0.12, w:3.7, fontSize:13, bold:true, color: C.textDark });
    s.addShape(pptx.ShapeType.rect, { x, y:y+0.55, w:4.0, h:0.04, fill:{ color: C.brand } });
    s.addText(p.desc, { x:x+0.15, y:y+0.72, w:3.7, fontSize:11.5, color: C.textMedium, lineSpacingMultiple:1.4 });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 4: PROBLEM STATEMENT
// ═══════════════════════════════════════════════════════════════
{
  let s = contentSlide("The Problem We Solve", "Market pain points for e-commerce merchants.");
  
  // Highlight bar
  s.addShape(pptx.ShapeType.rect, { x:0.35, y:1.05, w:12.7, h:0.65, fill:{ color: C.fillGray }, line:{ color: C.borderGray, pt:1 } });
  s.addShape(pptx.ShapeType.rect, { x:0.35, y:1.05, w:0.4, h:0.65, fill:{ color: C.brand } });
  s.addText("E-commerce brands lose billions due to product expectation mismatches, return logistics, and low cart conversions.", {
    x:0.9, y:1.13, w:11.9, fontSize:12.5, bold:true, color: C.textDark
  });
  
  let problems = [
    { icon:"👟", stat:"30–40%", unit:"Return Rate",  title:"Expectation Vs Reality Gap", desc:"Online buyers cannot examine material textures or gauge scale. This leads to massive return rates, particularly in footwear, bags, and home furnishings." },
    { icon:"📊", stat:"<1%",    unit:"Adoption",     title:"Technical & Cost Barrier",   desc:"While conversion lift is proven, small and mid-sized brands are locked out. They lack internal developers to build custom 3D web engines." },
    { stat:"$10K",   unit:"Per Month",    title:"Prohibitive Enterprise Pricing", desc:"Incumbents like Threekit charge enterprise licensing fees. SMB merchants cannot afford to deploy these features." },
    { icon:"🎨", stat:"80%",    unit:"No Models",    title:"The Asset Creation Bottleneck",   desc:"Merchants struggle to create high-quality 3D assets. Without .glb/USDZ model files, configurator software is unusable." },
  ];
  
  problems.forEach((p, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = 0.35 + col * 6.55;
    let y = 2.05 + row * 2.3;
    card(s, x, y, 6.2, 2.1, C.white);
    
    s.addShape(pptx.ShapeType.rect, { x, y, w:1.6, h:2.1, fill:{ color: C.fillGray } });
    s.addText(p.stat, { x, y:y+0.35, w:1.6, fontSize:22, bold:true, color: C.brand, align:"center" });
    s.addText(p.unit, { x, y:y+0.9, w:1.6, fontSize:10.5, color: C.textMedium, align:"center" });
    
    s.addShape(pptx.ShapeType.rect, { x:x+1.6, y, w:0.04, h:2.1, fill:{ color: C.brand } });
    
    let prefix = p.icon ? p.icon + " " : "";
    s.addText(prefix + p.title, { x:x+1.75, y:y+0.12, w:4.3, fontSize:13, bold:true, color: C.textDark });
    s.addText(p.desc, { x:x+1.75, y:y+0.52, w:4.3, h:1.4, fontSize:11, color: C.textMedium, lineSpacingMultiple:1.4 });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 5: OUR SOLUTION
// ═══════════════════════════════════════════════════════════════
{
  let s = contentSlide("Our Solution", "Cognaq — Product configuration + AR + 3D modeling as a service.");
  
  s.addShape(pptx.ShapeType.rect, { x:0.35, y:1.05, w:12.7, h:0.65, fill:{ color: C.white }, line:{ color: C.brand, pt:1 } });
  s.addText("Cognaq bridges the market gap: A high-fidelity, native Shopify integration at an affordable subscription price.", {
    x:0.55, y:1.13, w:12.3, fontSize:12.5, bold:true, color: C.brand
  });
  
  // Step workflow
  let flow = [
    { label:"1. Onboard",    icon:"🏢" },
    { label:"2. Create 3D",  icon:"🎨" },
    { label:"3. Set Rules",  icon:"⚙️" },
    { label:"4. Embed App",  icon:"🚀" },
    { label:"5. Customise",  icon:"🧊" },
    { label:"6. View AR",    icon:"🪄" },
    { label:"7. Checkout",   icon:"🛒" },
  ];
  flow.forEach((f, i) => {
    let x = 0.35 + i * 1.83;
    s.addShape(pptx.ShapeType.rect, { x, y:2.0, w:1.6, h:0.9, fill:{ color: C.white }, line:{ color: C.borderGray, pt:1 } });
    s.addShape(pptx.ShapeType.rect, { x, y:2.0, w:0.06, h:0.9, fill:{ color: C.brand } });
    s.addText(f.icon + "  " + f.label, { x, y:2.1, w:1.6, h:0.7, fontSize:10.5, bold:true, color: C.textDark, align:"center" });
    if (i < flow.length - 1) s.addText("→", { x:x+1.62, y:2.25, w:0.2, fontSize:16, bold:true, color: C.brand, align:"center" });
  });
  
  // KPI Benefit Boxes
  let benefits = [
    { stat:"94%", label:"⚡ Higher Conversions", sub:"average lift observed on Shopify pages with interactive AR" },
    { stat:"38%", label:"📉 Fewer Returns",       sub:"when shoppers inspect textures and proportions in 3D" },
    { stat:"10×", label:"💵 Lower Platform Cost", sub:"makes high-fidelity configurators viable for SMB merchants" },
    { stat:"0",   label:"🛠️ Code Required",       sub:"one-click app store install auto-injects configurator" },
  ];
  benefits.forEach((b, i) => {
    let x = 0.35 + i * 3.25;
    card(s, x, 3.15, 3.0, 3.5, C.white);
    s.addShape(pptx.ShapeType.rect, { x, y:3.15, w:3.0, h:0.06, fill:{ color: C.brand } });
    s.addText(b.stat, { x:x+0.1, y:3.3, w:2.8, fontSize:36, bold:true, color: C.brand, align:"center" });
    s.addText(b.label, { x:x+0.1, y:4.1, w:2.8, fontSize:12.5, bold:true, color: C.textDark, align:"center" });
    s.addShape(pptx.ShapeType.rect, { x:x+0.6, y:4.6, w:1.8, h:0.04, fill:{ color: C.borderGray } });
    s.addText(b.sub, { x:x+0.1, y:4.72, w:2.8, fontSize:10.5, color: C.textMedium, align:"center", lineSpacingMultiple:1.4 });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 6: TARGET MARKET
// ═══════════════════════════════════════════════════════════════
{
  let s = contentSlide("Target Verticals & Opportunity", "Focusing on visual-first product categories.");
  let industries = [
    { icon:"👟", name:"Footwear & Sneakers",    why:"Try-on AR displays shoes on feet; real-time colourway swaps." },
    { icon:"👔", name:"Apparel & Fashion",       why:"Virtual try-on matches garments; material textures render in-camera." },
    { icon:"👜", name:"Bags & Accessories",      why:"360° product rotations; configures clasps, straps, and leather types." },
    { icon:"🛋️", name:"Furniture & Home Décor", why:"Room placement AR ensures size fits; modular frames and fabrics config." },
    { icon:"💍", name:"Jewelry & Watches",       why:"High-res details; custom metal types, bands, and gemstone swaps." },
    { icon:"🏠", name:"Interior Architecture",  why:"Configures custom fixtures, cabinetry, and home decor layout in AR." },
  ];
  industries.forEach((ind, i) => {
    let col = i % 3;
    let row = Math.floor(i / 3);
    let x = 0.35 + col * 4.25;
    let y = 1.15 + row * 2.1;
    card(s, x, y, 3.95, 1.9, C.white);
    s.addShape(pptx.ShapeType.rect, { x, y, w:3.95, h:0.55, fill:{ color: C.fillGray } });
    s.addShape(pptx.ShapeType.rect, { x, y, w:0.07, h:1.9, fill:{ color: C.brand } });
    s.addText(ind.icon + "  " + ind.name, { x:x+0.18, y:y+0.12, w:3.6, fontSize:12.5, bold:true, color: C.textDark });
    s.addText("Why 3D/AR is critical:", { x:x+0.18, y:y+0.72, w:3.6, fontSize:10, bold:true, color: C.brand });
    s.addText(ind.why, { x:x+0.18, y:y+0.98, w:3.6, fontSize:11.5, color: C.textMedium, lineSpacingMultiple:1.3 });
  });
  
  // Stat strip (gray background)
  s.addShape(pptx.ShapeType.rect, { x:0.1, y:5.55, w:"99%", h:1.35, fill:{ color: C.white }, line:{ color: C.borderGray, pt:1 } });
  s.addShape(pptx.ShapeType.rect, { x:0.1, y:5.55, w:0.08, h:1.35, fill:{ color: C.brand } });
  let stats = [
    { val:"$8.7B", label:"Global 3D/AR Commerce\nMarket Value (2026)" },
    { val:"34% CAGR", label:"Annual growth rate of\nAR tech in retail space" },
    { val:"2.9B",  label:"Active AR consumers\nglobally by 2026" },
    { val:"$550B", label:"Value lost in e-commerce\nproduct returns annually" },
  ];
  stats.forEach((st, i) => {
    let x = 0.5 + i * 3.25;
    if (i > 0) s.addShape(pptx.ShapeType.rect, { x, y:5.68, w:0.02, h:1.1, fill:{ color: C.borderGray } });
    s.addText(st.val, { x:x+0.1, y:5.62, w:3.0, fontSize:22, bold:true, color: C.brand, align:"center" });
    s.addText(st.label, { x:x+0.1, y:6.1, w:3.0, fontSize:10.5, color: C.textMedium, align:"center", lineSpacingMultiple:1.3 });
  });
}

// ═══════════════════════════════════════════════════════════════
// SECTION: AR DEEP DIVE
// ═══════════════════════════════════════════════════════════════
sectionSlide("Augmented Reality\nTechnology Deep Dive", "05");

// ─── AR SLIDE A: ROOM AR ─────────────────────────────────────
{
  let s = contentSlide("AR Experience 1: View in Your Space", "WebXR-powered — no app download required.");
  
  // Left Panel
  card(s, 0.35, 1.05, 7.5, 5.75, C.white);
  s.addShape(pptx.ShapeType.rect, { x:0.35, y:1.05, w:7.5, h:0.55, fill:{ color: C.fillGray } });
  s.addShape(pptx.ShapeType.rect, { x:0.35, y:1.05, w:0.07, h:5.75, fill:{ color: C.brand } });
  s.addText("🏠  View in Your Space — Room AR Workflow", { x:0.6, y:1.12, w:6.9, fontSize:13.5, bold:true, color: C.textDark });
  
  let steps = [
    "Shopper clicks the 'View in AR' button on the product page.",
    "Browser camera permissions are requested (Safari iOS / Chrome Android).",
    "On-device WebXR session detects horizontal surfaces (floors/tables).",
    "The 3D model (.glb file) resolves and positions at 1:1 metric scale.",
    "User inspects product placement, shadows, orientation in their space.",
    "Configurations swap instantly in the AR scene if the user toggles options.",
  ];
  steps.forEach((st, i) => {
    s.addShape(pptx.ShapeType.rect, { x:0.55, y:1.8+i*0.68, w:0.4, h:0.4, fill:{ color: C.fillGray }, line:{ color: C.brand, pt:1 } });
    s.addText((i+1).toString(), { x:0.55, y:1.85+i*0.68, w:0.4, fontSize:12, bold:true, color: C.brand, align:"center" });
    s.addText(st, { x:1.15, y:1.85+i*0.68, w:6.5, fontSize:11.5, color: C.textDark });
  });
  s.addShape(pptx.ShapeType.rect, { x:0.55, y:5.95, w:7.1, h:0.55, fill:{ color: C.white }, line:{ color: C.borderGray, pt:1 } });
  s.addText("✅  Best for: Furniture, Rugs, Tableware, Floor Lamps, Large Appliances", { x:0.7, y:6.02, w:6.8, fontSize:11, color: C.brand, bold:true });
  
  // Right Panel
  card(s, 8.1, 1.05, 5.1, 5.75, C.white);
  s.addShape(pptx.ShapeType.rect, { x:8.1, y:1.05, w:5.1, h:0.55, fill:{ color: C.fillGray } });
  s.addShape(pptx.ShapeType.rect, { x:8.1, y:1.05, w:0.07, h:5.75, fill:{ color: C.brand } });
  s.addText("⚙️  Technology Engine", { x:8.35, y:1.12, w:4.7, fontSize:13.5, bold:true, color: C.textDark });
  
  let tech = [
    ["WebXR Device API",   "Standard web specs for AR hardware integration without native apps."],
    ["Three.js",           "WebGL graphics wrapper managing projection grids and shadows."],
    ["glTF / GLB Format",  "Highly efficient, compressed transmission format for 3D web delivery."],
    ["Apple USDZ format",  "Quick Look model bridge for iOS integrations (native AR Quick Look)."],
    ["Mesh Draco Comp.",   "Decompresses geometry dynamically, preserving page speed KPIs."],
  ];
  tech.forEach((t, i) => {
    s.addShape(pptx.ShapeType.rect, { x:8.25, y:1.8+i*0.9, w:0.06, h:0.7, fill:{ color: C.brand } });
    s.addText(t[0], { x:8.45, y:1.85+i*0.9, w:4.6, fontSize:12, bold:true, color: C.textDark });
    s.addText(t[1], { x:8.45, y:2.12+i*0.9, w:4.6, fontSize:10.5, color: C.textMedium });
  });
}

// ─── AR SLIDE B: VTO ────────────────────────────────────────
{
  let s = contentSlide("AR Experience 2: Virtual Try-On (VTO)", "Camera-overlay AR with on-device computer vision.");
  
  // Left Panel
  card(s, 0.35, 1.05, 7.5, 5.75, C.white);
  s.addShape(pptx.ShapeType.rect, { x:0.35, y:1.05, w:7.5, h:0.55, fill:{ color: C.fillGray } });
  s.addShape(pptx.ShapeType.rect, { x:0.35, y:1.05, w:0.07, h:5.75, fill:{ color: C.brand } });
  s.addText("👟  Virtual Try-On (VTO) Camera Pipeline", { x:0.6, y:1.12, w:6.9, fontSize:13.5, bold:true, color: C.textDark });
  
  let steps2 = [
    "Shopper enables frontend camera stream access via WebRTC.",
    "TensorFlow.js executes pipeline inference inside browser sandbox.",
    "Google MediaPipe framework locates key skeleton tracking landmarks.",
    "3D bounding boxes wrap detected joints (e.g. feet, wrists, face).",
    "Configured 3D asset overlays body landmarks with real-time scaling.",
  ];
  steps2.forEach((st, i) => {
    s.addShape(pptx.ShapeType.rect, { x:0.55, y:1.8+i*0.8, w:0.4, h:0.4, fill:{ color: C.fillGray }, line:{ color: C.brand, pt:1 } });
    s.addText((i+1).toString(), { x:0.55, y:1.85+i*0.8, w:0.4, fontSize:12, bold:true, color: C.brand, align:"center" });
    s.addText(st, { x:1.15, y:1.85+i*0.8, w:6.5, fontSize:11.5, color: C.textDark });
  });
  // Tracking targets
  s.addShape(pptx.ShapeType.rect, { x:0.55, y:5.9, w:7.1, h:0.6, fill:{ color: C.fillGray }, line:{ color: C.borderGray, pt:1 } });
  s.addText("Tracking Targets:", { x:0.7, y:6.05, w:1.6, fontSize:10.5, bold:true, color: C.brand });
  let targets = ["Footwear", "Wristwear", "Eyewear", "Headwear"];
  targets.forEach((t, i) => {
    s.addShape(pptx.ShapeType.rect, { x:2.3+i*1.3, y:6.0, w:1.15, h:0.4, fill:{ color: C.white }, line:{ color:C.borderGray, pt:0.5 } });
    s.addText("• " + t, { x:2.3+i*1.3, y:6.08, w:1.15, fontSize:9.5, color: C.textDark, align:"center", bold:true });
  });
  
  // Right Panel - tech
  card(s, 8.1, 1.05, 5.1, 5.75, C.white);
  s.addShape(pptx.ShapeType.rect, { x:8.1, y:1.05, w:5.1, h:0.55, fill:{ color: C.fillGray } });
  s.addShape(pptx.ShapeType.rect, { x:8.1, y:1.05, w:0.07, h:5.75, fill:{ color: C.brand } });
  s.addText("⚙️  Technology Engine", { x:8.35, y:1.12, w:4.7, fontSize:13.5, bold:true, color: C.textDark });
  
  let tech2 = [
    ["MediaPipe Models",   "On-device landmark models for foot detection and face geometry mesh."],
    ["TensorFlow.js",       "Compiles WebGL shaders for high-performance neural runtimes."],
    ["WebGL Compositor",   "Combines raw camera inputs and 3D overlays without latency."],
    ["WebRTC Media Stream","Exposes camera frames to processing canvases at low power."],
    ["Draco Weight Comp.", "Keeps binary neural weights small to load quickly over cellular."],
  ];
  tech2.forEach((t, i) => {
    s.addShape(pptx.ShapeType.rect, { x:8.25, y:1.8+i*0.9, w:0.06, h:0.7, fill:{ color: C.brand } });
    s.addText(t[0], { x:8.45, y:1.85+i*0.9, w:4.6, fontSize:12, bold:true, color: C.textDark });
    s.addText(t[1], { x:8.45, y:2.12+i*0.9, w:4.6, fontSize:10.5, color: C.textMedium });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE: TECHNOLOGY STACK
// ═══════════════════════════════════════════════════════════════
{
  let s = contentSlide("Full Technology Stack", "Every layer of Cognaq's architecture.");
  let layers = [
    { label:"Shopper\nExperience",    items:[["Next.js 14","App router framework"],["React Three Fiber","3D rendering manager"],["Three.js / WebGL","Canvas WebGL runner"],["WebXR API","AR spatial engine"],["MediaPipe + TF.js","Machine learning VTO"],["Vanilla CSS","Floating UI panels"]] },
    { label:"Merchant\nDashboard",    items:[["Next.js Pages","Dashboard pages framework"],["Supabase Auth","Session authentication"],["React Hook Form","Form validation state"],["Recharts","Metric analytics UI"],["Zustand","Client state storage"],["TypeScript","Interface type safety"]] },
    { label:"Backend API",            items:[["Node.js + Express","REST/GraphQL backend"],["PostgreSQL","Database relations"],["Supabase BaaS","Backend utility services"],["Prisma ORM","Database schema layer"],["Stripe SDK","SaaS billing engine"],["GraphQL","Unified client calls"]] },
    { label:"DevOps &\nCloud",         items:[["Vercel","Static & API deploy"],["AWS S3 / R2","Bucket storage for models"],["Cloudflare CDN","Mesh model delivery"],["GitHub Actions","CI/CD code workflows"],["Docker Engine","Container build packaging"],["AWS / GCP","Target cloud infrastructure"]] },
    { label:"E-Commerce\nCore",       items:[["Shopify Admin API","GraphQL product data sync"],["App Bridge","Embed store dashboard"],["OAuth 2.0","Tenant auth handshake"],["Shopify Webhooks","Order capture events"],["GLTF/USDZ Export","Dynamic AR model export"],["REST & GraphQL","Gateway communication"]] },
  ];
  layers.forEach((l, i) => {
    let x = 0.35 + i * 2.6;
    s.addShape(pptx.ShapeType.rect, { x, y:1.05, w:2.45, h:0.75, fill:{ color: C.fillGray } });
    s.addShape(pptx.ShapeType.rect, { x, y:1.05, w:0.06, h:0.75, fill:{ color: C.brand } });
    s.addText(l.label, { x:x+0.1, y:1.1, w:2.28, h:0.65, fontSize:11, bold:true, color: C.textDark, align:"center" });
    l.items.forEach((it, j) => {
      let iy = 1.9 + j * 0.84;
      card(s, x, iy, 2.45, 0.75, C.white);
      s.addShape(pptx.ShapeType.rect, { x, y:iy, w:0.06, h:0.75, fill:{ color: C.brand } });
      s.addText(it[0], { x:x+0.15, y:iy+0.06, w:2.2, fontSize:11, bold:true, color: C.textDark });
      s.addText(it[1], { x:x+0.15, y:iy+0.4, w:2.2, fontSize:10, color: C.textMedium });
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE: PROJECT ROADMAP
// ═══════════════════════════════════════════════════════════════
{
  let s = contentSlide("Project Roadmap", "5 phases — from MVP to scale.");
  let phases = [
    { no:"01", name:"3D Viewer\nMVP",       weeks:"Weeks 1–2",  status:"✓ COMPLETE",  done:true,  items:["Next.js + Three.js core setup","3D model loader (.glb format)","Modular part/colour swapping","Floating glassmorphism UI","Calculated dynamic price engine","GitHub workflow integration"] },
    { no:"02", name:"SaaS\nDashboard",      weeks:"Weeks 3–6",  status:"🔜 NEXT",     done:false, items:["Secure merchant login (JWT)","Model upload & asset manager","Material & pricing rule builder","E-commerce product config","Visual embed code widget","Usage analytics system"] },
    { no:"03", name:"Shopify\nIntegration", weeks:"Weeks 7–9",  status:"📋 PLANNED",  done:false, items:["Shopify app builder setup","Injectable custom theme widget","Dynamic variant mapping to cart","Shopify App Store validation","Webhook system for updates","App Bridge administration"] },
    { no:"04", name:"AR\nExperiences",      weeks:"Weeks 10–14",status:"📋 PLANNED",  done:false, items:["WebXR Room view layout","USDZ iOS pipeline compiler","Android AR Core assets","MediaPipe tracking engine","Camera permission UI config","VTO client validation"] },
    { no:"05", name:"Infrastructure",       weeks:"Weeks 15–20",status:"🚀 FUTURE",   done:false, items:["AWS/GCP server migration","Cloud CDN edge routing","Supabase tenant partition","White-label domain rules","Stripe billing subscription","Developer public API release"] },
  ];
  
  // Timeline connector
  s.addShape(pptx.ShapeType.rect, { x:0.65, y:1.62, w:12.3, h:0.04, fill:{ color: C.borderGray } });
  
  phases.forEach((p, i) => {
    let x = 0.35 + i * 2.6;
    let accentColor = p.done ? C.brand : C.textMedium;
    // Timeline dot
    s.addShape(pptx.ShapeType.ellipse, { x:x+0.92, y:1.44, w:0.38, h:0.38, fill:{ color: p.done ? C.brand : C.white }, line:{ color: accentColor, pt:1.5 } });
    s.addText(p.done ? "✓" : p.no, { x:x+0.92, y:1.48, w:0.38, fontSize:10.5, bold:true, color: p.done ? C.white : C.textMedium, align:"center" });
    
    // Card
    s.addShape(pptx.ShapeType.rect, { x, y:1.85, w:2.45, h:5.0, fill:{ color: C.white }, line:{ color: p.done ? C.brand : C.borderGray, pt: p.done ? 1.5 : 0.75 } });
    s.addShape(pptx.ShapeType.rect, { x, y:1.85, w:2.45, h:0.95, fill:{ color: p.done ? C.brand : C.fillGray } });
    s.addShape(pptx.ShapeType.rect, { x, y:1.85, w:0.06, h:5.0, fill:{ color: p.done ? C.brand : C.borderGray } });
    
    s.addText("Phase " + p.no, { x:x+0.12, y:1.9, w:2.2, fontSize:9, bold:true, color: p.done ? C.white : C.textLight });
    s.addText(p.name, { x:x+0.12, y:2.05, w:2.2, fontSize:12.5, bold:true, color: p.done ? C.white : C.textDark, lineSpacingMultiple:1.2 });
    
    // Status
    s.addShape(pptx.ShapeType.rect, { x:x+0.12, y:2.9, w:2.2, h:0.35, fill:{ color: p.done ? C.brand : C.white }, line:{ color: C.borderGray, pt:0.5 } });
    s.addText(p.status, { x:x+0.12, y:2.95, w:2.2, fontSize:9, bold:true, color: p.done ? C.white : C.textMedium, align:"center" });
    s.addText(p.weeks, { x:x+0.12, y:3.36, w:2.2, fontSize:9.5, color: C.textLight, italic:true });
    
    // Items
    p.items.forEach((item, j) => {
      s.addShape(pptx.ShapeType.rect, { x:x+0.12, y:3.75+j*0.48, w:0.06, h:0.3, fill:{ color: C.brand } });
      s.addText(item, { x:x+0.25, y:3.78+j*0.48, w:2.1, fontSize:9.5, color: C.textDark });
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE: 3D ASSET PIPELINE
// ═══════════════════════════════════════════════════════════════
{
  let s = contentSlide("3D Asset Creation Pipeline", "End-to-end digitisation service — from design files to live web modules.");
  
  let steps = [
    { no:"1", title:"Client Briefing",     desc:"Merchant provides flat product photos, material fabric swatches, measurements, and style guides." },
    { no:"2", title:"3D Modelling",        desc:"Partner artists construct high-accuracy polygon models in Blender/Maya using real measurements." },
    { no:"3", title:"Rigging & Materials", desc:"Models compile with PBR textures. Components rig as separate nodes for modular swaps." },
    { no:"4", title:"Mesh Optimization",   desc:"Draco algorithms compress geometry, targeting files under 5MB for fast loading on phones." },
    { no:"5", title:"Upload & Rules",      desc:"Merchant publishes assets to the dashboard, setting material constraints and price add-ons." },
    { no:"6", title:"Widget Go Live",      desc:"Configurator widget embeds into Shopify page, loading instantly on product variants." },
  ];
  
  // Connect line
  s.addShape(pptx.ShapeType.rect, { x:0.65, y:2.48, w:12.2, h:0.04, fill:{ color: C.borderGray } });
  
  steps.forEach((st, i) => {
    let x = 0.35 + i * 2.18;
    s.addShape(pptx.ShapeType.ellipse, { x:x+0.62, y:2.22, w:0.52, h:0.52, fill:{ color: C.white }, line:{ color: C.brand, pt:1.5 } });
    s.addText(st.no, { x:x+0.62, y:2.3, w:0.52, fontSize:14, bold:true, color: C.brand, align:"center" });
    
    card(s, x, 2.95, 2.02, 3.9, C.white);
    s.addShape(pptx.ShapeType.rect, { x, y:2.95, w:2.02, h:0.65, fill:{ color: C.fillGray } });
    s.addShape(pptx.ShapeType.rect, { x, y:2.95, w:0.07, h:0.65, fill:{ color: C.brand } });
    
    s.addText(st.title, { x:x+0.12, y:3.0, w:1.82, h:0.55, fontSize:11.5, bold:true, color: C.textDark });
    s.addShape(pptx.ShapeType.rect, { x:x+0.1, y:3.6, w:1.82, h:0.03, fill:{ color: C.brand } });
    
    s.addText(st.desc, { x:x+0.1, y:3.7, w:1.82, h:2.95, fontSize:10.5, color: C.textMedium, lineSpacingMultiple:1.4 });
  });
  
  s.addShape(pptx.ShapeType.rect, { x:0.35, y:6.65, w:12.7, h:0.42, fill:{ color: C.fillGray }, line:{ color: C.borderGray, pt:0.5 } });
  s.addShape(pptx.ShapeType.rect, { x:0.35, y:6.65, w:0.35, h:0.42, fill:{ color: C.brand } });
  s.addText("Asset Studio Fees:  $150 (simple models)  |  $250–$400 (modular variables)  |  $400–$600 (complex rigs)", {
    x:0.8, y:6.7, w:12.2, fontSize:11, bold:true, color: C.textDark
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE: COMPETITOR ANALYSIS
// ═══════════════════════════════════════════════════════════════
{
  let s = contentSlide("Competitor Analysis", "How Cognaq compares with market options.");
  let cols = ["Platform","Price / Month","Shopify App","Room AR","Virtual Try-On","Asset Service","Target Segment"];
  let colW  = [2.0,1.7,1.4,1.2,1.6,1.5,2.4];
  let rows = [
    ["Threekit",     "$2,000 – $10k", "Via API",  "Partial",  "None",  "None",  "Large Enterprises"],
    ["Zakeke",       "$99 – $599",    "Native",   "Partial",  "Basic", "None",  "SMB Merchants"],
    ["Angle 3D",     "$29 – $199",    "Native",   "Yes",      "None",  "None",  "Small Boutiques"],
    ["Mimeeq",       "$499+",         "Native",   "Limited",  "None",  "None",  "Furniture Brands"],
    ["Sketchfab",    "$80 – $900",    "None",     "Yes",      "None",  "None",  "3D Asset Showcase"],
    ["COGNAQ",       "Affordable",    "Native",   "Full",     "Full",  "YES",   "SMB to Mid-Market"],
  ];
  // Header
  let xH = 0.35;
  s.addShape(pptx.ShapeType.rect, { x:0.35, y:1.05, w:12.7, h:0.6, fill:{ color: C.textDark } });
  s.addShape(pptx.ShapeType.rect, { x:0.35, y:1.05, w:0.08, h:0.6, fill:{ color: C.brand } });
  cols.forEach((col, i) => {
    s.addText(col, { x:xH+0.06, y:1.12, w:colW[i]-0.1, fontSize:10.5, bold:true, color:C.white, align:"center" });
    xH += colW[i];
  });
  // Rows
  rows.forEach((row, ri) => {
    let isUs = ri === rows.length - 1;
    let xR = 0.35;
    row.forEach((cell, ci) => {
      let rowY = 1.68 + ri * 0.72;
      let bg = isUs ? C.fillGray : (ri % 2 === 0 ? C.white : C.bgLight);
      let lineColor = isUs ? C.brand : C.borderGray;
      let textColor = isUs ? (ci === 0 ? C.brand : C.textDark) : C.textMedium;
      s.addShape(pptx.ShapeType.rect, { x:xR, y:rowY, w:colW[ci], h:0.7, fill:{ color: bg }, line:{ color: lineColor, pt: isUs ? 1.5 : 0.5 } });
      s.addText(cell, { x:xR+0.06, y:rowY+0.18, w:colW[ci]-0.1, fontSize:isUs ? 11 : 10.5, bold:isUs, color:textColor, align:"center" });
      xR += colW[ci];
    });
  });
  s.addShape(pptx.ShapeType.rect, { x:0.35, y:6.65, w:12.7, h:0.42, fill:{ color: C.white }, line:{ color: C.brand, pt:1 } });
  s.addText("Cognaq differentiates by combining full software utility with hands-on 3D asset pipeline services.", {
    x:0.55, y:6.7, w:12.4, fontSize:11, bold:true, color: C.brand
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE: BUSINESS MODEL
// ═══════════════════════════════════════════════════════════════
{
  let s = contentSlide("Business Model & Pricing Tiers", "SaaS subscription options and asset creation services.");
  let tiers = [
    { name:"Starter",      price:"$49",   per:"/mo", highlight:false, features:["5 active products","3D Configurator only","Native Shopify integration","1,000 viewer sessions/mo","Standard email support"] },
    { name:"Growth",       price:"$149",  per:"/mo", highlight:false, features:["25 active products","3D Configurator + Room AR","Dynamic pricing config","Dashboard analytics UI","24h priority support"] },
    { name:"Professional", price:"$399",  per:"/mo", highlight:true,  features:["Unlimited products","3D Configurator + AR + VTO","Custom branding white-label","Custom API endpoint access","Dedicated developer SLA"] },
    { name:"Enterprise",   price:"Custom",per:"",    highlight:false, features:["Unlimited products","On-premise infrastructure","Uptime SLA guarantee","Custom merchant integrations","Dedicated accounts manager"] },
  ];
  tiers.forEach((t, i) => {
    let x = 0.35 + i * 3.25;
    let cardY = 1.1;
    if (t.highlight) {
      s.addShape(pptx.ShapeType.rect, { x:x-0.08, y:0.95, w:3.16, h:5.9, fill:{ color: C.brand } });
      s.addShape(pptx.ShapeType.rect, { x:x+0.3, y:0.95, w:2.4, h:0.38, fill:{ color: C.textDark } });
      s.addText("RECOMMENDED TIER", { x:x+0.3, y:1.0, w:2.4, h:0.3, fontSize:9, bold:true, color:C.white, align:"center", charSpacing:1 });
      cardY = 1.35;
    }
    let bg = t.highlight ? C.white : C.white;
    s.addShape(pptx.ShapeType.rect, { x, y:cardY, w:3.0, h:t.highlight ? 5.45 : 5.65, fill:{ color: bg }, line:{ color: t.highlight ? C.brand : C.borderGray, pt:1.5 } });
    s.addShape(pptx.ShapeType.rect, { x, y:cardY, w:3.0, h:1.0, fill:{ color: t.highlight ? C.textDark : C.fillGray } });
    s.addText(t.name, { x:x+0.12, y:cardY+0.15, w:2.75, fontSize:15, bold:true, color: t.highlight ? C.brand : C.textDark, align:"center" });
    s.addText(t.price, { x:x+0.12, y:cardY+1.2, w:2.75, fontSize:28, bold:true, color: C.brand, align:"center" });
    if (t.per) s.addText(t.per, { x:x+0.12, y:cardY+1.9, w:2.75, fontSize:12, color: C.textMedium, align:"center" });
    s.addShape(pptx.ShapeType.rect, { x:x+0.6, y:cardY+2.3, w:1.8, h:0.04, fill:{ color: C.brand } });
    t.features.forEach((f, j) => {
      s.addText("•  " + f, { x:x+0.15, y:cardY+2.5+j*0.53, w:2.7, fontSize:11, color: C.textMedium });
    });
  });
  s.addShape(pptx.ShapeType.rect, { x:0.1, y:6.88, w:"99%", h:0.12, fill:{ color: C.brand } });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE: REQUIREMENTS
// ═══════════════════════════════════════════════════════════════
{
  let s = contentSlide("Technical & Platform Requirements", "Engineering parameters for the Cognaq platform.");
  let reqs = [
    { title:"⚡  Performance",  items:["Viewer load times under 3s on standard cellular connections.","60 FPS canvas redraw speed on mid-range Android/iOS.","Edge CDN configurations to load model bundles quickly."] },
    { title:"🪄  AR Platform",  items:["WebXR device integration on Safari iOS 12+ and Chrome.","USDZ generation pipeline for instant Apple Quick Look.","In-browser TensorFlow weight loading target under 2MB."] },
    { title:"🔒  Security",     items:["Token rotations with JWT authorization for dashboards.","Shopify App credentials stored securely; OAuth handshake.","GDPR-compliant cookie limits and transaction logging."] },
    { title:"📈  Scalability",  items:["Database structures support isolated multi-tenant schemas.","Serverless function scaling for global API endpoint routing.","Migrate platform infrastructure to AWS at 500+ tenants."] },
    { title:"🛒  E-Commerce",   items:["Direct cart synchronization on custom options select.","Shopify variant IDs map to material/colour properties.","JSON webhooks trigger orders with specifications sheet pdf."] },
    { title:"🎨  3D Models",    items:["Draco compression keeps geometry bundles under 5MB.","Physically Based Rendering (PBR) materials for realistic view.","Initial pipeline testing verifies geometry coordinates."] },
  ];
  reqs.forEach((r, i) => {
    let col = i % 3;
    let row = Math.floor(i / 3);
    let x = 0.35 + col * 4.28;
    let y = 1.1 + row * 2.75;
    card(s, x, y, 3.95, 2.6, C.white);
    s.addShape(pptx.ShapeType.rect, { x, y, w:3.95, h:0.58, fill:{ color: C.fillGray } });
    s.addShape(pptx.ShapeType.rect, { x, y, w:0.07, h:2.6, fill:{ color: C.brand } });
    s.addText(r.title, { x:x+0.18, y:y+0.12, w:3.65, fontSize:13, bold:true, color: C.textDark });
    r.items.forEach((item, j) => {
      s.addShape(pptx.ShapeType.rect, { x:x+0.18, y:y+0.75+j*0.6, w:0.06, h:0.35, fill:{ color: C.brand } });
      s.addText(item, { x:x+0.32, y:y+0.75+j*0.6, w:3.5, fontSize:10.5, color: C.textMedium });
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE: SUCCESS METRICS & DELIVERABLES
// ═══════════════════════════════════════════════════════════════
{
  let s = contentSlide("Deliverables & Success Metrics", "Core launch components and performance parameters.");
  let deliverables = [
    ["Shopify App Listing",          "Public app listing with OAuth setup for organic shop installs."],
    ["Merchant Admin Panel",        "Self-serve panel handles upload, variant maps, and statistics."],
    ["Embedded 3D Configurator",    "Widget embeds natively in product detail pages; runs 60FPS."],
    ["WebXR Space AR",              "'View in Your Room' functions natively inside mobile browser."],
    ["MediaPipe Try-On AR",         "On-device VTO overlays watches, glasses, and footwear instantly."],
    ["Asset Creation Engine",       "Partner pipeline handles design file to optimised GLB builds."],
    ["Stripe Billing Gateway",      "Subscription tier payment management with custom upgrades."],
    ["Variant Inventory Sync",      "Configurator choices connect directly to product inventory codes."],
  ];
  deliverables.forEach((d, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = 0.35 + col * 6.55;
    let y = 1.1 + row * 1.05;
    card(s, x, y, 6.2, 0.92, C.white);
    s.addShape(pptx.ShapeType.rect, { x, y, w:0.75, h:0.92, fill:{ color: C.brand } });
    s.addText((i+1).toString(), { x, y:y+0.25, w:0.75, fontSize:18, bold:true, color:C.white, align:"center" });
    s.addText(d[0], { x:x+0.9, y:y+0.08, w:5.1, fontSize:13, bold:true, color: C.textDark });
    s.addText(d[1], { x:x+0.9, y:y+0.45, w:5.1, fontSize:11, color: C.textMedium });
  });
  
  // KPI Bar (Light Gray / White borders)
  s.addShape(pptx.ShapeType.rect, { x:0.1, y:5.42, w:"99%", h:1.45, fill:{ color: C.white }, line:{ color: C.borderGray, pt:1 } });
  s.addShape(pptx.ShapeType.rect, { x:0.1, y:5.42, w:0.08, h:1.45, fill:{ color: C.brand } });
  s.addText("Key Performance Indicators (KPIs):", { x:0.35, y:5.52, w:5.0, fontSize:12, bold:true, color: C.brand });
  
  let kpis = [
    { val:"50+",    label:"Shopify installs\nin first 90 days" },
    { val:"1,000+", label:"Monthly active\nAR sessions" },
    { val:">25%",   label:"Average conversion\nlift for merchants" },
    { val:"<3s",    label:"Configurator widget\nmobile load speed" },
    { val:"$10K+",  label:"MRR targets by\nmonth 6 launch" },
  ];
  kpis.forEach((k, i) => {
    let x = 0.35 + i * 2.6;
    s.addShape(pptx.ShapeType.rect, { x, y:5.85, w:2.5, h:0.9, fill:{ color: C.fillGray }, line:{ color: C.borderGray, pt:0.5 } });
    s.addText(k.val, { x, y:5.88, w:2.5, fontSize:18, bold:true, color: C.textDark, align:"center" });
    s.addText(k.label, { x, y:6.2, w:2.5, fontSize:9.5, color: C.textMedium, align:"center", lineSpacingMultiple:1.3 });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE: NEXT STEPS (CLOSING)
// ═══════════════════════════════════════════════════════════════
{
  let s = pptx.addSlide();
  s.background = { color: C.bgLight };
  
  // Left Panel - split theme
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.2, h:"100%", fill:{ color: C.brand } });
  s.addShape(pptx.ShapeType.rect, { x:0.2, y:0, w:5.0, h:"100%", fill:{ color: C.white } });
  s.addShape(pptx.ShapeType.rect, { x:5.2, y:0, w:0.03, h:"100%", fill:{ color: C.borderGray } });
  
  s.addText("🚀", { x:0.8, y:1.5, w:3.8, h:1.2, fontSize:64, align:"center" });
  s.addText("What's Next", { x:0.8, y:2.8, w:3.8, h:0.8, fontSize:32, bold:true, color: C.brand, align:"center" });
  s.addShape(pptx.ShapeType.rect, { x:1.2, y:3.8, w:3.0, h:0.04, fill:{ color: C.brand } });
  s.addText("Immediate engineering action plan\nto execute platform release.", { x:0.8, y:4.1, w:3.8, fontSize:12, color: C.textMedium, align:"center", lineSpacingMultiple:1.5 });
  
  // Right Panel - next items
  s.addText("Immediate Action Items", { x:5.8, y:0.8, w:7.2, fontSize:22, bold:true, color: C.textDark });
  s.addShape(pptx.ShapeType.rect, { x:5.8, y:1.3, w:7.0, h:0.04, fill:{ color: C.brand } });
  
  let nexts = [
    ["Phase 2 Kickoff",           "Develop SaaS dashboard interface: merchant login & rule builders."],
    ["3D Artist Network Setup",   "Onboard and align with 3D studios for standardized GLB production."],
    ["AR WebXR Prototyping",      "Assemble initial WebXR spatial viewer for pilot furniture merchant."],
    ["Shopify App Scaffolding",   "Register Shopify Developer tenant and initialize App Bridge."],
    ["Stripe Subscription API",   "Integrate subscription checkout webhooks before merchant onboarding."],
    ["Pilot Merchant Selection",  "Select 3 active merchants for initial beta testing run."],
  ];
  nexts.forEach((n, i) => {
    s.addShape(pptx.ShapeType.rect, { x:5.8, y:1.5 + i*0.8, w:0.4, h:0.4, fill:{ color: C.brand } });
    s.addText((i+1).toString(), { x:5.8, y:1.55+i*0.8, w:0.4, fontSize:12, bold:true, color:C.white, align:"center" });
    s.addText(n[0], { x:6.35, y:1.5 + i*0.8, w:7.0, fontSize:13, bold:true, color: C.textDark });
    s.addText(n[1], { x:6.35, y:1.75 + i*0.8, w:7.0, fontSize:11, color: C.textMedium });
  });
  
  s.addShape(pptx.ShapeType.rect, { x:5.8, y:6.4, w:7.2, h:0.04, fill:{ color: C.borderGray } });
  s.addText("github.com/Vaidyakrushna/Cognaq  |  Version 3.0  |  July 2026", {
    x:5.8, y:6.5, w:7.2, fontSize:10, color: C.textLight, italic:true
  });
}

// ─── GENERATE FILE ─────────────────────────────────────────────
pptx.writeFile({ fileName:"Cognaq_3D_Configurator_Documentation.pptx" })
  .then(f => console.log("Saved: " + f))
  .catch(err => console.error("Error:", err));
