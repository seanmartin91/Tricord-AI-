const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TableOfContents, ExternalHyperlink
} = require('docx');
const fs = require('fs');

// ── colour palette ──────────────────────────────────────────────
const NAVY   = "1B3A6B";
const TEAL   = "0E7490";
const LIGHT  = "E8F4F8";
const ACCENT = "F0A500";
const WHITE  = "FFFFFF";
const LGREY  = "F5F5F5";
const MGREY  = "D0D0D0";
const DGREY  = "444444";
const BLACK  = "111111";

// ── helpers ─────────────────────────────────────────────────────
const border = (c = MGREY) => ({ style: BorderStyle.SINGLE, size: 1, color: c });
const borders = (c = MGREY) => ({ top: border(c), bottom: border(c), left: border(c), right: border(c) });
const noBorders = () => {
  const nb = { style: BorderStyle.NONE, size: 0, color: WHITE };
  return { top: nb, bottom: nb, left: nb, right: nb };
};

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: "Arial", size: 32, bold: true, color: NAVY })],
    spacing: { before: 360, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 4 } },
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: TEAL })],
    spacing: { before: 280, after: 120 },
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: DGREY })],
    spacing: { before: 200, after: 80 },
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Arial", size: 22, color: BLACK, ...opts })],
    spacing: { before: 60, after: 80 },
  });
}

function bullet(text, level = 0, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    children: [new TextRun({ text, font: "Arial", size: 22, color: BLACK, ...opts })],
    spacing: { before: 40, after: 40 },
  });
}

function numbered(text, level = 0, opts = {}) {
  return new Paragraph({
    numbering: { reference: "numbers", level },
    children: [new TextRun({ text, font: "Arial", size: 22, color: BLACK, ...opts })],
    spacing: { before: 40, after: 40 },
  });
}

function spacer(lines = 1) {
  return new Paragraph({ children: [new TextRun({ text: "", size: 22 })], spacing: { before: 0, after: lines * 120 } });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// highlight box (shaded paragraph via single-cell table)
function callout(text, fillColor = LIGHT, textColor = NAVY) {
  const b = borders(TEAL);
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: b,
        width: { size: 9360, type: WidthType.DXA },
        shading: { fill: fillColor, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 180, right: 180 },
        children: [new Paragraph({
          children: [new TextRun({ text, font: "Arial", size: 22, color: textColor, bold: false })],
          spacing: { before: 0, after: 0 }
        })]
      })]
    })]
  });
}

// simple 2-col table helper
function twoColTable(rows, col1W = 3200, col2W = 6160) {
  const total = col1W + col2W;
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: [col1W, col2W],
    rows: rows.map(([left, right, isHeader]) => new TableRow({
      tableHeader: !!isHeader,
      children: [
        new TableCell({
          borders: borders(MGREY),
          width: { size: col1W, type: WidthType.DXA },
          shading: { fill: isHeader ? NAVY : LGREY, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: left, font: "Arial", size: 20, bold: !!isHeader, color: isHeader ? WHITE : NAVY })],
            spacing: { before: 0, after: 0 }
          })]
        }),
        new TableCell({
          borders: borders(MGREY),
          width: { size: col2W, type: WidthType.DXA },
          shading: { fill: isHeader ? NAVY : WHITE, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: right, font: "Arial", size: 20, bold: !!isHeader, color: isHeader ? WHITE : BLACK })],
            spacing: { before: 0, after: 0 }
          })]
        }),
      ]
    }))
  });
}

// 3-col table
function threeColTable(rows, col1W = 2500, col2W = 3400, col3W = 3460) {
  const total = col1W + col2W + col3W;
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: [col1W, col2W, col3W],
    rows: rows.map(([c1, c2, c3, isHeader]) => new TableRow({
      tableHeader: !!isHeader,
      children: [c1, c2, c3].map((txt, i) => {
        const w = [col1W, col2W, col3W][i];
        return new TableCell({
          borders: borders(MGREY),
          width: { size: w, type: WidthType.DXA },
          shading: { fill: isHeader ? NAVY : (i === 0 ? LGREY : WHITE), type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: txt, font: "Arial", size: 20, bold: !!isHeader, color: isHeader ? WHITE : BLACK })],
            spacing: { before: 0, after: 0 }
          })]
        });
      })
    }))
  });
}

// ════════════════════════════════════════════════════════════════
// DOCUMENT CONTENT
// ════════════════════════════════════════════════════════════════

const coverChildren = [
  spacer(4),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "AI CONSULTING BUSINESS", font: "Arial", size: 52, bold: true, color: WHITE })],
    spacing: { before: 0, after: 160 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Sales Process & Marketing Plan", font: "Arial", size: 36, color: WHITE })],
    spacing: { before: 0, after: 200 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "━━━━━━━━━━━━━━━━━━━━", font: "Arial", size: 28, color: ACCENT })],
    spacing: { before: 0, after: 200 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "AI Consulting  ·  AI Roadmap  ·  App Building  ·  Agent Building", font: "Arial", size: 24, color: LIGHT, italics: true })],
    spacing: { before: 0, after: 480 },
  }),
  spacer(2),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Prepared: May 2026", font: "Arial", size: 22, color: LIGHT })],
    spacing: { before: 0, after: 80 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Confidential — Internal Use Only", font: "Arial", size: 20, color: ACCENT, italics: true })],
    spacing: { before: 0, after: 0 },
  }),
];

// ── Part 1: Sales Process ───────────────────────────────────────
const salesChildren = [
  // TOC Page
  pageBreak(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "TABLE OF CONTENTS", font: "Arial", size: 28, bold: true, color: NAVY })],
    spacing: { before: 200, after: 160 },
  }),
  new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-2" }),

  // ── SALES PROCESS ───────────────────────────────────────────
  pageBreak(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "PART ONE", font: "Arial", size: 24, bold: true, color: ACCENT })],
    spacing: { before: 200, after: 80 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Structured Sales Process", font: "Arial", size: 40, bold: true, color: NAVY })],
    spacing: { before: 0, after: 240 },
  }),

  // ── 1. ICP ──────────────────────────────────────────────────
  h1("1. Ideal Client Profile (ICP)"),
  body("Your strongest early wins will come from businesses that feel AI is inevitable but have no internal capability to act on it. You bring both the strategic clarity (consulting + roadmap) and the technical execution (app + agent building) — that combination is rare and highly valuable to the following profiles."),
  spacer(),

  h2("1.1 Primary ICP — The Overworked SMB Owner"),
  twoColTable([
    ["Attribute", "Description", true],
    ["Company Size", "5–50 employees; £500K–£10M annual revenue"],
    ["Industries", "Professional services (accountants, solicitors, recruiters, estate agents, coaches), trades businesses, e-commerce, marketing agencies"],
    ["Geography", "UK (primary), Canada/Ontario (secondary, leveraging existing Ontario network)"],
    ["Tech Maturity", "Uses Microsoft 365 or Google Workspace; no dedicated IT/tech team; not yet using AI beyond ChatGPT ad hoc"],
    ["Pain Points", "Repetitive admin tasks eating founder time, inconsistent client onboarding, staff turnover causing knowledge loss, can't compete on headcount"],
    ["Budget Range", "£500–£5,000/month for consulting or £3,000–£25,000 for a project deliverable"],
    ["Decision Maker", "The founder or a director who also handles operations — no procurement committee"],
  ]),
  spacer(),

  h2("1.2 Secondary ICP — The Growth-Stage Scale-Up"),
  twoColTable([
    ["Attribute", "Description", true],
    ["Company Size", "50–200 employees; Series A or equivalent revenue (£2M–£20M)"],
    ["Industries", "SaaS, fintech, healthtech, HR tech, any sector undergoing digital transformation"],
    ["Tech Maturity", "Has a product but no AI capabilities; CTO or Head of Product present but lacking AI expertise"],
    ["Pain Points", "Competitors launching AI features, investors asking about AI strategy, internal data underutilised"],
    ["Budget Range", "£5,000–£50,000 for roadmap + build engagements"],
    ["Decision Maker", "CEO, CTO, or CPO; typically a 2–4 week sales cycle with one stakeholder presentation"],
  ]),
  spacer(),

  h2("1.3 ICP Disqualifiers — Who NOT to Target"),
  body("Avoid these prospects to protect your time and close rate:"),
  bullet("Enterprises with existing AI teams or vendor relationships (long sales cycles, procurement gates)"),
  bullet("Businesses that have never heard of AI and express zero curiosity (education cost is too high)"),
  bullet("Clients who want to own IP from day one and have no budget for recurring work"),
  bullet("Anyone primarily price-shopping without a real problem to solve"),
  spacer(),

  // ── 2. Lead Generation ──────────────────────────────────────
  pageBreak(),
  h1("2. Lead Generation Methods"),

  h2("2.1 Outbound — Proactive Prospecting"),
  h3("LinkedIn Direct Outreach (Priority #1)"),
  body("LinkedIn is your primary hunting ground. The goal is to start conversations, not pitch immediately."),
  numbered("Identify prospects using LinkedIn Sales Navigator or manual search: filter by company size (11–200), job title (Founder, MD, CEO, Operations Director, Head of [anything]), and industry verticals from the ICP above."),
  numbered("Warm the connection by engaging with their posts (genuine comments, not likes) for 3–5 days before connecting."),
  numbered('Send a connection request with a personalised note: "Hi [Name] — I noticed you [specific observation about their post/company]. I help businesses like yours use AI to [specific outcome]. Would love to connect." No pitch yet.'),
  numbered("After connection is accepted, send Message 1 (see Follow-Up Cadence section) — a curiosity-opener, not a sales message."),
  numbered("If no response in 5 days, send Message 2. If no response, move on. Never send more than 3 messages."),
  spacer(),

  h3("Cold Email (Supporting Channel)"),
  body("Use cold email to reach prospects who are less active on LinkedIn, particularly accountants, solicitors, and trades businesses."),
  bullet("Source emails via Apollo.io, Hunter.io, or LinkedIn profile data (where publicly visible)"),
  bullet("Subject lines that work: 'Quick question about your admin workload', 'Saving [Company Name] 5 hours/week', 'AI for [Industry] — is it worth it?'"),
  bullet("Keep emails under 100 words. One clear ask. No attachments on first contact."),
  bullet("Personalise the first line with something specific about their business"),
  spacer(),

  h2("2.2 Inbound — Content-Driven Leads"),
  h3("LinkedIn Content (Daily)"),
  body("Consistent LinkedIn posting is your highest-leverage inbound activity. See the full 4-week content calendar in the Marketing Plan section. Key post formats:"),
  bullet("'Here's what I built this week and what it cost the client in time saved' — builder credibility"),
  bullet("'3 AI mistakes I see SMBs make' — authority + relatability"),
  bullet("'I built [WorkDelta / CompMom / Privlio] in X days — here's how' — proof of capability"),
  bullet("'AI audit checklist for [industry]' posts with downloadable lead magnet CTA"),
  spacer(),

  h3("Website & SEO (Medium Term)"),
  body("A clean, fast website (2–3 pages) builds credibility when prospects Google you after seeing your LinkedIn. Essential pages: Home (positioning + CTA), Services (four service lines with clear outcomes), Portfolio (WorkDelta, CompMom, Privlio, fridge scanner — with screenshots and metrics), Contact (simple form + booking link)."),
  spacer(),

  h2("2.3 Referral — Highest Conversion Channel"),
  body("Referrals will be your best leads. Budget 20% of your time to nurturing the referral network."),
  bullet("After every completed project, explicitly ask: 'Do you know any other business owners who are trying to figure out AI?'"),
  bullet("Create a simple referral incentive: £200 credit on their next invoice for any referral that converts"),
  bullet("Stay warm with past clients: share a relevant AI article or insight once a month — no ask, just value"),
  bullet("Tap your existing Ontario network (from your MSP/tech background) and UK professional network immediately on launch"),
  bullet("Consider strategic alliances with web designers, business coaches, and accountants — they serve the same SMB clients and can refer bidirectionally"),
  spacer(),

  // ── 3. Discovery Call ───────────────────────────────────────
  pageBreak(),
  h1("3. Discovery Call Framework"),
  callout("Goal of the discovery call: Understand their world deeply enough to prescribe the right solution confidently. You are a doctor diagnosing before prescribing — never pitch a service line until you know what they actually need.", LIGHT, NAVY),
  spacer(),

  h2("3.1 Call Structure (45–60 minutes)"),
  twoColTable([
    ["Phase", "Content & Duration", true],
    ["Opening (5 min)", "Thank them. Set the agenda: 'I'd love to spend most of this call understanding your business — then I'll share what I think could genuinely help.' Signal that you're not going to pitch immediately."],
    ["Their World (15 min)", "Open-ended questions about their business — see question bank below."],
    ["Pain & Priority (15 min)", "Dig into the 1–2 biggest problems surfaced. Quantify pain where possible."],
    ["Previous Attempts (5 min)", "What have they tried? What didn't work? Why?"],
    ["Vision (5 min)", "What does 'success' look like 12 months from now? What's blocking it?"],
    ["Recap & Next Steps (10 min)", "Reflect back what you heard. Share your initial thinking. Agree a clear next step."],
  ]),
  spacer(),

  h2("3.2 Discovery Question Bank"),
  h3("Business Context"),
  bullet("Tell me about your business — what do you do and who do you serve?"),
  bullet("What does a typical week look like for you as the founder/director?"),
  bullet("Where does most of your team's time go day-to-day?"),
  spacer(),

  h3("AI Awareness & Readiness"),
  bullet("Have you started using AI tools at all — even just ChatGPT?"),
  bullet("What's your honest view on AI right now — exciting, overwhelming, or irrelevant to you?"),
  bullet("Is there pressure from customers, competitors, or investors to 'do something with AI'?"),
  spacer(),

  h3("Pain Points"),
  bullet("If you could wave a magic wand and eliminate one painful, repetitive process, what would it be?"),
  bullet("Where do mistakes or inconsistencies most often happen in your operations?"),
  bullet("What's currently costing you the most in either time or money?"),
  bullet("Do you have data (customer data, documents, inventory, communications) that isn't being used?"),
  spacer(),

  h3("Quantifying the Pain"),
  bullet("How many hours a week does [painful process] take across your team?"),
  bullet("What does that cost you in salary terms? What's the impact on the business if it stays unsolved?"),
  bullet("Have you lost clients or deals because of [problem]?"),
  spacer(),

  h3("Decision & Budget"),
  bullet("Who else would be involved in deciding to move forward on something like this?"),
  bullet("Do you have a rough budget in mind for this kind of work, or would that be shaped by what the solution looks like?"),
  bullet("What's your timeline — is this urgent, or something you're planning for?"),
  spacer(),

  h2("3.3 What to Listen For"),
  callout("Green flags: They mention specific processes that eat time. They express frustration with repetition. They compare themselves to competitors who 'have figured out AI.' They mention data they're not using. They have a clear timeline or trigger event (new hire, contract renewal, funding round).", "E8F5E9", "1B5E20"),
  spacer(),
  callout("Amber flags: Vague pain ('we just want to be more efficient'). No clear budget signal. Multiple stakeholders with no clear lead. Previous failed tech project making them cautious. Address these head-on before investing in a proposal.", LIGHT, NAVY),
  spacer(),

  // ── 4. Proposal Structure ───────────────────────────────────
  pageBreak(),
  h1("4. Proposal Structure & Pricing"),
  body("Proposals should be short (4–6 pages), visual, and outcome-focused. Lead with what the client gets, not what you do. Always personalise the first page with their name, company, and the exact problem you discussed."),
  spacer(),

  h2("4.1 Universal Proposal Structure"),
  numbered("Cover Page — Client name, their company logo (if you can source it), proposed solution title, date"),
  numbered("The Situation — Mirror back their problem in their words (shows you listened)"),
  numbered("The Recommended Solution — What you'll do and why it fits their situation"),
  numbered("What's Included — Clear scope, deliverables, timeline"),
  numbered("Investment — Pricing (one option or good/better/best)"),
  numbered("Why Us — Brief credibility section: your built products, your approach"),
  numbered("Next Steps — One clear action (sign, schedule a call, reply to approve)"),
  spacer(),

  h2("4.2 Service Line Pricing"),

  h3("Service Line 1: AI Consulting"),
  body("Strategic advisory — helps clients understand AI's potential for their business and make informed decisions."),
  twoColTable([
    ["Package", "What's Included / Price", true],
    ["AI Health Check (one-off)", "2-hour deep-dive workshop + written report covering: current AI readiness, top 3 opportunities, risks and considerations, recommended next steps. Price: £750–£1,200"],
    ["Advisory Retainer — Starter", "4 hours/month of strategic advisory (calls, Slack/email Q&A, document review). Price: £800/month"],
    ["Advisory Retainer — Growth", "8 hours/month + monthly strategy session + review of any AI initiatives in progress. Price: £1,500/month"],
    ["Advisory Day", "Full day on-site or virtual deep dive into one business problem. Price: £1,800–£2,500/day"],
  ]),
  spacer(),

  h3("Service Line 2: AI Roadmap ('AI Road')"),
  body("A structured, tailored AI adoption plan that gives the client a clear path to implementation — without needing to become a technical expert themselves."),
  twoColTable([
    ["Package", "What's Included / Price", true],
    ["Standard AI Roadmap", "Discovery workshop (2–3 hrs) + full written roadmap document covering: prioritised AI opportunities, recommended tools/platforms, implementation timeline (3–12 months), ROI estimates, quick wins vs. longer-term plays. Delivered in 1–2 weeks. Price: £2,500–£4,000"],
    ["Roadmap + Kickoff", "Standard roadmap + 90-day kickoff support (monthly check-ins, tool selection guidance, vendor introductions). Price: £5,000–£7,500"],
    ["Enterprise Roadmap", "For larger organisations: multi-department assessment, stakeholder interviews, board-ready presentation, change management considerations. Price: £8,000–£15,000"],
  ]),
  spacer(),

  h3("Service Line 3: App Building"),
  body("Custom AI-powered applications built to solve specific business problems. Your portfolio (WorkDelta, CompMom, Privlio, fridge scanner concept) demonstrates real capability here."),
  twoColTable([
    ["Package", "What's Included / Price", true],
    ["Mini App / Proof of Concept", "A focused single-feature app to prove concept and get stakeholder buy-in. Typically 2–3 weeks. Price: £2,000–£5,000"],
    ["Standard Web App", "Full AI-powered web application with auth, database, AI features, and deployment. 4–10 weeks. Examples: internal productivity tools, client portals, AI-assisted workflow tools. Price: £8,000–£20,000"],
    ["Complex / Multi-Feature App", "Larger-scope builds with integrations, advanced AI features, custom UI. 10–20 weeks. Price: £20,000–£50,000"],
    ["App Maintenance Retainer", "Ongoing hosting, updates, bug fixes, and minor feature additions post-launch. Price: £300–£800/month"],
  ]),
  spacer(),

  h3("Service Line 4: Agent Building"),
  body("Custom AI agents and automations that run independently to handle business processes — saving significant human time."),
  twoColTable([
    ["Package", "What's Included / Price", true],
    ["Single Automation / Workflow", "One AI-powered automation (e.g., email triage, document summarisation, lead qualification). 1–2 weeks. Price: £1,500–£3,500"],
    ["Agent Bundle (3 automations)", "Three interconnected automations built around one core business process. 3–4 weeks. Price: £6,000–£10,000"],
    ["Custom AI Agent", "A sophisticated autonomous agent capable of handling multi-step tasks, integrating with external APIs, making decisions. 4–8 weeks. Price: £8,000–£18,000"],
    ["Agent Maintenance", "Monitoring, prompt tuning, updates, and expansion of existing agents. Price: £250–£600/month"],
  ]),
  spacer(),

  h2("4.3 Packaging Tip: The Starter Bundle"),
  callout("For new clients unsure where to start: offer the 'AI Kickstart Package' — AI Health Check + Standard AI Roadmap bundled at £3,200 (saving ~£500). This low-risk entry point generates qualified leads for App or Agent Building projects downstream.", LIGHT, NAVY),
  spacer(),

  // ── 5. Follow-Up Cadence ────────────────────────────────────
  pageBreak(),
  h1("5. Follow-Up Cadence"),
  body("Most sales are lost not because the prospect said no, but because the conversation went cold. Consistent, value-adding follow-up keeps you front of mind without being pushy."),
  spacer(),

  h2("5.1 LinkedIn Outreach Sequence"),
  twoColTable([
    ["Touch", "Message & Timing", true],
    ["Connection Request", "Personalised note mentioning something specific about them. No pitch."],
    ["Message 1 (Day 1 after connect)", "Value-first opener: Share a relevant insight or observation about their industry + AI. End with a soft question. Example: 'I've been working with a few [industry] businesses on reducing admin time with AI — are you finding that's a challenge at [Company]?'"],
    ["Message 2 (Day 6 if no reply)", "Different angle: 'I built a quick AI audit framework for [industry] businesses — happy to share if it would be useful?' Offer something free."],
    ["Message 3 (Day 14 if no reply)", "Graceful exit: 'I'll leave it here — not for everyone at the moment. If AI ever moves up the priority list, feel free to reach out.' This often gets responses."],
  ]),
  spacer(),

  h2("5.2 Post-Discovery Call Cadence"),
  twoColTable([
    ["Timing", "Action", true],
    ["Same day", "Send a personalised thank-you email with a 2–3 bullet summary of what you discussed. Confirm next steps."],
    ["24–48 hours", "Send the proposal (never later than 48 hours after a discovery call — urgency drops fast)."],
    ["Day 5 (if no response)", "Follow-up email: 'Just checking in — happy to jump on a 15-min call if any questions came up.'"],
    ["Day 10 (if no response)", "Add something new: a case study, article, or specific idea you thought of for their business."],
    ["Day 18 (if no response)", "Final follow-up: 'I want to respect your time — shall I close this off for now, or would a different timing work better?'"],
    ["Day 30+", "Move to low-touch nurture: monthly LinkedIn comment on their posts, occasional value share."],
  ]),
  spacer(),

  h2("5.3 Past Client Nurture (Monthly)"),
  bullet("Month 1 post-project: Check-in on results — 'How is [the thing you built] working out?'"),
  bullet("Month 2: Share a relevant article, tool, or AI update with no ask"),
  bullet("Month 3: Soft ask — 'We've been building some interesting agent automations lately — happy to share what's been working if useful?'"),
  bullet("Quarterly: Invite to any webinars, LinkedIn Lives, or events you run"),
  spacer(),

  // ── 6. Objection Handling ───────────────────────────────────
  pageBreak(),
  h1("6. Objection Handling"),
  body("The most common objections to AI consulting are predictable. Here's how to handle them with empathy and evidence — never argue, always redirect to their outcome."),
  spacer(),

  h2('"It\'s too expensive"'),
  callout("Hear it as: 'I don\'t yet see the ROI clearly enough to justify this.' Your job is to quantify the value, not defend the price.", LGREY, DGREY),
  spacer(),
  body("Response framework:"),
  numbered("Acknowledge: 'That's a fair concern — this is a real investment.'"),
  numbered("Quantify the cost of inaction: 'You mentioned [problem] takes your team 15 hours a week. At an average of £25/hour, that's £375/week, or nearly £20K a year — just in time cost, not counting the mistakes or the stress.'"),
  numbered("Reframe ROI: 'If this automation saves even half that, it pays for itself in under 6 months and keeps paying every month after.'"),
  numbered("Offer a lower-risk entry: 'If budget is a constraint right now, the AI Health Check at £750 gives you a clear picture of where the ROI is — no big commitment.'"),
  spacer(),

  h2('"We\'re not ready for AI yet"'),
  callout("Hear it as: 'I\'m not sure where to start and don\'t want to get it wrong.' This is actually your perfect prospect — the roadmap service was built for them.", LGREY, DGREY),
  spacer(),
  body("Response:"),
  body("'That's exactly why businesses hire me — not to implement AI chaotically, but to build a clear, realistic plan before spending a penny on tools or tech. The AI Roadmap is designed for businesses at exactly your stage. It tells you what to do, in what order, and what to avoid. Most clients say it's the most useful thing they've done to get clarity on their next 12 months.'"),
  spacer(),

  h2('"I don\'t trust AI — it makes things up"'),
  body("Response:"),
  body("'That's a legitimate concern, and it's why the design of AI systems matters so much. I don't build AI tools that guess or hallucinate critical information — the solutions I build have clear guardrails, human oversight steps, and are designed around specific, constrained tasks where AI is genuinely reliable. Happy to show you an example of how that looks in practice.'"),
  spacer(),

  h2('"We tried AI before and it didn\'t work"'),
  body("Response:"),
  body("'Tell me more about what you tried — I'd genuinely like to understand what went wrong. In my experience, most failed AI implementations come down to one of three things: wrong tool for the job, no clear process definition before automation, or no change management for the team. I'd want to make sure we don't repeat the same mistake.'"),
  spacer(),

  h2('"I need to think about it / speak to my partner"'),
  body("Response:"),
  body("'Of course — this is a real decision and I don't want you to rush it. Can I ask: is there anything in the proposal that's unclear, or a specific concern I can address before you have that conversation? I'd rather iron out questions now than leave you guessing.'"),
  body("Then: agree a specific follow-up date/time before ending the conversation. 'When should I check back in — would Thursday work?'"),
  spacer(),

  // ── 7. Closing Process ──────────────────────────────────────
  pageBreak(),
  h1("7. Closing Process"),
  callout("The best close happens naturally when discovery was thorough, the proposal reflects their exact situation, and the follow-up was consistent. Never use pressure tactics — they kill referrals.", LIGHT, NAVY),
  spacer(),

  h2("7.1 Verbal Close (End of Proposal Call)"),
  body("After walking through the proposal, instead of asking 'What do you think?' (weak), use:"),
  bullet("'Based on what we've discussed, does this feel like it addresses the core problem you described?'"),
  bullet("'Is there anything that would need to change for this to feel right for you?'"),
  bullet("'What would need to be true for you to feel confident moving forward?'"),
  spacer(),
  body("If they're positive but hesitant:"),
  bullet("'What's the one thing that would make this a clear yes for you?' — isolates the real blocker"),
  bullet("'Would a phased start help — we begin with [smaller scope] and expand from there?'"),
  spacer(),

  h2("7.2 Written Close (After Proposal Sent)"),
  bullet("Use a clear e-signature tool (DocuSign, PandaDoc, or HelloSign) — friction kills deals"),
  bullet("Include a clear expiry date on pricing: 'This proposal is valid for 21 days from the date above'"),
  bullet("Make the next step one click: a link to approve, sign, or book a call to finalise"),
  spacer(),

  h2("7.3 Starter Incentive (Optional)"),
  body("For hesitant prospects where budget is tight, offer a 'quick start' bonus: 'If we kick off before [date], I'll include a free AI tools audit (normally £300) to make sure you're on the right platforms from day one.' Creates urgency without discounting."),
  spacer(),

  // ── 8. Onboarding ───────────────────────────────────────────
  pageBreak(),
  h1("8. Post-Sale Onboarding"),
  body("A smooth onboarding experience sets the tone for the entire client relationship — and makes referrals far more likely. Clients remember how they felt in the first two weeks, not just the end result."),
  spacer(),

  h2("8.1 Day 1: Welcome & Kickoff"),
  numbered("Send a personalised welcome email within 24 hours of signing. Include: what happens next, your communication preferences, and a warm human note."),
  numbered("Send the onboarding questionnaire (see below) — keep it short (5–8 questions)."),
  numbered("Schedule the kickoff call for within 5 business days."),
  spacer(),

  h2("8.2 Onboarding Questionnaire (All Service Lines)"),
  twoColTable([
    ["Question", "Why It Matters", true],
    ["How do you prefer to communicate — Slack, email, WhatsApp, calls?", "Reduces friction throughout the project"],
    ["Who is the main point of contact on your side?", "Avoids bottlenecks"],
    ["What are your working hours / when are you typically available?", "Sets realistic expectations"],
    ["Is there anything that's changed since we last spoke?", "Surfaces new priorities or concerns early"],
    ["What would make this engagement a 10/10 for you?", "Anchors success criteria upfront"],
    ["Are there any upcoming events, deadlines, or pressures I should know about?", "Avoids nasty surprises mid-project"],
  ]),
  spacer(),

  h2("8.3 Kickoff Call Agenda"),
  numbered("Welcome + relationship building (5 min)"),
  numbered("Confirm goals and success metrics — what does 'done' look like? (10 min)"),
  numbered("Walk through the project plan / roadmap (15 min)"),
  numbered("Agree communication cadence: weekly update? Fortnightly check-in? (5 min)"),
  numbered("Any questions or concerns to address upfront? (10 min)"),
  numbered("Confirm next milestone and who's responsible for what (5 min)"),
  spacer(),

  h2("8.4 During the Project"),
  bullet("Send a brief weekly status update (email or Slack) — 3–5 bullet points: done this week, next week, any blockers"),
  bullet("Flag any scope changes immediately with a clear impact assessment before acting"),
  bullet("Share progress wins as they happen (a screenshot, a quick Loom demo, a metric) — keeps excitement high"),
  spacer(),

  h2("8.5 Project Close"),
  bullet("Deliver a handover document: what was built, how to use it, how to maintain it, who to contact for support"),
  bullet("Schedule a 30-day post-project check-in at the point of delivery"),
  bullet("Ask for a testimonial/case study within 1 week of delivery (while the experience is fresh)"),
  bullet("Ask for a referral: 'Is there anyone in your network who might benefit from something similar?'"),
  spacer(),
];

// ── Part 2: Marketing Plan ──────────────────────────────────────
const marketingChildren = [
  pageBreak(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "PART TWO", font: "Arial", size: 24, bold: true, color: ACCENT })],
    spacing: { before: 200, after: 80 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Marketing Plan", font: "Arial", size: 40, bold: true, color: NAVY })],
    spacing: { before: 0, after: 240 },
  }),

  // ── 9. Positioning ──────────────────────────────────────────
  h1("9. Positioning & Key Differentiator"),

  h2("9.1 Positioning Statement"),
  callout("For SMB owners and growth-stage founders who know AI is important but don't know where to start, we are the AI consulting and build partner that combines strategic clarity with hands-on delivery — so you get a plan you can trust AND execution you can see. Unlike large consultancies that advise from a distance, or freelancers who only code, we do both: from boardroom strategy to production-ready app.", LIGHT, NAVY),
  spacer(),

  h2("9.2 The Core Differentiator: Builder-Consultant"),
  body("The single biggest thing that separates this business from competitors is the combination of strategic thinking AND practical building in one person/team. This is rare and must be front and centre in all marketing."),
  spacer(),
  body("Most AI consultants have never built a real product. Most developers don't understand business strategy. You sit at the intersection — and the portfolio proves it:"),
  spacer(),
  twoColTable([
    ["Product Built", "What It Demonstrates", true],
    ["WorkDelta.ai", "Understanding of workplace productivity problems; ability to build and ship a real SaaS product"],
    ["CompMom (dance competition app)", "Ability to identify a niche problem and build a purpose-specific app from scratch"],
    ["Fridge AI Scanner concept", "Creative application of computer vision / AI to consumer problems — innovation mindset"],
    ["Privlio (secure contact management)", "Understanding of data privacy concerns; enterprise-grade thinking at SMB scale"],
  ]),
  spacer(),
  body("Use these in every proposal, website, and LinkedIn post. They are your proof points, not just projects."),
  spacer(),

  h2("9.3 Messaging Pillars"),
  twoColTable([
    ["Pillar", "What to Say", true],
    ["Clarity", "'We cut through the AI hype and tell you exactly what matters for your business — no jargon, no fluff.'"],
    ["Credibility", "'We don't just advise — we build. Our portfolio of real AI products proves we can deliver, not just recommend.'"],
    ["Speed", "'Most clients see their first tangible AI result within 30 days of working with us.'"],
    ["Safety", "'We design AI systems with human oversight built in — no black boxes, no surprises.'"],
  ]),
  spacer(),

  // ── 10. Target Sectors ──────────────────────────────────────
  pageBreak(),
  h1("10. Target Sectors"),
  body("Focus initial marketing effort on sectors where AI ROI is clearest, decision cycles are short, and word-of-mouth travels fast. Below are the top five recommended sectors."),
  spacer(),

  threeColTable([
    ["Sector", "AI Opportunity", "Relevant Service Line", true],
    ["Professional Services (accountants, solicitors, recruiters)", "Document processing, client onboarding automation, meeting notes, contract review", "Consulting, AI Road, Agent Building"],
    ["Marketing & PR Agencies", "AI content workflows, competitor research agents, client reporting automation, SEO tools", "App Building, Agent Building, Consulting"],
    ["Estate Agents & Property", "Listing copy generation, lead qualification agents, viewing scheduling automation", "Agent Building, App Building"],
    ["E-commerce & Retail", "Product description AI, customer service agents, inventory prediction, review analysis", "Agent Building, App Building"],
    ["Trades & Field Services (HVAC, construction, facilities)", "Job scheduling agents, quote automation, compliance document management", "Agent Building, AI Road"],
  ]),
  spacer(),

  h2("10.1 Sector Entry Strategy"),
  body("Pick one sector to lead with for the first 90 days. The recommendation is Professional Services — specifically accountants and recruiters — because:"),
  bullet("High volume of repetitive document-heavy tasks (ideal for agent automation)"),
  bullet("Already familiar with paying for professional services at premium rates"),
  bullet("Network density: one happy client refers to 3–5 others in the same profession"),
  bullet("Real-world relevance: your Privlio app (contact management) speaks directly to their data hygiene needs"),
  spacer(),

  // ── 11. Channel Strategy ────────────────────────────────────
  pageBreak(),
  h1("11. Channel Strategy"),

  h2("11.1 LinkedIn (Primary Channel — 60% of effort)"),
  body("LinkedIn is where your ICP lives, makes decisions, and is most receptive to professional content. It is your most important marketing asset in year one."),
  twoColTable([
    ["Activity", "Frequency / Detail", true],
    ["Profile optimisation", "One-time: Headline = 'I help SMBs use AI to save time and grow. Consultant + Builder.' Banner = your services. About section = story-led, builder credibility, CTA."],
    ["Original posts", "5x per week: mix of builder stories, insight posts, case studies, and opinion pieces (see content calendar)"],
    ["Comments on ICP posts", "10–15 per day: genuine, substantive comments on posts by founders and operators in your target sectors"],
    ["DM outreach", "10–15 new personalised connection requests per day, follow-up sequence (see sales process)"],
    ["LinkedIn newsletter", "Weekly or fortnightly: 'The AI Practitioner' — one practical AI insight or case study per issue"],
  ]),
  spacer(),

  h2("11.2 Website (Credibility Hub — 20% of effort)"),
  bullet("Keep it simple: Home, Services, Portfolio, About, Contact — 5 pages maximum"),
  bullet("Homepage hero: clear positioning statement + one CTA ('Book a Free AI Health Check Call')"),
  bullet("Portfolio page: screenshots, problem/solution/result format for each product built"),
  bullet("Embed a Calendly or TidyCal booking link everywhere — remove friction from the first contact"),
  bullet("Add live case studies as they develop — even anonymised, outcomes-focused summaries"),
  spacer(),

  h2("11.3 Referrals (Highest Value — ongoing)"),
  bullet("Structured referral ask at project close (scripted — see Onboarding section)"),
  bullet("Alliance partners: web designers, business coaches, HR consultants, accountants — offer reciprocal referrals"),
  bullet("Join 2–3 local business communities (Chamber of Commerce, BNI, or sector-specific networks)"),
  spacer(),

  h2("11.4 Speaking & Events (Brand Building — occasional)"),
  bullet("Pitch to speak at local business events, Chamber meetings, and sector conferences on 'AI for [sector]'"),
  bullet("Run quarterly free LinkedIn Lives: '30-minute AI audit of a real SMB' — live format builds trust fast"),
  bullet("Guest podcast appearances on small business and entrepreneurship shows"),
  spacer(),

  // ── 12. Content Calendar ────────────────────────────────────
  pageBreak(),
  h1("12. LinkedIn Content Calendar — Weeks 1–4"),
  body("Post consistently for 4 weeks to seed authority before pushing outbound hard. Each post should be written in first person, specific, and end with a question or CTA. Aim for 150–300 words per post."),
  spacer(),

  h2("Week 1 — Establishing Who You Are"),
  twoColTable([
    ["Day", "Post Topic & Format", true],
    ["Monday", "LAUNCH POST: 'I'm starting an AI consulting business — here's why, and what I've already built' — introduce the four service lines, mention WorkDelta + CompMom + Privlio as proof points. Personal, honest, founder story tone."],
    ["Tuesday", "Builder post: 'What I learned building WorkDelta.ai — the AI productivity tool I shipped in [X weeks]' — specific lessons, one insight for SMB owners about workplace AI."],
    ["Wednesday", "Insight post: 'The #1 mistake I see businesses make when they first try AI' — relatable, concise, ends with: 'What's your experience been?' to drive comments."],
    ["Thursday", "Proof of concept: 'I built a fridge AI scanner concept in a weekend. Here's how it works and what it taught me about consumer AI.' — techy but explained simply."],
    ["Friday", "Soft CTA: 'If you run a business and you're curious about where AI could actually save you time, I offer a free 30-minute AI Clarity Call. No pitch, just honest advice. Drop a comment or DM me to book one.'"],
  ]),
  spacer(),

  h2("Week 2 — Building Credibility with Specifics"),
  twoColTable([
    ["Day", "Post Topic & Format", true],
    ["Monday", "Case study (can be anonymised): 'A [sector] business was spending 12 hours a week on [task]. We automated it in 3 weeks. Here's what we did and what it cost them.' — specific numbers build trust."],
    ["Tuesday", "Education post: '5 AI tools every SMB should know about in 2026 (and 2 they should avoid)' — practical, opinionated, shareable."],
    ["Wednesday", "Myth-busting: 'AI will replace your staff' — the truth is more interesting.' — address the fear, position AI as a force multiplier."],
    ["Thursday", "Behind the build: 'Building Privlio — why data privacy is the next AI frontier for SMBs' — links your product to a trend."],
    ["Friday", "Lead magnet post: 'I just published a free AI Readiness Assessment — 10 questions to tell you exactly where your business stands with AI. Link in comments.'"],
  ]),
  spacer(),

  h2("Week 3 — Sector-Specific Authority"),
  twoColTable([
    ["Day", "Post Topic & Format", true],
    ["Monday", "'If you're a recruiter reading this, here are the 3 AI tools that would save your team 6+ hours a week.' — direct address to a sector, highly shareable within that network."],
    ["Tuesday", "Process post: 'How I run an AI discovery session — the exact questions I ask in the first 45 minutes with a new client.' — behind the curtain content."],
    ["Wednesday", "Opinion: 'AI Roadmaps are the new business plan — and most businesses still don't have one.' — provocative, starts conversations."],
    ["Thursday", "Client outcome (anonymised): 'A marketing agency was manually writing 40 client reports per month. They now do it in 2 hours with an AI agent we built. Thread below...'"],
    ["Friday", "Engagement driver: 'Rate your business on AI readiness 1–10. 1 = never thought about it. 10 = already using it daily. I'll respond to every comment with one suggestion.'"],
  ]),
  spacer(),

  h2("Week 4 — Offers & Social Proof"),
  twoColTable([
    ["Day", "Post Topic & Format", true],
    ["Monday", "Testimonial/result share: 'Feedback from a recent client — [quote or paraphrased outcome].' — human proof."],
    ["Tuesday", "Service explainer: 'What an AI Roadmap actually looks like — I'll show you a real one (anonymised).' — attach or screenshot the deliverable format."],
    ["Wednesday", "'The AI Agent Building checklist — 5 things to figure out before you build anything.' — practical, saves prospects from common mistakes."],
    ["Thursday", "Personal story: 'Why I left [previous context] to build AI tools full-time. The real reason.' — vulnerability + clarity builds connection."],
    ["Friday", "CTA post: 'I have 3 client slots available in June. If you're an SMB owner who wants a clear AI strategy and someone who can actually build it — DM me or book a call. Details in comments.'"],
  ]),
  spacer(),

  // ── 13. Lead Magnets ────────────────────────────────────────
  pageBreak(),
  h1("13. Lead Magnet Ideas"),
  body("Lead magnets convert your content audience into email subscribers and discovery call bookings. Prioritise speed to create — simple is better than polished."),
  spacer(),

  threeColTable([
    ["Lead Magnet", "Format", "Funnel Stage", true],
    ["AI Readiness Assessment (10 questions)", "PDF or typeform quiz with scored output", "Awareness — top of funnel"],
    ["Free AI Audit Call (30 minutes)", "Calendly booking link — no-cost session", "Consideration — moves to discovery"],
    ["'AI for [Industry]' Guide (e.g., AI for Accountants)", "PDF: 5 pages, specific tools, use cases, ROI examples", "Awareness — sector-specific"],
    ["AI Tools Comparison Sheet", "PDF or Notion page: best tools by category, pros/cons, price", "Consideration — positions you as guide"],
    ["AI ROI Calculator", "Simple web tool or Excel: input hours + salary, output ROI estimate", "Consideration — quantifies value"],
    ["'What AI Can (and Can't) Do for Your Business' Webinar", "60-min live or pre-recorded, hosted on LinkedIn Live or Zoom", "Awareness + email capture"],
    ["Automation Idea Bank (sector-specific)", "PDF: 20 AI automation ideas for [sector] businesses", "Awareness — highly shareable"],
  ]),
  spacer(),
  callout("Priority: Build the AI Readiness Assessment first (can be done with Typeform in a day). It doubles as a discovery tool — the answers give you rich intel before a sales call.", LIGHT, NAVY),
  spacer(),

  // ── 14. 90-Day Launch Plan ──────────────────────────────────
  pageBreak(),
  h1("14. 90-Day Launch Plan"),
  body("The first 90 days are about building momentum, not perfection. Focus on consistency over quality, conversations over content, and learning over optimising."),
  spacer(),

  h2("Month 1 — Foundation (Weeks 1–4)"),
  h3("Week 1: Set Up"),
  bullet("Finalise service line scope, pricing, and proposal templates"),
  bullet("Optimise LinkedIn profile: headline, banner, about section, featured posts"),
  bullet("Register domain and set up simple landing page (Framer, Webflow, or Squarespace)"),
  bullet("Set up CRM (free HubSpot or Notion CRM) to track leads"),
  bullet("Create booking link (Calendly or TidyCal)"),
  bullet("Write the Week 1 LinkedIn posts in advance"),
  spacer(),

  h3("Week 2: Launch Content & First Outreach"),
  bullet("Publish Week 1 and 2 LinkedIn posts (5/week)"),
  bullet("Start daily LinkedIn engagement: 15 comments on ICP posts per day"),
  bullet("Send 10 personalised connection requests per day"),
  bullet("Create AI Readiness Assessment (Typeform or Google Forms)"),
  bullet("Target: 50 new connections by end of week"),
  spacer(),

  h3("Week 3: First Conversations"),
  bullet("Follow up on accepted connections with Message 1 sequence"),
  bullet("Book first 3–5 free AI Clarity Calls"),
  bullet("Publish Week 3 LinkedIn content"),
  bullet("Create sector-specific 'AI for Accountants' guide (first lead magnet)"),
  bullet("Reach out to 5 people in existing network about launching the business"),
  spacer(),

  h3("Week 4: First Proposals"),
  bullet("Run discovery calls and send first proposals"),
  bullet("Publish Week 4 LinkedIn content"),
  bullet("Set up email newsletter (Beehiiv or Substack) — publish first issue"),
  bullet("Review and iterate: what post format got the most engagement? What questions came up most in calls?"),
  spacer(),

  h2("Month 2 — Momentum (Weeks 5–8)"),
  bullet("Target: 2–3 paying clients by end of month 2"),
  bullet("Continue daily LinkedIn posting and outreach"),
  bullet("Run first LinkedIn Live: 'Live AI Audit of a Real SMB' (recruit a friendly contact to be the subject)"),
  bullet("Build out the website portfolio page with the first completed project"),
  bullet("Start building referral partnerships: approach 3 web designers, 2 business coaches, 1 accountant"),
  bullet("Publish AI ROI Calculator as second lead magnet"),
  bullet("Send first email newsletter issue to subscribers"),
  bullet("Adjust ICP based on who's actually responding — double down on what's working"),
  spacer(),

  h2("Month 3 — Scaling What Works (Weeks 9–12)"),
  bullet("Target: 4–6 active clients; £5,000–£15,000 monthly revenue"),
  bullet("Build and launch first sector-specific case study (detailed, with permission from client)"),
  bullet("Apply to speak at one local business event or industry meetup"),
  bullet("Review outreach sequences: which messages get the best response? Rewrite the others."),
  bullet("Introduce the first upsell: if a consulting client doesn't have a roadmap, propose one; if a roadmap client hasn't asked about building, propose it"),
  bullet("Set up a simple client portal (Notion or ClickUp) for project management — signals professional operation"),
  bullet("Hire or contract a VA for 5 hours/week if admin is becoming a bottleneck"),
  bullet("Review pricing: are clients accepting without hesitation? If so, increase by 15–20%"),
  spacer(),

  h2("14.1 90-Day KPIs to Track"),
  twoColTable([
    ["Metric", "Target by Day 90", true],
    ["LinkedIn connections (net new)", "500+"],
    ["Discovery calls booked", "20–30"],
    ["Proposals sent", "10–15"],
    ["Paying clients", "3–6"],
    ["Monthly recurring revenue", "£3,000–£10,000"],
    ["Email subscribers", "100+"],
    ["LinkedIn post impressions (cumulative)", "50,000+"],
    ["Referral conversations initiated", "10+"],
  ]),
  spacer(),

  // ── 15. Quick Reference ─────────────────────────────────────
  pageBreak(),
  h1("15. Quick Reference: Pricing Summary"),
  twoColTable([
    ["Service", "Starting Price", true],
    ["AI Health Check (one-off)", "£750"],
    ["Advisory Retainer — Starter", "£800/month"],
    ["Advisory Retainer — Growth", "£1,500/month"],
    ["Advisory Day", "£1,800/day"],
    ["Standard AI Roadmap", "£2,500"],
    ["Roadmap + Kickoff Support", "£5,000"],
    ["Enterprise AI Roadmap", "£8,000"],
    ["Mini App / Proof of Concept", "£2,000"],
    ["Standard Web App", "£8,000"],
    ["Complex Multi-Feature App", "£20,000"],
    ["App Maintenance Retainer", "£300/month"],
    ["Single Automation / Workflow", "£1,500"],
    ["Agent Bundle (3 automations)", "£6,000"],
    ["Custom AI Agent", "£8,000"],
    ["Agent Maintenance", "£250/month"],
    ["AI Kickstart Bundle (Health Check + Roadmap)", "£3,200 (bundled)"],
  ]),
  spacer(),

  callout("Pricing is in GBP. For Canadian/Ontario clients, apply a 1.7x multiplier to convert to CAD and remain competitive in that market. Review and adjust pricing quarterly — especially if close rates exceed 50% (a sign you're underpriced).", LIGHT, NAVY),
  spacer(),

  pageBreak(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "— End of Document —", font: "Arial", size: 22, color: MGREY, italics: true })],
    spacing: { before: 480, after: 80 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Built with purpose. Built to ship.", font: "Arial", size: 20, color: ACCENT, bold: true })],
    spacing: { before: 0, after: 0 },
  }),
];

// ════════════════════════════════════════════════════════════════
// ASSEMBLE DOCUMENT
// ════════════════════════════════════════════════════════════════

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }, {
          level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1080, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
    ]
  },
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22, color: BLACK } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: NAVY },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: TEAL },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: DGREY },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 }
      },
    ]
  },
  sections: [
    // ── Cover section (navy background via shaded table approach) ──
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 0, right: 0, bottom: 0, left: 0 }
        }
      },
      children: [
        new Table({
          width: { size: 12240, type: WidthType.DXA },
          columnWidths: [12240],
          rows: [new TableRow({
            children: [new TableCell({
              borders: noBorders(),
              width: { size: 12240, type: WidthType.DXA },
              shading: { fill: NAVY, type: ShadingType.CLEAR },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 2160, bottom: 2160, left: 1440, right: 1440 },
              children: coverChildren
            })]
          })]
        }),
      ]
    },
    // ── Main content section ──
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            children: [
              new TextRun({ text: "AI Consulting Business — Sales Process & Marketing Plan", font: "Arial", size: 18, color: TEAL }),
            ],
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: TEAL, space: 4 } },
            spacing: { before: 0, after: 120 }
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [
              new TextRun({ text: "Confidential — ", font: "Arial", size: 18, color: MGREY }),
              new TextRun({ text: "Page ", font: "Arial", size: 18, color: MGREY }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: MGREY }),
              new TextRun({ text: " of ", font: "Arial", size: 18, color: MGREY }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 18, color: MGREY }),
            ],
            alignment: AlignmentType.RIGHT,
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: TEAL, space: 4 } },
            spacing: { before: 80, after: 0 }
          })]
        })
      },
      children: [...salesChildren, ...marketingChildren]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/sessions/elegant-gifted-bardeen/mnt/outputs/AI_Consulting_Sales_Marketing_Plan.docx', buffer);
  console.log('Document written successfully');
}).catch(err => {
  console.error('Error:', err.message);
  process.exit(1);
});
