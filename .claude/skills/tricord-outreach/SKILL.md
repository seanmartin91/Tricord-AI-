---
name: tricord-outreach
description: >
  Run a full multi-touch outreach campaign to find and contact new customers for
  TRICORD I.T Solutions (tricordit.ca) — an Ontario MSP whose primary hook is AI
  Strategy & Governance. Use this skill whenever the user asks about finding prospects,
  customer outreach, lead generation, cold email, LinkedIn prospecting, Mailchimp
  campaigns, Hunter.io lookups, Outlook drafts, or LinkedIn messages for Tricord.
  Target sectors include law firms, manufacturers, CPA/accounting firms, pharma/life
  sciences, engineering/architecture, insurance brokerages, and healthcare groups —
  all across Ontario. The skill finds real targets via web search, leads with the AI
  governance angle, verifies emails via Hunter.io, creates Outlook Web drafts and
  LinkedIn connection requests via Chrome, queues Mailchimp campaigns, and delivers
  a full tracking spreadsheet — all for Dwight to review before sending.
---

# TRICORD Outreach Campaign

You are running a senior-led customer acquisition campaign for **TRICORD I.T Solutions**.
The primary hook for all outreach is **AI Strategy & Governance** — the sharpest,
most urgent pain point across every target sector in 2026.

**Read ALL reference files before writing any content:**
- `references/company-profile.md` — Tricord's services, differentiators, Dwight's voice
- `references/ai-strategy.md` — The AI strategy hook, three-lane model, sector angles **(read first for messaging)**
- `references/manufacturing-icp.md` — Manufacturing targets, triggers, research tips
- `references/legal-icp.md` — Law firm targets, triggers, Hunter.io tips
- `references/additional-sectors.md` — CPA firms, pharma, engineering, insurance, healthcare ICPs

---

## Campaign Scope

**Default campaign (no user specification): 10 prospects across 2 sectors**
- 5 manufacturing prospects
- 5 legal prospects

**If the user specifies a sector, focus, or headcount:** adjust accordingly.
All seven supported sectors are listed below — pick the most relevant per the user's request.

### Supported Sectors
| Sector | Reference File | Primary AI Hook |
|--------|---------------|-----------------|
| Manufacturing | manufacturing-icp.md | No coherent AI strategy; OT/IT/AI convergence |
| Legal (law firms) | legal-icp.md | Shadow AI on client files; law society guidance |
| CPA / Accounting | additional-sectors.md | CPA Ontario guidance; client financial data residency |
| Pharma / Life Sciences | additional-sectors.md | GxP validation + AI; clinical data residency |
| Engineering / Architecture | additional-sectors.md | Professional liability + AI output; PEO/OAA gap |
| Insurance Brokerages | additional-sectors.md | FSRA guidance coming; E&O renewal pressure |
| Healthcare / Allied Health | additional-sectors.md | PHIPA + AI scribe tools; EMR AI features |

---

## Step 1 — Find Prospects via Web Search

Use web search to find **real, named Ontario companies**. Tailor searches to the
sectors selected. Example angles by sector:

- **Manufacturing**: `Ontario manufacturer "Industry 4.0" OR AI automation 2025 2026`
- **Legal**: `"Ontario law firm" AI policy OR Copilot 2025 site:linkedin.com`
- **CPA**: `"Ontario CPA firm" OR "Ontario accounting firm" AI Caseware OR Xero`
- **Pharma**: `Ontario pharmaceutical biotech OR CRO AI machine learning 2025`
- **Engineering**: `Ontario engineering firm AI design tools OR BIM 2025`
- **Insurance**: `Ontario insurance brokerage OR MGA "Applied Epic" OR AI`
- **Healthcare**: `Ontario "multi-site" OR "multi-location" clinic AI scribe OR EMR`

For each prospect, research:
- Company size, sector, city
- News: AI adoption, expansion, merger, regulatory pressure, hiring surge
- Named contact (name + title — needed for Hunter.io and LinkedIn)
- Website domain (needed for Hunter.io)
- Any public signal that AI or governance is on their radar

---

## Step 2 — Find Email Addresses via Hunter.io

For each prospect, call the **Hunter.io Domain Search API**:

```
GET https://api.hunter.io/v2/domain-search?domain={domain}&api_key={key}
```

Ask the user for their Hunter.io API key if not provided. Use `web_fetch`.

1. Extract the domain from the prospect's website
2. Call domain-search
3. Filter for senior decision-maker titles (President, CEO, Owner, Managing Partner,
   COO, VP Operations, Practice Manager, Principal, Medical Director, etc.)
4. Record email + confidence score (%)

If no match: mark "LinkedIn outreach only" — don't skip the prospect.

---

## Step 3 — Draft the Outreach Sequences

For each prospect, write a **4-touch sequence**. Lead every Touch 1 with the
AI governance angle specific to their sector (see reference files for exact hooks).

| Touch | Channel | When |
|-------|---------|------|
| 1 | Email — AI governance hook specific to their sector | Day 0 |
| 2 | LinkedIn — connection request + note | Day 3 |
| 3 | Email — different angle (data residency, regulatory body, insurance) | Day 7 |
| 4 | Email — short, no-pressure close | Day 14 |

### Dwight's voice (non-negotiable)
- Direct, senior peer-to-peer — not a vendor, not a consultant doing a pitch
- Touch 1 must contain one specific detail about this company's situation
- Pain anchored on what the prospect is risking right now without AI governance
- CTA: "Would a 15-minute AI strategy call make sense?" — never "demo" or "meeting"
- Emails under 200 words; LinkedIn notes under 150 characters
- Never: "I hope this email finds you well", "touch base", "synergy",
  "value proposition", "reach out to learn more"

### Touch 3 pivot angles by sector
- **Legal** → Canadian data residency + US CLOUD Act implications for client files
- **Manufacturing** → Government AI modernisation incentives, or cyber insurance angle
- **CPA** → CPA Ontario guidance gap; audit trail requirements for AI-generated content
- **Pharma** → Health Canada GxP validation requirements for AI in regulated processes
- **Engineering** → Professional liability exposure; PEO/OAA policy gap
- **Insurance** → FSRA guidance coming; E&O renewal documentation
- **Healthcare** → IPC scrutiny; PHIPA audit exposure from AI scribe tools

### Touch 4 (short close)
3–5 sentences max. "I'll leave it here — if the timing isn't right, no worries.
The offer of a 15-minute AI strategy call stands whenever it's useful."

---

## Step 4 — Create Outlook Web Drafts (via Chrome)

For each prospect with a verified email, use Chrome to create a draft in Outlook Web
(https://outlook.office.com):

1. Screenshot to confirm user is signed in — if not, ask them to sign in first
2. For each prospect: New Mail → To, Subject (Touch 1), Body (Touch 1)
3. Save Draft — do NOT Send
4. Confirm draft saved before moving to the next

**Ask upfront:** "Should I create Outlook drafts for all verified prospects, or
just the top priority ones?"

---

## Step 5 — LinkedIn Connection Requests (via Chrome)

For each prospect with a named contact, use Chrome to send the Touch 2 connection
request:

1. Screenshot to confirm user is signed in to LinkedIn — ask if not
2. Search for contact by name + company
3. Open profile → Connect → Add a note → paste Touch 2 message (≤150 chars)
4. Send invitation
5. Confirm before moving to next contact

**Before starting:** "Ready for me to send LinkedIn requests? I'll confirm each one
before it goes out." Wait for explicit user confirmation.

---

## Step 6 — Push to Mailchimp

For prospects with verified emails:

1. `campaign_planner` — draft Touch 1 (subject, preview text, body)
2. `save_to_mailchimp` — save as draft, do NOT send
3. Record campaign ID in tracking spreadsheet

If Mailchimp errors: skip, note it, use Outlook drafts as fallback.

---

## Step 7 — Build the Campaign Tracking Spreadsheet

Use the xlsx skill to create an Excel file in the outputs folder.

**Sheet 1 — Prospect Tracker**

| Company | Sector | Contact Name | Title | Email | Confidence % | LinkedIn URL | AI Signal (research note) | Regulatory Body | Outlook Draft | Mailchimp ID | LinkedIn Sent | T3 Due | T4 Due | Response | Notes |

- Pre-fill everything found; leave action/tracking columns blank
- Conditional formatting: row turns green when Response is filled
- Freeze top row; auto-fit columns
- Sort by: closest regulatory deadline or warmest AI signal first

**Sheet 2 — Message Library**

All messages (prospects × 4 touches), one per row:
| Company | Sector | Touch # | Channel | Subject / Note | Body | Suggested Send Date |

---

## Step 8 — Deliver the Package

1. `computer://` link to the Excel file
2. Quick table: prospects found / emails verified / Outlook drafts / LinkedIn sent / Mailchimp queued
3. **Top 2–3 priority prospects** — warmest AI signal, most urgent regulatory pressure, or timing event
4. Any LinkedIn-only prospects (no email found)

Keep it tight — Dwight is technical and busy.

---

## Quality bar

Before finishing:
- Every Touch 1 leads with a sector-specific AI governance angle?
- Every Touch 1 has a company-specific detail (not just the name)?
- Touch 3 uses a different angle than Touch 1?
- Subject lines are distinct and specific across all prospects?
- Outlook drafts saved (not sent)?
- LinkedIn only sent after explicit user confirmation?
- Mailchimp drafts saved (not sent)?
- Spreadsheet has Regulatory Body column, freeze panes, conditional formatting?

Fix anything that fails before delivering.
