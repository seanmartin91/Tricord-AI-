const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname)));

// ── Existing routes ──────────────────────────────────────────────────────────

app.post('/api/contact', (req, res) => {
  const { firstName, lastName, email, company, industry, challenge } = req.body;
  if (!firstName || !email) return res.status(400).json({ error: 'Name and email are required.' });
  const entry = { timestamp: new Date().toISOString(), firstName, lastName, email, company, industry, challenge };
  fs.appendFileSync(path.join(__dirname, 'leads.jsonl'), JSON.stringify(entry) + '\n');
  console.log(`[lead] ${email} — ${company || 'no company'}`);
  res.json({ ok: true });
});

app.get('/api/leads', (req, res) => {
  if (req.headers['authorization'] !== `Bearer ${process.env.ADMIN_TOKEN || 'tricord-admin'}`)
    return res.status(401).json({ error: 'Unauthorized' });
  const f = path.join(__dirname, 'leads.jsonl');
  if (!fs.existsSync(f)) return res.json([]);
  res.json(fs.readFileSync(f, 'utf8').trim().split('\n').filter(Boolean).map(l => JSON.parse(l)));
});

// ── API key status (booleans only — never exposes values) ─────────────────────

app.get('/api/status', (req, res) => {
  res.json({
    hubspot:   !!process.env.HUBSPOT_API_KEY,
    hunter:    !!process.env.HUNTER_API_KEY,
    anthropic: !!process.env.ANTHROPIC_API_KEY,
  });
});

// ── HubSpot proxy ─────────────────────────────────────────────────────────────

app.get('/api/hubspot/contacts', async (req, res) => {
  const key = process.env.HUBSPOT_API_KEY;
  if (!key) return res.status(503).json({ error: 'HUBSPOT_API_KEY not configured' });
  const props = 'firstname,lastname,email,company,jobtitle,lifecyclestage,hs_lead_status';
  try {
    const r = await fetch(
      `https://api.hubapi.com/crm/v3/objects/contacts?limit=50&properties=${props}&archived=false`,
      { headers: { Authorization: `Bearer ${key}` } }
    );
    res.json(await r.json());
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/hubspot/search', async (req, res) => {
  const key = process.env.HUBSPOT_API_KEY;
  if (!key) return res.status(503).json({ error: 'HUBSPOT_API_KEY not configured' });
  const { query } = req.body;
  const props = ['firstname','lastname','email','company','jobtitle','lifecyclestage'];
  try {
    const r = await fetch('https://api.hubapi.com/crm/v3/objects/contacts/search', {
      method: 'POST',
      headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        query,
        properties: props,
        limit: 30,
      }),
    });
    res.json(await r.json());
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/hubspot/contact', async (req, res) => {
  const key = process.env.HUBSPOT_API_KEY;
  if (!key) return res.status(503).json({ error: 'HUBSPOT_API_KEY not configured' });
  try {
    const r = await fetch('https://api.hubapi.com/crm/v3/objects/contacts', {
      method: 'POST',
      headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties: req.body }),
    });
    res.json(await r.json());
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Hunter.io proxy ───────────────────────────────────────────────────────────

app.post('/api/hunter/search', async (req, res) => {
  const key = process.env.HUNTER_API_KEY;
  if (!key) return res.status(503).json({ error: 'HUNTER_API_KEY not configured' });
  const { domain } = req.body;
  if (!domain) return res.status(400).json({ error: 'domain required' });
  try {
    const r = await fetch(
      `https://api.hunter.io/v2/domain-search?domain=${encodeURIComponent(domain)}&api_key=${key}&limit=10`
    );
    res.json(await r.json());
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/hunter/verify', async (req, res) => {
  const key = process.env.HUNTER_API_KEY;
  if (!key) return res.status(503).json({ error: 'HUNTER_API_KEY not configured' });
  const { email } = req.body;
  if (!email) return res.status(400).json({ error: 'email required' });
  try {
    const r = await fetch(
      `https://api.hunter.io/v2/email-verifier?email=${encodeURIComponent(email)}&api_key=${key}`
    );
    res.json(await r.json());
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Anthropic AI proxy ────────────────────────────────────────────────────────

const TRICORD_SYSTEM = `You are an AI outreach specialist for TRICORD I.T Solutions (tricordit.ca) — an Ontario MSP specialising in AI Strategy & Governance for mid-market companies.

TRICORD PROFILE
- Founder: Dwight G., 20+ years senior I.T. architect
- Location: Milton, Ontario
- Primary hook: AI Strategy & Governance — helping Ontario businesses adopt AI safely with proper governance, data-residency, and compliance frameworks
- Services: AI Strategy Engagement (Phase 1 fixed-fee 4–8 wk), Fractional CIO, Managed IT
- Target sectors: Law firms, manufacturers, CPA/accounting, pharma/life sciences, engineering/architecture, insurance brokerages, healthcare

OUTREACH STYLE
- Lead with the prospect's pain, not TRICORD's services
- Sector-specific AI hooks (law: shadow AI on client files; manufacturing: OT/IT/AI convergence; CPA: CPA Ontario guidance + data residency; etc.)
- Short, confident, no jargon. Dwight's voice: direct, experienced, peer-level
- Always personalise to a named contact and company

When writing emails: subject ≤ 50 chars, body ≤ 150 words for Touch 1.
When writing LinkedIn messages: ≤ 300 chars for connection request, ≤ 500 for follow-up.`;

app.post('/api/ai/chat', async (req, res) => {
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return res.status(503).json({ error: 'ANTHROPIC_API_KEY not configured' });
  const { messages } = req.body;
  if (!messages?.length) return res.status(400).json({ error: 'messages required' });
  try {
    // Dynamic import for ESM-only SDK versions; fall back to require
    let Anthropic;
    try { Anthropic = require('@anthropic-ai/sdk'); } catch { Anthropic = (await import('@anthropic-ai/sdk')).default; }
    const client = new Anthropic({ apiKey: key });
    const resp = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 1500,
      system: [{ type: 'text', text: TRICORD_SYSTEM, cache_control: { type: 'ephemeral' } }],
      messages,
    });
    res.json({ content: resp.content[0].text, usage: resp.usage });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Page routes ───────────────────────────────────────────────────────────────

app.get('/dashboard', (_req, res) => res.sendFile(path.join(__dirname, 'dashboard.html')));
app.get('/outreach', (_req, res) => res.sendFile(path.join(__dirname, 'outreach.html')));
app.get('/eval',     (_req, res) => res.sendFile(path.join(__dirname, 'eval_review.html')));
app.get('*',         (_req, res) => res.sendFile(path.join(__dirname, 'index.html')));

app.listen(PORT, () => console.log(`Tricord AI running on http://localhost:${PORT}`));
