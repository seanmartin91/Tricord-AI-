const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname)));

app.post('/api/contact', (req, res) => {
  const { firstName, lastName, email, company, industry, challenge } = req.body;

  if (!firstName || !email) {
    return res.status(400).json({ error: 'Name and email are required.' });
  }

  const entry = {
    timestamp: new Date().toISOString(),
    firstName,
    lastName,
    email,
    company,
    industry,
    challenge,
  };

  // Append to leads.jsonl for persistence without a database
  const leadsFile = path.join(__dirname, 'leads.jsonl');
  fs.appendFileSync(leadsFile, JSON.stringify(entry) + '\n');

  console.log(`[lead] ${email} — ${company || 'no company'}`);
  res.json({ ok: true });
});

app.get('/api/leads', (req, res) => {
  const authHeader = req.headers['authorization'];
  if (authHeader !== `Bearer ${process.env.ADMIN_TOKEN || 'tricord-admin'}`) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  const leadsFile = path.join(__dirname, 'leads.jsonl');
  if (!fs.existsSync(leadsFile)) return res.json([]);
  const leads = fs.readFileSync(leadsFile, 'utf8')
    .trim().split('\n').filter(Boolean)
    .map(line => JSON.parse(line));
  res.json(leads);
});

app.get('/dashboard', (req, res) => {
  res.sendFile(path.join(__dirname, 'dashboard.html'));
});

app.get('/eval', (req, res) => {
  res.sendFile(path.join(__dirname, 'eval_review.html'));
});

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, () => {
  console.log(`Tricord AI running on http://localhost:${PORT}`);
});
