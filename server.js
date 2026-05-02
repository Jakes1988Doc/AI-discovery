// ============================================================
//  Business Discovery Agent — Backend (Render.com)
// ============================================================
//  Handles:
//    1. Proxying calls to the Anthropic Claude API
//    2. Saving completed transcripts to a Google Sheet
//
//  Environment variables required (set in Render dashboard):
//    ANTHROPIC_API_KEY     -- your Anthropic API key (sk-ant-...)
//    GOOGLE_SHEETS_ID      -- the long ID from your Sheet's URL
//    GOOGLE_SERVICE_ACCOUNT_JSON  -- the entire service-account JSON, as a single line
//    ALLOWED_ORIGIN        -- e.g. https://yourname.github.io  (keeps random sites
//                              from hammering your backend)
//    PORT                  -- Render sets this automatically; default 3000 locally
// ============================================================

const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');

const app = express();
app.use(express.json({ limit: '2mb' }));

// CORS — only allow your frontend origin
const allowedOrigin = process.env.ALLOWED_ORIGIN || '*';
app.use(cors({ origin: allowedOrigin === '*' ? true : allowedOrigin }));

const PORT = process.env.PORT || 3000;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const SHEET_ID = process.env.GOOGLE_SHEETS_ID;
const SVC_JSON = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

const SYSTEM_PROMPT = `You are a skilled business discovery interviewer conducting a voice conversation with a business owner. Your goal is to deeply understand their business, current systems, manual processes, and pain points so that someone else can later identify AI automation opportunities.

ABSOLUTE RULES:
1. NEVER recommend, suggest, or hint at solutions, tools, or automations. Not even subtly.
2. NEVER say things like "you could try", "have you considered", "AI could help with that", or "many businesses use X for that".
3. Your only job is to ASK and LISTEN. Capture context, not give advice.
4. Stay curious. Dig into specifics — numbers, frequencies, who is involved, what breaks, when it started.
5. Use natural conversational language. You are speaking, not writing. Keep questions to 1-2 sentences max.
6. Reference earlier things they said when relevant (e.g. "you mentioned booking takes 3 hours — is that one person doing all of it?").
7. If an answer is vague, ask for a specific example. If an answer is rich, move on.

You will be told the current discovery topic and given the conversation history. Decide whether to:
- Ask a follow-up on the same topic (if there's more to uncover)
- Move to the next topic (if the current one feels well-explored, usually after 2-4 exchanges per topic)

Respond in JSON only, no markdown fences, no preamble:
{
  "next_question": "the exact words you will speak",
  "action": "followup" or "next_topic" or "wrap_up",
  "reasoning": "one short sentence on why"
}`;

// ============================================================
//  Health check
// ============================================================
app.get('/', (req, res) => {
  res.json({ ok: true, service: 'discovery-agent-backend', time: new Date().toISOString() });
});

// ============================================================
//  POST /api/next-question
//  Generates the next dynamic question via Claude.
// ============================================================
app.post('/api/next-question', async (req, res) => {
  if (!ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY not configured' });
  }

  const {
    messages = [],
    currentTopic = '',
    topicIndex = 0,
    totalTopics = 7,
    exchangesOnTopic = 1,
    remainingTopics = [],
    isLastTopic = false
  } = req.body || {};

  const directive = `CURRENT TOPIC: ${currentTopic} (topic ${topicIndex + 1} of ${totalTopics})
EXCHANGES ON THIS TOPIC SO FAR: ${exchangesOnTopic}
REMAINING TOPICS AFTER THIS: ${remainingTopics.length ? remainingTopics.join(', ') : 'none — this is the last topic'}

Based on the conversation history, decide whether to ask another follow-up on "${currentTopic}" or move to the next topic. Aim for 2-4 exchanges per topic — fewer if the answer was rich, more if it was thin. ${isLastTopic ? 'This is the LAST topic. After 2-3 exchanges here, set action to "wrap_up".' : ''}

Respond with JSON only.`;

  const apiMessages = [...messages, { role: 'user', content: directive }];

  try {
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5-20250929',
        max_tokens: 400,
        system: SYSTEM_PROMPT,
        messages: apiMessages
      })
    });
    if (!r.ok) {
      const errText = await r.text();
      console.error('Anthropic error', r.status, errText);
      return res.status(500).json({ error: 'Anthropic API error', detail: r.status });
    }
    const data = await r.json();
    const text = data.content
      .filter(c => c.type === 'text')
      .map(c => c.text)
      .join('')
      .trim();
    const cleaned = text.replace(/```json\s*|\s*```/g, '').trim();
    let parsed;
    try {
      parsed = JSON.parse(cleaned);
    } catch (e) {
      console.warn('Failed to parse JSON, using fallback', text);
      parsed = { next_question: 'Can you tell me more about that?', action: 'followup', reasoning: 'JSON parse fallback' };
    }
    res.json(parsed);
  } catch (err) {
    console.error('next-question error', err);
    res.status(500).json({ error: 'Server error generating next question' });
  }
});

// ============================================================
//  POST /api/save-transcript
//  Appends the completed call to the configured Google Sheet.
//  Sheet should have a header row like:
//  Timestamp | Client | Business | Duration (s) | Topics covered | Transcript
// ============================================================
app.post('/api/save-transcript', async (req, res) => {
  if (!SHEET_ID || !SVC_JSON) {
    return res.status(500).json({ error: 'Google Sheets not configured' });
  }

  const {
    clientName = '',
    businessName = '',
    startTime = null,
    endTime = null,
    durationSeconds = 0,
    topicsCovered = [],
    transcript = []
  } = req.body || {};

  // Format the transcript as a single readable cell
  const transcriptText = transcript.map(t => {
    const role = t.role === 'agent' ? 'AGENT' : 'CLIENT';
    const topic = t.topic ? ' (' + t.topic + ')' : '';
    return '[' + (t.time || '') + '] ' + role + topic + ': ' + (t.text || '');
  }).join('\n\n');

  try {
    const credentials = JSON.parse(SVC_JSON);
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });
    const sheets = google.sheets({ version: 'v4', auth });

    const row = [
      new Date().toISOString(),
      clientName,
      businessName,
      durationSeconds,
      topicsCovered.join(', '),
      transcriptText
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: 'Sheet1!A1',
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] }
    });

    res.json({ ok: true });
  } catch (err) {
    console.error('save-transcript error', err);
    res.status(500).json({ error: 'Could not save to Google Sheets', detail: err.message });
  }
});

// ============================================================
app.listen(PORT, () => {
  console.log('Discovery agent backend listening on port', PORT);
});
