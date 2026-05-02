// ============================================================
//  Business Discovery Agent — Backend (Render.com) — v2
// ============================================================
//  Handles:
//    1. Proxying calls to the Anthropic Claude API
//    2. Generating a structured Word document summary per call
//    3. Saving raw transcript to Google Sheets
//    4. Emailing the Word summary to the client
//
//  Environment variables required (set in Render dashboard):
//    ANTHROPIC_API_KEY            — your Anthropic key (sk-ant-...)
//    GOOGLE_SHEETS_ID             — long ID from your Sheet's URL
//    GOOGLE_SERVICE_ACCOUNT_JSON  — full service-account JSON contents
//    ALLOWED_ORIGIN               — your GitHub Pages URL (no trailing slash)
//    RESEND_API_KEY               — get free at https://resend.com (3000 emails/mo free)
//    FROM_EMAIL                   — e.g. "AI Discovery <onboarding@resend.dev>"
//                                    OR a verified domain address
//    BCC_EMAIL                    — (optional) your own email to receive a copy of every summary
//    PORT                         — set automatically by Render
// ============================================================

const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  LevelFormat, BorderStyle, PageOrientation
} = require('docx');

const app = express();
app.use(express.json({ limit: '2mb' }));

const allowedOrigin = process.env.ALLOWED_ORIGIN || '*';
app.use(cors({ origin: allowedOrigin === '*' ? true : allowedOrigin }));

const PORT = process.env.PORT || 3000;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const SHEET_ID = process.env.GOOGLE_SHEETS_ID;
const SVC_JSON = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const RESEND_API_KEY = process.env.RESEND_API_KEY;
const FROM_EMAIL = process.env.FROM_EMAIL || 'AI Discovery <onboarding@resend.dev>';
const BCC_EMAIL = process.env.BCC_EMAIL || '';

const INTERVIEWER_PROMPT = `You are a skilled business discovery interviewer conducting a voice conversation with a business owner. Your goal is to deeply understand their business, current systems, manual processes, and pain points so that someone else can later identify AI automation opportunities.

ABSOLUTE RULES:
1. NEVER recommend, suggest, or hint at solutions, tools, or automations. Not even subtly.
2. NEVER say things like "you could try", "have you considered", "AI could help with that", or "many businesses use X for that".
3. Your only job is to ASK and LISTEN. Capture context, not give advice.
4. Stay curious. Dig into specifics — numbers, frequencies, who is involved, what breaks, when it started.
5. Use natural conversational language. You are speaking, not writing. Keep questions to 1-2 sentences max.
6. Reference earlier things they said when relevant.
7. If an answer is vague, ask for a specific example. If an answer is rich, move on.

Respond in JSON only:
{
  "next_question": "the exact words you will speak",
  "action": "followup" or "next_topic" or "wrap_up",
  "reasoning": "one short sentence on why"
}`;

const SUMMARY_PROMPT = `You are summarising a business discovery interview. The interview was conducted by a voice agent that asked the business owner about their operations, systems, manual work, pain points, and priorities.

Produce a structured JSON summary with these exact fields. Be specific — pull out numbers, tool names, frequencies, and direct phrasing. Do NOT recommend solutions. This is purely a capture document.

Respond ONLY with valid JSON, no markdown fences:
{
  "executiveSummary": "2-3 sentence high-level overview of the business and its key challenges",
  "businessOverview": "What they do, who their customers are, team size, business model",
  "currentSystems": ["list of tools, software, and systems they currently use"],
  "manualWork": ["list of specific manual or repetitive tasks identified, with frequency/time-cost where stated"],
  "painPoints": ["list of specific pain points and frustrations, with context on impact"],
  "customerInteractions": "Summary of how they handle customer inquiries, support, and follow-ups",
  "priorities": ["list of stated priorities and what they wish was different"],
  "notableQuotes": ["2-4 direct verbatim quotes that capture the business owner's voice"],
  "automationOpportunities": ["objective list of areas where automation COULD apply, framed as observations not recommendations — e.g. 'Quoting currently takes 3 hours per week and is done manually in Excel'"]
}`;

// ============================================================
//  Health check
// ============================================================
app.get('/', (req, res) => {
  res.json({
    ok: true,
    service: 'discovery-agent-backend',
    version: '2.0',
    time: new Date().toISOString()
  });
});

// ============================================================
//  POST /api/next-question
// ============================================================
app.post('/api/next-question', async (req, res) => {
  if (!ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY not configured' });
  }

  const {
    messages = [], currentTopic = '', topicIndex = 0, totalTopics = 7,
    exchangesOnTopic = 1, remainingTopics = [], isLastTopic = false
  } = req.body || {};

  const directive = `CURRENT TOPIC: ${currentTopic} (topic ${topicIndex + 1} of ${totalTopics})
EXCHANGES ON THIS TOPIC SO FAR: ${exchangesOnTopic}
REMAINING TOPICS AFTER THIS: ${remainingTopics.length ? remainingTopics.join(', ') : 'none — this is the last topic'}

Based on the conversation history, decide whether to ask another follow-up on "${currentTopic}" or move to the next topic. Aim for 2-4 exchanges per topic — fewer if rich, more if thin. ${isLastTopic ? 'This is the LAST topic. After 2-3 exchanges, set action to "wrap_up".' : ''}

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
        system: INTERVIEWER_PROMPT,
        messages: apiMessages
      })
    });
    if (!r.ok) {
      const errText = await r.text();
      console.error('Anthropic error', r.status, errText);
      return res.status(500).json({ error: 'Anthropic API error', detail: r.status });
    }
    const data = await r.json();
    const text = data.content.filter(c => c.type === 'text').map(c => c.text).join('').trim();
    const cleaned = text.replace(/```json\s*|\s*```/g, '').trim();
    let parsed;
    try {
      parsed = JSON.parse(cleaned);
    } catch (e) {
      console.warn('JSON parse fallback', text);
      parsed = { next_question: 'Can you tell me more about that?', action: 'followup' };
    }
    res.json(parsed);
  } catch (err) {
    console.error('next-question error', err);
    res.status(500).json({ error: 'Server error' });
  }
});

// ============================================================
//  Generate structured summary via Claude
// ============================================================
async function generateSummary(transcript, clientName, businessName) {
  const transcriptText = transcript.map(t => {
    const role = t.role === 'agent' ? 'AGENT' : 'CLIENT';
    return role + ': ' + t.text;
  }).join('\n\n');

  const r = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-5-20250929',
      max_tokens: 2500,
      system: SUMMARY_PROMPT,
      messages: [{
        role: 'user',
        content: `Client: ${clientName}\nBusiness: ${businessName}\n\nTRANSCRIPT:\n\n${transcriptText}\n\nReturn JSON only.`
      }]
    })
  });
  if (!r.ok) throw new Error('Summary API failed: ' + r.status);
  const data = await r.json();
  const text = data.content.filter(c => c.type === 'text').map(c => c.text).join('').trim();
  const cleaned = text.replace(/```json\s*|\s*```/g, '').trim();
  return JSON.parse(cleaned);
}

// ============================================================
//  Build a Word document from the structured summary
// ============================================================
function buildSummaryDocx(summary, meta) {
  const { clientName, businessName, dateStr, durationStr } = meta;

  // Helper to create a section heading
  const heading = (text) => new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 360, after: 120 },
    children: [new TextRun({ text, bold: true, size: 28, color: '1F2937' })]
  });

  // Helper to create a paragraph of body text
  const body = (text) => new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, size: 22, color: '374151' })]
  });

  // Helper to create a bulleted item
  const bullet = (text) => new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 22, color: '374151' })]
  });

  // Helper for quote-style
  const quote = (text) => new Paragraph({
    spacing: { before: 80, after: 80 },
    indent: { left: 360 },
    border: { left: { style: BorderStyle.SINGLE, size: 16, color: '4A8CFF', space: 8 } },
    children: [new TextRun({ text: '"' + text + '"', italics: true, size: 22, color: '4B5563' })]
  });

  const children = [];

  // Title
  children.push(new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { after: 80 },
    children: [new TextRun({ text: 'AI Discovery Call Summary', bold: true, size: 40, color: '0F172A' })]
  }));
  children.push(new Paragraph({
    spacing: { after: 360 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '4A8CFF', space: 4 } },
    children: [new TextRun({ text: businessName, size: 28, color: '4A8CFF', bold: true })]
  }));

  // Meta block
  children.push(new Paragraph({
    spacing: { after: 80 },
    children: [
      new TextRun({ text: 'Client: ', bold: true, size: 22, color: '374151' }),
      new TextRun({ text: clientName, size: 22, color: '374151' })
    ]
  }));
  children.push(new Paragraph({
    spacing: { after: 80 },
    children: [
      new TextRun({ text: 'Date: ', bold: true, size: 22, color: '374151' }),
      new TextRun({ text: dateStr, size: 22, color: '374151' })
    ]
  }));
  children.push(new Paragraph({
    spacing: { after: 240 },
    children: [
      new TextRun({ text: 'Call duration: ', bold: true, size: 22, color: '374151' }),
      new TextRun({ text: durationStr, size: 22, color: '374151' })
    ]
  }));

  // Executive summary
  if (summary.executiveSummary) {
    children.push(heading('Executive Summary'));
    children.push(body(summary.executiveSummary));
  }

  // Business overview
  if (summary.businessOverview) {
    children.push(heading('Business Overview'));
    children.push(body(summary.businessOverview));
  }

  // Current systems
  if (summary.currentSystems && summary.currentSystems.length) {
    children.push(heading('Current Systems & Tools'));
    summary.currentSystems.forEach(s => children.push(bullet(s)));
  }

  // Manual work
  if (summary.manualWork && summary.manualWork.length) {
    children.push(heading('Manual & Repetitive Work'));
    summary.manualWork.forEach(s => children.push(bullet(s)));
  }

  // Pain points
  if (summary.painPoints && summary.painPoints.length) {
    children.push(heading('Pain Points'));
    summary.painPoints.forEach(s => children.push(bullet(s)));
  }

  // Customer interactions
  if (summary.customerInteractions) {
    children.push(heading('Customer Interactions'));
    children.push(body(summary.customerInteractions));
  }

  // Priorities
  if (summary.priorities && summary.priorities.length) {
    children.push(heading('Priorities & Stated Goals'));
    summary.priorities.forEach(s => children.push(bullet(s)));
  }

  // Automation opportunities (observations only, never recommendations)
  if (summary.automationOpportunities && summary.automationOpportunities.length) {
    children.push(heading('Areas Worth Exploring'));
    children.push(body('The following observations from the call are areas where automation could potentially apply. These are descriptions of the current state, not recommendations:'));
    summary.automationOpportunities.forEach(s => children.push(bullet(s)));
  }

  // Notable quotes
  if (summary.notableQuotes && summary.notableQuotes.length) {
    children.push(heading('Notable Quotes'));
    summary.notableQuotes.forEach(q => children.push(quote(q)));
  }

  // Footer
  children.push(new Paragraph({
    spacing: { before: 480 },
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: 'E5E7EB', space: 4 } },
    children: [new TextRun({ text: 'Generated by AI Discovery — this is a capture document, not professional advice.', italics: true, size: 18, color: '9CA3AF' })]
  }));

  const doc = new Document({
    creator: 'AI Discovery',
    title: 'AI Discovery Call Summary - ' + businessName,
    styles: {
      default: { document: { run: { font: 'Calibri', size: 22 } } }
    },
    numbering: {
      config: [{
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '•',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: children
    }]
  });

  return Packer.toBuffer(doc);
}

// ============================================================
//  Send the Word document via Resend
// ============================================================
async function emailSummary(toEmail, clientName, businessName, docBuffer) {
  if (!RESEND_API_KEY) {
    console.warn('RESEND_API_KEY not set — skipping email send');
    return { skipped: true };
  }

  const firstName = (clientName || '').split(' ')[0] || 'there';
  const subject = 'Your AI Discovery summary — ' + businessName;
  const htmlBody = `
    <div style="font-family: -apple-system, sans-serif; max-width: 560px; margin: 0 auto; color: #374151;">
      <h2 style="color: #0F172A; margin-bottom: 16px;">Hi ${firstName},</h2>
      <p style="font-size: 15px; line-height: 1.6;">
        Thanks for taking the time today. Attached is your written summary from our discovery call about <strong>${businessName}</strong>.
      </p>
      <p style="font-size: 15px; line-height: 1.6;">
        It captures your current systems, manual workload, pain points, and stated priorities — straight from the call, in your own words.
      </p>
      <p style="font-size: 15px; line-height: 1.6;">
        It's yours to keep, share, or use however you like. There are no recommendations in this document — just an honest map of where things stand.
      </p>
      <p style="font-size: 15px; line-height: 1.6; margin-top: 24px;">
        All the best,<br/>
        <strong>The AI Discovery team</strong>
      </p>
      <hr style="border: none; border-top: 1px solid #E5E7EB; margin: 32px 0;" />
      <p style="font-size: 12px; color: #9CA3AF;">
        This summary was generated automatically from your call transcript. If anything looks off, just reply to this email.
      </p>
    </div>
  `;

  const filename = 'ai-discovery-summary-' +
    (businessName || 'call').replace(/[^a-z0-9]+/gi, '-').toLowerCase() +
    '-' + new Date().toISOString().slice(0, 10) + '.docx';

  const payload = {
    from: FROM_EMAIL,
    to: [toEmail],
    subject: subject,
    html: htmlBody,
    attachments: [{
      filename: filename,
      content: docBuffer.toString('base64')
    }]
  };
  if (BCC_EMAIL) payload.bcc = [BCC_EMAIL];

  const r = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + RESEND_API_KEY
    },
    body: JSON.stringify(payload)
  });
  if (!r.ok) {
    const errText = await r.text();
    console.error('Resend error', r.status, errText);
    throw new Error('Email send failed: ' + r.status);
  }
  return await r.json();
}

// ============================================================
//  POST /api/save-transcript
//  Saves transcript to Sheet, generates Word summary, emails it
// ============================================================
app.post('/api/save-transcript', async (req, res) => {
  const {
    clientName = '', businessName = '', email = '',
    startTime = null, endTime = null, durationSeconds = 0,
    topicsCovered = [], transcript = []
  } = req.body || {};

  // 1. Save raw transcript to Google Sheets
  const transcriptText = transcript.map(t => {
    const role = t.role === 'agent' ? 'AGENT' : 'CLIENT';
    const topic = t.topic ? ' (' + t.topic + ')' : '';
    return '[' + (t.time || '') + '] ' + role + topic + ': ' + (t.text || '');
  }).join('\n\n');

  let sheetSaved = false;
  if (SHEET_ID && SVC_JSON) {
    try {
      const credentials = JSON.parse(SVC_JSON);
      const auth = new google.auth.GoogleAuth({
        credentials,
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
      });
      const sheets = google.sheets({ version: 'v4', auth });
      const row = [
        new Date().toISOString(),
        clientName, businessName, email,
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
      sheetSaved = true;
    } catch (err) {
      console.error('Sheet save error', err.message);
    }
  }

  // 2. Generate structured summary + Word document + email
  let emailed = false;
  let summaryError = null;
  if (email && transcript.length > 1) {
    try {
      const summary = await generateSummary(transcript, clientName, businessName);
      const docBuffer = await buildSummaryDocx(summary, {
        clientName,
        businessName,
        dateStr: new Date(startTime || Date.now()).toLocaleString('en-GB', { dateStyle: 'long', timeStyle: 'short' }),
        durationStr: Math.floor(durationSeconds / 60) + ' min ' + (durationSeconds % 60) + ' sec'
      });
      const emailResult = await emailSummary(email, clientName, businessName, docBuffer);
      emailed = !emailResult.skipped;
    } catch (err) {
      console.error('Summary/email error', err);
      summaryError = err.message;
    }
  }

  res.json({
    ok: sheetSaved || emailed,
    sheetSaved,
    emailed,
    summaryError
  });
});

// ============================================================
app.listen(PORT, () => {
  console.log('Discovery agent backend v2 listening on port', PORT);
});
