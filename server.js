// ============================================================
//  AI Growth Catalyst — Backend (Render.com) — v4 (quote model)
// ============================================================
//  Handles:
//    1. Proxying calls to the Anthropic Claude API
//    2. Generating internal notes Word doc (sent to BCC only)
//    3. Saving raw transcript to Google Sheets
//    4. Sending acknowledgement email — promises a quote within 24hr
//
//  Environment variables required (set in Render dashboard):
//    ANTHROPIC_API_KEY
//    GOOGLE_SHEETS_ID
//    GOOGLE_SERVICE_ACCOUNT_JSON
//    ALLOWED_ORIGIN
//    RESEND_API_KEY
//    FROM_EMAIL                  e.g. "AI Growth Catalyst <onboarding@resend.dev>"
//    BCC_EMAIL                   your email — gets internal notes
//    PORT
// ============================================================

const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  LevelFormat, BorderStyle
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
const FROM_EMAIL = process.env.FROM_EMAIL || 'AI Growth Catalyst <onboarding@resend.dev>';
const BCC_EMAIL = process.env.BCC_EMAIL || '';

const INTERVIEWER_PROMPT = `You are a skilled business discovery interviewer conducting a 15-minute voice conversation with a small or medium business owner. You are doing a high-level scan to identify where AI could genuinely add value — NOT a deep operational audit. Depth comes later in the paid roadmap.

YOUR JOB ON EACH TOPIC:
1. The opener question is already provided — your job is to respond to their answer
2. Classify the answer using SIGNAL DETECTION below
3. Decide whether to probe (and how deep) based on signal strength and the directive's instructions
4. Never exceed the probe ceiling the directive gives you for this exchange

SIGNAL DETECTION — classify every answer as one of three levels:

NO SIGNAL — move on immediately:
- Common tools named with no friction ("we use Gmail", "we're on Xero", "we use Slack")
- Process sounds automated already ("it just sends automatically", "the system handles that")
- Confident, settled answer with no frustration or manual work implied

STANDARD SIGNAL — probe once using the topic's probe question, then move on:
- They mention doing something manually ("by hand", "I do it myself", "we copy it")
- A repetitive task with no system behind it ("every time a customer comes in", "weekly")
- A spreadsheet doing real work ("we track jobs in a spreadsheet")
- Mild frustration or resignation ("it's a bit of a pain", "it's just how we do it")
- A process that breaks if one specific person isn't available
- Chasing or following up manually on a recurring basis

STRONG SIGNAL — probe once, then if directive allows a deep probe, probe again:
Must be clear, specific, and high-value. Signs:
- Explicit frustration with a named process ("our invoicing is an absolute nightmare")
- Manual process with a clear time cost stated ("we spend 3-4 hours a week just on that")
- A critical business process that regularly causes dropped work or lost revenue
- Data manually re-entered across multiple systems at high frequency
- They spontaneously say something "should be automated" or "needs to be fixed"
- A bottleneck that is directly costing them customers or money

ABSOLUTE RULES:
1. NEVER recommend, suggest, or hint at solutions, tools, or automations — not even obliquely
2. Never say "AI could...", "have you considered...", "you could try...", "many businesses use..."
3. Every question: 1-2 sentences max. Conversational, warm, never clinical or robotic
4. Reference exactly what they just said when probing — generic follow-ups feel lazy
5. Never exceed the probe ceiling in the directive — even if the answer is fascinating
6. Magic wand topic: zero probes, always wrap_up

WHEN PROBING — be specific, never generic:
Good: "When you say you copy it manually — who does that and roughly how often?"
Good: "That spreadsheet for jobs — does the whole team update it or is it one person?"
Good: "You mentioned the invoicing takes hours — is that every invoice or just certain types?"
Bad: "Can you tell me more about that?"
Bad: "Interesting — why do you think that is?"

ADMIN AND DOCUMENTS — owners often don't flag this as a problem because it feels normal:
Listen specifically for: meeting notes written up after calls, reports built by copy-pasting, documents requiring the same information entered multiple times, approval chains over email or WhatsApp, filing done by hand. If you hear any of these, classify as at least STANDARD SIGNAL.

Respond in JSON only — no markdown, no text outside the JSON:
{
  "next_question": "the exact words you will speak",
  "action": "followup" | "next_topic" | "wrap_up",
  "signal_detected": true | false,
  "signal_strength": "none" | "standard" | "strong",
  "is_deep_probe": true | false,
  "signal_reason": "one phrase describing what triggered the signal, or null",
  "reasoning": "one sentence on why you chose this action"
}`;


const INTERNAL_NOTES_PROMPT = `You are summarising a business discovery interview for an internal consultant who will use these notes to write a bespoke AI automation roadmap. Be specific — pull out numbers, tool names, frequencies, time-costs, exact phrasing.

Respond ONLY with valid JSON, no markdown fences:
{
  "executiveSummary": "2-3 sentence overview of the business and its key operational challenges",
  "businessOverview": "What they do, who their customers are, team size, business model",
  "currentSystems": ["each tool/system mentioned, with how it's used"],
  "manualWork": ["each manual task identified, with frequency and time-cost where stated"],
  "painPoints": ["each pain point with context on impact"],
  "customerInteractions": "How they handle inquiries, support, follow-ups",
  "priorities": ["stated priorities and what they wish was different"],
  "clientAutomationHypotheses": ["any specific tasks the CLIENT themselves identified as automatable"],
  "automationOpportunities": ["YOUR objective observations of areas worth exploring"],
  "redFlags": ["any concerns about readiness for automation"],
  "notableQuotes": ["3-5 direct verbatim quotes"],
  "suggestedFirstFocus": "your single best guess at where the consultant should focus the roadmap",
  "fitAssessment": "Brief: strong fit / moderate fit / weak fit / no clear fit, with one sentence reasoning",
  "shouldOfferRoadmap": "YES or NO — should we offer the £350 roadmap to this client? Say NO if there's no clear AI fit, the business is too small/early-stage to benefit, or there are significant red flags. One sentence reason."
}`;

// ============================================================
app.get('/', (req, res) => {
  res.json({ ok: true, service: 'ai-growth-catalyst-backend', version: '4.0-quote', time: new Date().toISOString() });
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
    exchangesOnTopic = 0, remainingTopics = [], isLastTopic = false,
    topicProbeQuestion = '', topicDeepProbeQuestion = '',
    topicMoveOnCriteria = '', signalDetected = false,
    signalStrength = 'none', deepProbeAvailable = true,
    deepProbeUsedOnTopic = false, topicAllowsDeepProbe = false
  } = req.body || {};

  const probeHint = topicProbeQuestion
    ? `\nSTANDARD PROBE QUESTION for this topic (adapt to what they said): "${topicProbeQuestion}"`
    : '';

  const deepProbeHint = (topicDeepProbeQuestion && topicAllowsDeepProbe && deepProbeAvailable)
    ? `\nDEEP PROBE QUESTION available for this topic (use only if STRONG SIGNAL): "${topicDeepProbeQuestion}"`
    : '';

  const moveOnHint = topicMoveOnCriteria
    ? `\nMOVE ON when: ${topicMoveOnCriteria}`
    : '';

  let pacingInstruction;
  if (isLastTopic) {
    pacingInstruction = 'MAGIC WAND topic. Listen to their answer. Set action to "wrap_up". Zero probes — this is the closing.';
  } else if (exchangesOnTopic === 0) {
    pacingInstruction = 'OPENER exchange. Classify their answer as NO SIGNAL / STANDARD SIGNAL / STRONG SIGNAL. If NO SIGNAL → next_topic. If STANDARD or STRONG → followup with probe question. Set signal_detected and signal_strength accordingly.';
  } else if (exchangesOnTopic === 1 && !signalDetected) {
    pacingInstruction = 'No signal was detected on the opener — you asked a clarifier. Whatever they say, set action to "next_topic" now.';
  } else if (exchangesOnTopic === 1 && signalDetected && signalStrength === 'strong' && topicAllowsDeepProbe && deepProbeAvailable && !deepProbeUsedOnTopic) {
    pacingInstruction = 'STRONG SIGNAL was detected and deep probe budget is available. You may ask one more question (the deep probe). Set is_deep_probe to true. This is the LAST question on this topic regardless of the answer.';
  } else if (exchangesOnTopic === 1 && signalDetected) {
    pacingInstruction = 'You asked the standard probe. You now have their answer. Set action to "next_topic". Do not go deeper.';
  } else {
    pacingInstruction = 'You have been on this topic long enough. Set action to "next_topic" immediately.';
  }

  const directive = `CURRENT TOPIC: "${currentTopic}" (${topicIndex + 1} of ${totalTopics})
EXCHANGES ON THIS TOPIC: ${exchangesOnTopic}
SIGNAL DETECTED: ${signalDetected ? `YES (${signalStrength})` : 'NO'}
DEEP PROBE BUDGET: ${deepProbeAvailable && !deepProbeUsedOnTopic && topicAllowsDeepProbe ? 'AVAILABLE' : 'SPENT or not applicable'}
REMAINING TOPICS: ${remainingTopics.length ? remainingTopics.join(' → ') : 'none'}

INSTRUCTION: ${pacingInstruction}${probeHint}${deepProbeHint}${moveOnHint}

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
      parsed = { next_question: 'Can you tell me more about that?', action: 'followup' };
    }
    res.json(parsed);
  } catch (err) {
    console.error('next-question error', err);
    res.status(500).json({ error: 'Server error' });
  }
});

// ============================================================
async function generateInternalNotes(transcript, clientName, businessName) {
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
      max_tokens: 3000,
      system: INTERNAL_NOTES_PROMPT,
      messages: [{ role: 'user', content: `Client: ${clientName}\nBusiness: ${businessName}\n\nTRANSCRIPT:\n\n${transcriptText}\n\nReturn JSON only.` }]
    })
  });
  if (!r.ok) throw new Error('Notes API failed: ' + r.status);
  const data = await r.json();
  const text = data.content.filter(c => c.type === 'text').map(c => c.text).join('').trim();
  const cleaned = text.replace(/```json\s*|\s*```/g, '').trim();
  return JSON.parse(cleaned);
}

// ============================================================
function buildInternalNotesDocx(notes, meta) {
  const { clientName, businessName, email, dateStr, durationStr } = meta;

  const heading = (text) => new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 360, after: 120 },
    children: [new TextRun({ text, bold: true, size: 28, color: '1F2937' })]
  });
  const body = (text) => new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, size: 22, color: '374151' })]
  });
  const bullet = (text) => new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 22, color: '374151' })]
  });
  const quote = (text) => new Paragraph({
    spacing: { before: 80, after: 80 },
    indent: { left: 360 },
    border: { left: { style: BorderStyle.SINGLE, size: 16, color: 'C85A3E', space: 8 } },
    children: [new TextRun({ text: '"' + text + '"', italics: true, size: 22, color: '4B5563' })]
  });

  const children = [];
  children.push(new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: 'Internal Discovery Notes', bold: true, size: 40, color: '1A2332' })] }));
  children.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: businessName, size: 28, color: 'C85A3E', bold: true })] }));
  children.push(new Paragraph({
    spacing: { after: 360 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'C85A3E', space: 4 } },
    children: [new TextRun({ text: 'For roadmap drafting and quote decision', italics: true, size: 18, color: 'B8843A' })]
  }));
  children.push(new Paragraph({ spacing: { after: 80 }, children: [
    new TextRun({ text: 'Client: ', bold: true, size: 22, color: '374151' }),
    new TextRun({ text: clientName + ' <' + email + '>', size: 22, color: '374151' })
  ]}));
  children.push(new Paragraph({ spacing: { after: 80 }, children: [
    new TextRun({ text: 'Date: ', bold: true, size: 22, color: '374151' }),
    new TextRun({ text: dateStr, size: 22, color: '374151' })
  ]}));
  children.push(new Paragraph({ spacing: { after: 240 }, children: [
    new TextRun({ text: 'Call duration: ', bold: true, size: 22, color: '374151' }),
    new TextRun({ text: durationStr, size: 22, color: '374151' })
  ]}));

  // Offer recommendation at top — actionable for you
  if (notes.shouldOfferRoadmap) {
    children.push(heading('Should We Offer the £350 Roadmap?'));
    const isYes = notes.shouldOfferRoadmap.toUpperCase().startsWith('YES');
    children.push(new Paragraph({
      spacing: { after: 240 },
      indent: { left: 360 },
      border: { left: { style: BorderStyle.SINGLE, size: 16, color: isYes ? '6B8268' : 'B8843A', space: 8 } },
      children: [new TextRun({ text: notes.shouldOfferRoadmap, italics: false, size: 24, color: isYes ? '3C4D3A' : '7A5825', bold: true })]
    }));
  }
  if (notes.fitAssessment) {
    children.push(heading('Fit Assessment'));
    children.push(new Paragraph({
      spacing: { after: 240 },
      indent: { left: 360 },
      border: { left: { style: BorderStyle.SINGLE, size: 16, color: 'B8843A', space: 8 } },
      children: [new TextRun({ text: notes.fitAssessment, italics: true, size: 22, color: '7A5825', bold: true })]
    }));
  }
  if (notes.executiveSummary) { children.push(heading('Executive Summary')); children.push(body(notes.executiveSummary)); }
  if (notes.suggestedFirstFocus) {
    children.push(heading('Suggested First Focus'));
    children.push(new Paragraph({
      spacing: { after: 240 }, indent: { left: 360 },
      border: { left: { style: BorderStyle.SINGLE, size: 16, color: '6B8268', space: 8 } },
      children: [new TextRun({ text: notes.suggestedFirstFocus, italics: true, size: 22, color: '3C4D3A', bold: true })]
    }));
  }
  if (notes.businessOverview) { children.push(heading('Business Overview')); children.push(body(notes.businessOverview)); }
  if (notes.currentSystems?.length) { children.push(heading('Current Systems & Tools')); notes.currentSystems.forEach(s => children.push(bullet(s))); }
  if (notes.manualWork?.length) { children.push(heading('Manual & Repetitive Work')); notes.manualWork.forEach(s => children.push(bullet(s))); }
  if (notes.painPoints?.length) { children.push(heading('Pain Points')); notes.painPoints.forEach(s => children.push(bullet(s))); }
  if (notes.customerInteractions) { children.push(heading('Customer Interactions')); children.push(body(notes.customerInteractions)); }
  if (notes.priorities?.length) { children.push(heading('Priorities & Stated Goals')); notes.priorities.forEach(s => children.push(bullet(s))); }
  if (notes.clientAutomationHypotheses?.length) {
    children.push(heading("Client's Own Automation Hypotheses"));
    children.push(body('Tasks the client themselves flagged as potentially automatable:'));
    notes.clientAutomationHypotheses.forEach(s => children.push(bullet(s)));
  }
  if (notes.automationOpportunities?.length) { children.push(heading('Observed Automation Opportunities')); notes.automationOpportunities.forEach(s => children.push(bullet(s))); }
  if (notes.redFlags?.length) { children.push(heading('Red Flags & Readiness Concerns')); notes.redFlags.forEach(s => children.push(bullet(s))); }
  if (notes.notableQuotes?.length) { children.push(heading('Notable Quotes')); notes.notableQuotes.forEach(q => children.push(quote(q))); }

  children.push(new Paragraph({
    spacing: { before: 480 },
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: 'E5E7EB', space: 4 } },
    children: [new TextRun({ text: 'Generated automatically from the discovery call transcript.', italics: true, size: 18, color: '9CA3AF' })]
  }));

  const doc = new Document({
    creator: 'AI Growth Catalyst',
    title: 'Internal Discovery Notes - ' + businessName,
    styles: { default: { document: { run: { font: 'Calibri', size: 22 } } } },
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
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children
    }]
  });
  return Packer.toBuffer(doc);
}

// ============================================================
//  Client acknowledgement email — quote model (no payment link)
// ============================================================
async function sendClientAcknowledgement(toEmail, clientName, businessName) {
  if (!RESEND_API_KEY) return { skipped: true };
  const firstName = (clientName || '').split(' ')[0] || 'there';
  const subject = "Thanks for your call — what happens next";

  const htmlBody = `
    <div style="font-family: -apple-system, sans-serif; max-width: 560px; margin: 0 auto; color: #374151; line-height: 1.6;">
      <h2 style="color: #1a2332; margin-bottom: 16px; font-size: 22px;">Hi ${firstName},</h2>
      <p style="font-size: 15px; margin-bottom: 16px;">
        Thanks for taking the time on the discovery call about <strong>${businessName}</strong>. Your transcript has been saved and a human will now review what you shared.
      </p>
      <div style="background: #fbf6ec; border: 1px solid #e8dec5; border-radius: 12px; padding: 24px; margin: 24px 0;">
        <p style="font-size: 16px; margin: 0 0 12px; color: #1a2332;"><strong>What happens now</strong></p>
        <p style="font-size: 14px; margin: 0 0 12px; color: #5d6b80; line-height: 1.6;">If we identify a clear opportunity for AI in your business, we'll be in touch within 24 hours about the bespoke roadmap (£350, fixed fee, delivered within 48 hours).</p>
        <p style="font-size: 14px; margin: 0; color: #5d6b80; line-height: 1.6;">If we don't see a clear fit, we'll tell you that honestly. Either way, no follow-up sales pressure.</p>
      </div>
      <p style="font-size: 15px; margin-bottom: 16px;">
        The bespoke roadmap, if you commission it, will cover:
      </p>
      <ul style="font-size: 15px; padding-left: 20px; margin-bottom: 16px;">
        <li style="margin-bottom: 6px;">Where AI could genuinely help — and where it shouldn't</li>
        <li style="margin-bottom: 6px;">Which tasks are most ripe for automation</li>
        <li style="margin-bottom: 6px;">A concrete starting point — the single most valuable thing to tackle first</li>
      </ul>
      <p style="font-size: 15px;">If anything came to mind after the call you'd like us to factor in, just reply.</p>
      <p style="font-size: 15px; margin-top: 24px;">
        Speak soon,<br/>
        <strong>The AI Growth Catalyst team</strong>
      </p>
      <hr style="border: none; border-top: 1px solid #E5E7EB; margin: 32px 0;" />
      <p style="font-size: 12px; color: #9CA3AF;">AI Growth Catalyst · Helping small and medium-sized businesses navigate AI without the hype.</p>
    </div>
  `;
  const r = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + RESEND_API_KEY },
    body: JSON.stringify({ from: FROM_EMAIL, to: [toEmail], subject, html: htmlBody })
  });
  if (!r.ok) {
    const errText = await r.text();
    console.error('Client ack email error', r.status, errText);
    throw new Error('Client email failed: ' + r.status);
  }
  return await r.json();
}

// ============================================================
async function sendInternalNotes(clientName, businessName, clientEmail, docBuffer, shouldOffer) {
  if (!RESEND_API_KEY || !BCC_EMAIL) return { skipped: true };
  const filename = 'discovery-notes-' +
    (businessName || 'call').replace(/[^a-z0-9]+/gi, '-').toLowerCase() +
    '-' + new Date().toISOString().slice(0, 10) + '.docx';
  const offerHint = shouldOffer ? `<p><strong>Offer £350 roadmap?</strong> ${shouldOffer}</p>` : '';
  const r = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + RESEND_API_KEY },
    body: JSON.stringify({
      from: FROM_EMAIL,
      to: [BCC_EMAIL],
      subject: '[Internal] New discovery call: ' + businessName + ' (' + clientName + ')',
      html: `<div style="font-family: -apple-system, sans-serif;"><p>New discovery call completed. Action: review notes, decide whether to offer the £350 roadmap within 24 hours.</p><p><strong>Client:</strong> ${clientName}<br/><strong>Business:</strong> ${businessName}<br/><strong>Email:</strong> ${clientEmail}</p>${offerHint}<p>Full notes attached.</p></div>`,
      attachments: [{ filename, content: docBuffer.toString('base64') }]
    })
  });
  if (!r.ok) {
    const errText = await r.text();
    console.error('Internal email error', r.status, errText);
    throw new Error('Internal email failed: ' + r.status);
  }
  return await r.json();
}

// ============================================================
//  GET /api/speak-test  — diagnostic endpoint
// ============================================================
app.get('/api/speak-test', async (req, res) => {
  const ELEVENLABS_API_KEY = process.env.ELEVENLABS_API_KEY;
  const VOICE_ID = process.env.ELEVENLABS_VOICE_ID || 'hpp4J3VqNfWAUOO0d1Us';

  if (!ELEVENLABS_API_KEY) {
    return res.json({ status: 'error', reason: 'ELEVENLABS_API_KEY not set in environment' });
  }

  try {
    const r = await fetch(`https://api.elevenlabs.io/v1/text-to-speech/${VOICE_ID}`, {
      method: 'POST',
      headers: {
        'xi-api-key': ELEVENLABS_API_KEY,
        'Content-Type': 'application/json',
        'Accept': 'audio/mpeg'
      },
      body: JSON.stringify({
        text: 'Hello this is a test.',
        model_id: 'eleven_turbo_v2',
        voice_settings: { stability: 0.5, similarity_boost: 0.8 }
      })
    });

    const detail = r.ok ? null : await r.text();
    return res.json({
      status: r.ok ? 'success' : 'error',
      http_status: r.status,
      content_type: r.headers.get('Content-Type'),
      voice_id: VOICE_ID,
      key_prefix: ELEVENLABS_API_KEY.substring(0, 8) + '...',
      elevenlabs_error: detail || null
    });
  } catch (err) {
    return res.json({ status: 'exception', error: err.message });
  }
});

// ============================================================
//  POST /api/speak  — ElevenLabs TTS proxy
//  Keeps ELEVENLABS_API_KEY server-side, never exposed to browser
// ============================================================
app.post('/api/speak', async (req, res) => {
  const { text } = req.body || {};
  if (!text) return res.status(400).json({ error: 'No text provided' });

  const ELEVENLABS_API_KEY = process.env.ELEVENLABS_API_KEY;
  const VOICE_ID = process.env.ELEVENLABS_VOICE_ID || 'hpp4J3VqNfWAUOO0d1Us';

  if (!ELEVENLABS_API_KEY) {
    return res.status(503).json({ error: 'ElevenLabs not configured', fallback: true });
  }

  try {
    const r = await fetch(`https://api.elevenlabs.io/v1/text-to-speech/${VOICE_ID}/stream`, {
      method: 'POST',
      headers: {
        'xi-api-key': ELEVENLABS_API_KEY,
        'Content-Type': 'application/json',
        'Accept': 'audio/mpeg'
      },
      body: JSON.stringify({
        text: text,
        model_id: 'eleven_turbo_v2',
        voice_settings: {
          stability: 0.45,
          similarity_boost: 0.85,
          style: 0.2,
          use_speaker_boost: true
        }
      })
    });

    if (!r.ok) {
      const err = await r.text();
      console.error('ElevenLabs error', r.status, err);
      return res.status(500).json({
        error: 'ElevenLabs API error',
        detail: err,
        fallback: true
      });
    }

    // Stream audio directly back to browser
    res.setHeader('Content-Type', 'audio/mpeg');
    res.setHeader('Transfer-Encoding', 'chunked');
    const reader = r.body.getReader();
    const pump = async () => {
      const { done, value } = await reader.read();
      if (done) { res.end(); return; }
      res.write(Buffer.from(value));
      await pump();
    };
    await pump();

  } catch (err) {
    console.error('/api/speak error', err);
    res.status(500).json({ error: 'Server error', detail: err.message, fallback: true });
  }
});

// ============================================================
//  POST /api/save-transcript
// ============================================================
app.post('/api/save-transcript', async (req, res) => {
  const {
    clientName = '', businessName = '', email = '',
    startTime = null, endTime = null, durationSeconds = 0,
    topicsCovered = [], transcript = []
  } = req.body || {};

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
      // Columns: A Timestamp | B Client | C Business | D Email | E Duration | F Topics | G Transcript | H Status
      const row = [
        new Date().toISOString(),
        clientName, businessName, email,
        durationSeconds,
        topicsCovered.join(', '),
        transcriptText,
        'NEW — quote pending review'
      ];
      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID, range: 'Sheet1!A1',
        valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
        requestBody: { values: [row] }
      });
      sheetSaved = true;
    } catch (err) {
      console.error('Sheet save error', err.message);
    }
  }

  let clientEmailed = false;
  let internalEmailed = false;
  let summaryError = null;

  if (email && transcript.length > 1) {
    try {
      const ackResult = await sendClientAcknowledgement(email, clientName, businessName);
      clientEmailed = !ackResult.skipped;
    } catch (err) {
      console.error('Client ack failed', err);
      summaryError = err.message;
    }
    try {
      const notes = await generateInternalNotes(transcript, clientName, businessName);
      const docBuffer = await buildInternalNotesDocx(notes, {
        clientName, businessName, email,
        dateStr: new Date(startTime || Date.now()).toLocaleString('en-GB', { dateStyle: 'long', timeStyle: 'short' }),
        durationStr: Math.floor(durationSeconds / 60) + ' min ' + (durationSeconds % 60) + ' sec'
      });
      const internalResult = await sendInternalNotes(clientName, businessName, email, docBuffer, notes.shouldOfferRoadmap);
      internalEmailed = !internalResult.skipped;
    } catch (err) {
      console.error('Internal notes/email error', err);
      summaryError = (summaryError ? summaryError + ' | ' : '') + err.message;
    }
  }

  res.json({ ok: sheetSaved || clientEmailed, sheetSaved, clientEmailed, internalEmailed, summaryError });
});

app.listen(PORT, () => {
  console.log('AI Growth Catalyst backend v4-quote listening on port', PORT);
});
