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
const crypto = require('crypto');
const cors = require('cors');
const { google } = require('googleapis');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  LevelFormat, BorderStyle
} = require('docx');

const app = express();
app.use(express.json({ limit: '2mb' }));
app.use(express.urlencoded({ extended: false }));

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

YOUR JOB ON EACH EXCHANGE:
1. Engage genuinely with what they just said — reflect back the implication, not just the fact
2. Classify the answer using SIGNAL DETECTION
3. Either probe or move to the next topic based on signal strength and the directive's instructions
4. Never exceed the probe ceiling the directive gives you

ENGAGEMENT — this is the most important rule in the whole prompt:
You are not a question machine. You are a sharp, curious consultant who listens carefully and shows it.
Every response must feel like a real person heard them — not a bot waiting for its turn to speak.

HOW TO ENGAGE (reflect the implication, not just the fact):

WEAK — restates the fact or echoes their words back:
"That's a decent chunk." (echoing "decent chunk" back at them)
"Right, so the spreadsheet is doing a lot of work there."
"Got it, so the invoicing sits with you personally."

STRONG — reflects the implication or operational consequence:
"So a meaningful slice of the week is going on things that aren't actually the client work itself."
"So essentially if you're not there, the invoicing doesn't happen — everything funnels back to you."
"That spreadsheet is doing a job that probably should have a proper system behind it."
"So the process works, but only because someone is manually holding it together every time."
"It sounds like that's one of those things that's fine until it isn't."
"Right — so the tool does most of it, but there's a manual step that someone has to remember to do."

The difference: WEAK echoes or restates what they said. STRONG draws out what it means — the real operational consequence of what they described.

RULES FOR ENGAGEMENT:
- 1-2 sentences before the next question — never more
- Vary your openers: "So essentially...", "It sounds like...", "Right —", "That makes sense —", "So the picture there is...", "So a meaningful slice of..."
- Never echo their exact words back — if they said "decent chunk", don't say "that's a decent chunk"
- Never use: "Great!", "Wonderful!", "Fantastic!", "That's really interesting" — sycophantic filler
- Never hint at solutions: "that's exactly the kind of thing that AI could..." is strictly forbidden
- For NO SIGNAL answers (standard tools, no friction), keep it brief: "Makes sense." or "Got it." — don't over-engage with boring answers
- When moving between topics, give only a brief acknowledgement — the system adds the next topic's opening question itself, so never introduce the next topic yourself
- Never skip the engagement entirely when moving between topics — even a brief "Makes sense." before the next opener is better than nothing

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

Respond in JSON only — no markdown, no text outside the JSON. Keep "action" FIRST:
{
  "action": "followup" | "next_topic" | "wrap_up",
  "signal_detected": true | false,
  "signal_strength": "none" | "standard" | "strong",
  "is_deep_probe": true | false,
  "next_question": "see below"
}

next_question depends on action:
- followup: engagement + the probe question as one natural spoken response (engagement 1-2 sentences max, then the question).
- next_topic or wrap_up: ONLY a brief spoken acknowledgement of what they just said (one short sentence, e.g. "Makes sense." or a sharper reflection of the implication). Do NOT ask any question and do NOT introduce the next topic — the system appends the next topic's opening question itself.`;


const INTERNAL_NOTES_PROMPT = `You are summarising a business discovery interview for an internal consultant who will use these notes to write a bespoke AI automation roadmap, and who will also send the client a free written snapshot that previews the full roadmap. Be specific — pull out numbers, tool names, frequencies, time-costs, exact phrasing. Record everything by calling the record_notes tool.

Field guide:
- executiveSummary: 2-3 sentence overview of the business and its key operational challenges
- businessOverview: what they do, who their customers are, team size, business model
- currentSystems: each tool/system mentioned, with how it is used
- manualWork: each manual task identified, with frequency and time-cost where stated
- painPoints: each pain point with context on impact
- customerInteractions: how they handle inquiries, support, follow-ups
- priorities: stated priorities and what they wish was different
- clientAutomationHypotheses: tasks the CLIENT themselves identified as automatable
- automationOpportunities: YOUR objective observations of areas worth exploring (short one-liners)
- opportunities: the same opportunities as structured items, most valuable first (max 6). For each: name; whatItIs (the problem, not a solution); evidence (1-3 short VERBATIM client quotes); statedFigures (only numbers the client actually said, e.g. "about a day a week across the team", otherwise "none stated"); systemsInvolved; impact (high/medium/low); effort (high/medium/low, your rough judgement); unknowns (what the call did not establish)
- gapQuestions: 3-6 questions to ask the client by email to firm up a roadmap (volumes, tool costs, who owns what, constraints)
- redFlags: concerns about readiness for automation
- notableQuotes: 3-5 direct verbatim quotes
- suggestedFirstFocus: your single best guess at where the consultant should focus the roadmap
- fitAssessment: brief — strong fit / moderate fit / weak fit / no clear fit, with one sentence reasoning
- shouldOfferRoadmap: YES or NO — should we offer the £350 roadmap? Say NO if there is no clear AI fit, the business is too small/early-stage to benefit, or there are significant red flags. One sentence reason.

CLIENT RECAP FIELDS (these are used in an email that goes to the client, so be strict):
- recapFit: strong, moderate, weak or none (your fit judgement, same as above)
- snapshotChecks: INTERNAL list of anything in the snapshot the consultant should verify before it goes to the client (estimates, tool examples, assumptions)

THE CLIENT SNAPSHOT (the "snapshot" object). This is the free written snapshot emailed to the client. It is a CONDENSED PREVIEW of the full roadmap, using the same sections and the same wording, so the client can see the roadmap is simply the fuller version. It must be specific to THIS call and cover the issues the client actually raised. Be strict: it goes to the client.
- businessLine: one line, e.g. "Residential estate agency · 9 staff" (only facts the client stated)
- executiveSummary: 2-3 sentences. What the business is, the main operational friction heard on the call, and the single highest-value starting point.
- painPoints: the TOP 3 issues the client raised, ranked by impact. Each has: title (short, plain, include their figure if they gave one, e.g. "Follow-up depends on one person"), description (1-2 sentences using their facts), quote (a short VERBATIM client quote, or empty string if none fits).
- startingPoint: the single best first step, the same opportunity the full roadmap would put first. Fields: name (e.g. "Automated enquiry follow-up"), whatChanges (2-3 plain sentences: what would be different; generic tool types only, nothing to buy), timeToLive (e.g. "3-5 days"), estimatedSaving (hours or money per week/month, ONLY from figures the client stated, shown as an estimate with "~"; if no usable figure, "To be confirmed"), teamEffort (e.g. "~2 hours"), workflowDisruption (None / Low / Medium), whyThisFirst (2 sentences).
Never invent facts, tools they use, customer counts or savings. Plain English, UK spelling, no hype.`;

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
    deepProbeUsedOnTopic = false, topicAllowsDeepProbe = false,
    topicAllowsProbe = true
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
  } else if (topicAllowsProbe === false) {
    pacingInstruction = 'This topic allows NO probe. Whatever they said, set action to "next_topic" now, with a brief acknowledgement.';
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
        // Haiku 4.5 is ~2-3x faster than Sonnet for this short JSON task; set NEXT_Q_MODEL=claude-sonnet-4-5-20250929 to revert
        model: process.env.NEXT_Q_MODEL || 'claude-haiku-4-5-20251001',
        max_tokens: 220,
        // Cached system prompt = faster time-to-first-token on every turn after the first
        system: [{ type: 'text', text: INTERVIEWER_PROMPT, cache_control: { type: 'ephemeral' } }],
        messages: apiMessages
      })
    });
    if (!r.ok) {
      const errText = await r.text();
      console.error('Anthropic error', r.status, errText);
      return res.status(500).json({ error: 'Anthropic API error', status: r.status, detail: String(errText).slice(0, 300) });
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
    res.json({ ...parsed, v: 2, _model: data.model, _cache_read: (data.usage && data.usage.cache_read_input_tokens) || 0 });
  } catch (err) {
    console.error('next-question error', err);
    res.status(500).json({ error: 'Server error' });
  }
});

// ============================================================
const NOTES_TOOL = {
  name: 'record_notes',
  description: 'Record the structured internal notes for this discovery call.',
  input_schema: {
    type: 'object',
    properties: {
      executiveSummary: { type: 'string' },
      businessOverview: { type: 'string' },
      currentSystems: { type: 'array', items: { type: 'string' } },
      manualWork: { type: 'array', items: { type: 'string' } },
      painPoints: { type: 'array', items: { type: 'string' } },
      customerInteractions: { type: 'string' },
      priorities: { type: 'array', items: { type: 'string' } },
      clientAutomationHypotheses: { type: 'array', items: { type: 'string' } },
      automationOpportunities: { type: 'array', items: { type: 'string' } },
      redFlags: { type: 'array', items: { type: 'string' } },
      notableQuotes: { type: 'array', items: { type: 'string' } },
      suggestedFirstFocus: { type: 'string' },
      fitAssessment: { type: 'string' },
      shouldOfferRoadmap: { type: 'string' },
      opportunities: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            name: { type: 'string' },
            whatItIs: { type: 'string' },
            evidence: { type: 'array', items: { type: 'string' } },
            statedFigures: { type: 'string' },
            systemsInvolved: { type: 'array', items: { type: 'string' } },
            impact: { type: 'string', enum: ['high', 'medium', 'low'] },
            effort: { type: 'string', enum: ['high', 'medium', 'low'] },
            unknowns: { type: 'array', items: { type: 'string' } }
          },
          required: ['name', 'whatItIs', 'impact', 'effort']
        }
      },
      gapQuestions: { type: 'array', items: { type: 'string' } },
      recapFit: { type: 'string', enum: ['strong', 'moderate', 'weak', 'none'] },
      snapshotChecks: { type: 'array', items: { type: 'string' } },
      snapshot: {
        type: 'object',
        properties: {
          businessLine: { type: 'string' },
          executiveSummary: { type: 'string' },
          painPoints: {
            type: 'array',
            items: {
              type: 'object',
              properties: { title: { type: 'string' }, description: { type: 'string' }, quote: { type: 'string' } },
              required: ['title', 'description']
            }
          },
          startingPoint: {
            type: 'object',
            properties: {
              name: { type: 'string' }, whatChanges: { type: 'string' }, timeToLive: { type: 'string' },
              estimatedSaving: { type: 'string' }, teamEffort: { type: 'string' },
              workflowDisruption: { type: 'string' }, whyThisFirst: { type: 'string' }
            },
            required: ['name', 'whatChanges', 'timeToLive', 'estimatedSaving', 'teamEffort', 'workflowDisruption', 'whyThisFirst']
          }
        },
        required: ['executiveSummary', 'painPoints', 'startingPoint']
      }
    },
    required: ['executiveSummary', 'suggestedFirstFocus', 'fitAssessment', 'shouldOfferRoadmap', 'recapFit', 'snapshot']
  }
};

async function generateInternalNotes(transcript, clientName, businessName) {
  const transcriptText = transcript.map(t => {
    const role = t.role === 'agent' ? 'AGENT' : 'CLIENT';
    return role + ': ' + t.text;
  }).join('\n\n');

  const models = [process.env.NOTES_MODEL || 'claude-sonnet-5-5', 'claude-haiku-4-5-20251001'];
  let lastErr = 'no model tried';
  for (const model of models) {
    try {
      const r = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': ANTHROPIC_API_KEY,
          'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify({
          model,
          max_tokens: 6000,
          system: INTERNAL_NOTES_PROMPT,
          tools: [NOTES_TOOL],
          tool_choice: { type: 'tool', name: 'record_notes' },
          messages: [{ role: 'user', content: `Client: ${clientName}\nBusiness: ${businessName}\n\nTRANSCRIPT:\n\n${transcriptText}\n\nCall the record_notes tool with your notes.` }]
        })
      });
      if (!r.ok) {
        const body = await r.text();
        lastErr = 'Notes API ' + model + ' ' + r.status + ': ' + body.slice(0, 200);
        console.error(lastErr);
        continue;
      }
      const data = await r.json();
      if (data.stop_reason === 'max_tokens') console.error('Notes reply hit max_tokens on ' + model);
      const tool = (data.content || []).find(c => c.type === 'tool_use' && c.name === 'record_notes');
      if (tool && tool.input && typeof tool.input === 'object') return tool.input;
      throw new Error('no tool_use block in reply (stop_reason ' + data.stop_reason + ')');
    } catch (err) {
      lastErr = 'Notes ' + model + ' failed: ' + err.message;
      console.error(lastErr);
    }
  }
  throw new Error(lastErr);
}

// ============================================================
function toText(v) {
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'object') return Object.values(v).map(toText).filter(Boolean).join(' — ');
  return String(v);
}
function toList(x) { return Array.isArray(x) ? x.map(toText).filter(Boolean) : (x ? [toText(x)] : []); }
function normaliseStartingPoint(sp) {
  sp = sp || {};
  return { name: toText(sp.name), whatChanges: toText(sp.whatChanges), timeToLive: toText(sp.timeToLive),
    estimatedSaving: toText(sp.estimatedSaving), teamEffort: toText(sp.teamEffort),
    workflowDisruption: toText(sp.workflowDisruption), whyThisFirst: toText(sp.whyThisFirst) };
}
function normalisePainPoints(pp) {
  return (Array.isArray(pp) ? pp : []).filter(o => o && typeof o === 'object')
    .map(o => ({ title: toText(o.title), description: toText(o.description), quote: toText(o.quote) }));
}
function normaliseNotes(n) {
  const out = {};
  Object.keys(n || {}).forEach(function(k) {
    const v = n[k];
    if (k === 'snapshot') {
      const sn = v && typeof v === 'object' ? v : {};
      out[k] = { businessLine: toText(sn.businessLine), executiveSummary: toText(sn.executiveSummary),
        painPoints: normalisePainPoints(sn.painPoints), startingPoint: normaliseStartingPoint(sn.startingPoint) };
    } else if (k === 'opportunities') {
      out[k] = (Array.isArray(v) ? v : []).filter(function(o) { return o && typeof o === 'object'; }).map(function(o) {
        const lst = function(x) { return Array.isArray(x) ? x.map(toText).filter(Boolean) : (x ? [toText(x)] : []); };
        return {
          name: toText(o.name), whatItIs: toText(o.whatItIs), statedFigures: toText(o.statedFigures),
          impact: toText(o.impact), effort: toText(o.effort),
          evidence: lst(o.evidence), systemsInvolved: lst(o.systemsInvolved), unknowns: lst(o.unknowns)
        };
      });
    } else {
      out[k] = Array.isArray(v) ? v.map(toText).filter(Boolean) : toText(v);
    }
  });
  return out;
}

// ============================================================
//  Tier 1 snapshot — condensed preview of the roadmap (same sections, same metrics)
// ============================================================
function escHtml(t) {
  return String(t == null ? '' : t).replace(/[&<>"']/g, function(c) {
    return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c];
  });
}
function snapshotUsable(rawNotes) {
  const n = normaliseNotes(rawNotes);
  const sn = n.snapshot;
  return !!(sn && sn.executiveSummary && sn.painPoints.length >= 2 && sn.startingPoint.name && sn.startingPoint.whatChanges &&
            (n.recapFit === 'strong' || n.recapFit === 'moderate'));
}
function buildSnapshotEmail(rawNotes, clientName, businessName) {
  const n = normaliseNotes(rawNotes);
  const sn = n.snapshot, sp = sn.startingPoint;
  const first = (clientName || '').split(' ')[0] || 'there';
  const dateStr = new Date().toLocaleString('en-GB', { month: 'long', year: 'numeric' });
  const oppCount = (n.opportunities || []).length;
  const subject = 'Your AI Opportunity Snapshot — ' + (businessName || 'your business');
  const tile = (label, value) => '<td style="width:50%;padding:6px"><div style="background:#fff;border:1px solid #e8dec5;border-radius:10px;padding:12px"><div style="font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:#9a8a6a">' + escHtml(label) + '</div><div style="font-size:16px;color:#1a2332;font-weight:600;margin-top:2px">' + escHtml(value || '—') + '</div></div></td>';
  const sectionHead = (num, title) => '<p style="margin:28px 0 8px;font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#c85a3e"><strong>' + num + ' · ' + escHtml(title) + '</strong></p>';
  const pain = sn.painPoints.slice(0, 3).map(function(p, i) {
    return '<div style="margin:0 0 14px"><div style="font-size:15px;color:#1a2332"><strong>' + (i + 1) + '. ' + escHtml(p.title) + '</strong></div>' +
      (p.description ? '<div style="font-size:14px;color:#5d6b80;margin-top:2px">' + escHtml(p.description) + '</div>' : '') +
      (p.quote ? '<div style="font-size:14px;font-style:italic;color:#4b5563;border-left:3px solid #c85a3e;padding-left:10px;margin-top:6px">“' + escHtml(p.quote) + '”</div>' : '') + '</div>';
  }).join('');
  const html = `
    <div style="font-family:-apple-system,Segoe UI,sans-serif;max-width:600px;margin:0 auto;color:#374151;line-height:1.6">
      <p style="font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#9a8a6a;margin:0">AI Growth Catalyst · ${escHtml(dateStr)}</p>
      <h1 style="font-size:26px;color:#1a2332;margin:6px 0 2px">AI Opportunity Snapshot</h1>
      <p style="font-size:17px;color:#c85a3e;margin:0"><strong>${escHtml(businessName)}</strong></p>
      ${sn.businessLine ? '<p style="font-size:13px;color:#5d6b80;margin:2px 0 0">' + escHtml(sn.businessLine) + '</p>' : ''}
      <p style="font-size:15px;margin:18px 0 0">Hi ${escHtml(first)}, thanks for the call. Here is your free snapshot: a short version of the full roadmap, covering what you told us and the one place we would start.</p>
      ${sectionHead('01', 'Executive summary')}
      <p style="font-size:15px;margin:0">${escHtml(sn.executiveSummary)}</p>
      ${sectionHead('02', 'What we heard')}
      ${pain}
      ${sectionHead('03', 'Recommended starting point')}
      <div style="background:#fbf6ec;border:1px solid #e8dec5;border-radius:12px;padding:16px">
        <p style="margin:0 0 4px;font-size:12px;letter-spacing:.06em;text-transform:uppercase;color:#b8843a"><strong>Quick win #1</strong></p>
        <p style="margin:0 0 8px;font-size:18px;color:#1a2332"><strong>${escHtml(sp.name)}</strong></p>
        <p style="margin:0 0 12px;font-size:14px;color:#374151">${escHtml(sp.whatChanges)}</p>
        <table role="presentation" style="width:100%;border-collapse:collapse"><tr>${tile('Time to live', sp.timeToLive)}${tile('Estimated saving', sp.estimatedSaving)}</tr><tr>${tile('Team effort', sp.teamEffort)}${tile('Workflow disruption', sp.workflowDisruption)}</tr></table>
        ${sp.whyThisFirst ? '<p style="margin:10px 0 0;font-size:14px;color:#374151"><strong>Why this first:</strong> ' + escHtml(sp.whyThisFirst) + '</p>' : ''}
      </div>
      <div style="margin-top:24px;padding:16px;border:1px dashed #c9bfa5;border-radius:12px;background:#faf7f0">
        <p style="margin:0 0 6px;font-size:14px;color:#1a2332"><strong>This snapshot covers your single strongest opportunity.</strong></p>
        <p style="margin:0;font-size:14px;color:#5d6b80">${oppCount > 1 ? 'We identified ' + oppCount + ' opportunities on the call. ' : ''}The full bespoke roadmap sets out every one of them on a priority matrix, with a phased plan and honest caveats on where AI won't help. It is £350, a fixed fee, delivered within 48 hours of payment, and only if you choose it. No obligation either way.</p>
      </div>
      <p style="font-size:15px;margin-top:20px">If anything here doesn't ring true, or you thought of something after the call, just reply and tell us.</p>
      <p style="font-size:15px;margin-top:20px">Speak soon,<br/><strong>The AI Growth Catalyst team</strong></p>
      <hr style="border:none;border-top:1px solid #E5E7EB;margin:28px 0"/>
      <p style="font-size:12px;color:#9CA3AF">AI Growth Catalyst · Helping small and medium-sized businesses navigate AI without the hype.</p>
    </div>`;
  const text = 'Subject: ' + subject + '\n\nAI Opportunity Snapshot — ' + businessName + (sn.businessLine ? '\n' + sn.businessLine : '') +
    '\n\nHi ' + first + ', thanks for the call. Here is your free snapshot.\n\n01 EXECUTIVE SUMMARY\n' + sn.executiveSummary +
    '\n\n02 WHAT WE HEARD\n' + sn.painPoints.slice(0, 3).map((p, i) => (i + 1) + '. ' + p.title + (p.description ? ' — ' + p.description : '') + (p.quote ? '\n   "' + p.quote + '"' : '')).join('\n') +
    '\n\n03 RECOMMENDED STARTING POINT — Quick win #1: ' + sp.name + '\n' + sp.whatChanges +
    '\nTime to live: ' + sp.timeToLive + '\nEstimated saving: ' + sp.estimatedSaving + '\nTeam effort: ' + sp.teamEffort + '\nWorkflow disruption: ' + sp.workflowDisruption +
    (sp.whyThisFirst ? '\nWhy this first: ' + sp.whyThisFirst : '') +
    '\n\nThis snapshot covers your single strongest opportunity.' + (oppCount > 1 ? ' We identified ' + oppCount + ' on the call.' : '') +
    ' The full bespoke roadmap (£350, fixed fee, delivered within 48 hours of payment, only if you choose it) sets out every opportunity on a priority matrix, with a phased plan and honest caveats.\n\nSpeak soon,\nThe AI Growth Catalyst team';
  return { subject, html, text };
}

async function sendClientSnapshot(toEmail, rawNotes, clientName, businessName) {
  if (!RESEND_API_KEY) return { skipped: true };
  const { subject, html } = buildSnapshotEmail(rawNotes, clientName, businessName);
  const r = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + RESEND_API_KEY },
    body: JSON.stringify({ from: FROM_EMAIL, to: [toEmail], subject, html })
  });
  if (!r.ok) {
    const errText = await r.text();
    console.error('Client snapshot email error', r.status, errText);
    throw new Error('Client snapshot failed: ' + r.status + ' ' + errText.slice(0, 200));
  }
  return await r.json();
}

function buildInternalNotesDocx(rawNotes, meta) {
  const notes = normaliseNotes(rawNotes);
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
  if (notes.opportunities && notes.opportunities.length) {
    children.push(heading('Scored Opportunities'));
    notes.opportunities.forEach(function(o, i) {
      children.push(new Paragraph({ spacing: { before: 160, after: 60 }, children: [
        new TextRun({ text: (i + 1) + '. ' + o.name, bold: true, size: 24, color: '1A2332' }),
        new TextRun({ text: '   Impact: ' + (o.impact || '?') + ' · Effort: ' + (o.effort || '?'), size: 20, color: 'B8843A', bold: true })
      ]}));
      if (o.whatItIs) children.push(body(o.whatItIs));
      if (o.statedFigures) children.push(bullet('Stated figures: ' + o.statedFigures));
      if (o.systemsInvolved.length) children.push(bullet('Systems: ' + o.systemsInvolved.join(', ')));
      o.evidence.forEach(function(q) { children.push(quote(q)); });
      if (o.unknowns.length) children.push(bullet('Still unknown: ' + o.unknowns.join('; ')));
    });
  }
  if (notes.gapQuestions && notes.gapQuestions.length) {
    children.push(heading('Gaps To Ask About (follow-up email)'));
    notes.gapQuestions.forEach(function(q) { children.push(bullet(q)); });
  }
  if (snapshotUsable(rawNotes)) {
    children.push(heading('Draft Client Snapshot (Tier 1)'));
    children.push(body('Check before sending. Same sections and metrics as the full roadmap.'));
    buildSnapshotEmail(rawNotes, meta.clientName, meta.businessName).text.split('\n').forEach(function(line) { children.push(body(line)); });
    if (notes.snapshotChecks && notes.snapshotChecks.length) {
      children.push(heading('Check Before Sending'));
      notes.snapshotChecks.forEach(function(q) { children.push(bullet(q)); });
    }
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
        <li style="margin-bottom: 6px;">A ranked plan: every opportunity we found, how each would work and what it would involve, and the order to tackle them in</li>
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
    throw new Error('Client email failed: ' + r.status + ' ' + errText.slice(0, 200));
  }
  return await r.json();
}


// ============================================================
//  Draft roadmap (the £350 product) — generated for YES-verdict calls, reviewed by you before it goes out
// ============================================================
const ROADMAP_PROMPT = `You are a senior, vendor-neutral AI and automation consultant writing a draft bespoke roadmap for a small or medium business owner, based on a discovery call. A human consultant reviews and edits your draft before the client sees it, so be useful, specific and honest rather than impressive.

FORMAT: the roadmap follows the same seven sections as our published example, and it is the FULL version of the free snapshot the client already received. The snapshot's pain points and starting point are in the notes (notes.snapshot). Keep your starting point consistent with that snapshot (same opportunity, same metrics) unless the transcript clearly shows a better one; if you change it, say why in reviewerChecks.

WRITE FOR THE BUSINESS OWNER: plain English, UK spelling, no jargon, no hype. Warm but professional.

GROUND EVERYTHING IN THE CALL:
- Use the client's own words, tools, team size and figures. Never invent a number, a tool they use, a customer count or a saving.
- Savings and impact come from figures the client actually gave (simple arithmetic is fine, shown with "~"). Where no figure was given, say so and describe the impact qualitatively.
- Timelines and effort are realistic estimates, labelled as estimates. Do NOT quote build prices or software subscription prices.
- "Industry data suggests…" claims only where you are confident, and list every such claim in reviewerChecks.

SECTIONS:
01 executiveSummary: one paragraph: the business, the headline friction, how many opportunities you found, and the single best starting point.
02 keyFigures (3-4 stated facts as value + label, e.g. "9" / "Team members") and currentSystems (a short paragraph on the tools and how they connect).
03 painPoints: EVERY pain point raised, ranked by impact, each with title, description and, where one exists, a verbatim client quote.
04 opportunities: EVERY AI opportunity you can justify from the call (aim for 6-12, never pad). For each: name, quadrant (quick_win = high impact and easy; challenge = high impact and difficult; fill_in = low impact and easy; low_priority = low impact and difficult) and a one-line note.
05 startingPoint: the one thing to do first: name, whatChanges, timeToLive, estimatedSaving, teamEffort, workflowDisruption, whyThisFirst.
06 phases: 3 phases (quick wins, strategic, optimisation), each with phase, timeframe, title, description and items.
07 caveats: honest caveats, including where AI will NOT help or where a task is risky (title + detail). Include at least one.

HONESTY: if an opportunity is weak, small or risky, say so. Do not oversell AI.

reviewerChecks (INTERNAL, never shown to the client): every claim, estimate, tool example or assumption the human consultant should verify before sending, and anything the call did not give enough evidence for.

Record the roadmap by calling the record_roadmap tool.`;

const ROADMAP_TOOL = {
  name: 'record_roadmap',
  description: 'Record the draft roadmap for this client.',
  input_schema: {
    type: 'object',
    properties: {
      executiveSummary: { type: 'string' },
      keyFigures: { type: 'array', items: { type: 'object', properties: { value: { type: 'string' }, label: { type: 'string' } }, required: ['value', 'label'] } },
      currentSystems: { type: 'string' },
      painPoints: {
        type: 'array',
        items: { type: 'object', properties: { title: { type: 'string' }, description: { type: 'string' }, quote: { type: 'string' } }, required: ['title', 'description'] }
      },
      opportunities: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            name: { type: 'string' },
            quadrant: { type: 'string', enum: ['quick_win', 'challenge', 'fill_in', 'low_priority'] },
            note: { type: 'string' }
          },
          required: ['name', 'quadrant']
        }
      },
      startingPoint: {
        type: 'object',
        properties: {
          name: { type: 'string' }, whatChanges: { type: 'string' }, timeToLive: { type: 'string' },
          estimatedSaving: { type: 'string' }, teamEffort: { type: 'string' },
          workflowDisruption: { type: 'string' }, whyThisFirst: { type: 'string' }
        },
        required: ['name', 'whatChanges', 'timeToLive', 'estimatedSaving', 'teamEffort', 'workflowDisruption', 'whyThisFirst']
      },
      phases: {
        type: 'array',
        items: {
          type: 'object',
          properties: { phase: { type: 'string' }, timeframe: { type: 'string' }, title: { type: 'string' }, description: { type: 'string' }, items: { type: 'array', items: { type: 'string' } } },
          required: ['phase', 'title']
        }
      },
      caveats: { type: 'array', items: { type: 'object', properties: { title: { type: 'string' }, detail: { type: 'string' } }, required: ['title', 'detail'] } },
      reviewerChecks: { type: 'array', items: { type: 'string' } }
    },
    required: ['executiveSummary', 'painPoints', 'opportunities', 'startingPoint', 'phases', 'caveats', 'reviewerChecks']
  }
};

async function generateRoadmap(rawNotes, transcript, clientName, businessName) {
  const transcriptText = transcript.map(t => (t.role === 'agent' ? 'AGENT' : 'CLIENT') + ': ' + t.text).join('\n\n');
  const models = [process.env.ROADMAP_MODEL || 'claude-sonnet-5-5', 'claude-haiku-4-5-20251001'];
  let lastErr = 'no model tried';
  for (const model of models) {
    try {
      const r = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-api-key': ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
        body: JSON.stringify({
          model,
          max_tokens: 8000,
          system: ROADMAP_PROMPT,
          tools: [ROADMAP_TOOL],
          tool_choice: { type: 'tool', name: 'record_roadmap' },
          messages: [{ role: 'user', content:
            'Client: ' + clientName + '\nBusiness: ' + businessName +
            '\n\nINTERNAL DISCOVERY NOTES (structured, from the call):\n' + JSON.stringify(rawNotes, null, 1) +
            '\n\nFULL TRANSCRIPT:\n\n' + transcriptText +
            '\n\nCall the record_roadmap tool with the draft roadmap.' }]
        })
      });
      if (!r.ok) {
        const body = await r.text();
        lastErr = 'Roadmap API ' + model + ' ' + r.status + ': ' + body.slice(0, 200);
        console.error(lastErr);
        continue;
      }
      const data = await r.json();
      if (data.stop_reason === 'max_tokens') console.error('Roadmap reply hit max_tokens on ' + model);
      const tool = (data.content || []).find(c => c.type === 'tool_use' && c.name === 'record_roadmap');
      if (tool && tool.input && typeof tool.input === 'object') return tool.input;
      throw new Error('no tool_use block in reply (stop_reason ' + data.stop_reason + ')');
    } catch (err) {
      lastErr = 'Roadmap ' + model + ' failed: ' + err.message;
      console.error(lastErr);
    }
  }
  throw new Error(lastErr);
}

function normaliseRoadmap(rm) {
  return {
    executiveSummary: toText(rm.executiveSummary),
    keyFigures: (Array.isArray(rm.keyFigures) ? rm.keyFigures : []).filter(o => o && typeof o === 'object').map(o => ({ value: toText(o.value), label: toText(o.label) })),
    currentSystems: toText(rm.currentSystems),
    painPoints: normalisePainPoints(rm.painPoints),
    opportunities: (Array.isArray(rm.opportunities) ? rm.opportunities : []).filter(o => o && typeof o === 'object').map(o => ({ name: toText(o.name), quadrant: toText(o.quadrant), note: toText(o.note) })),
    startingPoint: normaliseStartingPoint(rm.startingPoint),
    phases: (Array.isArray(rm.phases) ? rm.phases : []).filter(o => o && typeof o === 'object').map(o => ({ phase: toText(o.phase), timeframe: toText(o.timeframe), title: toText(o.title), description: toText(o.description), items: toList(o.items) })),
    caveats: (Array.isArray(rm.caveats) ? rm.caveats : []).filter(o => o && typeof o === 'object').map(o => ({ title: toText(o.title), detail: toText(o.detail) })),
    reviewerChecks: toList(rm.reviewerChecks)
  };
}

function buildRoadmapDocx(rawRoadmap, meta) {
  const rm = normaliseRoadmap(rawRoadmap);
  const { clientName, businessName, dateStr } = meta;
  const num = (n, t, sub) => [
    new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 440, after: 60 }, children: [
      new TextRun({ text: n + '  ', bold: true, size: 22, color: 'C85A3E' }), new TextRun({ text: t, bold: true, size: 32, color: '1A2332' })] }),
    ...(sub ? [new Paragraph({ spacing: { after: 140 }, children: [new TextRun({ text: sub, italics: true, size: 20, color: '6B7280' })] })] : [])
  ];
  const sub = (text) => new Paragraph({ spacing: { before: 200, after: 60 }, children: [new TextRun({ text, bold: true, size: 24, color: '1A2332' })] });
  const label = (l, t) => new Paragraph({ spacing: { after: 100 }, children: [
    new TextRun({ text: l + ' ', bold: true, size: 22, color: '1F2937' }), new TextRun({ text: t, size: 22, color: '374151' })] });
  const body = (text) => new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text, size: 22, color: '374151' })] });
  const bullet = (text) => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80 }, children: [new TextRun({ text, size: 22, color: '374151' })] });
  const quote = (text) => new Paragraph({ spacing: { before: 60, after: 100 }, indent: { left: 360 },
    border: { left: { style: BorderStyle.SINGLE, size: 16, color: 'C85A3E', space: 8 } }, children: [new TextRun({ text: '"' + text + '"', italics: true, size: 22, color: '4B5563' })] });

  const c = [];
  c.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: 'BESPOKE AI ROADMAP', bold: true, size: 18, color: 'B8843A' })] }));
  c.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: 'Where AI actually fits at ' + businessName + '.', bold: true, size: 44, color: '1A2332' })] }));
  c.push(new Paragraph({ spacing: { after: 300 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'C85A3E', space: 4 } },
    children: [new TextRun({ text: 'Prepared for ' + clientName + ' · ' + dateStr + ' · AI Growth Catalyst', italics: true, size: 18, color: 'B8843A' })] }));

  c.push(...num('01', 'Executive summary', 'The headline findings from the discovery call.')); c.push(body(rm.executiveSummary));
  c.push(...num('02', 'Business overview', 'How the business operates today, based on the discovery call.'));
  rm.keyFigures.forEach(f => c.push(bullet(f.value + ' — ' + f.label)));
  if (rm.currentSystems) { c.push(sub('Current systems & tools')); c.push(body(rm.currentSystems)); }
  c.push(...num('03', 'Identified pain points', 'The operational frustrations described during the call, ranked by impact.'));
  rm.painPoints.forEach((p, i) => { c.push(sub((i + 1) + '. ' + p.title)); if (p.description) c.push(body(p.description)); if (p.quote) c.push(quote(p.quote)); });

  c.push(...num('04', 'Priority matrix', 'Each opportunity by ease of implementation and potential impact.'));
  const groups = [['quick_win', 'Quick wins — high impact, easy. Do first.'], ['challenge', 'Challenges — high impact, difficult. Plan.'],
    ['fill_in', 'Fill-ins — low impact, easy. If time.'], ['low_priority', 'Low priority — low impact, difficult. Avoid.']];
  groups.forEach(function(g) {
    const items = rm.opportunities.filter(o => o.quadrant === g[0]);
    if (!items.length) return;
    c.push(sub(g[1])); items.forEach(o => c.push(bullet(o.name + (o.note ? ' — ' + o.note : ''))));
  });

  const sp = rm.startingPoint;
  c.push(...num('05', 'Recommended starting point', 'If you do nothing else from this roadmap, do this.'));
  c.push(sub('Quick win #1: ' + sp.name)); c.push(body(sp.whatChanges));
  c.push(label('Time to live:', sp.timeToLive)); c.push(label('Estimated saving:', sp.estimatedSaving));
  c.push(label('Team effort:', sp.teamEffort)); c.push(label('Workflow disruption:', sp.workflowDisruption));
  if (sp.whyThisFirst) c.push(label('Why this first:', sp.whyThisFirst));

  c.push(...num('06', 'Implementation phases', 'A realistic sequence for tackling the rest of the roadmap.'));
  rm.phases.forEach(ph => { c.push(sub(ph.phase + (ph.timeframe ? ' · ' + ph.timeframe : '') + (ph.title ? ' — ' + ph.title : ''))); if (ph.description) c.push(body(ph.description)); ph.items.forEach(t => c.push(bullet(t))); });

  c.push(...num('07', 'Honest caveats', 'Things to be aware of before implementing any of this.'));
  rm.caveats.forEach(cv => { c.push(sub(cv.title)); c.push(body(cv.detail)); });
  c.push(new Paragraph({ spacing: { before: 400 }, border: { top: { style: BorderStyle.SINGLE, size: 6, color: 'E5E7EB', space: 6 } }, children: [new TextRun({
    text: 'This roadmap is yours to keep and use however you like. If you would like us to build any of it, we quote per project, only where the roadmap shows a clear return, and you own everything we build. Time and saving estimates are based on what you told us and will be firmed up when the work is scoped.', italics: true, size: 20, color: '6B7280' })] }));

  const doc = new Document({
    creator: 'AI Growth Catalyst', title: 'Roadmap - ' + businessName,
    styles: { default: { document: { run: { font: 'Calibri', size: 22 } } } },
    numbering: { config: [{ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children: c }]
  });
  return Packer.toBuffer(doc);
}

async function sendRoadmapDraft(clientName, businessName, clientEmail, docBuffer, reviewerChecks) {
  if (!RESEND_API_KEY || !BCC_EMAIL) return { skipped: true };
  const filename = 'roadmap-draft-' + (businessName || 'client').replace(/[^a-z0-9]+/gi, '-').toLowerCase() + '-' + new Date().toISOString().slice(0, 10) + '.docx';
  const checks = (reviewerChecks || []).map(c => '<li>' + escHtml(c) + '</li>').join('');
  const r = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + RESEND_API_KEY },
    body: JSON.stringify({
      from: FROM_EMAIL, to: [BCC_EMAIL],
      subject: '[Roadmap draft] ' + businessName + ' (' + clientName + ')',
      html: '<div style="font-family: -apple-system, sans-serif;"><p><strong>Draft £350 roadmap attached for ' + escHtml(businessName) + '</strong> (' + escHtml(clientName) + ', ' + escHtml(clientEmail) + '). It has NOT been sent to the client. Edit it, then send it yourself.</p>' +
        '<p><strong>Check before sending:</strong></p><ul>' + (checks || '<li>No flags raised — still read it through.</li>') + '</ul></div>',
      attachments: [{ filename, content: docBuffer.toString('base64') }]
    })
  });
  if (!r.ok) {
    const errText = await r.text();
    console.error('Roadmap draft email error', r.status, errText);
    throw new Error('Roadmap email failed: ' + r.status + ' ' + errText.slice(0, 200));
  }
  return await r.json();
}

// Runs after the HTTP response has gone, so the call page never waits on it
async function runRoadmapJob(notes, transcript, meta) {
  try {
    const rm = await generateRoadmap(notes, transcript, meta.clientName, meta.businessName);
    const buf = await buildRoadmapDocx(rm, meta);
    await sendRoadmapDraft(meta.clientName, meta.businessName, meta.email, buf, normaliseRoadmap(rm).reviewerChecks);
    console.log('Roadmap draft emailed for', meta.businessName);
  } catch (err) {
    console.error('Roadmap job failed', err);
    try {
      if (RESEND_API_KEY && BCC_EMAIL) {
        await fetch('https://api.resend.com/emails', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + RESEND_API_KEY },
          body: JSON.stringify({ from: FROM_EMAIL, to: [BCC_EMAIL],
            subject: '[Roadmap draft FAILED] ' + meta.businessName,
            html: '<p>The draft roadmap could not be generated: ' + escHtml(err.message) + '</p><p>The discovery notes email is unaffected.</p>' })
        });
      }
    } catch (e2) { console.error('Roadmap failure email failed', e2); }
  }
}

// ============================================================
async function sendInternalNotes(clientName, businessName, clientEmail, docBuffer, shouldOffer, opts) {
  if (!RESEND_API_KEY || !BCC_EMAIL) return { skipped: true };
  const ext = (opts && opts.ext) || 'docx';
  const extra = (opts && opts.extraHtml) || '';
  const filename = 'discovery-notes-' +
    (businessName || 'call').replace(/[^a-z0-9]+/gi, '-').toLowerCase() +
    '-' + new Date().toISOString().slice(0, 10) + '.' + ext;
  const offerHint = shouldOffer ? `<p><strong>Offer £350 roadmap?</strong> ${shouldOffer}</p>` : '';
  const r = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + RESEND_API_KEY },
    body: JSON.stringify({
      from: FROM_EMAIL,
      to: [BCC_EMAIL],
      subject: '[Internal] New discovery call: ' + businessName + ' (' + clientName + ')',
      html: `<div style="font-family: -apple-system, sans-serif;">${extra}<p>New discovery call completed. Action: review notes, decide whether to offer the £350 roadmap within 24 hours.</p><p><strong>Client:</strong> ${clientName}<br/><strong>Business:</strong> ${businessName}<br/><strong>Email:</strong> ${clientEmail}</p>${offerHint}<p>Full notes attached.</p>${(opts && opts.afterHtml) || ''}</div>`,
      attachments: [{ filename, content: docBuffer.toString('base64') }]
    })
  });
  if (!r.ok) {
    const errText = await r.text();
    console.error('Internal email error', r.status, errText);
    throw new Error('Internal email failed: ' + r.status + ' ' + errText.slice(0, 200));
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
const speakHandler = async (req, res) => {
  const text = (req.method === 'GET' ? req.query.text : (req.body || {}).text);
  if (!text || typeof text !== 'string') return res.status(400).json({ error: 'No text provided' });
  if (text.length > 2000) return res.status(400).json({ error: 'Text too long' });

  const ELEVENLABS_API_KEY = process.env.ELEVENLABS_API_KEY;
  const VOICE_ID = process.env.ELEVENLABS_VOICE_ID || 'hpp4J3VqNfWAUOO0d1Us';
  // eleven_flash_v2_5 is ElevenLabs' lowest-latency model; set ELEVENLABS_MODEL=eleven_turbo_v2 to revert
  const MODEL_ID = process.env.ELEVENLABS_MODEL || 'eleven_flash_v2_5';

  if (!ELEVENLABS_API_KEY) {
    return res.status(503).json({ error: 'ElevenLabs not configured', fallback: true });
  }

  const upstream = new AbortController();
  res.on('close', () => upstream.abort()); // client stopped listening -> stop generating

  try {
    const r = await fetch(`https://api.elevenlabs.io/v1/text-to-speech/${VOICE_ID}/stream`, {
      method: 'POST',
      signal: upstream.signal,
      headers: {
        'xi-api-key': ELEVENLABS_API_KEY,
        'Content-Type': 'application/json',
        'Accept': 'audio/mpeg'
      },
      body: JSON.stringify({
        text: text,
        model_id: MODEL_ID,
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
      return res.status(500).json({ error: 'ElevenLabs API error', detail: err, fallback: true });
    }

    res.setHeader('Content-Type', 'audio/mpeg');
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('X-Accel-Buffering', 'no');
    // Stream chunks straight through so the browser can start playing immediately
    const { Readable } = require('stream');
    Readable.fromWeb(r.body).on('error', () => res.end()).pipe(res);

  } catch (err) {
    if (err.name === 'AbortError') return;
    console.error('/api/speak error', err);
    if (!res.headersSent) res.status(500).json({ error: 'Server error', detail: err.message, fallback: true });
    else res.end();
  }
};
app.post('/api/speak', speakHandler);
app.get('/api/speak', speakHandler);


// ============================================================
//  On-demand roadmap — a signed link in the internal email lets you generate a draft for ANY saved call
//  (e.g. when the notes said NO but you disagree). Needs ADMIN_KEY set on Render.
// ============================================================
const PUBLIC_URL = (process.env.PUBLIC_URL || 'https://ai-discovery-tx39.onrender.com').replace(/\/$/, '');
function signRow(row) {
  return crypto.createHmac('sha256', process.env.ADMIN_KEY || '').update('roadmap:' + row).digest('hex');
}
function validSig(row, sig) {
  if (!process.env.ADMIN_KEY || !sig) return false;
  const a = Buffer.from(signRow(row)), b = Buffer.from(String(sig));
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}
function roadmapLinkHtml(row) {
  if (!row || !process.env.ADMIN_KEY) return '';
  const url = PUBLIC_URL + '/admin/roadmap?row=' + row + '&sig=' + signRow(row);
  return '<p style="margin-top:12px"><a href="' + url + '">Generate a draft roadmap for this call</a> (opens a confirm page; the draft arrives by email about a minute later).</p>';
}
function parseTranscriptText(text) {
  return String(text || '').split(/\n\n(?=\[)/).map(function(block) {
    const m = block.match(/^\[(.*?)\] (AGENT|CLIENT)(?: \((.*?)\))?: ([\s\S]*)$/);
    return m ? { role: m[2] === 'AGENT' ? 'agent' : 'user', time: m[1], topic: m[3] || '', text: m[4] } : null;
  }).filter(Boolean);
}
function adminPage(title, bodyHtml) {
  return '<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + escHtml(title) + '</title>' +
    '<body style="font-family:-apple-system,sans-serif;max-width:520px;margin:12vh auto;padding:0 20px;color:#1a2332;line-height:1.6">' + bodyHtml + '</body>';
}
app.get('/admin/roadmap', function(req, res) {
  const row = parseInt(req.query.row, 10);
  if (!row || row < 2 || !validSig(row, req.query.sig)) return res.status(403).send(adminPage('Not allowed', '<h2>Link not valid</h2><p>This link is invalid, or ADMIN_KEY is not set on the server.</p>'));
  res.send(adminPage('Generate roadmap', '<h2>Generate a draft roadmap?</h2><p>This reads call row ' + row + ' from your sheet, rebuilds the notes and writes a draft roadmap. It costs a few pence in AI usage and is emailed to you. Nothing is sent to the client.</p>' +
    '<form method="POST" action="/admin/roadmap"><input type="hidden" name="row" value="' + row + '"><input type="hidden" name="sig" value="' + escHtml(req.query.sig) + '">' +
    '<button style="font:inherit;padding:10px 18px;border:0;border-radius:8px;background:#c85a3e;color:#fff;cursor:pointer">Generate draft roadmap</button></form>'));
});
app.post('/admin/roadmap', async function(req, res) {
  const row = parseInt(req.body.row, 10);
  if (!row || row < 2 || !validSig(row, req.body.sig)) return res.status(403).send(adminPage('Not allowed', '<h2>Link not valid</h2>'));
  if (!(SHEET_ID && SVC_JSON)) return res.status(500).send(adminPage('Not available', '<h2>Google Sheet is not configured on the server</h2>'));
  let r;
  try {
    const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(SVC_JSON), scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const got = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Sheet1!A' + row + ':H' + row });
    r = (got.data.values || [])[0];
  } catch (err) {
    console.error('Admin roadmap sheet read failed', err.message);
    return res.status(500).send(adminPage('Error', '<h2>Could not read the sheet</h2><p>' + escHtml(err.message) + '</p>'));
  }
  if (!r || !r[6]) return res.status(404).send(adminPage('Not found', '<h2>No call found in row ' + row + '</h2>'));
  const meta = { clientName: r[1] || '', businessName: r[2] || '', email: r[3] || '',
    dateStr: new Date(r[0] || Date.now()).toLocaleString('en-GB', { dateStyle: 'long' }) };
  const transcript = parseTranscriptText(r[6]);
  if (transcript.length < 2) return res.status(422).send(adminPage('Too short', '<h2>That call has too little transcript to build a roadmap from</h2>'));
  res.send(adminPage('Queued', '<h2>Roadmap queued</h2><p>Drafting a roadmap for <strong>' + escHtml(meta.businessName) + '</strong>. It will arrive in your inbox in about a minute, or you will get a "[Roadmap draft FAILED]" email with the reason.</p>'));
  generateInternalNotes(transcript, meta.clientName, meta.businessName)
    .then(function(notes) { return runRoadmapJob(notes, transcript, meta); })
    .catch(function(err) {
      console.error('Admin roadmap notes failed', err);
      return runRoadmapJob({ note: 'notes step failed: ' + err.message }, transcript, meta);
    });
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
  let sheetRow = null;
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
      const appendRes = await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID, range: 'Sheet1!A1',
        valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
        requestBody: { values: [row] }
      });
      sheetSaved = true;
      try {
        const m = String((appendRes.data.updates || {}).updatedRange || '').match(/!A(\d+)/);
        if (m) sheetRow = parseInt(m[1], 10);
      } catch (e) {}
    } catch (err) {
      console.error('Sheet save error', err.message);
    }
  }

  let clientEmailed = false;
  let internalEmailed = false;
  let summaryError = null;
  let roadmapJob = null;

  if (email && transcript.length > 1) {
    let notes = null;
    let notesError = null;
    try {
      notes = await generateInternalNotes(transcript, clientName, businessName);
    } catch (err) {
      console.error('Notes generation failed', err);
      notesError = err.message;
      summaryError = err.message;
    }

    // Draft roadmap: only for calls where the notes say YES (or always, with ROADMAP_ALWAYS=true). Runs after the response.
    if (notes && (process.env.ROADMAP_ALWAYS === 'true' || /^\s*yes/i.test(String(notes.shouldOfferRoadmap || '')))) {
      roadmapJob = function() {
        runRoadmapJob(notes, transcript, {
          clientName, businessName, email,
          dateStr: new Date(startTime || Date.now()).toLocaleString('en-GB', { dateStyle: 'long' })
        });
      };
    }

    // Client email: with AUTO_SEND_SNAPSHOT=true and a usable strong/moderate-fit snapshot, send the Tier 1 snapshot;
    // otherwise (default) send the generic acknowledgement and leave the snapshot as a draft for you.
    const autoSnapshot = process.env.AUTO_SEND_SNAPSHOT === 'true' && notes && snapshotUsable(notes);
    try {
      const sent = autoSnapshot
        ? await sendClientSnapshot(email, notes, clientName, businessName)
        : await sendClientAcknowledgement(email, clientName, businessName);
      clientEmailed = !sent.skipped;
    } catch (err) {
      console.error('Client email failed', err);
      summaryError = (summaryError ? summaryError + ' | ' : '') + err.message;
    }

    try {
      if (!notes) throw new Error(notesError || 'no notes');
      const docBuffer = await buildInternalNotesDocx(notes, {
        clientName, businessName, email,
        dateStr: new Date(startTime || Date.now()).toLocaleString('en-GB', { dateStyle: 'long', timeStyle: 'short' }),
        durationStr: Math.floor(durationSeconds / 60) + ' min ' + (durationSeconds % 60) + ' sec'
      });
      let afterHtml = '';
      if (snapshotUsable(notes)) {
        const d = buildSnapshotEmail(notes, clientName, businessName);
        afterHtml = '<hr/><p><strong>' + (autoSnapshot ? 'Sent to the client automatically' : 'Draft client snapshot — not sent. Review, then forward or paste') + ':</strong></p>' +
          '<p><em>Subject: ' + escHtml(d.subject) + '</em></p><div style="border:1px solid #e5e7eb;border-radius:8px;padding:12px">' + d.html + '</div>'
        const sc = normaliseNotes(notes).snapshotChecks || [];
        if (sc.length) afterHtml += '<p><strong>Check before sending:</strong></p><ul>' + sc.map(c => '<li>' + escHtml(c) + '</li>').join('') + '</ul>';
      } else {
        afterHtml = '<hr/><p><em>No client snapshot drafted — fit was weak or the AI had too little to go on. The client received the generic acknowledgement.</em></p>';
      }
      if (!roadmapJob) afterHtml += roadmapLinkHtml(sheetRow);
      const internalResult = await sendInternalNotes(clientName, businessName, email, docBuffer, notes.shouldOfferRoadmap, { afterHtml });
      internalEmailed = !internalResult.skipped;
      if (internalResult.skipped) summaryError = (summaryError ? summaryError + ' | ' : '') + 'Internal email skipped: RESEND_API_KEY or BCC_EMAIL not set on the server';
    } catch (err) {
      console.error('Internal notes/email error', err);
      summaryError = (summaryError ? summaryError + ' | ' : '') + err.message;
      // Fallback: if the Word doc could not be made or sent, still email the raw transcript so the call is never lost
      if (!internalEmailed) {
        try {
          const raw = Buffer.from('Client: ' + clientName + '\nBusiness: ' + businessName + '\nEmail: ' + email + '\n\n' + transcriptText, 'utf8');
          const fb = await sendInternalNotes(clientName, businessName, email, raw, null, {
            ext: 'txt',
            extraHtml: '<p style="color:#b45309"><strong>Note:</strong> the Word summary could not be generated or sent (' + String(err.message).replace(/[<>]/g, '') + '). Raw transcript attached instead.</p>'
          });
          internalEmailed = !fb.skipped;
        } catch (err2) {
          console.error('Fallback transcript email failed', err2);
          summaryError += ' | fallback: ' + err2.message;
        }
      }
    }
  }

  res.json({ ok: sheetSaved || clientEmailed, sheetSaved, clientEmailed, internalEmailed, summaryError, roadmapQueued: !!roadmapJob });
  if (roadmapJob) roadmapJob();
});

// ============================================================
//  POST /api/contact  — contact form submissions
// ============================================================
app.post('/api/contact', async (req, res) => {
  const { name = '', business = '', email = '', message = '' } = req.body || {};

  if (!name || !email || !message) {
    return res.status(400).json({ error: 'name, email and message are required' });
  }

  // Always notify us via BCC_EMAIL
  if (RESEND_API_KEY && BCC_EMAIL) {
    try {
      await fetch('https://api.resend.com/emails', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${RESEND_API_KEY}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          from: FROM_EMAIL,
          to: [BCC_EMAIL],
          subject: `New contact form message from ${name}${business ? ' · ' + business : ''}`,
          html: `
            <div style="font-family: -apple-system, sans-serif; max-width: 560px; margin: 0 auto; color: #374151; line-height: 1.6;">
              <h2 style="color: #1a2332; margin-bottom: 16px;">Contact form message</h2>
              <p><strong>Name:</strong> ${name}</p>
              ${business ? `<p><strong>Business:</strong> ${business}</p>` : ''}
              <p><strong>Email:</strong> <a href="mailto:${email}">${email}</a></p>
              <div style="background: #f9f9f6; border: 1px solid #e8e8e0; border-radius: 10px; padding: 16px; margin: 16px 0;">
                <p style="margin: 0; white-space: pre-wrap;">${message.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</p>
              </div>
              <p style="font-size: 12px; color: #94a3b8;">Sent via aigrowthcatalyst.org/contact.html</p>
            </div>
          `
        })
      });
    } catch (err) {
      console.error('contact notification email error', err);
    }
  }

  // Send acknowledgement to the person who filled in the form
  if (RESEND_API_KEY && email) {
    const firstName = name.split(' ')[0] || 'there';
    try {
      await fetch('https://api.resend.com/emails', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${RESEND_API_KEY}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          from: FROM_EMAIL,
          to: [email],
          subject: 'Thanks for getting in touch',
          html: `
            <div style="font-family: -apple-system, sans-serif; max-width: 560px; margin: 0 auto; color: #374151; line-height: 1.6;">
              <h2 style="color: #1a2332; margin-bottom: 16px;">Hi ${firstName},</h2>
              <p style="font-size: 15px; margin-bottom: 16px;">
                Thanks for reaching out. We've received your message and will get back to you within one business day.
              </p>
              <p style="font-size: 15px; margin-bottom: 16px;">
                In the meantime, if you're curious whether AI could help your business, the free 15-minute discovery call is the fastest way to find out — no pitch, no obligation.
              </p>
              <p style="margin: 24px 0;">
                <a href="https://aigrowthcatalyst.org/call/" style="background: #0f172a; color: #fafaf7; padding: 12px 24px; border-radius: 100px; text-decoration: none; font-size: 14px; font-weight: 500;">Book the free discovery call</a>
              </p>
              <p style="font-size: 15px; margin-top: 24px;">
                Speak soon,<br/>
                <strong>The AI Growth Catalyst team</strong>
              </p>
            </div>
          `
        })
      });
    } catch (err) {
      console.error('contact ack email error', err);
    }
  }

  res.json({ ok: true });
});

app.listen(PORT, () => {
  console.log('AI Growth Catalyst backend v4-quote listening on port', PORT);
});
