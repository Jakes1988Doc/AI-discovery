// =============================================================
//  Business Discovery Voice Agent — Frontend (v3)
// =============================================================
const BACKEND_URL = "https://ai-discovery-tx39.onrender.com";
// =============================================================

const TOPICS = [
  {
    id: 'business_snapshot',
    label: 'Business snapshot',
    opener: "To kick us off — what does the business do, roughly how many people are in the team, and what's the main way revenue comes in?",
    probeQuestion: "And how long have you been running it, and is it mainly you driving the day-to-day or do you have people managing different areas?",
    moveOnCriteria: "you know what they do, team size, and revenue model",
    allowProbe: false
  },
  {
    id: 'tech_stack',
    label: 'Tech stack',
    opener: "What are the main tools and systems the business runs on day-to-day — things like how you manage customers or jobs, accounting, communications, taking orders, anything like that?",
    probeQuestion: "Is there anything in that list where you feel like you're working around the tool rather than with it — like it doesn't quite do what you need so you've had to add a workaround?",
    moveOnCriteria: "you've heard their core tools — CRM, accounting, comms, job management — and any friction with those tools",
    allowProbe: true
  },
  {
    id: 'process_walkthrough',
    label: 'Process walkthrough',
    opener: "Pick one process that happens regularly in the business — could be taking a new order, handling an enquiry, sending an invoice, anything — and just walk me through what actually happens, step by step, from start to finish.",
    probeQuestion: "Where in that process does someone have to stop and do something manually that feels like it should just happen automatically?",
    moveOnCriteria: "they've narrated a real process and you've identified at least one manual step or handoff",
    allowProbe: true
  },
  {
    id: 'time_and_pain',
    label: 'Time and pain points',
    opener: "If you think about a typical week — roughly how many hours would you say go into things that aren't actually the core work? I mean admin, chasing people, data entry, copying things from one place to another, that kind of thing.",
    probeQuestion: "Is there one specific task in that list where you've genuinely thought 'I can't believe we're still doing this by hand'?",
    moveOnCriteria: "you've heard either a time estimate or a specific frustrating manual task",
    allowProbe: true
  },
  {
    id: 'customer_journey',
    label: 'Customer journey',
    opener: "How does a new customer typically come to you, and what happens from that first contact through to them paying and coming back — is that process pretty consistent or does it vary?",
    probeQuestion: "And after the job or sale is done — do you have a consistent way of following up with customers, or does that depend on who's available at the time?",
    moveOnCriteria: "you understand how enquiries come in, how they're handled, and whether follow-up is systematic or ad-hoc",
    allowProbe: true
  },
  {
    id: 'magic_wand',
    label: 'Magic wand',
    opener: "Last one — if I could fix one thing in your business tomorrow and it would just work, no cost, no effort on your part — what would it be?",
    probeQuestion: null,
    moveOnCriteria: "they've answered",
    allowProbe: false
  }
];

let state = {
  started: false,
  callOver: false,
  currentTopicIdx: 0,
  exchangesOnTopic: 0,
  questionsAsked: 0,
  topicsCovered: new Set(),
  transcript: [],
  startTime: null,
  timerInterval: null,
  recognition: null,
  isListening: false,
  isSpeaking: false,
  finalTranscript: '',
  silenceTimer: null,
  clientName: '',
  businessName: '',
  email: '',
  signalDetectedOnCurrentTopic: false
};

const dot = document.getElementById('dot');
const statusEl = document.getElementById('status');
const currentQEl = document.getElementById('current-q-text');
const topicPill = document.getElementById('topic-pill');
const transcriptEl = document.getElementById('transcript');
const warnEl = document.getElementById('warn');
const okMsgEl = document.getElementById('ok-msg');
const qCountEl = document.getElementById('q-count');
const topicCountEl = document.getElementById('topic-count');
const durationEl = document.getElementById('duration');
const textInput = document.getElementById('text-input');
const clientNameInput = document.getElementById('client-name');
const businessNameInput = document.getElementById('business-name');
const emailInput = document.getElementById('email');
const nameSection = document.getElementById('name-section');

const startBtn = document.getElementById('start-btn');
const stopBtn = document.getElementById('stop-btn');
const sendTextBtn = document.getElementById('send-text-btn');
const exportBtn = document.getElementById('export-btn');
const resetBtn = document.getElementById('reset-btn');

function setStatus(text, mode) {
  statusEl.textContent = text;
  dot.className = 'dot' + (mode ? ' ' + mode : '');
}
function showWarn(msg) { warnEl.textContent = msg; }
function clearWarn() { warnEl.textContent = ''; }
function showOk(msg, autoclear) {
  okMsgEl.textContent = msg;
  if (autoclear) setTimeout(function() { okMsgEl.textContent = ''; }, 6000);
}
function timestamp() {
  return new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
}
function escapeHtml(s) {
  return s.replace(/[&<>"']/g, function(c) {
    return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c];
  });
}
function isValidEmail(e) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);
}

function updateStats() {
  qCountEl.textContent = state.questionsAsked;
  topicCountEl.textContent = state.topicsCovered.size + ' / ' + TOPICS.length;
  if (TOPICS[state.currentTopicIdx]) {
    topicPill.textContent = TOPICS[state.currentTopicIdx].label;
  }
}
function updateDuration() {
  if (!state.startTime) return;
  const secs = Math.round((Date.now() - state.startTime) / 1000);
  const m = Math.floor(secs / 60);
  const s = secs % 60;
  durationEl.textContent = m + ':' + (s < 10 ? '0' : '') + s;
}
function renderTranscript() {
  if (state.transcript.length === 0) {
    transcriptEl.innerHTML = '<p class="transcript-empty">The conversation transcript will appear here as you talk.</p>';
    return;
  }
  transcriptEl.innerHTML = state.transcript.map(function(t) {
    const cls = t.role === 'agent' ? 'turn-agent' : 'turn-user';
    const label = t.role === 'agent' ? 'Agent' : (state.clientName || 'You');
    return '<div class="turn ' + cls + '">' +
      '<div class="turn-meta">' + label + ' · ' + t.time + (t.topic ? ' · ' + t.topic : '') + '</div>' +
      '<p class="turn-text">' + escapeHtml(t.text) + '</p>' +
      '</div>';
  }).join('');
  transcriptEl.scrollTop = transcriptEl.scrollHeight;
}

function speak(text, onDone) {
  if (!('speechSynthesis' in window)) { if (onDone) onDone(); return; }
  window.speechSynthesis.cancel();
  const utter = new SpeechSynthesisUtterance(text);
  utter.rate = 1.0; utter.pitch = 1.0; utter.volume = 1.0;
  const voices = window.speechSynthesis.getVoices();
  const preferred =
    voices.find(function(v) { return /en-ZA/i.test(v.lang); }) ||
    voices.find(function(v) { return /en-GB/i.test(v.lang) && /female|samantha|kate|karen|tessa|fiona|serena/i.test(v.name); }) ||
    voices.find(function(v) { return /en-GB/i.test(v.lang); }) ||
    voices.find(function(v) { return /en-US/i.test(v.lang) && /female|samantha|karen|allison|ava|nicky/i.test(v.name); }) ||
    voices.find(function(v) { return /en-US/i.test(v.lang); }) ||
    voices.find(function(v) { return v.lang.startsWith('en'); });
  if (preferred) utter.voice = preferred;
  state.isSpeaking = true;
  setStatus('Agent speaking…', 'speaking');
  utter.onend  = function() { state.isSpeaking = false; if (onDone) onDone(); };
  utter.onerror = function() { state.isSpeaking = false; if (onDone) onDone(); };
  window.speechSynthesis.speak(utter);
}

function pushAgent(text) {
  const topicLabel = TOPICS[state.currentTopicIdx] ? TOPICS[state.currentTopicIdx].label : '';
  currentQEl.textContent = text;
  state.transcript.push({ role: 'agent', text: text, time: timestamp(), topic: topicLabel });
  state.questionsAsked++;
  updateStats();
  renderTranscript();
}
function pushUser(text) {
  const topicLabel = TOPICS[state.currentTopicIdx] ? TOPICS[state.currentTopicIdx].label : '';
  state.transcript.push({ role: 'user', text: text, time: timestamp(), topic: topicLabel });
  state.topicsCovered.add(state.currentTopicIdx);
  updateStats();
  renderTranscript();
}

async function generateNextQuestion() {
  setStatus('Agent thinking…', 'thinking');
  const currentTopic = TOPICS[state.currentTopicIdx];
  const isLastTopic = state.currentTopicIdx >= TOPICS.length - 1;
  const remainingTopics = TOPICS.slice(state.currentTopicIdx + 1).map(function(t) { return t.label; });

  const messagesForApi = state.transcript.map(function(t) {
    return { role: t.role === 'agent' ? 'assistant' : 'user', content: t.text };
  });

  try {
    const response = await fetch(BACKEND_URL.replace(/\/$/, '') + '/api/next-question', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        messages: messagesForApi,
        currentTopic: currentTopic.label,
        topicIndex: state.currentTopicIdx,
        totalTopics: TOPICS.length,
        exchangesOnTopic: state.exchangesOnTopic,
        remainingTopics: remainingTopics,
        isLastTopic: isLastTopic,
        signalDetected: state.signalDetectedOnCurrentTopic,
        topicProbeQuestion: currentTopic.probeQuestion || '',
        topicMoveOnCriteria: currentTopic.moveOnCriteria || '',
        topicAllowsProbe: currentTopic.allowProbe !== false
      })
    });
    if (!response.ok) throw new Error('Backend ' + response.status);
    const decision = await response.json();

    // Track signal detection returned by the agent
    if (decision.signal_detected === true) {
      state.signalDetectedOnCurrentTopic = true;
    }

    return decision;
  } catch (err) {
    console.error('Backend error', err);
    showWarn('Could not reach the backend. Falling back to next topic.');
    const nextIdx = state.currentTopicIdx + 1;
    if (nextIdx >= TOPICS.length) {
      return { next_question: "Thanks for sharing all of this. Is there anything else you'd like me to know before we wrap up?", action: 'wrap_up' };
    }
    return { next_question: TOPICS[nextIdx].opener, action: 'next_topic' };
  }
}

async function handleAnswerAndContinue() {
  const currentTopic = TOPICS[state.currentTopicIdx];

  // Hard guard: if this topic doesn't allow probes and we've had the opener exchange,
  // force next_topic without even calling the backend
  if (!currentTopic.allowProbe && state.exchangesOnTopic >= 1) {
    const nextIdx = state.currentTopicIdx + 1;
    if (nextIdx >= TOPICS.length) {
      const closing = "Thanks so much for your time today. That's genuinely useful context. We'll put together your personalised roadmap and be in touch within 48 hours.";
      pushAgent(closing);
      speak(closing);
      state.callOver = true;
      return;
    }
    state.currentTopicIdx = nextIdx;
    state.exchangesOnTopic = 0;
    state.signalDetectedOnCurrentTopic = false;
    const nextTopic = TOPICS[state.currentTopicIdx];
    pushAgent(nextTopic.opener);
    speak(nextTopic.opener, function() { startListening(); });
    return;
  }

  const decision = await generateNextQuestion();
  clearWarn();

  if (decision.action === 'next_topic') {
    state.currentTopicIdx = Math.min(state.currentTopicIdx + 1, TOPICS.length - 1);
    state.exchangesOnTopic = 0;
    state.signalDetectedOnCurrentTopic = false;
  } else if (decision.action === 'wrap_up') {
    pushAgent(decision.next_question);
    speak(decision.next_question, function() { startListening(); });
    state.callOver = true;
    return;
  } else {
    // followup
    state.exchangesOnTopic++;
  }

  pushAgent(decision.next_question);
  speak(decision.next_question, function() { startListening(); });
}

async function saveTranscriptToBackend() {
  setStatus('Saving your transcript…', 'thinking');
  try {
    const response = await fetch(BACKEND_URL.replace(/\/$/, '') + '/api/save-transcript', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        clientName: state.clientName,
        businessName: state.businessName,
        email: state.email,
        startTime: state.startTime ? new Date(state.startTime).toISOString() : null,
        endTime: new Date().toISOString(),
        durationSeconds: state.startTime ? Math.round((Date.now() - state.startTime) / 1000) : 0,
        topicsCovered: Array.from(state.topicsCovered).map(function(i) { return TOPICS[i].label; }),
        transcript: state.transcript
      })
    });
    if (!response.ok) throw new Error('Save failed: ' + response.status);
    showOk("All saved. We'll be in touch at " + state.email + " shortly.", true);
    setStatus('Done. Thanks!', null);
  } catch (err) {
    console.error('Save error', err);
    showWarn('Could not save automatically. Use Download to keep a local copy of your transcript.');
  }
}

function startListening() {
  if (!state.recognition) {
    setStatus('Mic not available — type your answer below.', null);
    textInput.disabled = false; sendTextBtn.disabled = false; textInput.focus();
    return;
  }
  state.finalTranscript = '';
  try {
    state.recognition.start();
    state.isListening = true;
    setStatus("Listening… (pause when you're done)", 'live');
    textInput.disabled = false; sendTextBtn.disabled = false;
  } catch (e) {
    setStatus('Mic busy — try the text input.', null);
  }
}
function stopListening() {
  if (state.recognition && state.isListening) {
    try { state.recognition.stop(); } catch (e) {}
    state.isListening = false;
  }
  if (state.silenceTimer) { clearTimeout(state.silenceTimer); state.silenceTimer = null; }
}
async function captureAnswer(text) {
  text = text.trim();
  if (!text) return;
  pushUser(text);
  setStatus('Got it. Thinking about what to ask next…', 'thinking');

  if (state.callOver) {
    const closing = "That's everything I wanted to ask. Thanks so much for taking the time. A human will now review what you've shared. If we identify a clear fit for AI in your business, we'll be in touch about the £350 bespoke roadmap. No pressure either way.";
    pushAgent(closing);
    await saveTranscriptToBackend();
    speak(closing);
    setStatus('Interview complete.', null);
    stopBtn.disabled = true;
    return;
  }
  await handleAnswerAndContinue();
}

function setupRecognition() {
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SR) {
    showWarn("Your browser doesn't support speech recognition. Use Chrome or Edge for voice, or use the text input below.");
    return;
  }
  const rec = new SR();
  rec.continuous = true; rec.interimResults = true; rec.lang = 'en-US';
  rec.onresult = function(ev) {
    for (let i = ev.resultIndex; i < ev.results.length; i++) {
      const t = ev.results[i][0].transcript;
      if (ev.results[i].isFinal) state.finalTranscript += t + ' ';
    }
    if (state.silenceTimer) clearTimeout(state.silenceTimer);
    state.silenceTimer = setTimeout(function() {
      if (state.finalTranscript.trim()) {
        stopListening();
        captureAnswer(state.finalTranscript);
      }
    }, 2200);
  };
  rec.onerror = function(ev) {
    if (ev.error === 'not-allowed') {
      showWarn('Microphone permission denied. Allow mic access or use the text input.');
      setStatus('Mic blocked — use text input.', null);
    } else if (ev.error !== 'no-speech') {
      setStatus('Mic issue: ' + ev.error + '. Use text input.', null);
    }
    state.isListening = false;
  };
  rec.onend = function() { state.isListening = false; };
  state.recognition = rec;
}

function startCall() {
  const cn = (clientNameInput.value || '').trim();
  const bn = (businessNameInput.value || '').trim();
  const em = (emailInput.value || '').trim();
  if (!cn || !bn || !em) {
    showWarn('Please enter your name, business name, and email before starting.');
    return;
  }
  if (!isValidEmail(em)) {
    showWarn('That email address looks invalid. Please double-check it.');
    return;
  }
  state.clientName = cn;
  state.businessName = bn;
  state.email = em;
  clearWarn();

  state.started = true;
  state.startTime = Date.now();
  state.timerInterval = setInterval(updateDuration, 1000);

  startBtn.disabled = true;
  stopBtn.disabled = false;
  exportBtn.disabled = false;
  clientNameInput.disabled = true;
  businessNameInput.disabled = true;
  emailInput.disabled = true;
  nameSection.style.opacity = '0.5';

  state.currentTopicIdx = 0;
  state.exchangesOnTopic = 0;
  state.signalDetectedOnCurrentTopic = false;
  const opener = "Hi " + state.clientName.split(' ')[0] + ", thanks for taking the time today. This is a quick fifteen-minute call to get a high-level picture of your business — I'll ask a few broad questions across six areas, with no recommendations from me on the call and no pressure to take a paid roadmap afterwards. Let's just have a useful conversation. " + TOPICS[0].opener;
  pushAgent(opener);
  speak(opener, function() { startListening(); });
}

async function endCallEarly() {
  stopListening();
  if (state.timerInterval) clearInterval(state.timerInterval);
  const closing = "Thanks for your time. We'll review what you've shared. If we see a clear fit for AI in your business, we'll be in touch about the £350 roadmap. No obligation either way.";
  pushAgent(closing);
  await saveTranscriptToBackend();
  speak(closing);
  setStatus('Call ended.', null);
  stopBtn.disabled = true;
}

function exportTranscript() {
  const lines = [
    'Business Discovery Interview Transcript',
    'Client: ' + state.clientName,
    'Business: ' + state.businessName,
    'Email: ' + state.email,
    'Date: ' + new Date().toLocaleString(),
    'Duration: ' + durationEl.textContent,
    'Topics covered: ' + state.topicsCovered.size + ' of ' + TOPICS.length,
    ''
  ];
  state.transcript.forEach(function(t) {
    lines.push('[' + t.time + '] ' + (t.role === 'agent' ? 'AGENT' : 'CLIENT') + (t.topic ? ' (' + t.topic + ')' : '') + ':');
    lines.push(t.text); lines.push('');
  });
  const blob = new Blob([lines.join('\n')], { type: 'text/plain' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'discovery-' + (state.businessName || 'call').replace(/[^a-z0-9]+/gi, '-').toLowerCase() + '-' + new Date().toISOString().slice(0,10) + '.txt';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function resetAll() {
  stopListening();
  if (state.timerInterval) clearInterval(state.timerInterval);
  if ('speechSynthesis' in window) window.speechSynthesis.cancel();
  state = {
    started: false, callOver: false, currentTopicIdx: 0, exchangesOnTopic: 0,
    questionsAsked: 0, topicsCovered: new Set(), transcript: [],
    startTime: null, timerInterval: null, recognition: state.recognition,
    isListening: false, isSpeaking: false, finalTranscript: '', silenceTimer: null,
    clientName: '', businessName: '', email: ''
  };
  currentQEl.textContent = "When you press start, the agent will ask its first question and listen for your reply.";
  topicPill.textContent = 'Intro';
  setStatus('Enter your details and press start when ready.', null);
  durationEl.textContent = '0:00';
  updateStats(); renderTranscript();
  startBtn.disabled = false; stopBtn.disabled = true;
  textInput.disabled = true; sendTextBtn.disabled = true;
  exportBtn.disabled = true;
  clientNameInput.disabled = false; businessNameInput.disabled = false; emailInput.disabled = false;
  nameSection.style.opacity = '1';
  textInput.value = '';
  clientNameInput.value = ''; businessNameInput.value = ''; emailInput.value = '';
  clearWarn(); okMsgEl.textContent = '';
}

startBtn.addEventListener('click', startCall);
stopBtn.addEventListener('click', endCallEarly);
exportBtn.addEventListener('click', exportTranscript);
resetBtn.addEventListener('click', resetAll);
sendTextBtn.addEventListener('click', function() {
  const v = textInput.value.trim();
  if (!v) return;
  stopListening();
  captureAnswer(v);
  textInput.value = '';
});
textInput.addEventListener('keypress', function(e) { if (e.key === 'Enter') sendTextBtn.click(); });

setupRecognition();
if ('speechSynthesis' in window) {
  window.speechSynthesis.getVoices();
  window.speechSynthesis.onvoiceschanged = function() { window.speechSynthesis.getVoices(); };
}
