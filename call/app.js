// =============================================================
//  Business Discovery Voice Agent — Frontend
// =============================================================
//  IMPORTANT: After deploying your Render backend, update this URL
//  to point to your Render service. Keep the trailing slash.
// =============================================================
const BACKEND_URL = "https://ai-discovery-tx39.onrender.com";
// =============================================================

const TOPICS = [
  { id: 'business_overview',     label: 'Business overview',    opener: "To start, can you tell me about your business — what you do, who your customers are, and roughly how big the team is?" },
  { id: 'daily_operations',      label: 'Daily operations',     opener: "Walk me through a typical day or week in the business. What are the main activities and who does what?" },
  { id: 'systems_tools',         label: 'Systems and tools',    opener: "What software, systems, or tools do you currently use to run the business? Things like CRM, accounting, email, project management." },
  { id: 'manual_work',           label: 'Manual work',          opener: "What tasks do you or your team do manually that feel like they could or should be automated?" },
  { id: 'pain_points',           label: 'Pain points',          opener: "What are the biggest pain points or frustrations in the business right now? The things that genuinely slow you down or stress you out." },
  { id: 'customer_interactions', label: 'Customer interactions',opener: "Tell me about how you interact with customers. Inquiries, support, follow-ups, that sort of thing." },
  { id: 'priorities_goals',      label: 'Priorities and goals', opener: "If you could wave a magic wand and have one thing in your business automated or improved tomorrow, what would it be?" }
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
  businessName: ''
};

// ----- DOM refs -----
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
const nameSection = document.getElementById('name-section');

const startBtn = document.getElementById('start-btn');
const stopBtn = document.getElementById('stop-btn');
const sendTextBtn = document.getElementById('send-text-btn');
const exportBtn = document.getElementById('export-btn');
const resetBtn = document.getElementById('reset-btn');

// ----- Helpers -----
function setStatus(text, mode) {
  statusEl.textContent = text;
  dot.className = 'dot' + (mode ? ' ' + mode : '');
}
function showWarn(msg) { warnEl.textContent = msg; }
function clearWarn() { warnEl.textContent = ''; }
function showOk(msg, autoclear) {
  okMsgEl.textContent = msg;
  if (autoclear) setTimeout(function() { okMsgEl.textContent = ''; }, 4000);
}
function timestamp() {
  return new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
}
function escapeHtml(s) {
  return s.replace(/[&<>"']/g, function(c) {
    return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c];
  });
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

// ----- Speech (browser built-in for now; swap to ElevenLabs later) -----
function speak(text, onDone) {
  if (!('speechSynthesis' in window)) { if (onDone) onDone(); return; }
  window.speechSynthesis.cancel();
  const utter = new SpeechSynthesisUtterance(text);
  utter.rate = 1.0; utter.pitch = 1.0; utter.volume = 1.0;
  const voices = window.speechSynthesis.getVoices();
  // Prefer a clear English voice — try in order: SA, GB, US, AU, then any English
  const preferred =
    voices.find(function(v) { return /en-ZA/i.test(v.lang); }) ||
    voices.find(function(v) { return /en-GB/i.test(v.lang) && /female|samantha|kate|karen|tessa|fiona|serena/i.test(v.name); }) ||
    voices.find(function(v) { return /en-GB/i.test(v.lang); }) ||
    voices.find(function(v) { return /en-US/i.test(v.lang) && /female|samantha|karen|allison|ava|nicky/i.test(v.name); }) ||
    voices.find(function(v) { return /en-US/i.test(v.lang); }) ||
    voices.find(function(v) { return v.lang.startsWith('en'); });
  if (preferred) utter.voice = preferred;

  state.isSpeaking = true;
  setStatus('Agent speaking...', 'speaking');
  utter.onend  = function() { state.isSpeaking = false; if (onDone) onDone(); };
  utter.onerror = function() { state.isSpeaking = false; if (onDone) onDone(); };
  window.speechSynthesis.speak(utter);
}

// ----- Transcript helpers -----
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

// ----- Backend calls -----
async function generateNextQuestion() {
  setStatus('Agent thinking...', 'thinking');
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
        isLastTopic: isLastTopic
      })
    });
    if (!response.ok) throw new Error('Backend ' + response.status);
    const data = await response.json();
    return data;
  } catch (err) {
    console.error('Backend error', err);
    showWarn('Could not reach the backend. Falling back to next topic.');
    const nextIdx = state.currentTopicIdx + 1;
    if (nextIdx >= TOPICS.length) {
      return { next_question: "Thanks for sharing all of this. Is there anything I haven't asked about that you think I should know?", action: 'wrap_up' };
    }
    return { next_question: TOPICS[nextIdx].opener, action: 'next_topic' };
  }
}

async function handleAnswerAndContinue() {
  const decision = await generateNextQuestion();
  clearWarn();

  if (decision.action === 'next_topic') {
    state.currentTopicIdx = Math.min(state.currentTopicIdx + 1, TOPICS.length - 1);
    state.exchangesOnTopic = 1;
  } else if (decision.action === 'wrap_up') {
    pushAgent(decision.next_question);
    speak(decision.next_question, function() { startListening(); });
    state.callOver = true;
    return;
  } else {
    state.exchangesOnTopic++;
  }

  pushAgent(decision.next_question);
  speak(decision.next_question, function() { startListening(); });
}

async function saveTranscriptToBackend() {
  setStatus('Saving transcript...', 'thinking');
  try {
    const response = await fetch(BACKEND_URL.replace(/\/$/, '') + '/api/save-transcript', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        clientName: state.clientName,
        businessName: state.businessName,
        startTime: state.startTime ? new Date(state.startTime).toISOString() : null,
        endTime: new Date().toISOString(),
        durationSeconds: state.startTime ? Math.round((Date.now() - state.startTime) / 1000) : 0,
        topicsCovered: Array.from(state.topicsCovered).map(function(i) { return TOPICS[i].label; }),
        transcript: state.transcript
      })
    });
    if (!response.ok) throw new Error('Save failed: ' + response.status);
    showOk('Transcript saved successfully.', true);
    setStatus('Done. Transcript saved to Google Sheets.', null);
  } catch (err) {
    console.error('Save error', err);
    showWarn('Could not save transcript automatically. Use Download to keep a local copy.');
  }
}

// ----- Speech recognition -----
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
  setStatus('Got it. Thinking about what to ask next...', 'thinking');

  if (state.callOver) {
    const closing = "That's everything I wanted to ask. Thanks so much for taking the time — this gives a really clear picture. Your answers have been saved.";
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

// ----- Lifecycle -----
function startCall() {
  const cn = (clientNameInput.value || '').trim();
  const bn = (businessNameInput.value || '').trim();
  if (!cn || !bn) {
    showWarn('Please enter your name and business name before starting.');
    return;
  }
  state.clientName = cn;
  state.businessName = bn;
  clearWarn();

  state.started = true;
  state.startTime = Date.now();
  state.timerInterval = setInterval(updateDuration, 1000);

  startBtn.disabled = true;
  stopBtn.disabled = false;
  exportBtn.disabled = false;
  clientNameInput.disabled = true;
  businessNameInput.disabled = true;
  nameSection.style.opacity = '0.5';

  state.currentTopicIdx = 0;
  state.exchangesOnTopic = 1;
  const opener = "Hi " + state.clientName.split(' ')[0] + ", thanks for taking the time today. " + TOPICS[0].opener;
  pushAgent(opener);
  speak(opener, function() { startListening(); });
}

async function endCallEarly() {
  stopListening();
  if (state.timerInterval) clearInterval(state.timerInterval);
  const closing = "Thanks for your time — your answers have been saved.";
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
    clientName: '', businessName: ''
  };
  currentQEl.textContent = "When you press start, the agent will ask its first question and listen for your reply.";
  topicPill.textContent = 'Intro';
  setStatus('Enter your details and press start when ready.', null);
  durationEl.textContent = '0:00';
  updateStats(); renderTranscript();
  startBtn.disabled = false; stopBtn.disabled = true;
  textInput.disabled = true; sendTextBtn.disabled = true;
  exportBtn.disabled = true;
  clientNameInput.disabled = false; businessNameInput.disabled = false;
  nameSection.style.opacity = '1';
  textInput.value = '';
  clientNameInput.value = ''; businessNameInput.value = '';
  clearWarn(); okMsgEl.textContent = '';
}

// ----- Wire up events -----
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

// Initialise
setupRecognition();
if ('speechSynthesis' in window) {
  window.speechSynthesis.getVoices();
  window.speechSynthesis.onvoiceschanged = function() { window.speechSynthesis.getVoices(); };
}
