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
    probeQuestion: "And is it mainly you driving the day-to-day, or do you have people managing different areas of the business?",
    moveOnCriteria: "you know what they do, team size, and revenue model",
    allowProbe: false,
    allowDeepProbe: false
  },
  {
    id: 'tech_stack',
    label: 'Tech stack',
    opener: "What are the main tools and systems the business runs on day-to-day — things like how you manage customers or jobs, accounting, communications, taking orders, anything like that?",
    probeQuestion: "Is there anything in that list where you feel like you're working around the tool rather than with it — like it doesn't quite do what you need so you've had to add a workaround?",
    deepProbeQuestion: "And when that workaround breaks down — what actually happens? Does work get missed, or does someone have to step in manually each time?",
    moveOnCriteria: "you've heard their core tools and any friction or workarounds",
    allowProbe: true,
    allowDeepProbe: true
  },
  {
    id: 'process_walkthrough',
    label: 'Process walkthrough',
    opener: "Pick one process that happens regularly in the business — could be taking a new order, handling an enquiry, sending an invoice, anything — and just walk me through what actually happens, step by step, from start to finish.",
    probeQuestion: "Where in that process does someone have to stop and do something manually that feels like it should just happen automatically?",
    deepProbeQuestion: "And roughly how often does that happen — is that a daily thing, or more like every time a new order comes in?",
    moveOnCriteria: "they've narrated a real process and you've identified at least one manual step",
    allowProbe: true,
    allowDeepProbe: true
  },
  {
    id: 'time_and_pain',
    label: 'Time and pain points',
    opener: "If you think about a typical week — roughly how many hours would you say go into things that aren't the core work? Admin, chasing people, data entry, copying things between systems, that kind of thing.",
    probeQuestion: "Is there one specific task in that list where you've genuinely thought 'I can't believe we're still doing this by hand'?",
    deepProbeQuestion: "And is that something one person owns, or does it fall to whoever's available at the time?",
    moveOnCriteria: "you've heard a time estimate or a specific frustrating manual task",
    allowProbe: true,
    allowDeepProbe: true
  },
  {
    id: 'admin_tasks',
    label: 'Admin and documents',
    opener: "What about the pure admin side of things — I'm talking meeting notes, reports, filing, preparing documents, chasing approvals, anything that's more paperwork than actual work. Is any of that eating into your week in a way that frustrates you?",
    probeQuestion: "Is there a specific document or report that gets produced regularly that you think — if you're honest — probably takes longer than it should?",
    deepProbeQuestion: "And who ends up doing that — is it you personally, or does it get passed around depending on who has time?",
    moveOnCriteria: "you've heard whether admin is a real time drain and identified at least one specific admin task",
    allowProbe: true,
    allowDeepProbe: true
  },
  {
    id: 'customer_journey',
    label: 'Customer journey',
    opener: "How does a new customer typically come to you, and what happens from that first contact through to them paying and coming back — is that process pretty consistent or does it vary a lot?",
    probeQuestion: "And after the job or sale is done — do you have a consistent way of following up with customers, or does that depend on who's available at the time?",
    deepProbeQuestion: "When follow-up does happen — is that something you're doing personally, or is there someone in the team whose job that is?",
    moveOnCriteria: "you understand how enquiries come in, how they're handled, and whether follow-up is systematic or ad-hoc",
    allowProbe: true,
    allowDeepProbe: false
  },
  {
    id: 'magic_wand',
    label: 'Magic wand',
    opener: "Last one — if I could fix one thing in your business tomorrow and it would just work, no cost, no effort on your part — what would it be?",
    probeQuestion: null,
    deepProbeQuestion: null,
    moveOnCriteria: "they've answered",
    allowProbe: false,
    allowDeepProbe: false
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
  signalDetectedOnCurrentTopic: false,
  signalStrengthOnCurrentTopic: 'none', // 'none' | 'standard' | 'strong'
  deepProbeAvailable: true,             // one deep probe allowed per call
  deepProbeUsedOnTopic: false           // whether deep probe used on current topic
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

// ============================================================
//  speak() — ElevenLabs TTS via backend proxy, browser TTS fallback
// ============================================================
let currentAudio = null;

function speakFallback(text, onDone) {
  // Browser TTS fallback when ElevenLabs is unavailable
  if (!('speechSynthesis' in window)) { if (onDone) onDone(); return; }
  window.speechSynthesis.cancel();
  const utter = new SpeechSynthesisUtterance(text);
  utter.rate = 1.0; utter.pitch = 1.0; utter.volume = 1.0;
  const voices = window.speechSynthesis.getVoices();
  const preferred =
    voices.find(function(v) { return /en-GB/i.test(v.lang); }) ||
    voices.find(function(v) { return v.lang.startsWith('en'); });
  if (preferred) utter.voice = preferred;
  state.isSpeaking = true;
  setStatus('Agent speaking…', 'speaking');
  utter.onend  = function() { state.isSpeaking = false; if (onDone) onDone(); };
  utter.onerror = function() { state.isSpeaking = false; if (onDone) onDone(); };
  window.speechSynthesis.speak(utter);
}

async function speak(text, onDone) {
  // Stop any currently playing audio
  if (currentAudio) {
    currentAudio.pause();
    currentAudio = null;
  }
  if ('speechSynthesis' in window) window.speechSynthesis.cancel();

  state.isSpeaking = true;
  setStatus('Agent speaking…', 'speaking');

  try {
    const response = await fetch(BACKEND_URL.replace(/\/$/, '') + '/api/speak', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text: text })
    });

    // If backend signals fallback (ElevenLabs not configured or error)
    if (!response.ok || response.headers.get('Content-Type')?.includes('application/json')) {
      const data = await response.json().catch(() => ({}));
      if (data.fallback) {
        speakFallback(text, onDone);
        return;
      }
      throw new Error('Speak endpoint error');
    }

    // Stream audio from ElevenLabs response
    const arrayBuffer = await response.arrayBuffer();
    const audioBlob = new Blob([arrayBuffer], { type: 'audio/mpeg' });
    const audioUrl = URL.createObjectURL(audioBlob);
    const audio = new Audio(audioUrl);
    currentAudio = audio;

    audio.onended = function() {
      state.isSpeaking = false;
      currentAudio = null;
      URL.revokeObjectURL(audioUrl);
      if (onDone) onDone();
    };
    audio.onerror = function() {
      console.warn('Audio playback error — falling back to browser TTS');
      state.isSpeaking = false;
      currentAudio = null;
      URL.revokeObjectURL(audioUrl);
      speakFallback(text, onDone);
    };

    await audio.play();

  } catch (err) {
    console.warn('ElevenLabs speak failed, using browser TTS:', err);
    state.isSpeaking = false;
    speakFallback(text, onDone);
  }
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
        signalStrength: state.signalStrengthOnCurrentTopic,
        topicProbeQuestion: currentTopic.probeQuestion || '',
        topicDeepProbeQuestion: currentTopic.deepProbeQuestion || '',
        topicMoveOnCriteria: currentTopic.moveOnCriteria || '',
        topicAllowsProbe: currentTopic.allowProbe !== false,
        topicAllowsDeepProbe: currentTopic.allowDeepProbe === true,
        deepProbeAvailable: state.deepProbeAvailable,
        deepProbeUsedOnTopic: state.deepProbeUsedOnTopic
      })
    });
    if (!response.ok) throw new Error('Backend ' + response.status);
    const decision = await response.json();

    // Track signal state returned by agent
    if (decision.signal_detected === true) {
      state.signalDetectedOnCurrentTopic = true;
    }
    if (decision.signal_strength === 'strong' && state.deepProbeAvailable && currentTopic.allowDeepProbe) {
      state.signalStrengthOnCurrentTopic = 'strong';
    } else if (decision.signal_detected === true) {
      state.signalStrengthOnCurrentTopic = 'standard';
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

  // Hard guard: no-probe topics move on immediately after opener
  if (!currentTopic.allowProbe && state.exchangesOnTopic >= 1) {
    moveToNextTopic();
    return;
  }

  // Hard guard: deep probe spent or not allowed — force next_topic after standard probe
  const deepProbeEligible = currentTopic.allowDeepProbe &&
                             state.deepProbeAvailable &&
                             !state.deepProbeUsedOnTopic;

  // If we've already done a standard probe AND no deep probe is eligible, move on
  if (state.exchangesOnTopic >= 2 && !deepProbeEligible) {
    moveToNextTopic();
    return;
  }

  // If we've done a standard probe AND a deep probe, always move on
  if (state.exchangesOnTopic >= 3) {
    moveToNextTopic();
    return;
  }

  const decision = await generateNextQuestion();
  clearWarn();

  if (decision.action === 'next_topic') {
    moveToNextTopic();
    return;
  }

  if (decision.action === 'wrap_up') {
    pushAgent(decision.next_question);
    speak(decision.next_question, function() { startListening(); });
    state.callOver = true;
    return;
  }

  // followup — track whether this is a deep probe
  state.exchangesOnTopic++;
  if (decision.is_deep_probe === true) {
    state.deepProbeUsedOnTopic = true;
    state.deepProbeAvailable = false; // spend the global budget
  }

  pushAgent(decision.next_question);
  speak(decision.next_question, function() { startListening(); });
}

function moveToNextTopic() {
  const nextIdx = state.currentTopicIdx + 1;
  if (nextIdx >= TOPICS.length) {
    // Finished all topics — wrap up
    const closing = "That's really helpful — thank you. Is there anything else you feel is worth mentioning before we finish up?";
    pushAgent(closing);
    speak(closing, function() { startListening(); });
    state.callOver = true;
    return;
  }
  state.currentTopicIdx = nextIdx;
  state.exchangesOnTopic = 0;
  state.signalDetectedOnCurrentTopic = false;
  state.signalStrengthOnCurrentTopic = 'none';
  state.deepProbeUsedOnTopic = false;
  const nextTopic = TOPICS[state.currentTopicIdx];
  pushAgent(nextTopic.opener);
  speak(nextTopic.opener, function() { startListening(); });
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
    }, 1000);
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
  state.signalStrengthOnCurrentTopic = 'none';
  state.deepProbeAvailable = true;
  state.deepProbeUsedOnTopic = false;
  const opener = "Hi " + state.clientName.split(' ')[0] + ", thanks for taking the time today. This is a quick fifteen-minute call to get a high-level picture of your business — I'll ask a few broad questions across six areas, with no recommendations from me on the call and no pressure to take a paid roadmap afterwards. Let's just have a useful conversation. " + TOPICS[0].opener;
  pushAgent(opener);
  speak(opener, function() { startListening(); });
}

async function endCallEarly() {
  stopListening();
  if (state.timerInterval) clearInterval(state.timerInterval);
  if (currentAudio) { currentAudio.pause(); currentAudio = null; }
  if ('speechSynthesis' in window) window.speechSynthesis.cancel();
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
  if (currentAudio) { currentAudio.pause(); currentAudio = null; }
  state = {
    started: false, callOver: false, currentTopicIdx: 0, exchangesOnTopic: 0,
    questionsAsked: 0, topicsCovered: new Set(), transcript: [],
    startTime: null, timerInterval: null, recognition: state.recognition,
    isListening: false, isSpeaking: false, finalTranscript: '', silenceTimer: null,
    clientName: '', businessName: '', email: '',
    signalDetectedOnCurrentTopic: false, signalStrengthOnCurrentTopic: 'none',
    deepProbeAvailable: true, deepProbeUsedOnTopic: false
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
