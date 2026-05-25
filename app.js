// ── IndexedDB ─────────────────────────────────────────────────
const DB_NAME = 'lt-words';
const DB_VERSION = 1;
let db;

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = e => {
      const d = e.target.result;
      if (!d.objectStoreNames.contains('cards')) {
        const store = d.createObjectStore('cards', { keyPath: 'id' });
        store.createIndex('deck', 'deck', { unique: false });
        store.createIndex('nextReview', 'nextReview', { unique: false });
      }
    };
    req.onsuccess = e => resolve(e.target.result);
    req.onerror = e => reject(e.target.error);
  });
}

function getAllCards() {
  return new Promise((resolve, reject) => {
    const req = db.transaction('cards', 'readonly').objectStore('cards').getAll();
    req.onsuccess = e => resolve(e.target.result);
    req.onerror = e => reject(e.target.error);
  });
}

function putCards(cards) {
  return new Promise((resolve, reject) => {
    const t = db.transaction('cards', 'readwrite');
    const store = t.objectStore('cards');
    t.oncomplete = resolve;
    t.onerror = e => reject(e.target.error);
    t.onabort = e => reject(e.target.error);
    for (const card of cards) store.put(card);
  });
}

function putCard(card) {
  return new Promise((resolve, reject) => {
    const req = db.transaction('cards', 'readwrite').objectStore('cards').put(card);
    req.onsuccess = () => resolve();
    req.onerror = e => reject(e.target.error);
  });
}

function clearCards() {
  return new Promise((resolve, reject) => {
    const req = db.transaction('cards', 'readwrite').objectStore('cards').clear();
    req.onsuccess = () => resolve();
    req.onerror = e => reject(e.target.error);
  });
}

// ── Leitner SRS ───────────────────────────────────────────────
// Intervals in days per box (index = box number 1-5)
const INTERVALS = [0, 1, 2, 4, 8, 16];

function startOfDay(offsetDays = 0) {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + offsetDays);
  return d.getTime();
}

function isDue(card) {
  return card.nextReview <= startOfDay();
}

function rateCard(card, knew) {
  if (knew) {
    card.leitnerBox = Math.min(5, card.leitnerBox + 1);
  } else {
    card.leitnerBox = 1;
  }
  card.nextReview = startOfDay(INTERVALS[card.leitnerBox]);
  card.totalReviews++;
  if (knew) card.correctReviews++;
}

// ── XLSX Parsing ──────────────────────────────────────────────
function isHeaderRow(row) {
  const first = String(row[0] || '').toLowerCase();
  return first.includes('lietuv') || first.includes('kalba') || first.includes('русский');
}

function parseSheets(workbook) {
  const cards = [];
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!rows.length) continue;

    const startIdx = isHeaderRow(rows[0]) ? 1 : 0;

    for (let i = startIdx; i < rows.length; i++) {
      const row = rows[i];
      const front = String(row[0] || '').trim();
      const back = String(row[1] || '').trim();
      if (!front || !back) continue;

      cards.push({
        id: crypto.randomUUID(),
        deck: sheetName,
        front,
        back,
        note: String(row[2] || '').trim(),
        imageBlob: null,
        leitnerBox: 1,
        nextReview: 0,       // 0 = immediately due
        totalReviews: 0,
        correctReviews: 0
      });
    }
  }
  return cards;
}

// ── State ─────────────────────────────────────────────────────
let allCards = [];
let studyQueue = [];
let currentIdx = 0;
let isFlipped = false;
let selectedDecks = new Set();
let sessionStats = { correct: 0, wrong: 0, startTime: 0 };
let imageMap = {};

// ── Screen Management ─────────────────────────────────────────
function showScreen(id) {
  document.querySelectorAll('.screen').forEach(s => s.classList.add('hidden'));
  document.getElementById(id).classList.remove('hidden');
}

// ── Import Screen ─────────────────────────────────────────────
async function handleFile(file) {
  const status = document.getElementById('import-status');
  status.className = 'status-msg';
  status.textContent = 'Reading file...';

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    const cards = parseSheets(workbook);

    if (!cards.length) throw new Error('No valid cards found in the file.');

    status.textContent = `Importing ${cards.length} cards…`;
    await clearCards();
    await putCards(cards);

    const deckCount = new Set(cards.map(c => c.deck)).size;
    status.textContent = `Imported ${cards.length} cards from ${deckCount} decks.`;
    setTimeout(showHome, 900);
  } catch (err) {
    status.className = 'status-msg error';
    status.textContent = 'Error: ' + err.message;
    console.error(err);
  }
}

// ── Home Screen ───────────────────────────────────────────────
async function showHome() {
  allCards = await getAllCards();

  if (!allCards.length) {
    showScreen('screen-import');
    return;
  }

  showScreen('screen-home');

  const deckNames = [...new Set(allCards.map(c => c.deck))];

  // Restore selection or default to all
  if (!selectedDecks.size) deckNames.forEach(d => selectedDecks.add(d));

  renderDeckSelector(deckNames);
  updateDueBadge();
}

function renderDeckSelector(deckNames) {
  const container = document.getElementById('deck-selector');
  container.innerHTML = '';

  for (const deck of deckNames) {
    const total = allCards.filter(c => c.deck === deck).length;
    const due = allCards.filter(c => c.deck === deck && isDue(c)).length;

    const label = document.createElement('label');
    label.className = 'deck-option' + (selectedDecks.has(deck) ? ' selected' : '');

    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = selectedDecks.has(deck);
    cb.addEventListener('change', () => {
      if (cb.checked) { selectedDecks.add(deck); label.classList.add('selected'); }
      else { selectedDecks.delete(deck); label.classList.remove('selected'); }
      updateDueBadge();
    });

    const nameSpan = document.createElement('span');
    nameSpan.className = 'deck-name';
    nameSpan.textContent = deck;

    const countSpan = document.createElement('span');
    countSpan.className = 'deck-counts';
    countSpan.innerHTML = `<span class="deck-due-count">${due}</span>&thinsp;/&thinsp;${total} due`;

    label.append(cb, nameSpan, countSpan);
    container.appendChild(label);
  }
}

function updateDueBadge() {
  const due = allCards.filter(c => selectedDecks.has(c.deck) && isDue(c)).length;
  const badge = document.getElementById('due-badge');
  badge.textContent = due
    ? `${due} card${due !== 1 ? 's' : ''} due today`
    : 'All caught up today!';
  badge.className = 'due-badge' + (due === 0 ? ' zero' : '');
}

// ── Study Screen ──────────────────────────────────────────────
function startStudy() {
  const due = allCards.filter(c => selectedDecks.has(c.deck) && isDue(c));

  if (!due.length) {
    alert('No cards due right now! Come back tomorrow or select more decks.');
    return;
  }

  // Sort: oldest nextReview first, shuffle within same day
  studyQueue = due.sort((a, b) => a.nextReview - b.nextReview);
  currentIdx = 0;
  sessionStats = { correct: 0, wrong: 0, startTime: Date.now() };

  showScreen('screen-study');
  showCard();
}

function showCard() {
  if (currentIdx >= studyQueue.length) {
    showComplete();
    return;
  }

  const card = studyQueue[currentIdx];
  const total = studyQueue.length;
  const cardEl = document.getElementById('card');

  // Progress
  document.getElementById('progress-bar').style.width = `${(currentIdx / total) * 100}%`;
  document.getElementById('progress-text').textContent = `${currentIdx + 1} / ${total}`;

  // Card content
  document.getElementById('card-front-word').textContent = card.front;
  document.getElementById('card-deck-label').textContent = card.deck;
  document.getElementById('card-back-word').textContent = card.back;
  document.getElementById('card-note').textContent = card.note || '';
  document.getElementById('card-box-indicator').textContent = `Box ${card.leitnerBox}`;

  // Image
  const imgEl = document.getElementById('card-front-img');
  const imgUrl = imageMap[card.front];
  if (imgUrl) {
    imgEl.src = imgUrl;
    imgEl.classList.remove('hidden');
    cardEl.classList.add('has-image');
  } else {
    imgEl.src = '';
    imgEl.classList.add('hidden');
    cardEl.classList.remove('has-image');
  }

  // Reset flip instantly (no animation — prevents Russian side flashing into view)
  cardEl.classList.add('no-transition');
  cardEl.classList.remove('flipped');
  cardEl.offsetHeight; // force reflow so the class takes effect before transition is restored
  cardEl.classList.remove('no-transition');
  isFlipped = false;

  // Show "Show answer" button
  document.getElementById('controls-show').classList.remove('hidden');
  document.getElementById('controls-rate').classList.add('hidden');
}

function showAnswer() {
  if (isFlipped) return;
  isFlipped = true;
  document.getElementById('card').classList.add('flipped');
  document.getElementById('controls-show').classList.add('hidden');
  document.getElementById('controls-rate').classList.remove('hidden');
}

function rateAndNext(knew) {
  const card = studyQueue[currentIdx];
  rateCard(card, knew);
  putCard(card); // fire-and-forget — UI stays snappy
  if (knew) sessionStats.correct++;
  else sessionStats.wrong++;
  currentIdx++;
  showCard();
}

// ── Complete Screen ───────────────────────────────────────────
function showComplete() {
  const total = sessionStats.correct + sessionStats.wrong;
  const pct = total ? Math.round((sessionStats.correct / total) * 100) : 0;
  const elapsed = Math.round((Date.now() - sessionStats.startTime) / 1000);
  const mins = Math.floor(elapsed / 60);
  const secs = elapsed % 60;
  const timeStr = mins > 0 ? `${mins}m ${secs}s` : `${secs}s`;

  document.getElementById('complete-stats').innerHTML = `
    <div class="stat-card">
      <div class="stat-value">${total}</div>
      <div class="stat-label">Cards studied</div>
    </div>
    <div class="stat-card">
      <div class="stat-value">${pct}%</div>
      <div class="stat-label">Correct</div>
    </div>
    <div class="stat-card">
      <div class="stat-value">${sessionStats.correct}</div>
      <div class="stat-label">Knew</div>
    </div>
    <div class="stat-card">
      <div class="stat-value">${timeStr}</div>
      <div class="stat-label">Time</div>
    </div>
  `;

  showScreen('screen-complete');
}

// ── Keyboard Shortcuts ────────────────────────────────────────
document.addEventListener('keydown', e => {
  const studyVisible = !document.getElementById('screen-study').classList.contains('hidden');
  if (!studyVisible) return;

  switch (e.key) {
    case ' ':
    case 'Enter':
      e.preventDefault();
      if (!isFlipped) showAnswer();
      break;
    case 'ArrowRight':
    case 'k':
    case 'K':
      if (isFlipped) rateAndNext(true);
      break;
    case 'ArrowLeft':
    case 'f':
    case 'F':
      if (isFlipped) rateAndNext(false);
      break;
  }
});

// ── Init ─────────────────────────────────────────────────────
async function init() {
  db = await openDB();
  fetch('images.json').then(r => r.json()).then(d => { imageMap = d; }).catch(() => { });

  // ── Import screen ──
  const dropZone = document.getElementById('drop-zone');

  dropZone.addEventListener('dragover', e => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });

  document.getElementById('file-input').addEventListener('change', e => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
  });

  // ── Home screen ──
  document.getElementById('btn-study').addEventListener('click', startStudy);
  document.getElementById('btn-import-new').addEventListener('click', () => showScreen('screen-import'));

  // ── Study screen ──
  document.getElementById('card').addEventListener('click', showAnswer);
  document.getElementById('btn-show').addEventListener('click', showAnswer);
  document.getElementById('btn-wrong').addEventListener('click', () => rateAndNext(false));
  document.getElementById('btn-know').addEventListener('click', () => rateAndNext(true));
  document.getElementById('btn-home').addEventListener('click', showHome);

  // ── Complete screen ──
  document.getElementById('btn-again').addEventListener('click', startStudy);
  document.getElementById('btn-home-complete').addEventListener('click', showHome);

  // ── Register Service Worker ──
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(() => { });
  }

  // ── Boot ──
  await showHome();
}

init();
