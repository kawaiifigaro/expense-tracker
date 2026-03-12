'use strict';

// ============================================================
// State
// ============================================================
let settings = { name: '', claudeApiKey: '', scriptUrl: '', sheetUrl: '' };
let currentImageData = null;
let currentMimeType  = 'image/jpeg';
let toastTimer       = null;

// ============================================================
// Init
// ============================================================
window.addEventListener('load', () => {
  loadSettings();
  if (isSetupComplete()) {
    showPage('main');
    document.getElementById('main-username').textContent = settings.name;
    renderRecentExpenses();
  } else {
    showPage('setup');
  }
});

// ============================================================
// Settings
// ============================================================
function loadSettings() {
  try {
    const s = localStorage.getItem('expense_settings');
    if (s) settings = { ...settings, ...JSON.parse(s) };
  } catch (e) { /* ignore */ }
}

function persistSettings() {
  localStorage.setItem('expense_settings', JSON.stringify(settings));
}

function isSetupComplete() {
  return !!(settings.name && settings.claudeApiKey && settings.scriptUrl);
}

function saveSetup() {
  const name      = val('setup-name');
  const apiKey    = val('setup-claude-key');
  const scriptUrl = val('setup-script-url');

  if (!name || !apiKey || !scriptUrl) {
    showToast('すべての必須項目を入力してください', 'error');
    return;
  }

  settings = { name, claudeApiKey: apiKey, scriptUrl, sheetUrl: '' };
  persistSettings();

  document.getElementById('main-username').textContent = name;
  showPage('main');
  renderRecentExpenses();
  showToast('設定を保存しました！', 'success');
}

function showSettings() {
  setVal('settings-name',       settings.name);
  setVal('settings-claude-key', settings.claudeApiKey);
  setVal('settings-script-url', settings.scriptUrl);
  setVal('settings-sheet-url',  settings.sheetUrl || '');
  document.getElementById('settings-modal').classList.remove('hidden');
}

function closeSettings() {
  document.getElementById('settings-modal').classList.add('hidden');
}

function handleModalClick(e) {
  if (e.target === document.getElementById('settings-modal')) closeSettings();
}

function updateSettings() {
  const name      = val('settings-name');
  const apiKey    = val('settings-claude-key');
  const scriptUrl = val('settings-script-url');
  const sheetUrl  = val('settings-sheet-url');

  if (!name || !apiKey || !scriptUrl) {
    showToast('必須項目を入力してください', 'error');
    return;
  }

  settings = { name, claudeApiKey: apiKey, scriptUrl, sheetUrl };
  persistSettings();
  closeSettings();
  document.getElementById('main-username').textContent = name;
  showToast('設定を更新しました', 'success');
}

// ============================================================
// GAS Code copy button
// ============================================================
function copyGasCode() {
  const code = document.getElementById('gas-code').textContent;
  navigator.clipboard.writeText(code).then(() => {
    showToast('コードをコピーしました！', 'success');
  }).catch(() => {
    // Fallback for older browsers
    const ta = document.createElement('textarea');
    ta.value = code;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    showToast('コードをコピーしました！', 'success');
  });
}

// ============================================================
// Camera / File
// ============================================================
function openCamera() {
  document.getElementById('file-input').click();
}

function handleFileSelect(event) {
  const file = event.target.files[0];
  event.target.value = '';
  if (!file) return;
  if (!file.type.startsWith('image/')) {
    showToast('画像ファイルを選択してください', 'error');
    return;
  }
  processImage(file);
}

async function processImage(file) {
  showLoading('画像を処理中...');
  try {
    const { base64, mimeType } = await resizeImage(file);
    currentImageData = base64;
    currentMimeType  = mimeType;

    document.getElementById('receipt-preview').innerHTML =
      `<img src="data:${mimeType};base64,${base64}" alt="領収書">`;

    showLoading('🤖 AIが領収書を解析中...');
    let extracted = null;
    try {
      extracted = await analyzeWithClaude(base64, mimeType);
    } catch (e) {
      console.warn('Claude failed:', e.message);
      showToast('AI解析に失敗しました。手動で入力してください', 'warning');
    }

    const today = new Date().toISOString().split('T')[0];
    setVal('review-date',        extracted?.date        || today);
    setVal('review-amount',      extracted?.amount      || '');
    setVal('review-payee',       extracted?.payee       || '');
    setVal('review-description', extracted?.description || '');
    setVal('review-notes',       '');
    setVal('review-person',      settings.name);

    if (extracted?.category) {
      const sel = document.getElementById('review-category');
      for (let i = 0; i < sel.options.length; i++) {
        if (sel.options[i].value === extracted.category) { sel.selectedIndex = i; break; }
      }
    }

    hideLoading();
    showPage('review');
  } catch (e) {
    hideLoading();
    showToast('エラー: ' + e.message, 'error');
  }
}

function resizeImage(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const MAX = 1600;
        let { width, height } = img;
        if (width > MAX || height > MAX) {
          if (width >= height) { height = Math.round(height * MAX / width); width = MAX; }
          else                 { width  = Math.round(width  * MAX / height); height = MAX; }
        }
        const canvas = document.createElement('canvas');
        canvas.width = width; canvas.height = height;
        canvas.getContext('2d').drawImage(img, 0, 0, width, height);
        const mime = file.type === 'image/png' ? 'image/png' : 'image/jpeg';
        canvas.toBlob(blob => {
          const r2 = new FileReader();
          r2.onload = ev => resolve({ base64: ev.target.result.split(',')[1], mimeType: mime });
          r2.onerror = reject;
          r2.readAsDataURL(blob);
        }, mime, mime === 'image/jpeg' ? 0.82 : undefined);
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ============================================================
// Claude Vision API
// ============================================================
async function analyzeWithClaude(base64, mimeType) {
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': settings.claudeApiKey,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json',
      'anthropic-dangerous-direct-browser-access': 'true',
    },
    body: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 512,
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mimeType, data: base64 } },
          { type: 'text', text:
`この領収書から以下のJSONのみを返してください（説明文不要）:
{"date":"YYYY-MM-DD","amount":数値,"payee":"店名","description":"内容","category":"交通費/接待費/消耗品費/会議費/通信費/出張費/書籍・資料費/その他のいずれか"}` },
        ],
      }],
    }),
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `API error ${res.status}`);
  }
  const data  = await res.json();
  const text  = data.content[0].text.trim();
  const match = text.match(/\{[\s\S]*\}/);
  if (!match) throw new Error('JSONが取得できませんでした');
  return JSON.parse(match[0]);
}

// ============================================================
// Submit → Google Apps Script
// ============================================================
async function submitExpense() {
  const date        = val('review-date');
  const amountStr   = val('review-amount');
  const payee       = val('review-payee');
  const description = val('review-description');
  const category    = document.getElementById('review-category').value;
  const notes       = val('review-notes');

  if (!date || !amountStr || !payee) {
    showToast('日付・金額・支払先は必須です', 'error');
    return;
  }
  const amount = parseInt(amountStr, 10);
  if (isNaN(amount) || amount <= 0) {
    showToast('金額を正しく入力してください', 'error');
    return;
  }

  showLoading('スプレッドシートに保存中...');
  try {
    const expense = { date, amount, payee, description, category, notes, person: settings.name };

    // text/plain で送信することでCORSプリフライトを回避
    const res = await fetch(settings.scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(expense),
    });

    // レスポンスが読めない場合もリクエストは送信済みなので成功扱い
    if (res.ok) {
      const result = await res.json().catch(() => ({ success: true }));
      if (!result.success) throw new Error(result.error || '保存に失敗しました');
    }

    addToHistory({ date, amount, payee, description, category });
    hideLoading();
    showSuccessPage({ date, amount, payee, description, category });
  } catch (e) {
    hideLoading();
    showToast('保存エラー: ' + e.message, 'error');
  }
}

// ============================================================
// Local History
// ============================================================
function addToHistory(expense) {
  const h = getHistory();
  h.unshift(expense);
  localStorage.setItem('expense_history', JSON.stringify(h.slice(0, 20)));
}

function getHistory() {
  try { return JSON.parse(localStorage.getItem('expense_history') || '[]'); } catch { return []; }
}

function renderRecentExpenses() {
  const history   = getHistory();
  const container = document.getElementById('recent-list');
  document.getElementById('recent-count').textContent = history.length + '件';

  if (!history.length) {
    container.innerHTML = '<p class="empty-state">まだ経費申請はありません</p>';
    return;
  }
  container.innerHTML = history.slice(0, 5).map(e => `
    <div class="expense-item">
      <div class="expense-main">
        <span class="expense-payee">${esc(e.payee)}</span>
        <span class="expense-amount">¥${Number(e.amount).toLocaleString()}</span>
      </div>
      <div class="expense-meta">
        <span>${esc(e.date)}</span>
        <span class="expense-category">${esc(e.category)}</span>
      </div>
    </div>`).join('');
}

// ============================================================
// UI
// ============================================================
function showPage(id) {
  document.querySelectorAll('.page').forEach(p => p.classList.add('hidden'));
  document.getElementById('page-' + id).classList.remove('hidden');
  window.scrollTo(0, 0);
  if (id === 'main') renderRecentExpenses();
}

function showSuccessPage(expense) {
  document.getElementById('success-details').innerHTML = `
    <div class="success-row"><span>支払先</span><strong>${esc(expense.payee)}</strong></div>
    <div class="success-row"><span>金額</span><strong>¥${Number(expense.amount).toLocaleString()}</strong></div>
    <div class="success-row"><span>日付</span><strong>${esc(expense.date)}</strong></div>
    <div class="success-row"><span>仕訳</span><strong>${esc(expense.category)}</strong></div>`;

  const sheetUrl = settings.sheetUrl || '#';
  document.getElementById('sheet-link').href = sheetUrl;
  if (!settings.sheetUrl) {
    document.getElementById('sheet-link').style.display = 'none';
  } else {
    document.getElementById('sheet-link').style.display = '';
  }
  showPage('success');
}

function showLoading(msg = '処理中...') {
  document.getElementById('loading-message').textContent = msg;
  document.getElementById('loading-overlay').classList.remove('hidden');
}

function hideLoading() {
  document.getElementById('loading-overlay').classList.add('hidden');
}

function showToast(message, type = 'info') {
  const t = document.getElementById('toast');
  t.textContent = message;
  t.className = `toast toast-${type}`;
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.add('hidden'), 3500);
}

function val(id)       { return document.getElementById(id).value.trim(); }
function setVal(id, v) { document.getElementById(id).value = v ?? ''; }
function esc(s) {
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
