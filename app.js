'use strict';

// ============================================================
// State
// ============================================================
let settings = { name: '', claudeApiKey: '', googleClientId: '', spreadsheetId: '' };
let currentImageData  = null;   // base64 string
let currentMimeType   = 'image/jpeg';
let tokenClient       = null;
let accessToken       = null;
let toastTimer        = null;

// ============================================================
// Initialization
// ============================================================
window.addEventListener('load', () => {
  loadSettings();
  if (isSetupComplete()) {
    showPage('main');
    document.getElementById('main-username').textContent = settings.name;
    initGoogleAuth();
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
    const saved = localStorage.getItem('expense_settings');
    if (saved) settings = { ...settings, ...JSON.parse(saved) };
  } catch (e) { /* ignore */ }
}

function persistSettings() {
  localStorage.setItem('expense_settings', JSON.stringify(settings));
}

function isSetupComplete() {
  return !!(settings.name && settings.claudeApiKey && settings.googleClientId && settings.spreadsheetId);
}

function saveSetup() {
  const name           = val('setup-name');
  const claudeApiKey   = val('setup-claude-key');
  const googleClientId = val('setup-google-client-id');
  const spreadsheetId  = val('setup-sheet-id');

  if (!name || !claudeApiKey || !googleClientId || !spreadsheetId) {
    showToast('すべての項目を入力してください', 'error');
    return;
  }

  settings = { name, claudeApiKey, googleClientId, spreadsheetId };
  persistSettings();

  document.getElementById('main-username').textContent = name;
  showPage('main');
  initGoogleAuth();
  renderRecentExpenses();
  showToast('設定を保存しました', 'success');
}

function showSettings() {
  setVal('settings-name',              settings.name);
  setVal('settings-claude-key',        settings.claudeApiKey);
  setVal('settings-google-client-id',  settings.googleClientId);
  setVal('settings-sheet-id',          settings.spreadsheetId);
  document.getElementById('settings-modal').classList.remove('hidden');
}

function closeSettings() {
  document.getElementById('settings-modal').classList.add('hidden');
}

function handleModalClick(e) {
  if (e.target === document.getElementById('settings-modal')) closeSettings();
}

function updateSettings() {
  const name           = val('settings-name');
  const claudeApiKey   = val('settings-claude-key');
  const googleClientId = val('settings-google-client-id');
  const spreadsheetId  = val('settings-sheet-id');

  if (!name || !claudeApiKey || !googleClientId || !spreadsheetId) {
    showToast('すべての項目を入力してください', 'error');
    return;
  }

  settings = { name, claudeApiKey, googleClientId, spreadsheetId };
  persistSettings();
  closeSettings();

  document.getElementById('main-username').textContent = name;
  // Reset google auth if client ID changed
  tokenClient  = null;
  accessToken  = null;
  document.getElementById('google-signin-btn').classList.remove('hidden');
  document.getElementById('google-signed-in').classList.add('hidden');
  initGoogleAuth();

  showToast('設定を更新しました', 'success');
}

// ============================================================
// Google OAuth (Identity Services – token model)
// ============================================================
function initGoogleAuth() {
  if (!settings.googleClientId) return;
  if (typeof google === 'undefined' || !google?.accounts?.oauth2) {
    setTimeout(initGoogleAuth, 400);
    return;
  }
  try {
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: settings.googleClientId,
      scope: 'https://www.googleapis.com/auth/spreadsheets',
      callback: onGoogleToken,
    });
  } catch (e) {
    console.error('Google Auth init error:', e);
  }
}

function onGoogleToken(response) {
  if (response.error) {
    showToast('Google認証エラー: ' + response.error, 'error');
    return;
  }
  accessToken = response.access_token;
  document.getElementById('google-signin-btn').classList.add('hidden');
  document.getElementById('google-signed-in').classList.remove('hidden');
  showToast('Googleアカウントと連携しました', 'success');
}

function signInGoogle() {
  if (!tokenClient) {
    showToast('Google認証を初期化中です。しばらく待ってから再試行してください', 'warning');
    return;
  }
  tokenClient.requestAccessToken();
}

// ============================================================
// Camera / File Upload
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

    // Preview
    const preview = document.getElementById('receipt-preview');
    preview.innerHTML = `<img src="data:${mimeType};base64,${base64}" alt="領収書">`;

    // Analyze with Claude
    showLoading('🤖 AIが領収書を解析中...');
    let extracted = null;
    try {
      extracted = await analyzeReceiptWithClaude(base64, mimeType);
    } catch (e) {
      console.warn('Claude analysis failed:', e.message);
      showToast('AI解析に失敗しました。手動で入力してください', 'warning');
    }

    // Populate form
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
    showToast('画像処理エラー: ' + e.message, 'error');
  }
}

// Resize image on canvas to keep it under 1600px and ~1MB
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
        canvas.width  = width;
        canvas.height = height;
        canvas.getContext('2d').drawImage(img, 0, 0, width, height);

        const mimeType = file.type === 'image/png' ? 'image/png' : 'image/jpeg';
        canvas.toBlob((blob) => {
          const r2 = new FileReader();
          r2.onload  = (ev) => resolve({ base64: ev.target.result.split(',')[1], mimeType });
          r2.onerror = reject;
          r2.readAsDataURL(blob);
        }, mimeType, mimeType === 'image/jpeg' ? 0.82 : undefined);
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
async function analyzeReceiptWithClaude(base64, mimeType) {
  const prompt = `この領収書・レシートの画像から以下の情報をJSON形式のみで返してください（説明文・マークダウン不要）:
{
  "date": "YYYY-MM-DD形式の日付（不明な場合は今日の日付）",
  "amount": 税込合計金額（数値のみ、カンマなし）,
  "payee": "店名・支払先",
  "description": "購入内容・品目の概要（簡潔に）",
  "category": "交通費/接待費/消耗品費/会議費/通信費/出張費/書籍・資料費/その他 のいずれか一つ"
}
JSONのみを返してください。`;

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
          { type: 'text',  text: prompt },
        ],
      }],
    }),
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Claude API error ${res.status}`);
  }

  const data  = await res.json();
  const text  = data.content[0].text.trim();
  const match = text.match(/\{[\s\S]*\}/);
  if (!match) throw new Error('JSONが見つかりませんでした');
  return JSON.parse(match[0]);
}

// ============================================================
// Google Sheets API (REST, direct fetch)
// ============================================================
async function sheetsApi(method, path, body = null) {
  if (!accessToken) throw new Error('Googleにサインインしてください');

  const url  = `https://sheets.googleapis.com/v4/spreadsheets/${settings.spreadsheetId}${path}`;
  const opts = {
    method,
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
  };
  if (body !== null) opts.body = JSON.stringify(body);

  const res = await fetch(url, opts);
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    if (res.status === 401) {
      accessToken = null;
      document.getElementById('google-signin-btn').classList.remove('hidden');
      document.getElementById('google-signed-in').classList.add('hidden');
      throw new Error('Google認証の有効期限が切れました。再度ログインしてください');
    }
    throw new Error(err.error?.message || `Sheets API error ${res.status}`);
  }
  return res.json();
}

async function getSheetNames() {
  const info = await sheetsApi('GET', '?fields=sheets.properties.title');
  return (info.sheets || []).map(s => s.properties.title);
}

async function createSheet(title) {
  await sheetsApi('POST', ':batchUpdate', {
    requests: [{ addSheet: { properties: { title } } }],
  });
}

async function writeHeaders(sheetTitle) {
  const range = encodeRange(`${sheetTitle}!A1:H1`);
  await sheetsApi('PUT', `/values/${range}?valueInputOption=USER_ENTERED`, {
    values: [['日付', '金額（円）', '支払先', '内容', '仕訳項目', '備考', '申請者', '提出日時']],
  });
}

async function appendExpenseRow(sheetTitle, row) {
  const range = encodeRange(`${sheetTitle}!A:H`);
  await sheetsApi(
    'POST',
    `/values/${range}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    { values: [row] }
  );
}

async function getColumnValues(sheetTitle, col) {
  const range = encodeRange(`${sheetTitle}!${col}:${col}`);
  const result = await sheetsApi('GET', `/values/${range}`).catch(() => ({}));
  return result.values || [];
}

// Rebuild 集計 sheet from scratch
async function updateSummarySheet(sheetNames) {
  const personSheets = sheetNames.filter(n => n !== '集計');

  const rows = [['担当者', '件数', '合計金額（円）', '最終更新']];
  let totalCount = 0, totalAmount = 0;

  for (const name of personSheets) {
    try {
      const amounts = (await getColumnValues(name, 'B')).slice(1); // skip header
      const count   = amounts.length;
      const amount  = amounts.reduce((s, r) => s + (Number(r[0]) || 0), 0);
      totalCount  += count;
      totalAmount += amount;
      rows.push([name, count, amount, new Date().toLocaleDateString('ja-JP')]);
    } catch (e) {
      rows.push([name, 0, 0, 'エラー']);
    }
  }

  rows.push(['', '', '', '']);
  rows.push(['合計', totalCount, totalAmount, new Date().toLocaleDateString('ja-JP')]);

  // Ensure 集計 sheet exists
  if (!sheetNames.includes('集計')) await createSheet('集計');

  // Clear existing data
  const clearRange = encodeRange(`集計!A1:D${rows.length + 10}`);
  await sheetsApi('POST', `/values/${clearRange}:clear`, {});

  // Write new data
  const writeRange = encodeRange(`集計!A1:D${rows.length}`);
  await sheetsApi('PUT', `/values/${writeRange}?valueInputOption=USER_ENTERED`, { values: rows });
}

function encodeRange(range) {
  // Encode the sheet name part but keep ! and column letters unencoded
  const bang = range.indexOf('!');
  if (bang === -1) return encodeURIComponent(range);
  const sheet  = range.slice(0, bang);
  const cells  = range.slice(bang + 1);
  return encodeURIComponent(sheet) + '!' + cells;
}

// ============================================================
// Submit Expense
// ============================================================
async function submitExpense() {
  if (!accessToken) {
    showToast('先にGoogleアカウントと連携してください', 'error');
    return;
  }

  const date        = val('review-date');
  const amountStr   = val('review-amount');
  const payee       = val('review-payee');
  const description = val('review-description');
  const category    = document.getElementById('review-category').value;
  const notes       = val('review-notes');
  const person      = settings.name;

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
    let sheetNames = await getSheetNames();

    // Create person sheet if needed
    if (!sheetNames.includes(person)) {
      showLoading(`「${person}」シートを作成中...`);
      await createSheet(person);
      await writeHeaders(person);
      sheetNames = await getSheetNames();
    }

    // Append row
    const now = new Date().toLocaleString('ja-JP');
    await appendExpenseRow(person, [date, amount, payee, description, category, notes, person, now]);

    // Update summary
    showLoading('集計シートを更新中...');
    const latestNames = await getSheetNames();
    await updateSummarySheet(latestNames);

    // Local history
    addToHistory({ date, amount, payee, description, category, person });

    hideLoading();
    showSuccessPage({ date, amount, payee, description, category });
  } catch (e) {
    hideLoading();
    showToast('保存エラー: ' + e.message, 'error');
  }
}

// ============================================================
// Local History (last 20 items)
// ============================================================
function addToHistory(expense) {
  const history = getHistory();
  history.unshift(expense);
  localStorage.setItem('expense_history', JSON.stringify(history.slice(0, 20)));
}

function getHistory() {
  try { return JSON.parse(localStorage.getItem('expense_history') || '[]'); } catch (e) { return []; }
}

function renderRecentExpenses() {
  const history   = getHistory();
  const container = document.getElementById('recent-list');
  const badge     = document.getElementById('recent-count');

  badge.textContent = history.length + '件';

  if (history.length === 0) {
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
        <span class="expense-date">${esc(e.date)}</span>
        <span class="expense-category">${esc(e.category)}</span>
      </div>
    </div>
  `).join('');
}

// ============================================================
// UI Helpers
// ============================================================
function showPage(id) {
  document.querySelectorAll('.page').forEach(p => p.classList.add('hidden'));
  document.getElementById('page-' + id).classList.remove('hidden');
  window.scrollTo(0, 0);
  if (id === 'main') renderRecentExpenses();
}

function showSuccessPage(expense) {
  document.getElementById('success-details').innerHTML = `
    <div class="success-detail-row">
      <span>支払先</span><strong>${esc(expense.payee)}</strong>
    </div>
    <div class="success-detail-row">
      <span>金額</span><strong>¥${Number(expense.amount).toLocaleString()}</strong>
    </div>
    <div class="success-detail-row">
      <span>日付</span><strong>${esc(expense.date)}</strong>
    </div>
    <div class="success-detail-row">
      <span>仕訳項目</span><strong>${esc(expense.category)}</strong>
    </div>
  `;
  document.getElementById('sheet-link').href =
    `https://docs.google.com/spreadsheets/d/${settings.spreadsheetId}/edit`;
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
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = `toast toast-${type}`;
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.add('hidden'), 3500);
}

// ============================================================
// Utilities
// ============================================================
function val(id)        { return document.getElementById(id).value.trim(); }
function setVal(id, v)  { document.getElementById(id).value = v ?? ''; }
function esc(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
