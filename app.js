(() => {
  'use strict';

  const STORAGE_KEY = 'biddingQuoteApp.v1';

  const state = {
    files: [],
    items: [],
    model: '',
    extractedText: '',
    form: {
      docTitle: '見積書',
      quoteNumber: '',
      quoteDate: '',
      validUntil: '',
      clientName: '',
      projectTitle: '',
      taxRate: 10,
      companyName: '',
      companyPerson: '',
      companyAddress: '',
      notes: ''
    }
  };

  let nextItemId = 1;
  let nextFileId = 1;

  const el = (id) => document.getElementById(id);

  const fileInput = el('fileInput');
  const fileListEl = el('fileList');
  const extractTextBtn = el('extractTextBtn');
  const extractStatus = el('extractStatus');
  const extractedTextArea = el('extractedText');

  const modelSelect = el('modelSelect');
  const aiExtractBtn = el('aiExtractBtn');
  const aiStatus = el('aiStatus');

  const itemsTableBody = el('itemsTableBody');
  const addItemBtn = el('addItemBtn');
  const exportCsvBtn = el('exportCsvBtn');
  const subtotalDisplay = el('subtotalDisplay');
  const taxRateDisplay = el('taxRateDisplay');
  const taxDisplay = el('taxDisplay');
  const grandTotalDisplay = el('grandTotalDisplay');

  const docTitleInput = el('docTitleInput');
  const quoteNumberInput = el('quoteNumberInput');
  const quoteDateInput = el('quoteDateInput');
  const validUntilInput = el('validUntilInput');
  const clientNameInput = el('clientNameInput');
  const projectTitleInput = el('projectTitleInput');
  const taxRateInput = el('taxRateInput');
  const companyNameInput = el('companyNameInput');
  const companyPersonInput = el('companyPersonInput');
  const companyAddressInput = el('companyAddressInput');
  const notesInput = el('notesInput');

  const generateQuoteBtn = el('generateQuoteBtn');
  const printQuoteBtn = el('printQuoteBtn');
  const resetAllBtn = el('resetAllBtn');
  const quotePreview = el('quotePreview');

  function formatNumber(n) {
    const v = Number.isFinite(n) ? n : 0;
    return v.toLocaleString('ja-JP', { maximumFractionDigits: 0 });
  }

  function setStatus(node, message, type) {
    node.textContent = message || '';
    node.className = 'status-msg' + (type ? ' ' + type : '');
  }

  function saveState() {
    const toSave = {
      items: state.items,
      model: state.model,
      extractedText: state.extractedText,
      form: state.form
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(toSave));
  }

  function loadState() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed.items)) {
        state.items = parsed.items;
        nextItemId = state.items.reduce((max, it) => Math.max(max, it.id + 1), 1);
      }
      if (typeof parsed.model === 'string' && parsed.model) state.model = parsed.model;
      if (typeof parsed.extractedText === 'string') state.extractedText = parsed.extractedText;
      if (parsed.form) Object.assign(state.form, parsed.form);
    } catch (e) {
      console.error('failed to load saved state', e);
    }
  }

  // ---- Step 1: file upload & text extraction ----

  function renderFileList() {
    fileListEl.innerHTML = '';
    state.files.forEach((entry) => {
      const li = document.createElement('li');
      const label = document.createElement('span');
      label.textContent = entry.file.name;
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'secondary';
      removeBtn.textContent = '削除';
      removeBtn.addEventListener('click', () => {
        state.files = state.files.filter((f) => f.id !== entry.id);
        renderFileList();
      });
      li.appendChild(label);
      li.appendChild(removeBtn);
      fileListEl.appendChild(li);
    });
  }

  fileInput.addEventListener('change', () => {
    Array.from(fileInput.files).forEach((file) => {
      state.files.push({ id: nextFileId++, file });
    });
    fileInput.value = '';
    renderFileList();
  });

  async function extractPdfText(file) {
    if (typeof window.pdfjsLib === 'undefined') {
      throw new Error('PDF読み込みライブラリの読み込みに失敗しました。ネットワーク接続をご確認の上、再読み込みしてください。');
    }
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
    const buffer = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: buffer }).promise;
    let text = '';
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      text += content.items.map((it) => it.str).join(' ') + '\n';
    }
    return text;
  }

  async function extractDocxText(file) {
    if (typeof window.mammoth === 'undefined') {
      throw new Error('Word読み込みライブラリの読み込みに失敗しました。ネットワーク接続をご確認の上、再読み込みしてください。');
    }
    const buffer = await file.arrayBuffer();
    const result = await window.mammoth.extractRawText({ arrayBuffer: buffer });
    return result.value;
  }

  async function extractExcelText(file) {
    if (typeof window.XLSX === 'undefined') {
      throw new Error('Excel読み込みライブラリの読み込みに失敗しました。ネットワーク接続をご確認の上、再読み込みしてください。');
    }
    const buffer = await file.arrayBuffer();
    const wb = window.XLSX.read(buffer, { type: 'array' });
    let text = '';
    wb.SheetNames.forEach((name) => {
      text += `--- ${name} ---\n`;
      text += window.XLSX.utils.sheet_to_csv(wb.Sheets[name]) + '\n';
    });
    return text;
  }

  async function extractTextFromFile(entry) {
    const file = entry.file;
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'pdf') return extractPdfText(file);
    if (ext === 'docx') return extractDocxText(file);
    if (ext === 'xlsx' || ext === 'xls') return extractExcelText(file);
    if (ext === 'txt' || ext === 'csv') return file.text();
    if (ext === 'doc') throw new Error('.doc形式（旧Word）には対応していません。.docx形式で保存し直してください。');
    throw new Error(`未対応のファイル形式です: .${ext}`);
  }

  extractTextBtn.addEventListener('click', async () => {
    if (state.files.length === 0) {
      setStatus(extractStatus, 'ファイルを選択してください。', 'error');
      return;
    }
    extractTextBtn.disabled = true;
    setStatus(extractStatus, 'テキストを抽出しています...', 'loading');
    const parts = [];
    for (const entry of state.files) {
      try {
        const text = await extractTextFromFile(entry);
        parts.push(`===== ${entry.file.name} =====\n${text.trim()}`);
      } catch (e) {
        parts.push(`===== ${entry.file.name} =====\n[抽出エラー: ${e.message}]`);
      }
    }
    extractedTextArea.value = parts.join('\n\n');
    state.extractedText = extractedTextArea.value;
    saveState();
    extractTextBtn.disabled = false;
    setStatus(extractStatus, 'テキストの抽出が完了しました。内容を確認してください。', 'success');
  });

  extractedTextArea.addEventListener('input', () => {
    state.extractedText = extractedTextArea.value;
    saveState();
  });

  // ---- Step 2: AI item extraction (runs fully in-browser via WebLLM, no API key) ----

  const WEBLLM_CDN_URL = 'https://cdn.jsdelivr.net/npm/@mlc-ai/web-llm/+esm';
  const PREFERRED_MODEL_SUBSTRINGS = [
    'Qwen2.5-1.5B-Instruct',
    'Qwen2.5-3B-Instruct',
    'Llama-3.2-3B-Instruct',
    'Phi-3.5-mini-instruct',
    'gemma-2-2b-it',
    'Qwen2.5-7B-Instruct',
    'Llama-3.1-8B-Instruct'
  ];

  let webllmModule = null;
  let engine = null;
  let loadedModelId = null;

  modelSelect.addEventListener('change', () => {
    state.model = modelSelect.value;
    saveState();
  });

  function populateModelSelect(modelList) {
    const chatModels = modelList.filter((m) => !/embed/i.test(m.model_id));
    chatModels.sort((a, b) => {
      const rank = (m) => {
        const idx = PREFERRED_MODEL_SUBSTRINGS.findIndex((s) => m.model_id.includes(s));
        return idx === -1 ? PREFERRED_MODEL_SUBSTRINGS.length : idx;
      };
      return rank(a) - rank(b) || a.model_id.localeCompare(b.model_id);
    });
    modelSelect.innerHTML = '';
    chatModels.forEach((m) => {
      const option = document.createElement('option');
      option.value = m.model_id;
      option.textContent = m.model_id;
      modelSelect.appendChild(option);
    });
    const preferredMatch = chatModels.find((m) => m.model_id === state.model);
    if (preferredMatch) {
      modelSelect.value = state.model;
    } else if (chatModels.length > 0) {
      modelSelect.value = chatModels[0].model_id;
      state.model = modelSelect.value;
    }
    modelSelect.disabled = false;
    aiExtractBtn.disabled = false;
  }

  async function ensureWebLLMLoaded() {
    if (webllmModule) return webllmModule;
    webllmModule = await import(WEBLLM_CDN_URL);
    return webllmModule;
  }

  if (!navigator.gpu) {
    setStatus(
      aiStatus,
      'このブラウザはWebGPUに対応していないため、ブラウザ内蔵AIを利用できません。Google ChromeまたはMicrosoft Edgeの最新版でお試しください。',
      'error'
    );
    modelSelect.innerHTML = '<option value="">利用不可</option>';
  } else {
    ensureWebLLMLoaded()
      .then((mod) => populateModelSelect(mod.prebuiltAppConfig.model_list))
      .catch((e) => {
        setStatus(aiStatus, `AIライブラリの読み込みに失敗しました: ${e.message}`, 'error');
        modelSelect.innerHTML = '<option value="">読み込み失敗</option>';
      });
  }

  async function ensureEngineForModel(modelId, onProgress) {
    if (engine && loadedModelId === modelId) return engine;
    const mod = await ensureWebLLMLoaded();
    engine = await mod.CreateMLCEngine(modelId, { initProgressCallback: onProgress });
    loadedModelId = modelId;
    return engine;
  }

  function buildExtractionPrompt(specText) {
    return `あなたは入札案件の仕様書を読み取り、購入すべき物品を抽出する専門家です。
以下の仕様書テキストから、購入・調達すべき物品（機器・資材・ソフトウェア等）を全て抽出してください。

出力は次のJSON配列のみとしてください。説明文、前置き、コードブロック記号(\`\`\`)は一切含めないでください。
[
  {"name": "品名", "model": "型番（不明な場合は空文字）", "quantity": 数量(数値。不明な場合は1), "unit": "単位（式・個・台など。不明な場合は空文字）"}
]

仕様書テキスト:
"""
${specText}
"""`;
  }

  function parseItemsFromAIText(raw) {
    let cleaned = raw.trim();
    cleaned = cleaned.replace(/^```(json)?/i, '').replace(/```$/, '').trim();
    const start = cleaned.indexOf('[');
    const end = cleaned.lastIndexOf(']');
    if (start === -1 || end === -1 || end < start) {
      throw new Error('AIの応答からJSONを読み取れませんでした。');
    }
    const jsonSlice = cleaned.slice(start, end + 1);
    const parsed = JSON.parse(jsonSlice);
    if (!Array.isArray(parsed)) throw new Error('AIの応答が配列形式ではありません。');
    return parsed;
  }

  async function runLocalLLM(prompt) {
    const modelId = modelSelect.value;
    if (!modelId) throw new Error('AIモデルを選択してください。');
    const activeEngine = await ensureEngineForModel(modelId, (report) => {
      const pct = Math.round((report.progress || 0) * 100);
      setStatus(aiStatus, report.text ? report.text : `モデルを読み込んでいます... ${pct}%`, 'loading');
    });
    setStatus(aiStatus, 'AIが仕様書を解析しています...', 'loading');
    const reply = await activeEngine.chat.completions.create({
      messages: [{ role: 'user', content: prompt }],
      max_tokens: 4096
    });
    const text = reply.choices && reply.choices[0] && reply.choices[0].message && reply.choices[0].message.content;
    if (!text) throw new Error('AIの応答にテキストが含まれていませんでした。');
    return text;
  }

  aiExtractBtn.addEventListener('click', async () => {
    const specText = extractedTextArea.value.trim();
    if (!specText) {
      setStatus(aiStatus, '先に仕様書のテキストを抽出または貼り付けてください。', 'error');
      return;
    }
    const MAX_CHARS = 8000;
    const truncated = specText.length > MAX_CHARS;
    const textForPrompt = truncated ? specText.slice(0, MAX_CHARS) : specText;

    aiExtractBtn.disabled = true;
    setStatus(aiStatus, 'AIモデルを準備しています（初回はダウンロードに時間がかかります）...', 'loading');
    try {
      const responseText = await runLocalLLM(buildExtractionPrompt(textForPrompt));
      const parsedItems = parseItemsFromAIText(responseText);
      const newItems = parsedItems.map((it) => ({
        id: nextItemId++,
        name: (it.name || '').toString(),
        model: (it.model || '').toString(),
        quantity: Number(it.quantity) || 1,
        unit: (it.unit || '').toString(),
        unitPrice: 0
      }));
      state.items = state.items.concat(newItems);
      renderItemsTable();
      saveState();
      setStatus(
        aiStatus,
        `${newItems.length}件の物品を抽出しました。${truncated ? '（テキストが長いため一部のみを解析しました）' : ''}`,
        'success'
      );
    } catch (e) {
      setStatus(aiStatus, e.message, 'error');
    } finally {
      aiExtractBtn.disabled = false;
    }
  });

  // ---- Step 3: item list & pricing ----

  function computeTotals() {
    const subtotal = state.items.reduce((sum, it) => sum + (Number(it.quantity) || 0) * (Number(it.unitPrice) || 0), 0);
    const taxRate = Number(state.form.taxRate) || 0;
    const tax = Math.floor(subtotal * (taxRate / 100));
    const total = subtotal + tax;
    return { subtotal, tax, total, taxRate };
  }

  function renderTotals() {
    const { subtotal, tax, total, taxRate } = computeTotals();
    subtotalDisplay.textContent = formatNumber(subtotal);
    taxRateDisplay.textContent = formatNumber(taxRate);
    taxDisplay.textContent = formatNumber(tax);
    grandTotalDisplay.textContent = formatNumber(total);
  }

  function renderItemsTable() {
    itemsTableBody.innerHTML = '';
    state.items.forEach((item) => itemsTableBody.appendChild(createItemRow(item)));
    renderTotals();
  }

  function createItemRow(item) {
    const tr = document.createElement('tr');

    const amountTd = document.createElement('td');
    amountTd.className = 'amount';

    function refreshAmount() {
      const amount = (Number(item.quantity) || 0) * (Number(item.unitPrice) || 0);
      amountTd.textContent = formatNumber(amount);
      renderTotals();
    }

    tr.appendChild(makeTextCell(item, 'name', 'text', refreshAmount));
    tr.appendChild(makeTextCell(item, 'model', 'text', refreshAmount));
    tr.appendChild(makeTextCell(item, 'quantity', 'number', refreshAmount));
    tr.appendChild(makeTextCell(item, 'unit', 'text', refreshAmount));
    tr.appendChild(makeTextCell(item, 'unitPrice', 'number', refreshAmount));

    refreshAmount();
    tr.appendChild(amountTd);

    const actionTd = document.createElement('td');
    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.className = 'danger';
    delBtn.textContent = '削除';
    delBtn.addEventListener('click', () => {
      state.items = state.items.filter((it) => it.id !== item.id);
      tr.remove();
      renderTotals();
      saveState();
    });
    actionTd.appendChild(delBtn);
    tr.appendChild(actionTd);

    return tr;
  }

  function makeTextCell(item, field, type, onChange) {
    const td = document.createElement('td');
    if (type === 'number') td.className = 'num';
    const input = document.createElement('input');
    input.type = 'text';
    if (type === 'number') input.className = 'num-input';
    input.value = item[field];
    input.addEventListener('input', () => {
      if (type === 'number') {
        item[field] = input.value === '' ? '' : Number(input.value.replace(/[^0-9.\-]/g, '')) || 0;
      } else {
        item[field] = input.value;
      }
      onChange();
      saveState();
    });
    td.appendChild(input);
    return td;
  }

  addItemBtn.addEventListener('click', () => {
    const item = { id: nextItemId++, name: '', model: '', quantity: 1, unit: '', unitPrice: 0 };
    state.items.push(item);
    itemsTableBody.appendChild(createItemRow(item));
    renderTotals();
    saveState();
  });

  exportCsvBtn.addEventListener('click', () => {
    const header = ['品名', '型番', '数量', '単位', '単価', '金額'];
    const rows = state.items.map((it) => {
      const amount = (Number(it.quantity) || 0) * (Number(it.unitPrice) || 0);
      return [it.name, it.model, it.quantity, it.unit, it.unitPrice, amount];
    });
    const csvLines = [header, ...rows].map((row) =>
      row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(',')
    );
    const csvContent = '﻿' + csvLines.join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '物品リスト.csv';
    a.click();
    URL.revokeObjectURL(url);
  });

  // ---- Step 4: quote form & generation ----

  function bindFormField(input, key, parser) {
    input.value = state.form[key] != null ? state.form[key] : '';
    input.addEventListener('input', () => {
      state.form[key] = parser ? parser(input.value) : input.value;
      if (key === 'taxRate') renderTotals();
      saveState();
    });
  }

  bindFormField(docTitleInput, 'docTitle');
  bindFormField(quoteNumberInput, 'quoteNumber');
  bindFormField(quoteDateInput, 'quoteDate');
  bindFormField(validUntilInput, 'validUntil');
  bindFormField(clientNameInput, 'clientName');
  bindFormField(projectTitleInput, 'projectTitle');
  bindFormField(taxRateInput, 'taxRate', (v) => Number(v) || 0);
  bindFormField(companyNameInput, 'companyName');
  bindFormField(companyPersonInput, 'companyPerson');
  bindFormField(companyAddressInput, 'companyAddress');
  bindFormField(notesInput, 'notes');

  function escapeHtml(str) {
    return String(str == null ? '' : str).replace(/[&<>"']/g, (c) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }

  function generateQuoteHtml() {
    const { subtotal, tax, total, taxRate } = computeTotals();
    const f = state.form;

    const rows = state.items.map((it) => {
      const amount = (Number(it.quantity) || 0) * (Number(it.unitPrice) || 0);
      return `<tr>
        <td>${escapeHtml(it.name)}</td>
        <td>${escapeHtml(it.model)}</td>
        <td class="num">${formatNumber(it.quantity)}</td>
        <td>${escapeHtml(it.unit)}</td>
        <td class="num">${formatNumber(it.unitPrice)}</td>
        <td class="num">${formatNumber(amount)}</td>
      </tr>`;
    }).join('');

    return `
      <div class="quote-doc-title">${escapeHtml(f.docTitle || '見積書')}</div>
      <div class="quote-top-row">
        <div class="quote-to">
          <div class="name">${escapeHtml(f.clientName || '　')} 御中</div>
          <div class="quote-meta" style="margin-top:8px;">件名: ${escapeHtml(f.projectTitle)}</div>
        </div>
        <div class="quote-issuer">${escapeHtml(f.companyName)}
${escapeHtml(f.companyAddress)}
${f.companyPerson ? '担当: ' + escapeHtml(f.companyPerson) : ''}</div>
      </div>
      <div class="quote-meta">
        見積番号: ${escapeHtml(f.quoteNumber)}　発行日: ${escapeHtml(f.quoteDate)}${f.validUntil ? '　有効期限: ' + escapeHtml(f.validUntil) : ''}
      </div>
      <div class="quote-total-banner">御見積金額　${formatNumber(total)} 円（税込）</div>
      <table class="quote-doc-table">
        <thead>
          <tr><th>品名</th><th>型番</th><th>数量</th><th>単位</th><th>単価（円）</th><th>金額（円）</th></tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
      <table class="quote-doc-totals">
        <tr><td class="label">小計</td><td class="val">${formatNumber(subtotal)} 円</td></tr>
        <tr><td class="label">消費税（${formatNumber(taxRate)}%）</td><td class="val">${formatNumber(tax)} 円</td></tr>
        <tr><td class="label">合計</td><td class="val">${formatNumber(total)} 円</td></tr>
      </table>
      ${f.notes ? `<div class="quote-notes">備考:\n${escapeHtml(f.notes)}</div>` : ''}
    `;
  }

  generateQuoteBtn.addEventListener('click', () => {
    if (state.items.length === 0) {
      alert('物品リストが空です。STEP2またはSTEP3で物品を追加してください。');
      return;
    }
    quotePreview.innerHTML = generateQuoteHtml();
    quotePreview.classList.add('visible');
    quotePreview.scrollIntoView({ behavior: 'smooth' });
  });

  printQuoteBtn.addEventListener('click', () => {
    if (!quotePreview.classList.contains('visible')) {
      alert('先に「見積書を作成」を押してください。');
      return;
    }
    window.print();
  });

  resetAllBtn.addEventListener('click', () => {
    if (!confirm('入力した全ての内容を削除します。よろしいですか？')) return;
    localStorage.removeItem(STORAGE_KEY);
    window.location.reload();
  });

  // ---- init ----

  function init() {
    loadState();
    renderFileList();
    extractedTextArea.value = state.extractedText;
    renderItemsTable();
    quoteDateInput.value = state.form.quoteDate || new Date().toISOString().slice(0, 10);
    state.form.quoteDate = quoteDateInput.value;
  }

  init();
})();
