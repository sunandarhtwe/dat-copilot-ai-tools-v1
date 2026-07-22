const documentsInput = document.getElementById('documents');
const sidebarToggle = document.getElementById('sidebarToggle');
const sidebar = document.getElementById('sidebar');
sidebarToggle.addEventListener('click', () => sidebar.classList.toggle('collapsed'));
const folderInput = document.getElementById('folderInput');
const fileList = document.getElementById('fileList');
const keywordEl = document.getElementById('keyword');
const functionNameEl = document.getElementById('functionName');
const analyzeBtn = document.getElementById('analyzeBtn');
const analyzeMessage = document.getElementById('analyzeMessage');
const testTypeGroup = document.getElementById('testTypeGroup');
const checklistCard = document.getElementById('checklistCard');
const checklistBody = document.getElementById('checklistBody');
const analysisSummary = document.getElementById('analysisSummary');
const analysisCoverageEl = document.getElementById('analysisCoverage');
const selectAllBtn = document.getElementById('selectAllBtn');
const selectNoneBtn = document.getElementById('selectNoneBtn');
const selectedCountEl = document.getElementById('selectedCount');
const generateBtn = document.getElementById('generateBtn');
const message = document.getElementById('message');
const caseBody = document.getElementById('caseBody');
const downloadBtn = document.getElementById('downloadBtn');
const log = document.getElementById('log');
const jaBtn = document.getElementById('jaBtn');
const enBtn = document.getElementById('enBtn');
const coveragePanel = document.getElementById('coveragePanel');
const coverageCards = document.getElementById('coverageCards');
const dirtyBadge = document.getElementById('dirtyBadge');
const saveEditBtn = document.getElementById('saveEditBtn');
const cancelEditBtn = document.getElementById('cancelEditBtn');
const addRowBtn = document.getElementById('addRowBtn');
const toggleExtraColsBtn = document.getElementById('toggleExtraColsBtn');
toggleExtraColsBtn.addEventListener('click', () => {
  showExtraColumns = !showExtraColumns;
  renderCases(latestCases, latestMeta);
});

// Core fields the AI is always asked for; any additional keys it returns
// (e.g. "userFlag") are appended after these, in first-seen order, both here
// and server-side (computeDynamicColumns in server.js mirrors this).
const CORE_FIELD_ORDER = ['no', 'category', 'testItem', 'precondition', 'steps', 'inputData', 'screenName', 'expectedResult', 'expectedScreenConfirmation', 'priority'];
// Mirrors server.js's GROUPED_FIELDS/buildColumnPlan — precondition/inputData
// can be a plain string OR an object of {label: value} pairs; when an
// object, they render as merged sub-columns instead of one crowded cell.
const GROUPED_FIELDS = ['precondition', 'inputData'];
function buildColumnPlan(cases) {
  const seen = new Set();
  const extra = [];
  const subKeySeen = { precondition: new Set(), inputData: new Set() };
  const subKeys = { precondition: [], inputData: [] };
  (cases || []).forEach(c => {
    Object.keys(c || {}).forEach(k => {
      if (!seen.has(k)) {
        seen.add(k);
        if (!CORE_FIELD_ORDER.includes(k)) extra.push(k);
      }
      if (GROUPED_FIELDS.includes(k) && c[k] && typeof c[k] === 'object' && !Array.isArray(c[k])) {
        Object.keys(c[k]).forEach(sk => {
          if (!subKeySeen[k].has(sk)) { subKeySeen[k].add(sk); subKeys[k].push(sk); }
        });
      }
    });
  });
  const orderedKeys = [...CORE_FIELD_ORDER.filter(k => seen.has(k)), ...extra];
  return orderedKeys.map(key => ({
    key,
    label: humanizeKey(key),
    subKeys: (GROUPED_FIELDS.includes(key) && subKeys[key].length) ? subKeys[key] : null
  }));
}
// Extra (non-core) fields the AI occasionally adds are hidden from the
// on-screen preview by default — end users just want the essential columns,
// not every ad-hoc technical field — but they're still kept in the data
// model and the Excel export always includes them in full.
let showExtraColumns = false;
function getVisibleColumns() {
  return showExtraColumns ? currentColumns : currentColumns.filter(c => CORE_FIELD_ORDER.includes(c));
}
function getVisiblePlan() {
  return showExtraColumns ? currentColumnPlan : currentColumnPlan.filter(g => CORE_FIELD_ORDER.includes(g.key));
}

const I18N = {
  ja: {
    appTitle: 'AI テストケース ジェネレーター', navGenerate: 'テストケース生成',
    uploadTitle: '設計書をアップロード',
    uploadFilesBtnTitle: 'ファイルを選択', uploadFilesBtnSub: '個別のファイルを選ぶ',
    uploadFolderBtnTitle: 'フォルダを選択', uploadFolderBtnSub: 'サブフォルダも含む',
    folderPickBtn: '📁 フォルダを選択（サブフォルダも含めて読み込みます）', clearFilesBtn: '🗑 ファイルをクリア', clearFilesConfirm: 'アップロード済みのファイルをすべて削除しますか？',
    step1: '資料アップロード', step2: '分析', step3: '生成',
    testTypeLabel: 'テストケース種別', testTypeIntegration: '結合テスト (Integration)', testTypeUnit: '単体テスト (Unit)', testTypeComprehensive: '総合テスト (Comprehensive)',
    keywordLabel: 'キーワード', keywordPlaceholder: '例：入力チェック、画面遷移、DB更新',
    functionNameLabel: '対象機能名（任意）', functionNamePlaceholder: '例：購入注文', metaFunctionName: '機能名',
    analyzeBtn: '🔍 Analysis（キーワードで資料を分析）', analyzing: '分析中...',
    checklistTitle: '分析結果 - 対象チェックポイントを選択', selectAllBtn: 'すべて選択', selectNoneBtn: '選択解除',
    checklistColCategory: 'カテゴリ', checklistColTitle: '項目', checklistColDetail: '詳細情報', checklistColExample: 'テストデータ例',
    filesReadTitle: '読み込んだファイル', rowsUnit: '行', categoryCountsTitle: 'カテゴリ別件数',
    chipFilterHint: '選択中のカテゴリのみ、下の一覧に表示されます', noCategorySelected: '表示するカテゴリを選択してください',
    filesCollapsedLabel: n => `${n}件のファイルを選択済み`, filesShowDetail: '詳細を表示', filesHideDetail: '詳細を閉じる',
    coverageOverall: '選択項目の反映率', coverageUncoveredTitle: '生成結果にまだ反映されていない項目',
    generateKeywordCoverageTitle: 'キーワード／機能名の反映率（生成結果ベース）',
    generateBtn: '✣ 選択項目でテストケースを生成', generating: '生成中...',
    previewTitle: '生成済みテストケース プレビュー', downloadBtn: '⬇ Excel ダウンロード',
    placeholder: 'ドキュメントをアップロードし、キーワードで分析してから、テストケースを生成してください。', logTitle: 'ログ', coverageTitle: 'カバレッジ（資料抽出候補との一致率）',
    executeTitle: '実行結果（Phase 2 予定）', executedSuites: '実行済みテストスイート', successRate: '成功率', playwrightPlan: 'Playwright 実行機能を追加予定',
    uploaded: n => `${n}件のドキュメントをアップロードしました。`, analyzingLog: 'キーワードで資料を分析中...',
    analysisCompleted: n => `分析完了：${n}件のチェックポイントが見つかりました。対象を選択して「テストケースを生成」をクリックしてください。`,
    generatingLog: 'テストケースを生成中...', generationCompleted: n => `Generation Completed\n${n}件のテストケースを生成しました。内容を確認の上、Excelとしてダウンロードしてください。`,
    metaKeyword: 'キーワード', metaDate: '作成日', metaAuthor: '作成者', metaTestType: 'テストケース種別', author: 'DAT-SONAR-FR', thPriority: 'Priority', thActions: '操作',
    unsavedBadge: '未保存の変更', saveBtn: '💾 保存', cancelBtn: 'キャンセル', saving: '保存中...', savedMsg: '変更を保存しました。', saveFailed: '保存に失敗しました', mustSaveFirst: '編集内容を保存してからダウンロードしてください。',
    translatingLog: '結果を翻訳中...', translateFailed: '翻訳に失敗しました', translateBlockedDirty: '言語を切り替える前に、編集内容を保存またはキャンセルしてください。',
    addRowBtn: '＋ 行を追加', deleteRowBtn: '削除', deleteRowConfirm: 'この行を削除しますか？',
    showExtraCols: '＋ 追加項目を表示', hideExtraCols: '－ 追加項目を非表示',
    selectedCountText: (n, total) => `${n} / ${total} 選択中`, needKeyword: 'キーワードを入力してください。', needSelection: 'チェックポイントを1つ以上選択してください。',
    analysisCoverageTitle: 'キーワード／機能名の反映率', analysisCoverageOverall: '反映率', analysisCoverageKeyword: 'キーワード', analysisCoverageFunctionName: '機能名', analysisCoverageUncovered: '資料内に見つからなかった用語', analysisCoverageEvidence: '反映根拠（どのチェックポイントに一致したか）'
  },
  en: {
    appTitle: 'AI Test Case Generator', navGenerate: 'Generate Test Cases',
    uploadTitle: 'Upload Design Documents',
    uploadFilesBtnTitle: 'Select Files', uploadFilesBtnSub: 'Choose individual files',
    uploadFolderBtnTitle: 'Select Folder', uploadFolderBtnSub: 'Includes sub-folders',
    folderPickBtn: '📁 Select a folder (sub-folders included)', clearFilesBtn: '🗑 Clear files', clearFilesConfirm: 'Remove all uploaded files?',
    step1: 'Upload Documents', step2: 'Analyze', step3: 'Generate',
    testTypeLabel: 'Test Case Type', testTypeIntegration: 'Integration Test', testTypeUnit: 'Unit Test', testTypeComprehensive: 'Comprehensive Test',
    keywordLabel: 'Keywords', keywordPlaceholder: 'e.g. input validation, screen transitions, DB updates',
    functionNameLabel: 'Target Function Name (optional)', functionNamePlaceholder: 'e.g. Purchase Order', metaFunctionName: 'Function Name',
    analyzeBtn: '🔍 Analysis (analyze documents by keyword)', analyzing: 'Analyzing...',
    checklistTitle: 'Analysis Results - Select Checkpoints to Cover', selectAllBtn: 'Select All', selectNoneBtn: 'Select None',
    checklistColCategory: 'Category', checklistColTitle: 'Item', checklistColDetail: 'Detail', checklistColExample: 'Example Test Data',
    filesReadTitle: 'Files Read', rowsUnit: 'rows', categoryCountsTitle: 'Checkpoints by Category',
    chipFilterHint: 'Only active categories are shown in the list below', noCategorySelected: 'Select at least one category to display',
    filesCollapsedLabel: n => `${n} file(s) selected`, filesShowDetail: 'Show details', filesHideDetail: 'Hide details',
    coverageOverall: 'Selected Checkpoints Covered', coverageUncoveredTitle: 'Not yet reflected in the generated cases',
    generateKeywordCoverageTitle: 'Keyword / Function Name Coverage (based on generated cases)',
    generateBtn: '✣ Generate Test Cases from Selection', generating: 'Generating...',
    previewTitle: 'Generated Test Case Preview', downloadBtn: '⬇ Download Excel',
    placeholder: 'Upload documents, analyze by keyword, then generate test cases.', logTitle: 'Log', coverageTitle: 'Coverage (match rate vs. extracted candidates)',
    executeTitle: 'Execution Result (Phase 2 Plan)', executedSuites: 'Executed Test Suites', successRate: 'Success Rate', playwrightPlan: 'Playwright execution feature will be added',
    uploaded: n => `Uploaded ${n} document(s).`, analyzingLog: 'Analyzing documents by keyword...',
    analysisCompleted: n => `Analysis complete: found ${n} checkpoint(s). Select the ones to cover and click Generate.`,
    generatingLog: 'Generating test cases...', generationCompleted: n => `Generation Completed\n${n} test cases were generated. Please review the content and download it as Excel.`,
    metaKeyword: 'Keyword', metaDate: 'Created Date', metaAuthor: 'Author', metaTestType: 'Test Case Type', author: 'DAT-SONAR-FR', thPriority: 'Priority', thActions: 'Actions',
    unsavedBadge: 'Unsaved changes', saveBtn: '💾 Save', cancelBtn: 'Cancel', saving: 'Saving...', savedMsg: 'Changes saved.', saveFailed: 'Save failed', mustSaveFirst: 'Please save your edits before downloading.',
    translatingLog: 'Translating results...', translateFailed: 'Translation failed', translateBlockedDirty: 'Please save or cancel your edits before switching language.',
    addRowBtn: '+ Add Row', deleteRowBtn: 'Delete', deleteRowConfirm: 'Delete this row?',
    showExtraCols: '+ Show extra columns', hideExtraCols: '− Hide extra columns',
    selectedCountText: (n, total) => `${n} / ${total} selected`, needKeyword: 'Please enter a keyword.', needSelection: 'Please select at least one checkpoint.',
    analysisCoverageTitle: 'Keyword / Function Name Coverage', analysisCoverageOverall: 'Covered', analysisCoverageKeyword: 'Keywords', analysisCoverageFunctionName: 'Function Name', analysisCoverageUncovered: 'Terms not found in the documents', analysisCoverageEvidence: 'Match evidence (which checkpoint each term matched)'
  }
};
let currentLang = localStorage.getItem('dat_ai_lang') || 'ja';

// Analyze-step state
let currentChecklist = []; // full checkpoint objects returned by /api/analyze

// Reflects actual progress (files uploaded / analyzed / generated) in the
// step indicator at the top of the screen — this used to be three hardcoded
// "active" steps that never changed regardless of what stage you were
// actually at.
function updateStepIndicator() {
  const step1 = document.getElementById('step1');
  const step2 = document.getElementById('step2');
  const step3 = document.getElementById('step3');
  [step1, step2, step3].forEach(s => s.classList.remove('active', 'done'));

  const hasFiles = !fileList.classList.contains('empty');
  const hasChecklist = currentChecklist.length > 0;
  const hasCases = latestCases.length > 0;

  if (hasCases) {
    step1.classList.add('done'); step2.classList.add('done'); step3.classList.add('active');
  } else if (hasChecklist || hasFiles) {
    step1.classList.add('done'); step2.classList.add('active');
  } else {
    step1.classList.add('active');
  }
}
let checklistById = new Map();
// Selection now lives in JS state, not just in checkbox DOM nodes, because
// clicking a category chip re-renders the table with only matching rows —
// if selection were DOM-only, filtering would silently lose your picks in
// whatever categories just got hidden.
let selectedIds = new Set();
let activeCategories = new Set(); // categories currently shown in the table; empty Set = "not yet loaded"

// Preview/result state
let latestCases = [];
let latestMeta = {};
let savedCases = [];
let isDirty = false;
let currentColumns = []; // dynamic column key list, computed from latestCases
let currentColumnPlan = []; // same, but with grouped-field sub-key detail for rendering

// Translation cache (per-language snapshots), same pattern as before
let resultsLang = 'ja';
let caseTranslations = {};

function t(key, ...args) { const v = I18N[currentLang][key]; return typeof v === 'function' ? v(...args) : (v || key); }

function applyLanguage(lang) {
  if (isDirty) latestCases = collectCasesFromTable();
  currentLang = lang; localStorage.setItem('dat_ai_lang', lang); document.documentElement.lang = lang;
  jaBtn.classList.toggle('active', lang === 'ja'); enBtn.classList.toggle('active', lang === 'en');
  document.querySelectorAll('[data-i18n]').forEach(el => { const v = t(el.dataset.i18n); if (typeof v === 'string') el.textContent = v; });
  document.querySelectorAll('[data-i18n-placeholder]').forEach(el => { const v = t(el.dataset.i18nPlaceholder); if (typeof v === 'string') el.placeholder = v; });
  if (!latestCases.length) renderPlaceholder(); else renderCases(latestCases, latestMeta);
  if (!analyzeBtn.disabled) analyzeBtn.textContent = t('analyzeBtn');
  if (!generateBtn.disabled || currentChecklist.length) generateBtn.textContent = t('generateBtn');
  updateDirtyUi();
}
jaBtn.addEventListener('click', () => switchUiAndResultsLanguage('ja'));
enBtn.addEventListener('click', () => switchUiAndResultsLanguage('en'));

async function switchUiAndResultsLanguage(lang) {
  applyLanguage(lang);
  if (!latestCases.length) return;
  if (isDirty) { showMessage('err', t('translateBlockedDirty')); return; }
  await translateAndRender(lang, { silent: false });
}

async function syncServerBaseline(cases) {
  try {
    const syncRes = await fetch('/api/update-cases', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ cases })
    });
    const syncData = await syncRes.json();
    if (syncRes.ok && syncData.ok) { downloadBtn.href = syncData.excelUrl; downloadBtn.classList.remove('disabled'); }
  } catch (_) { /* not fatal */ }
}

async function translateAndRender(targetLang, { silent = false } = {}) {
  if (targetLang === resultsLang) { renderCases(latestCases, latestMeta); return; }
  const cached = caseTranslations[targetLang];
  if (cached) {
    latestCases = cached.cases.map(c => ({ ...c }));
    savedCases = latestCases.map(c => ({ ...c }));
    latestMeta = { ...latestMeta, keyword: cached.meta.keyword };
    resultsLang = targetLang;
    renderCases(latestCases, latestMeta);
    setDirty(false);
    await syncServerBaseline(latestCases);
    return;
  }
  if (!silent) { clearMessage(); showMessage('ok', t('translatingLog')); }
  try {
    const res = await fetch('/api/translate-cases', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ cases: latestCases, meta: { keyword: latestMeta.keyword }, targetLang })
    });
    const data = await res.json(); if (!res.ok || !data.ok) throw new Error(data.message || 'Translate failed');
    caseTranslations[targetLang] = { cases: data.cases, meta: data.meta };
    latestCases = data.cases.map(c => ({ ...c }));
    savedCases = latestCases.map(c => ({ ...c }));
    latestMeta = { ...latestMeta, keyword: data.meta.keyword };
    resultsLang = targetLang;
    renderCases(latestCases, latestMeta);
    setDirty(false);
    if (!silent) clearMessage();
    await syncServerBaseline(latestCases);
  } catch (e) {
    if (!silent) showMessage('err', `${t('translateFailed')}: ${e.message || e}`);
  }
}

function switchView(view) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  const el = document.getElementById(`${view}View`);
  if (el) el.classList.add('active');
  document.querySelectorAll('.nav-item').forEach(n => n.classList.toggle('active', n.dataset.view === view));
}
document.querySelectorAll('.nav-item').forEach(btn => btn.addEventListener('click', () => switchView(btn.dataset.view)));

function selectedTestType() {
  const checked = testTypeGroup ? testTypeGroup.querySelector('input[name="testType"]:checked') : null;
  return checked ? checked.value : 'integration';
}

// --- Upload (files or a whole folder, sub-folders included) ----------------
async function uploadFileArray(files) {
  if (!files.length) return;
  const form = new FormData();
  files.forEach(f => {
    form.append('documents', f);
    form.append('relPaths', f.webkitRelativePath || f.name);
  });
  try {
    const res = await fetch('/api/upload', { method: 'POST', body: form });
    const data = await res.json();
    if (!res.ok || !data.ok) throw new Error(data.message || 'Upload failed');
    renderFileList(data.files || []);
    setLog(t('uploaded', (data.files || []).length));
  } catch (e) {
    // Upload feedback (including failures — e.g. a folder exceeding the
    // file-count limit) stays in the quiet log panel rather than the
    // prominent colored message box, which is reserved for Analyze/Generate
    // results. A big red alert box popping up mid-folder-upload reads as
    // more alarming than warranted, especially since partial uploads still
    // succeed.
    setLog(`Upload error: ${e.message || e}`);
  }
}

documentsInput.addEventListener('change', () => { uploadFileArray([...documentsInput.files]); documentsInput.value = ''; });
folderInput.addEventListener('change', () => { uploadFileArray([...folderInput.files]); folderInput.value = ''; });

function renderFileList(files) {
  if (!files.length) { fileList.innerHTML = ''; fileList.classList.add('empty'); updateStepIndicator(); return; }
  fileList.classList.remove('empty');
  // Collapsed by default — showing every uploaded file (which can be dozens
  // for a folder upload) up front was cluttered. Just the count, with an
  // expandable "show details" toggle for when the full list is actually
  // wanted (e.g. to confirm a specific file made it in, or to Clear).
  fileList.innerHTML = `
    <div class="file-summary" id="fileSummary">
      <div class="file-summary-head" id="fileSummaryHead">
        <span>${escapeHtml(t('filesCollapsedLabel', files.length))}</span>
        <span class="file-summary-toggle">
          <span class="file-summary-toggle-label">${escapeHtml(t('filesShowDetail'))}</span>
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><polyline points="6 9 12 15 18 9"/></svg>
        </span>
      </div>
      <div class="file-list-detail">
        <div class="file-list-header"><span>${files.length}</span><button id="clearFilesBtn" type="button">${escapeHtml(t('clearFilesBtn'))}</button></div>
        ${files.map(f => `<div class="file-item"><div class="file-icon">${iconFor(f.originalName)}</div><div><div class="file-name">${f.folder ? `<span class="file-folder">📁 ${escapeHtml(f.folder)}/</span>` : ''}${escapeHtml(f.originalName)}</div><div class="file-size">${formatSize(f.size)}</div></div><div class="check">✓</div></div>`).join('')}
      </div>
    </div>`;

  const head = document.getElementById('fileSummaryHead');
  const summary = document.getElementById('fileSummary');
  const toggleLabel = summary.querySelector('.file-summary-toggle-label');
  head.addEventListener('click', () => {
    const isOpen = summary.classList.toggle('open');
    toggleLabel.textContent = isOpen ? t('filesHideDetail') : t('filesShowDetail');
  });

  const clearBtn = document.getElementById('clearFilesBtn');
  if (clearBtn) clearBtn.addEventListener('click', async (e) => {
    e.stopPropagation();
    if (!confirm(t('clearFilesConfirm'))) return;
    await fetch('/api/clear-files', { method: 'POST' });
    renderFileList([]);
  });
  updateStepIndicator();
}

// --- Step: Analyze -----------------------------------------------------------
analyzeBtn.addEventListener('click', async () => {
  const keyword = keywordEl.value.trim();
  const functionName = functionNameEl.value.trim();
  if (!keyword) { showMessage2(analyzeMessage, 'err', t('needKeyword')); return; }
  clearMessage2(analyzeMessage);
  resetChecklist();
  clearPreview();
  analyzeBtn.disabled = true; const original = analyzeBtn.textContent; analyzeBtn.textContent = t('analyzing');
  setLog(t('analyzingLog'));
  try {
    const res = await fetch('/api/analyze', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ keyword, functionName, testType: selectedTestType() })
    });
    const data = await res.json(); if (!res.ok || !data.ok) throw new Error(data.message || 'Analyze failed');
    selectedIds = new Set(); activeCategories = new Set(); // fresh analysis — don't carry over stale selection state
    renderAnalysisSummary(data.parsedFiles || [], data.categoryCounts || []);
    renderAnalysisCoverage(data.analysisCoverage);
    renderChecklist(data.checklist || [], { isFreshLoad: true });
    checklistCard.classList.remove('hidden');
    showMessage2(analyzeMessage, 'ok', t('analysisCompleted', currentChecklist.length));
    setLog([`Source: ${data.source}`, ...(data.log || [])].join('\n'));
  } catch (e) {
    showMessage2(analyzeMessage, 'err', e.message || String(e));
    setLog(e.stack || String(e));
  } finally {
    analyzeBtn.disabled = false; analyzeBtn.textContent = original;
  }
});

function resetChecklist() {
  currentChecklist = []; checklistById = new Map();
  selectedIds = new Set(); activeCategories = new Set();
  checklistBody.innerHTML = '';
  checklistCard.classList.add('hidden');
  analysisSummary.innerHTML = '';
  analysisSummary.classList.add('hidden');
  analysisCoverageEl.innerHTML = '';
  analysisCoverageEl.classList.add('hidden');
  generateBtn.disabled = true;
  updateSelectedCount();
  updateStepIndicator();
}

// Shows what was actually read (file list with row counts), and how many
// checkpoints were found per category as CLICKABLE FILTER CHIPS — clicking
// a chip toggles that category in/out of the checklist table below (all
// categories active by default). Multiple categories can be active at once.
function renderAnalysisSummary(parsedFiles, categoryCounts) {
  if (!parsedFiles.length && !categoryCounts.length) { analysisSummary.classList.add('hidden'); return; }
  activeCategories = new Set(categoryCounts.map(c => c.category)); // all active by default

  const filesHtml = parsedFiles.length ? `
    <div class="analysis-summary-block">
      <div class="analysis-summary-title">${escapeHtml(t('filesReadTitle'))}</div>
      <ul class="analysis-file-list">
        ${parsedFiles.map(f => `<li>${f.ok ? '✓' : '✕'} ${escapeHtml(f.fileName)}${f.ok ? ` <span class="analysis-file-meta">(${f.rows} ${escapeHtml(t('rowsUnit'))}${f.documentType ? ', ' + escapeHtml(f.documentType) : ''})</span>` : ` <span class="analysis-file-meta err">${escapeHtml(f.error || '')}</span>`}</li>`).join('')}
      </ul>
    </div>` : '';
  const countsHtml = categoryCounts.length ? `
    <div class="analysis-summary-block">
      <div class="analysis-summary-title">${escapeHtml(t('categoryCountsTitle'))}</div>
      <div class="analysis-count-chips" id="categoryChipRow"></div>
      <div class="chip-hint">${escapeHtml(t('chipFilterHint'))}</div>
    </div>` : '';
  analysisSummary.innerHTML = filesHtml + countsHtml;
  analysisSummary.classList.remove('hidden');

  if (categoryCounts.length) renderCategoryChips(categoryCounts);
}

// Shows how many of the individual terms in Keywords + Function Name were
// actually found somewhere in the extracted checklist — a sanity check on
// the Analyze step itself (separate from the post-Generate coverage, which
// checks selected checkpoints against the generated cases instead).
function renderAnalysisCoverage(coverage) {
  if (!coverage || !coverage.totalTerms) { analysisCoverageEl.classList.add('hidden'); return; }
  const pct = v => v == null ? 'N/A' : `${v}%`;
  const byKeyword = coverage.items.filter(i => i.source === 'keyword');
  const byFunctionName = coverage.items.filter(i => i.source === 'functionName');
  const cards = [`<div><strong>${pct(coverage.coveragePercent)}</strong><span>${escapeHtml(t('analysisCoverageOverall'))} (${coverage.coveredTerms}/${coverage.totalTerms})</span></div>`];
  if (byKeyword.length) {
    const c = byKeyword.filter(i => i.covered).length;
    cards.push(`<div><strong>${pct(Math.round((c / byKeyword.length) * 100))}</strong><span>${escapeHtml(t('analysisCoverageKeyword'))} (${c}/${byKeyword.length})</span></div>`);
  }
  if (byFunctionName.length) {
    const c = byFunctionName.filter(i => i.covered).length;
    cards.push(`<div><strong>${pct(Math.round((c / byFunctionName.length) * 100))}</strong><span>${escapeHtml(t('analysisCoverageFunctionName'))} (${c}/${byFunctionName.length})</span></div>`);
  }
  const uncovered = coverage.items.filter(i => !i.covered);
  const uncoveredHtml = uncovered.length
    ? `<div class="coverage-uncovered"><b>${escapeHtml(t('analysisCoverageUncovered'))}</b><ul>${uncovered.map(i => `<li>${escapeHtml(i.term)}</li>`).join('')}</ul></div>`
    : '';
  // Evidence trail for covered terms — shows WHICH checkpoint each term
  // actually matched (plus a rough match strength), so the percentage
  // above isn't just a number to take on faith; it can be spot-checked.
  const coveredWithEvidence = coverage.items.filter(i => i.covered && i.matchedCheckpointTitle);
  const evidenceHtml = coveredWithEvidence.length
    ? `<div class="coverage-evidence"><b>${escapeHtml(t('analysisCoverageEvidence'))}</b><ul>${coveredWithEvidence.map(i =>
        `<li><span class="evidence-term">${escapeHtml(i.term)}</span> → ${escapeHtml(i.matchedCheckpointTitle)}${i.matchedCheckpointCategory ? ` <span class="evidence-cat">(${escapeHtml(i.matchedCheckpointCategory)})</span>` : ''} <span class="evidence-score">${i.matchScore}%</span></li>`
      ).join('')}</ul></div>`
    : '';
  analysisCoverageEl.innerHTML = `<div class="analysis-summary-title">${escapeHtml(t('analysisCoverageTitle'))}</div><div class="result-cards">${cards.join('')}</div>${evidenceHtml}${uncoveredHtml}`;
  analysisCoverageEl.classList.remove('hidden');
}

function renderCategoryChips(categoryCounts) {
  const row = document.getElementById('categoryChipRow');
  if (!row) return;
  row.innerHTML = categoryCounts.map(c => {
    const isActive = activeCategories.has(c.category);
    const label = currentLang === 'en' ? c.categoryLabelEn : c.categoryLabel;
    return `<button type="button" class="chip-filter ${isActive ? 'active' : ''}" data-cat="${escapeHtml(c.category)}">
      <span class="chip-dot"></span>${escapeHtml(label)} <b>${c.count}</b>
    </button>`;
  }).join('');
  row.querySelectorAll('.chip-filter').forEach(btn => {
    btn.addEventListener('click', () => {
      const cat = btn.dataset.cat;
      if (activeCategories.has(cat)) activeCategories.delete(cat); else activeCategories.add(cat);
      btn.classList.toggle('active');
      renderChecklist(currentChecklist);
    });
  });
}

function renderChecklist(checklist, { isFreshLoad = false } = {}) {
  currentChecklist = checklist;
  updateStepIndicator();
  checklistById = new Map(checklist.map(c => [c.id, c]));
  // Only initialize "everything selected" on an actual fresh Analyze result
  // (isFreshLoad: true, passed explicitly by the caller) — NOT whenever
  // selectedIds happens to be empty, which is indistinguishable from the
  // user deliberately clicking "Select None" and was the bug: clicking
  // Select None emptied selectedIds, then this re-render silently refilled
  // it right back to "all selected".
  if (isFreshLoad) {
    selectedIds = new Set(checklist.map(c => c.id));
    activeCategories = new Set(checklist.map(c => c.category));
  }

  const visible = checklist.filter(cp => activeCategories.has(cp.category));
  if (!visible.length) {
    checklistBody.innerHTML = `<tr class="checklist-empty-row"><td colspan="5">${escapeHtml(checklist.length ? t('noCategorySelected') : t('placeholder'))}</td></tr>`;
    updateSelectedCount();
    generateBtn.disabled = getSelectedCheckpoints().length === 0;
    return;
  }
  const groups = new Map();
  visible.forEach(cp => {
    const key = cp.categoryLabel || cp.category;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(cp);
  });
  const rowsHtml = [...groups.entries()].map(([catLabel, items]) => items.map((cp, i) => `
    <tr>
      <td class="ct-check-cell"><input type="checkbox" class="checklist-checkbox" data-id="${cp.id}" ${selectedIds.has(cp.id) ? 'checked' : ''} /></td>
      ${i === 0 ? `<td class="ct-category-cell" rowspan="${items.length}">${escapeHtml(catLabel)}</td>` : ''}
      <td class="ct-title-cell">${escapeHtml(cp.title || '')}</td>
      <td class="ct-detail-cell">${escapeHtml(cp.detail || '')}${cp.source ? `<span class="ct-source">${escapeHtml(cp.source)}</span>` : ''}</td>
      <td class="ct-example-cell">${cp.example ? escapeHtml(cp.example) : '<span class="ct-example-empty">—</span>'}</td>
    </tr>
  `).join('')).join('');
  checklistBody.innerHTML = rowsHtml;
  updateSelectedCount();
  generateBtn.disabled = getSelectedCheckpoints().length === 0;
}

checklistBody.addEventListener('change', (e) => {
  if (e.target.classList.contains('checklist-checkbox')) {
    const id = e.target.dataset.id;
    if (e.target.checked) selectedIds.add(id); else selectedIds.delete(id);
    updateSelectedCount();
    generateBtn.disabled = getSelectedCheckpoints().length === 0;
  }
});

// Select All / Select None act on whatever is currently VISIBLE (i.e.
// respects the active category filter) — not the full checklist — so
// narrowing to one category and clicking "Select All" doesn't silently
// re-select things in categories you'd deliberately filtered out.
selectAllBtn.addEventListener('click', () => {
  currentChecklist.filter(cp => activeCategories.has(cp.category)).forEach(cp => selectedIds.add(cp.id));
  renderChecklist(currentChecklist);
});
selectNoneBtn.addEventListener('click', () => {
  currentChecklist.filter(cp => activeCategories.has(cp.category)).forEach(cp => selectedIds.delete(cp.id));
  renderChecklist(currentChecklist);
});

function updateSelectedCount() {
  selectedCountEl.textContent = t('selectedCountText', selectedIds.size, currentChecklist.length);
}

function getSelectedCheckpoints() {
  return [...selectedIds].map(id => checklistById.get(id)).filter(Boolean);
}

// --- Step: Generate (from selected checkpoints) -----------------------------
generateBtn.addEventListener('click', async () => {
  const selected = getSelectedCheckpoints();
  if (!selected.length) { showMessage('err', t('needSelection')); return; }
  clearPreview();
  try { await fetch('/api/clear', { method: 'POST' }); } catch (_) { }
  clearMessage(); setGenerating(true); setLog(t('generatingLog'));
  try {
    const res = await fetch('/api/generate-from-selection', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ selectedCheckpoints: selected, keyword: keywordEl.value.trim(), functionName: functionNameEl.value.trim(), testType: selectedTestType() })
    });
    const data = await res.json(); if (!res.ok || !data.ok) throw new Error(data.message || 'Generate failed');
    latestCases = data.cases || []; latestMeta = data.meta || {};
    updateStepIndicator();
    savedCases = latestCases.map(c => ({ ...c }));
    resultsLang = 'ja';
    caseTranslations = { ja: { cases: savedCases.map(c => ({ ...c })), meta: { keyword: latestMeta.keyword } } };
    renderCases(latestCases, latestMeta);
    setDirty(false);
    downloadBtn.href = data.excelUrl; downloadBtn.classList.remove('disabled');
    renderCoverage(data.coverage);
    showMessage('ok', t('generationCompleted', latestCases.length));
    setLog([`Source: ${data.source}`, ...(data.log || [])].join('\n'));
    if (currentLang === 'en') await translateAndRender('en', { silent: true });
  } catch (e) {
    showMessage('err', e.message || String(e)); setLog(e.stack || String(e));
  } finally { setGenerating(false); }
});

function setGenerating(v) {
  generateBtn.disabled = v || getSelectedCheckpoints().length === 0;
  generateBtn.textContent = v ? t('generating') : t('generateBtn');
}

// --- Save / Cancel / Add Row / Delete Row -----------------------------------
saveEditBtn.addEventListener('click', async () => {
  const cases = collectCasesFromTable();
  saveEditBtn.disabled = true; const original = saveEditBtn.textContent; saveEditBtn.textContent = t('saving');
  try {
    const res = await fetch('/api/update-cases', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ cases })
    });
    const data = await res.json(); if (!res.ok || !data.ok) throw new Error(data.message || 'Save failed');
    latestCases = data.cases; savedCases = data.cases.map(c => ({ ...c }));
    renderCases(latestCases, latestMeta);
    downloadBtn.href = data.excelUrl; downloadBtn.classList.remove('disabled');
    setDirty(false);
    caseTranslations = { [resultsLang]: { cases: savedCases.map(c => ({ ...c })), meta: { keyword: latestMeta.keyword } } };
    showMessage('ok', t('savedMsg'));
  } catch (e) {
    showMessage('err', `${t('saveFailed')}: ${e.message || e}`);
  } finally {
    saveEditBtn.disabled = false; saveEditBtn.textContent = original;
  }
});

cancelEditBtn.addEventListener('click', () => {
  latestCases = savedCases.map(c => ({ ...c }));
  renderCases(latestCases, latestMeta);
  setDirty(false);
  clearMessage();
});

addRowBtn.addEventListener('click', () => {
  if (!latestCases.length && !caseBody.querySelector('tr[data-case-row]')) return;
  const cases = collectCasesFromTable();
  const blank = {}; currentColumns.forEach(col => { blank[col] = ''; });
  cases.push(blank);
  renumberCases(cases);
  latestCases = cases;
  renderCases(latestCases, latestMeta);
  setDirty(true);
});

caseBody.addEventListener('click', (e) => {
  const btn = e.target.closest('.delete-row-btn');
  if (!btn) return;
  if (!confirm(t('deleteRowConfirm'))) return;
  const idx = Number(btn.dataset.rowIndex);
  const cases = collectCasesFromTable();
  cases.splice(idx, 1);
  renumberCases(cases);
  latestCases = cases;
  renderCases(latestCases, latestMeta);
  setDirty(true);
});

function renumberCases(cases) {
  cases.forEach((c, i) => { c.no = `TC${String(i + 1).padStart(3, '0')}`; });
  return cases;
}

function clearPreview() {
  latestCases = []; latestMeta = {}; savedCases = []; currentColumns = []; currentColumnPlan = [];
  resultsLang = 'ja'; caseTranslations = {};
  renderPlaceholder();
  clearMessage();
  downloadBtn.href = '#'; downloadBtn.classList.add('disabled');
  if (coveragePanel) coveragePanel.classList.add('hidden');
  if (coverageCards) coverageCards.innerHTML = '';
  const coverageUncoveredEl = document.getElementById('coverageUncovered');
  if (coverageUncoveredEl) coverageUncoveredEl.innerHTML = '';
  setDirty(false);
  updateStepIndicator();
}

function renderPlaceholder() {
  caseBody.innerHTML = `<tr><td class="placeholder">${escapeHtml(t('placeholder'))}</td></tr>`;
}

function renderCoverage(coverage) {
  if (!coverage || !coveragePanel || !coverageCards) { if (coveragePanel) coveragePanel.classList.add('hidden'); return; }
  const pct = v => v == null ? 'N/A' : `${v}%`;
  const cards = [`<div><strong>${pct(coverage.coveragePercent)}</strong><span>${escapeHtml(t('coverageOverall'))} (${coverage.totalCovered}/${coverage.totalSelected})</span></div>`];
  const byCategory = new Map();
  (coverage.items || []).forEach(i => {
    const key = i.category || '';
    if (!byCategory.has(key)) byCategory.set(key, { total: 0, covered: 0 });
    const c = byCategory.get(key); c.total++; if (i.covered) c.covered++;
  });
  byCategory.forEach((c, cat) => {
    cards.push(`<div><strong>${pct(Math.round((c.covered / c.total) * 100))}</strong><span>${escapeHtml(cat)} (${c.covered}/${c.total})</span></div>`);
  });
  coverageCards.innerHTML = cards.join('');

  const uncovered = (coverage.items || []).filter(i => !i.covered);
  const coverageUncovered = document.getElementById('coverageUncovered');
  coverageUncovered.innerHTML = uncovered.length
    ? `<div class="coverage-uncovered"><b>${escapeHtml(t('coverageUncoveredTitle'))}</b><ul>${uncovered.map(i => `<li>${escapeHtml(i.title)}</li>`).join('')}</ul></div>`
    : '';

  // Keyword vs. Function Name term coverage, scored against the actual
  // generated cases — kept as a visibly separate block from the
  // checkpoint-level coverage above, since the two answer different
  // questions ("did my selections make it in" vs "did my search terms
  // themselves make it in", split by which input each term came from).
  const kfc = coverage.keywordFunctionCoverage;
  const kfcEl = document.getElementById('generateKeywordCoverage');
  if (kfcEl) {
    if (kfc && kfc.totalTerms) {
      const kfcCards = [];
      if (kfc.keyword.total) kfcCards.push(`<div><strong>${pct(kfc.keyword.coveragePercent)}</strong><span>${escapeHtml(t('analysisCoverageKeyword'))} (${kfc.keyword.covered}/${kfc.keyword.total})</span></div>`);
      if (kfc.functionName.total) kfcCards.push(`<div><strong>${pct(kfc.functionName.coveragePercent)}</strong><span>${escapeHtml(t('analysisCoverageFunctionName'))} (${kfc.functionName.covered}/${kfc.functionName.total})</span></div>`);
      const kfcUncovered = kfc.items.filter(i => !i.covered);
      const kfcUncoveredHtml = kfcUncovered.length
        ? `<div class="coverage-uncovered"><b>${escapeHtml(t('analysisCoverageUncovered'))}</b><ul>${kfcUncovered.map(i => `<li>${escapeHtml(i.term)}</li>`).join('')}</ul></div>` : '';
      kfcEl.innerHTML = `<div class="analysis-summary-title">${escapeHtml(t('generateKeywordCoverageTitle'))}</div><div class="result-cards">${kfcCards.join('')}</div>${kfcUncoveredHtml}`;
      kfcEl.classList.remove('hidden');
    } else {
      kfcEl.classList.add('hidden');
    }
  }
  coveragePanel.classList.remove('hidden');
}

function iconFor(name = '') { const n = name.toLowerCase(); if (n.endsWith('.xlsx') || n.endsWith('.xls')) return '📗'; if (n.endsWith('.docx') || n.endsWith('.doc')) return '📘'; if (n.endsWith('.pdf')) return '📕'; return '📄'; }
function formatSize(bytes = 0) { if (bytes < 1024) return `${bytes} B`; if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`; return `${(bytes / 1024 / 1024).toFixed(2)} MB`; }

// --- Dynamic-column result table --------------------------------------------
// Columns are computed from the union of keys actually present across the
// generated cases (core fields first in reading order, then any extra
// AI-added fields like "userFlag" in first-seen order) — never a fixed
// schema, mirroring computeDynamicColumns() in server.js.
function computeColumns(cases) {
  return buildColumnPlan(cases).map(g => g.key);
}

function humanizeKey(k) {
  if (k === 'no') return 'No';
  if (k === 'testItem') return 'Test Item';
  if (k === 'inputData') return 'Input Data';
  if (k === 'expectedResult') return 'Expected Result';
  if (k === 'expectedScreenConfirmation') return 'Expected Screen Confirmation';
  if (k === 'screenName') return 'Screen Name';
  if (k === 'precondition') return 'Precondition';
  if (k === 'category') return 'Category';
  if (k === 'steps') return 'Steps';
  if (k === 'priority') return t('thPriority');
  const spaced = String(k).replace(/([a-z0-9])([A-Z])/g, '$1 $2');
  return spaced.charAt(0).toUpperCase() + spaced.slice(1);
}

function renderCases(cases, meta = {}) {
  currentColumns = computeColumns(cases);
  currentColumnPlan = buildColumnPlan(cases);
  const metaColspan = Math.max(1, currentColumns.length);

  // Meta rows (keyword / test type / date / author) must appear FIRST,
  // then a blank spacer, then the column-title row(s), then the case rows —
  // all as plain <tr> inside the one <tbody>. (A separate <thead> was tried
  // earlier for the header row, but browsers always render <thead> before
  // <tbody> regardless of source order, which put column titles above the
  // keyword/test type/date/author rows — the opposite of what's intended.)
  const metaHtml = [
    [t('metaKeyword'), meta.keyword || keywordEl.value],
    ...(meta.functionName || functionNameEl.value.trim() ? [[t('metaFunctionName'), meta.functionName || functionNameEl.value]] : []),
    [t('metaTestType'), meta.testTypeLabel || ''],
    [t('metaDate'), meta.date || new Date().toISOString().slice(0, 10)],
    [t('metaAuthor'), t('author')]
  ].map(([k, v]) => `<tr><td class="meta-key">${escapeHtml(k)}</td><td class="meta-val" colspan="${metaColspan}">${escapeHtml(v)}</td></tr>`).join('');
  const blankHtml = `<tr><td colspan="${metaColspan + 1}"></td></tr>`;
  const visiblePlan = getVisiblePlan();

  // Two header rows for grouped fields (precondition/inputData when they
  // have sub-columns): row 1 has the group label spanning its sub-columns
  // (or spanning DOWN across both rows for plain single columns, so they
  // don't leave an empty cell underneath); row 2 has the sub-column labels.
  const hasAnyGroup = visiblePlan.some(g => g.subKeys);
  let headerHtml;
  if (hasAnyGroup) {
    const row1 = visiblePlan.map(g => g.subKeys
      ? `<th colspan="${g.subKeys.length}">${escapeHtml(g.label)}</th>`
      : `<th rowspan="2">${escapeHtml(g.label)}</th>`
    ).join('') + `<th class="h-actions" rowspan="2">${escapeHtml(t('thActions'))}</th>`;
    const row2 = visiblePlan.filter(g => g.subKeys).map(g => g.subKeys.map(sk => `<th class="sub-col-th">${escapeHtml(sk)}</th>`).join('')).join('');
    headerHtml = `<tr class="col-header-row">${row1}</tr><tr class="col-header-row sub-col-row">${row2}</tr>`;
  } else {
    headerHtml = `<tr class="col-header-row">${visiblePlan.map(g => `<th>${escapeHtml(g.label)}</th>`).join('')}<th class="h-actions">${escapeHtml(t('thActions'))}</th></tr>`;
  }

  function groupedCellValue(c, g, subKeyIndex) {
    const raw = c[g.key];
    if (raw && typeof raw === 'object' && !Array.isArray(raw)) return raw[g.subKeys[subKeyIndex]] ?? '';
    return subKeyIndex === 0 ? (raw ?? '') : '';
  }

  const rowsHtml = cases.map((c, idx) => (
    `<tr data-case-row="${idx}">` +
    visiblePlan.map(g => g.subKeys
      ? g.subKeys.map((sk, i) => `<td class="editable" data-field="${g.key}::${escapeHtml(sk)}" contenteditable="true">${escapeHtml(groupedCellValue(c, g, i))}</td>`).join('')
      : `<td class="editable" data-field="${g.key}" contenteditable="true">${escapeHtml((c[g.key] && typeof c[g.key] === 'object') ? Object.entries(c[g.key]).map(([k, v]) => `${k}: ${v}`).join(' / ') : c[g.key])}</td>`
    ).join('') +
    `<td class="actions-cell"><button type="button" class="delete-row-btn" data-row-index="${idx}" title="${escapeHtml(t('deleteRowBtn'))}">🗑</button></td>` +
    `</tr>`
  )).join('');

  caseBody.innerHTML = metaHtml + blankHtml + headerHtml + rowsHtml;
  addRowBtn.classList.toggle('hidden', !cases.length);
  const hasExtraColumns = currentColumns.some(c => !CORE_FIELD_ORDER.includes(c));
  toggleExtraColsBtn.classList.toggle('hidden', !hasExtraColumns);
  toggleExtraColsBtn.textContent = showExtraColumns ? t('hideExtraCols') : t('showExtraCols');
}

function collectCasesFromTable() {
  const rows = [...caseBody.querySelectorAll('tr[data-case-row]')];
  return rows.map(tr => {
    const idx = Number(tr.dataset.caseRow);
    // Start from the last-known full record (so columns hidden by the
    // extra-columns toggle aren't silently dropped — they just don't have
    // a cell to read a live edit from) then overlay whatever is actually
    // visible and editable in the DOM right now.
    const obj = { ...(latestCases[idx] || {}) };
    currentColumnPlan.forEach(g => {
      if (g.subKeys) {
        // Rebuild the nested object from its sub-column cells. Preserve
        // any sub-keys that exist in the original data but currently have
        // no visible cell (e.g. hidden by the extra-columns toggle — not
        // applicable to core grouped fields today, but kept consistent).
        const sub = { ...((obj[g.key] && typeof obj[g.key] === 'object') ? obj[g.key] : {}) };
        g.subKeys.forEach(sk => {
          const cell = tr.querySelector(`td[data-field="${g.key}::${sk}"]`);
          if (cell) sub[sk] = cell.textContent.trim();
        });
        obj[g.key] = sub;
      } else {
        const cell = tr.querySelector(`td[data-field="${g.key}"]`);
        if (cell) obj[g.key] = cell.textContent.trim();
      }
    });
    return obj;
  });
}

caseBody.addEventListener('input', (e) => {
  if (e.target && e.target.classList && e.target.classList.contains('editable')) setDirty(true);
});

function setDirty(v) { isDirty = v; updateDirtyUi(); }

function updateDirtyUi() {
  dirtyBadge.classList.toggle('hidden', !isDirty);
  saveEditBtn.classList.toggle('hidden', !isDirty);
  cancelEditBtn.classList.toggle('hidden', !isDirty);
  if (isDirty) {
    downloadBtn.classList.add('disabled');
    downloadBtn.title = t('mustSaveFirst');
  } else if (latestCases.length) {
    downloadBtn.classList.remove('disabled');
    downloadBtn.title = '';
  }
}

downloadBtn.addEventListener('click', (e) => {
  if (isDirty || downloadBtn.classList.contains('disabled')) {
    e.preventDefault();
    showMessage('err', t('mustSaveFirst'));
  }
});

function showMessage(type, text) { message.className = `message ${type}`; message.textContent = text; }
function clearMessage() { message.className = 'message hidden'; message.textContent = ''; }
function showMessage2(el, type, text) { el.className = `message ${type}`; el.textContent = text; }
function clearMessage2(el) { el.className = 'message hidden'; el.textContent = ''; }
function setLog(text) { log.textContent = text; }
function escapeHtml(s = '') { return String(s).replace(/[&<>'"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[c])); }

renderPlaceholder();
applyLanguage(currentLang);
updateStepIndicator();
