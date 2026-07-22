const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

loadDotEnv(path.join(__dirname, '.env'));

let XLSX = null;
try { XLSX = require('xlsx'); } catch (_) { XLSX = null; }

const app = express();
const PORT = process.env.PORT || 3000;
// --- Copilot Bridge configuration ------------------------------------------
// This server no longer calls OpenAI directly. Instead it talks to the local
// "DAT Copilot Bridge" VS Code extension (see /copilot-bridge folder), which
// runs inside VS Code, uses the vscode.lm API to reach the signed-in user's
// GitHub Copilot subscription, and exposes a small local HTTP endpoint.
// The extension MUST be running (VS Code open, extension activated) before
// clicking "Generate Test Cases", otherwise generation falls back to the
// document-based (non-AI) engine.
const COPILOT_BRIDGE_URL = process.env.COPILOT_BRIDGE_URL || 'http://127.0.0.1:4321/generate';
const COPILOT_MODEL_FAMILY = process.env.COPILOT_MODEL_FAMILY || ''; // e.g. 'gpt-4o', 'claude-3.5-sonnet' — empty = bridge default
const COPILOT_TIMEOUT_MS = Number(process.env.COPILOT_TIMEOUT_MS || 120000);
// --- Copilot prompt size budget ---------------------------------------------
// GitHub Copilot's chat models (reached via vscode.lm inside the bridge
// extension) have a much smaller per-message context window than a direct
// OpenAI API call — sending the same ~150k-character prompts that worked
// fine against gpt-4o's API used to work now fails with
// "Message exceeds token limit". These caps keep the combined prompt small
// enough to fit reliably. All are configurable via .env if a given Copilot
// model turns out to tolerate more (or less).
const COPILOT_MAX_DOC_CHARS = Number(process.env.COPILOT_MAX_DOC_CHARS || 9000);
const COPILOT_MAX_SAMPLE_CHARS = Number(process.env.COPILOT_MAX_SAMPLE_CHARS || 1200);
const COPILOT_MAX_KNOWLEDGE_CHARS = Number(process.env.COPILOT_MAX_KNOWLEDGE_CHARS || 5500);
const COPILOT_MAX_ANALYSIS_CHARS = Number(process.env.COPILOT_MAX_ANALYSIS_CHARS || 5000);
// Hard final safety cap applied to the fully-assembled prompt right before
// it's sent to the bridge, regardless of how the pieces above add up.
// Calibrated from a real observed failure: this account's Copilot model
// caps input at 12,078 tokens, and a 55,000-char prompt measured at 25,726
// tokens (~2.14 chars/token for this Japanese-heavy content). 20,000 chars
// lands around ~9,300 tokens, leaving headroom below the 12,078 limit.
const COPILOT_MAX_TOTAL_PROMPT_CHARS = Number(process.env.COPILOT_MAX_TOTAL_PROMPT_CHARS || 20000);
// ---------------------------------------------------------------------------

const UPLOAD_DIR = path.join(__dirname, 'uploads');
const GENERATED_DIR = path.join(__dirname, 'generated');
fs.mkdirSync(UPLOAD_DIR, { recursive: true });
fs.mkdirSync(GENERATED_DIR, { recursive: true });

app.use(cors());
app.use(express.json({ limit: '30mb' }));
app.use(express.static(path.join(__dirname, 'public')));
app.use('/generated', express.static(GENERATED_DIR));

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => {
    const safe = decodeFileName(file.originalname).replace(/[\\/:*?"<>|]/g, '_');
    cb(null, `${Date.now()}_${safe}`);
  }
});
const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(decodeFileName(file.originalname)).toLowerCase();
    const ok = ['.xlsx', '.xls', '.pdf', '.docx', '.doc', '.pptx', '.txt', '.md', '.csv'].includes(ext);
    // Folder uploads routinely include unrelated files (images, .git, thumbs.db,
    // media embedded in Office docs, etc). Passing an Error here would abort
    // the ENTIRE upload over a single unsupported file — instead, just skip
    // that one file (cb(null, false)) and keep the rest.
    cb(null, ok);
  },
  // `files` here must match (or exceed) the maxCount passed to
  // upload.array('documents', 200) below — multer enforces whichever is
  // LOWER, so a mismatch here silently caps every upload at this number
  // regardless of what the route allows. This was previously set to 20,
  // which broke any real folder upload with more than 20 files.
  limits: { fileSize: 80 * 1024 * 1024, files: 200 }
});
let uploadedFiles = [];
let latestCases = [];
let latestSavedCases = []; // last Saved/generated snapshot used by Cancel to revert edits
let latestMeta = {};
let latestCoverage = null; // Coverage Engine result for the most recent generation
let latestAnalysis = null; // most recent /api/analyze result (checklist + raw smartAnalysis)
let latestKnowledge = null; // deterministic extraction from the most recent /api/analyze, reused by Coverage Engine at generate time

// --- Test Case Type profiles ------------------------------------------------
// ② User selects one of these before Generate. Each profile changes wording
// injected into both AI prompts (analysis + case generation) so Copilot
// produces the right granularity/scope of test case.
const TEST_TYPE_PROFILES = {
  integration: {
    label: '結合テスト (Integration Test)',
    analysisHint: '画面間・機能間の連携、DB更新、外部/内部インターフェース呼び出し、業務フロー全体の整合性を重視してください。',
    caseHint: '結合テストとして、画面遷移・DB反映・他機能との連携・業務フロー全体を通したテストケースを作成してください。単一関数の内部ロジックだけを検証するテストは作らないでください。'
  },
  unit: {
    label: '単体テスト (Unit Test)',
    analysisHint: '個々の入力項目・関数・バリデーションロジック単位での境界値・異常値・条件分岐を重視してください。画面遷移やDB間連携などの結合観点は最小限にしてください。',
    caseHint: '単体テストとして、個々の入力項目・関数・条件分岐単位の境界値／異常値／同値分割テストケースを作成してください。ケースは1機能・1項目・1条件にできるだけ絞り込み、複数画面をまたぐ結合シナリオは作らないでください。'
  },
  comprehensive: {
    label: '総合テスト (Comprehensive Test)',
    analysisHint: '単体レベルの入力チェックから結合レベルの画面遷移・DB連携、さらに異常系・非機能観点（性能・排他制御・権限）まで幅広く整理してください。',
    caseHint: '総合テストとして、単体レベルの入力チェックから結合レベルの画面遷移・DB連携、業務シナリオ全体、異常系、権限/排他制御まで幅広く網羅するテストケースを作成してください。'
  }
};
// ---------------------------------------------------------------------------

// --- V2.2: In-memory Knowledge Base Cache ---------------------------------
// Caches parsed rows per file so re-generating with the same uploaded files
// does not re-parse the document every time. Cleared on server restart
// (in-memory only by design) and on graceful shutdown.
const knowledgeBaseCache = new Map(); // key: `${path}::${size}::${mtimeMs}` -> { rows, classification, cachedAt }

function cacheKeyFor(filePath) {
  try {
    const st = fs.statSync(filePath);
    return `${filePath}::${st.size}::${st.mtimeMs}`;
  } catch (_) {
    return `${filePath}::unknown`;
  }
}

function getCachedRows(filePath) {
  const key = cacheKeyFor(filePath);
  const hit = knowledgeBaseCache.get(key);
  return hit ? hit : null;
}

function setCachedRows(filePath, rows, classification) {
  const key = cacheKeyFor(filePath);
  knowledgeBaseCache.set(key, { rows, classification, cachedAt: Date.now() });
}

function clearKnowledgeBaseCache() {
  const size = knowledgeBaseCache.size;
  knowledgeBaseCache.clear();
  return size;
}
// ---------------------------------------------------------------------------

// --- V2.2: Document Classifier --------------------------------------------
// Classifies an uploaded design document into a category based on its file
// name and the text content extracted from it. Pure heuristic (no AI call)
// so it works even when the OpenAI API is not configured. The result is
// surfaced in the generation log AND fed into the AI prompt as context so
// the model can weight evidence according to the kind of document it came
// from (e.g. DB definition rows vs. screen design rows vs. error tables).
const DOCUMENT_TYPE_RULES = [
  { type: '基本設計書', re: /(基本設計書|概要設計書|基本設計)/ },
  { type: '詳細設計書', re: /(詳細設計書|詳細設計|機能設計書)/ },
  { type: 'DB定義書', re: /(DB定義書|テーブル定義書|テーブル仕様書|ER図|データ定義)/ },
  { type: '画面設計書', re: /(画面設計書|画面仕様書|画面遷移図|UI設計)/ },
  { type: 'エラーメッセージ一覧', re: /(エラーメッセージ一覧|メッセージ一覧|エラーコード一覧)/ },
  { type: 'テストケース（参考）', re: /(テストケース|テスト仕様書|試験項目)/ },
  { type: '提案書／その他資料', re: /(提案書|proposal|議事録|手順書)/i }
];

function classifyDocument(originalName, rows) {
  const sampleText = (rows || []).slice(0, 80).map(r => r.text).join(' ');
  const haystack = `${originalName} ${sampleText}`;

  const scored = DOCUMENT_TYPE_RULES.map(rule => ({
    type: rule.type,
    score: (haystack.match(rule.re) || []).length
  })).filter(r => r.score > 0);

  scored.sort((a, b) => b.score - a.score);

  // Secondary signal: row-shape heuristics when filename/content gave no hit.
  let type = scored.length ? scored[0].type : null;
  if (!type) {
    const dbLike = (rows || []).filter(r => /(DB|テーブル|カラム|レコード|型|PK|FK)/.test(r.text)).length;
    const screenLike = (rows || []).filter(r => /(画面|ボタン|メニュー|遷移)/.test(r.text)).length;
    const errorLike = (rows || []).filter(r => /(エラー|メッセージ|E\d{3,}|\b\d{4,5}\b)/.test(r.text)).length;
    const best = [['DB定義書', dbLike], ['画面設計書', screenLike], ['エラーメッセージ一覧', errorLike]]
      .sort((a, b) => b[1] - a[1])[0];
    type = best && best[1] > 0 ? best[0] : '一般資料';
  }

  const confidence = scored.length
    ? Math.min(1, scored[0].score / 5)
    : 0.3;

  return { fileName: originalName, documentType: type, confidence: Number(confidence.toFixed(2)) };
}
// ---------------------------------------------------------------------------

app.post('/api/upload', (req, res) => {
  upload.array('documents', 200)(req, res, (err) => {
    if (err) {
      const message = err.code === 'LIMIT_FILE_COUNT'
        ? 'アップロードできるファイル数の上限（200件）を超えています。フォルダを分けて複数回に分けてアップロードしてください。'
        : err.code === 'LIMIT_FILE_SIZE'
        ? 'ファイルサイズが上限（80MB/ファイル）を超えているファイルがあります。'
        : `アップロードに失敗しました: ${err.message || err}`;
      return res.status(400).json({ ok: false, message });
    }

    // Folder uploads (webkitdirectory on the frontend) send one relPaths[]
    // entry per file, in the same order as the files themselves, containing
    // each file's path relative to the selected folder root (e.g.
    // "要件定義/ログイン機能/spec.docx"). A plain multi-file selection has no
    // sub-folder structure, so relPaths falls back to just the file name.
    const rawRelPaths = req.body.relPaths;
    const relPaths = Array.isArray(rawRelPaths) ? rawRelPaths : (rawRelPaths ? [rawRelPaths] : []);
    const newFiles = (req.files || []).map((f, i) => {
      const relPath = decodeFileName(relPaths[i] || f.originalname);
      const folder = relPath.includes('/') ? relPath.slice(0, relPath.lastIndexOf('/')) : '';
      return {
        originalName: decodeFileName(f.originalname),
        relPath,
        folder,
        fileName: f.filename,
        path: f.path,
        size: f.size,
        mimetype: f.mimetype
      };
    });
    // Appends rather than replaces — picking individual files and then a whole
    // folder (or multiple folders) in separate steps should accumulate, not
    // overwrite each other. Use the Clear button to start over.
    uploadedFiles = uploadedFiles.concat(newFiles);
    res.json({ ok: true, files: uploadedFiles.map(publicFileInfo) });
  });
});

// Health endpoint — polled by the DAT Copilot Bridge extension's watchdog so
// it can automatically quit VS Code shortly after this server goes away
// (e.g. the user closed the cmd window without a clean Ctrl+C).
app.get('/api/health', (req, res) => res.json({ ok: true, service: 'dat-ai-testcase-generator', time: Date.now() }));

app.get('/api/files', (req, res) => res.json({ files: uploadedFiles.map(publicFileInfo) }));

app.post('/api/analyze', async (req, res) => {
  const { keyword = '', functionName = '', testType = 'integration' } = req.body || {};
  if (!keyword.trim()) return res.status(400).json({ ok: false, message: 'キーワードを入力してください。' });
  if (!TEST_TYPE_PROFILES[testType]) return res.status(400).json({ ok: false, message: `不正なtest case type: ${testType}` });

  // Function Name is optional, highest-priority context when present (like
  // the old "notes" slot) — combined with Keywords for document matching and
  // the AI prompt, but the two stay separate in the response/meta so the UI
  // can show and score them individually (requirement: coverage broken out
  // by keyword AND function name).
  const combinedKeyword = [functionName, keyword].filter(Boolean).join(' ').trim() || keyword;

  const log = [];
  // Every fresh Analyze starts from a clean slate: knowledge-base cache,
  // any previous analysis/checklist, any previously generated cases. This
  // guarantees each run re-reads and re-scores the actual current uploads
  // instead of silently reusing stale cached parses from an earlier run
  // (which was also a source of the duplicate-checkpoint problem below).
  const clearedCacheEntries = clearKnowledgeBaseCache();
  latestCases = []; latestSavedCases = []; latestMeta = {}; latestCoverage = null;
  latestAnalysis = null; latestKnowledge = null;
  log.push(`Knowledge Base Cache cleared (${clearedCacheEntries} entr${clearedCacheEntries === 1 ? 'y' : 'ies'}). Starting fresh analysis.`);

  let docResult, deterministicKnowledge;
  try {
    const effectiveFiles = uploadedFiles.length ? uploadedFiles : scanExistingUploads();
    if (!uploadedFiles.length && effectiveFiles.length) log.push(`Existing uploads folderから ${effectiveFiles.length} file(s) を読み込みました。`);
    docResult = await collectDocumentText(effectiveFiles, combinedKeyword, '', log);
    deterministicKnowledge = extractDeterministicKnowledge(docResult.allRows, combinedKeyword, '', docResult.documentClassifications);
  } catch (err) {
    console.error('[analyze] document parsing error:', err);
    return res.status(500).json({ ok: false, message: `資料の読み込み中にエラーが発生しました: ${err.message || err}`, log });
  }

  let smartAnalysis = null;
  let source = 'document-fallback';
  try {
    log.push(`GitHub Copilot Bridgeへ接続中: ${COPILOT_BRIDGE_URL}`);
    log.push(`Test Case Type: ${TEST_TYPE_PROFILES[testType].label}`);
    log.push('Analyzing: Business Rule / Error Message / Screen Flow / Validation extraction');
    smartAnalysis = await generateByCopilotObject(buildSmartAnalysisPrompt(keyword, functionName, docResult.relevantText || docResult.allText, deterministicKnowledge, '', testType));
    source = 'copilot-bridge';
    log.push(`Copilot Bridge: analysis completed${COPILOT_MODEL_FAMILY ? ` (model family: ${COPILOT_MODEL_FAMILY})` : ''}`);
  } catch (err) {
    log.push('Copilot Bridge not reachable/failed. Using document-extracted candidates only (no AI reasoning).');
    log.push(String(err.message || err));
    smartAnalysis = deterministicKnowledgeToAnalysis(deterministicKnowledge, combinedKeyword);
    source = 'document-fallback';
  }

  // Guards the checklist-building step too — a malformed AI response shape
  // here used to be able to throw past every try/catch and crash the whole
  // Node process (unhandled promise rejection), which showed up client-side
  // as a bare "Failed to fetch" with no explanation. Now it degrades to a
  // normal JSON error response instead, and the document-based analysis
  // isn't lost — the person can just retry Analyze.
  try {
    const rawChecklist = buildChecklistFromAnalysis(smartAnalysis);
    const checklist = dedupeChecklist(rawChecklist);
    if (rawChecklist.length !== checklist.length) {
      log.push(`Duplicate checkpoints removed: ${rawChecklist.length} → ${checklist.length} 件（重複 ${rawChecklist.length - checklist.length} 件を除去）`);
    }
    log.push(`Checklist: ${checklist.length} 件のチェックポイントを抽出しました。`);

    // Per-category counts so the checklist summary panel can show something
    // like "ビジネスルール: 5件 / エラーメッセージ: 14件 / ..." at a glance,
    // without the person having to scroll and count checkboxes themselves.
    // Derived from whatever categories actually appear in the checklist
    // (AI-chosen, not a fixed list) — preserves first-seen order so the
    // summary panel doesn't reshuffle between runs.
    const seenCategories = [];
    const categoryCountMap = new Map();
    checklist.forEach(c => {
      if (!categoryCountMap.has(c.category)) { categoryCountMap.set(c.category, { categoryLabel: c.categoryLabel, categoryLabelEn: c.categoryLabelEn, count: 0 }); seenCategories.push(c.category); }
      categoryCountMap.get(c.category).count++;
    });
    const categoryCounts = seenCategories.map(category => ({ category, ...categoryCountMap.get(category) }));

    // Coverage at the Analyze stage: of the terms in Keywords + Function
    // Name, how many actually show up somewhere in the extracted checklist.
    // This is separate from the post-Generate "selection coverage" — this
    // one answers "did my search terms actually find real content in the
    // documents", before any checkpoints are even selected.
    const analysisCoverage = buildAnalysisCoverage(keyword, functionName, checklist);
    log.push(`Keyword/Function Name Coverage: ${analysisCoverage.coveredTerms}/${analysisCoverage.totalTerms} (${analysisCoverage.coveragePercent}%)`);

    latestAnalysis = { keyword, functionName, testType, smartAnalysis, checklist, source, date: new Date().toISOString().slice(0, 10) };
    latestKnowledge = deterministicKnowledge;

    res.json({
      ok: true,
      source,
      log,
      checklist,
      categoryCounts,
      analysisCoverage,
      testType,
      testTypeLabel: TEST_TYPE_PROFILES[testType].label,
      matchedCount: docResult.matches.length,
      parsedFiles: docResult.parsedFiles,
      documentClassifications: docResult.documentClassifications
    });
  } catch (err) {
    console.error('[analyze] checklist-building error:', err);
    res.status(500).json({ ok: false, message: `チェックリスト作成中にエラーが発生しました: ${err.message || err}`, log });
  }
});

// ⑤ Generate: builds test cases scoped to ONLY the checkpoints the person
// selected from the Analyze step's checklist (not everything the analysis
// found) — selectedCheckpoints is exactly what /api/analyze returned,
// filtered down to the checked items, so no server-side session/cache is
// needed to correlate the two requests.
app.post('/api/generate-from-selection', async (req, res) => {
  const { selectedCheckpoints, keyword = '', functionName = '', testType = 'integration' } = req.body || {};
  if (!Array.isArray(selectedCheckpoints) || !selectedCheckpoints.length) {
    return res.status(400).json({ ok: false, message: 'チェックポイントを1つ以上選択してください。' });
  }
  if (!TEST_TYPE_PROFILES[testType]) return res.status(400).json({ ok: false, message: `不正なtest case type: ${testType}` });
  const combinedKeyword = [functionName, keyword].filter(Boolean).join(' ').trim() || keyword;

  const log = [];
  let generated = null;
  let source = 'document-fallback';
  try {
    log.push(`GitHub Copilot Bridgeへ接続中: ${COPILOT_BRIDGE_URL}`);
    const prompt = buildSelectionTestCasePrompt(combinedKeyword, testType, selectedCheckpoints);
    generated = await generateByCopilot(prompt);
    source = 'copilot-bridge-selection';
    log.push(`Copilot Bridge: request completed${COPILOT_MODEL_FAMILY ? ` (model family: ${COPILOT_MODEL_FAMILY})` : ''}`);
  } catch (err) {
    log.push('Copilot Bridge not reachable/failed. Building cases directly from selected checkpoints (no AI reasoning).');
    log.push(String(err.message || err));
    generated = buildCasesFromCheckpointsFallback(selectedCheckpoints, combinedKeyword);
  }

  // Everything past this point (dedupe, coverage scoring, Excel writing) has
  // to be defended too — this used to be unguarded, so any error here (bad
  // shape from an unusual Copilot response, an Excel-writer edge case, etc)
  // became an unhandled promise rejection that could take the whole Node
  // process down mid-request, which shows up client-side as a bare
  // "Failed to fetch" with zero explanation. Now it degrades to a normal
  // JSON error response instead.
  try {
    const beforeDedupeCount = Array.isArray(generated) ? generated.length : 0;
    latestCases = dedupeGeneratedCasesDynamic(normalizeDynamicCases(generated));
    if (beforeDedupeCount && beforeDedupeCount !== latestCases.length) {
      log.push(`Duplicate removal: ${beforeDedupeCount} → ${latestCases.length} 件（重複 ${beforeDedupeCount - latestCases.length} 件を除去）`);
    }
    latestMeta = {
      keyword, functionName, testType, testTypeLabel: TEST_TYPE_PROFILES[testType].label,
      date: new Date().toISOString().slice(0, 10), author: 'AI Test Case Generator (Copilot Bridge Smart QA)',
      source, selectedCount: selectedCheckpoints.length
    };
    latestSavedCases = latestCases;

    // Coverage here is scoped to what the person actually selected — not the
    // full document-wide extraction — so the percentage answers "did my
    // chosen checkpoints make it into the generated cases", which is the
    // meaningful question at this stage (comparing against everything the
    // Analyze step found, selected or not, would understate coverage on
    // purpose whenever someone deliberately narrows their selection).
    latestCoverage = buildSelectionCoverage(selectedCheckpoints, latestCases, keyword, functionName);
    log.push(`Selected Checkpoint Coverage: ${latestCoverage.totalCovered}/${latestCoverage.totalSelected} (${latestCoverage.coveragePercent}%)`);
    const uncovered = latestCoverage.items.filter(i => !i.covered);
    if (uncovered.length) log.push(`Not yet reflected in generated cases: ${uncovered.map(i => i.title).join(' / ')}`);

    const excelUrl = await writeExcelDynamic(latestCases, latestMeta, latestCoverage);
    res.json({ ok: true, source, log, meta: latestMeta, cases: latestCases, excelUrl, coverage: latestCoverage });
  } catch (err) {
    console.error('[generate-from-selection] post-processing error:', err);
    res.status(500).json({ ok: false, message: `テストケースの後処理中にエラーが発生しました: ${err.message || err}`, log });
  }
});

app.get('/api/download/latest', async (req, res) => {
  if (!latestCases.length) return res.status(404).send('No generated test cases.');
  const filePath = await writeExcelDynamic(latestCases, latestMeta, latestCoverage);
  res.download(path.join(__dirname, filePath.replace(/^\//, '')));
});

// Save: the frontend posts the full (edited) case table here after the
// user clicks "Save". Only once this succeeds does /api/download/latest (or
// the download button) reflect the edited content — this endpoint also
// becomes the new "last saved" baseline that Cancel reverts to. Cases keep
// whatever dynamic key set the frontend sends (columns are data-driven, not
// a fixed schema).
app.post('/api/update-cases', async (req, res) => {
  const { cases } = req.body || {};
  if (!Array.isArray(cases) || !cases.length) {
    return res.status(400).json({ ok: false, message: '保存するテストケースがありません。' });
  }
  latestCases = normalizeDynamicCases(cases);
  latestSavedCases = latestCases;
  const excelUrl = await writeExcelDynamic(latestCases, latestMeta, latestCoverage);
  res.json({ ok: true, cases: latestCases, excelUrl });
});

// Cancel: returns the last-saved baseline (last successful
// /api/generate-from-selection or /api/update-cases) so the frontend can
// restore the preview table to what it looked like before the user's
// unsaved edits.
app.get('/api/cases/saved', (req, res) => {
  res.json({ ok: true, cases: latestSavedCases, meta: latestMeta });
});

// Translates the currently displayed test case content (and keyword) between
// Japanese and English for the language toggle. Works on whatever dynamic
// key set the cases currently have. Requires the Copilot Bridge — there's no
// offline fallback for translating arbitrary AI-generated prose.
app.post('/api/translate-cases', async (req, res) => {
  const { cases, meta = {}, targetLang } = req.body || {};
  if (!Array.isArray(cases) || !cases.length) {
    return res.status(400).json({ ok: false, message: '翻訳するテストケースがありません。' });
  }
  if (targetLang !== 'en' && targetLang !== 'ja') {
    return res.status(400).json({ ok: false, message: 'targetLangは "en" または "ja" である必要があります。' });
  }
  try {
    const result = await translateCasesViaCopilot(cases, meta, targetLang);
    res.json({ ok: true, cases: result.cases, meta: result.meta });
  } catch (err) {
    res.status(502).json({ ok: false, message: `翻訳に失敗しました（DAT Copilot Bridgeの起動が必要です）: ${err.message || err}` });
  }
});

// ④ Clear in-memory state and Knowledge Base Cache on demand.
// Called by the frontend immediately before each /api/generate so each
// generation starts clean — no stale cases, no stale parsed-document cache.
app.post('/api/clear', (req, res) => {
  const cleared = clearKnowledgeBaseCache();
  latestCases = [];
  latestSavedCases = [];
  latestMeta = {};
  latestCoverage = null;
  latestAnalysis = null;
  latestKnowledge = null;
  res.json({ ok: true, clearedCacheEntries: cleared });
});

// Explicit "Clear files" action from the file list — separate from
// /api/clear above so that re-Analyzing with a new keyword doesn't
// accidentally wipe out documents the person still wants to reuse. Also
// deletes the actual uploaded files from disk (not just the in-memory list).
app.post('/api/clear-files', (req, res) => {
  uploadedFiles.forEach(f => { try { fs.unlinkSync(f.path); } catch (_) { /* already gone, ignore */ } });
  uploadedFiles = [];
  res.json({ ok: true });
});

function loadDotEnv(envPath) {
  if (!fs.existsSync(envPath)) return;
  const lines = fs.readFileSync(envPath, 'utf8').split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const idx = trimmed.indexOf('=');
    if (idx === -1) continue;
    const key = trimmed.slice(0, idx).trim();
    let value = trimmed.slice(idx + 1).trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    if (key && process.env[key] == null) process.env[key] = value;
  }
}

function publicFileInfo(f) {
  return { originalName: f.originalName, relPath: f.relPath || f.originalName, folder: f.folder || '', fileName: f.fileName, size: f.size, mimetype: f.mimetype };
}

function decodeFileName(name) {
  try {
    const decoded = Buffer.from(name, 'latin1').toString('utf8');
    return decoded.includes('�') ? name : decoded;
  } catch (_) { return name; }
}


function scanExistingUploads() {
  if (!fs.existsSync(UPLOAD_DIR)) return [];
  return fs.readdirSync(UPLOAD_DIR)
    .map(fileName => {
      const full = path.join(UPLOAD_DIR, fileName);
      const st = fs.statSync(full);
      if (!st.isFile()) return null;
      const originalName = fileName.replace(/^\d+_/, '');
      return { originalName, fileName, path: full, size: st.size, mimetype: '' };
    })
    .filter(Boolean);
}

async function collectGeneratedSampleText(log) {
  if (!fs.existsSync(GENERATED_DIR)) return '';
  const files = fs.readdirSync(GENERATED_DIR)
    .filter(n => /\.(xlsx|xls|json|txt|md)$/i.test(n))
    .map(n => ({ name: n, path: path.join(GENERATED_DIR, n), mtime: fs.statSync(path.join(GENERATED_DIR, n)).mtimeMs }))
    .sort((a,b) => b.mtime - a.mtime)
    .slice(0, 3);
  const chunks = [];
  for (const f of files) {
    try {
      const ext = path.extname(f.name).toLowerCase();
      if (ext === '.xlsx') {
        const rows = await parseFileToRows({ path: f.path, originalName: f.name }, ext);
        chunks.push(`【Generated sample: ${f.name}】\n` + rows.slice(0, 60).map(r => r.text).join('\n'));
      } else if (ext === '.xls' && XLSX) {
        const rows = parseWithXlsx(f.path);
        chunks.push(`【Generated sample: ${f.name}】\n` + rows.slice(0, 60).map(r => r.text).join('\n'));
      } else {
        chunks.push(`【Generated sample: ${f.name}】\n` + fs.readFileSync(f.path, 'utf8').slice(0, 8000));
      }
    } catch (e) {
      log.push(`Generated sample parse skipped ${f.name}: ${e.message}`);
    }
  }
  if (chunks.length) log.push(`Generated folder reference: ${chunks.length} file(s)`);
  return chunks.join('\n\n').slice(0, 20000);
}

async function collectDocumentText(files, keyword, notes, log) {
  const allRows = [];
  const parsedFiles = [];
  const documentClassifications = [];

  for (const file of files) {
    const ext = path.extname(file.originalName).toLowerCase();
    try {
      // V2.2: Knowledge Base Cache — reuse parsed rows + classification when
      // the same file (path+size+mtime) was already parsed before, instead
      // of re-parsing the document on every /api/generate call.
      const cached = getCachedRows(file.path);
      let rows, classification;
      if (cached) {
        rows = cached.rows;
        classification = cached.classification;
        log.push(`Cache hit ${file.originalName}: ${rows.length} row(s)/line(s) (Knowledge Base Cache)`);
      } else {
        rows = await parseFileToRows(file, ext);
        classification = classifyDocument(file.originalName, rows);
        setCachedRows(file.path, rows, classification);
        log.push(`Parsed ${file.folder ? file.folder + '/' : ''}${file.originalName}: ${rows.length} row(s)/line(s)`);
      }

      rows.forEach((r, idx) => allRows.push({ ...r, fileName: file.originalName, rowNo: idx + 1 }));
      parsedFiles.push({ fileName: file.originalName, ext, rows: rows.length, ok: true, documentType: classification.documentType });
      documentClassifications.push(classification);
      log.push(`Document Classifier: ${file.originalName} → ${classification.documentType} (confidence ${classification.confidence})`);
    } catch (e) {
      parsedFiles.push({ fileName: file.originalName, ext, rows: 0, ok: false, error: e.message });
      log.push(`Parse failed ${file.originalName}: ${e.message}`);
    }
  }

  const terms = buildSearchTerms(`${keyword} ${notes || ''}`);
  const scored = allRows.map(r => ({ ...r, score: scoreText(r.text, terms) })).filter(r => r.score > 0);
  scored.sort((a, b) => b.score - a.score);

  let matches = expandContextRows(allRows, scored.slice(0, 120), 3).slice(0, 220);
  if (matches.length === 0 && allRows.length) {
    matches = allRows.filter(r => isLikelySpecRow(r.text)).slice(0, 80);
  }
  if (matches.length === 0) matches = allRows.slice(0, 80);

  const allText = allRows.map(r => `[${r.fileName}${r.sheet ? ' / ' + r.sheet : ''} R${r.rowNo}] ${r.text}`).join('\n').slice(0, 80000);
  const relevantText = matches.map(r => `[${r.fileName}${r.sheet ? ' / ' + r.sheet : ''} R${r.rowNo}] ${r.text}`).join('\n').slice(0, 60000);
  log.push(`Keyword search matched: ${matches.length} row(s)/line(s)`);

  return { allRows, matches, allText, relevantText, parsedFiles, documentClassifications };
}

async function parseFileToRows(file, ext) {
  if (ext === '.txt' || ext === '.md' || ext === '.csv') {
    const text = fs.readFileSync(file.path, 'utf8');
    return text.split(/\r?\n/).map(x => cleanText(x)).filter(Boolean).map(text => ({ text }));
  }
  if (ext === '.docx') {
    const mammoth = require('mammoth');
    const result = await mammoth.extractRawText({ path: file.path });
    return result.value.split(/\r?\n/).map(x => cleanText(x)).filter(Boolean).map(text => ({ text }));
  }
  if (ext === '.doc') {
    // Legacy .doc is binary. This best-effort extraction may not be perfect.
    const buf = fs.readFileSync(file.path);
    const text = buf.toString('utf8').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, ' ');
    return text.split(/\r?\n|\s{3,}/).map(x => cleanText(x)).filter(x => x.length > 3).map(text => ({ text }));
  }

  if (ext === '.pptx') {
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(fs.readFileSync(file.path));
    const rows = [];
    const names = Object.keys(zip.files)
      .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n) || /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n))
      .sort((a,b) => a.localeCompare(b, undefined, { numeric: true }));
    for (const name of names) {
      const xml = await zip.files[name].async('string');
      const text = cleanText(xml
        .replace(/<a:br\s*\/?>/g, '\n')
        .replace(/<[^>]+>/g, ' ')
        .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
        .replace(/&quot;/g, '"').replace(/&#39;/g, "'"));
      if (text) rows.push({ sheet: name.replace(/^ppt\//, ''), text });
    }
    return rows;
  }

  if (ext === '.pdf') {
    const pdfParse = require('pdf-parse');
    const data = await pdfParse(fs.readFileSync(file.path));
    return data.text.split(/\r?\n/).map(x => cleanText(x)).filter(Boolean).map(text => ({ text }));
  }
  if (ext === '.xlsx') {
    // ExcelJS keeps formatting stable for xlsx.
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(file.path);
    const rows = [];
    wb.worksheets.forEach((ws) => {
      ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const values = row.values.slice(1).map(cellToText).filter(Boolean);
        const text = cleanText(values.join(' | '));
        if (text) rows.push({ sheet: ws.name, rowNo: rowNumber, text });
      });
    });
    return rows;
  }
  if (ext === '.xls') {
    if (!XLSX) throw new Error('To read .xls files, please run: npm install xlsx');
    return parseWithXlsx(file.path);
  }
  throw new Error(`Unsupported file extension: ${ext}`);
}

function parseWithXlsx(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const rows = [];
  wb.SheetNames.forEach(sheetName => {
    const sheet = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    data.forEach((arr, idx) => {
      const text = cleanText(arr.map(cellToText).filter(Boolean).join(' | '));
      if (text) rows.push({ sheet: sheetName, rowNo: idx + 1, text });
    });
  });
  return rows;
}

function cellToText(v) {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.text) return String(v.text);
    if (v.richText) return v.richText.map(x => x.text || '').join('');
    if (v.result != null) return String(v.result);
    if (v.formula) return String(v.formula);
    if (v instanceof Date) return v.toISOString().slice(0, 10);
  }
  return String(v);
}

function cleanText(s) {
  return String(s || '').replace(/\s+/g, ' ').trim();
}


function expandContextRows(allRows, matchedRows, radius = 2) {
  const keys = new Set();
  const result = [];
  const add = (r) => {
    const key = `${r.fileName}::${r.sheet || ''}::${r.rowNo}`;
    if (!keys.has(key)) { keys.add(key); result.push(r); }
  };
  for (const m of matchedRows) {
    for (const r of allRows) {
      if (r.fileName === m.fileName && (r.sheet || '') === (m.sheet || '') && Math.abs((r.rowNo || 0) - (m.rowNo || 0)) <= radius) add(r);
    }
  }
  return result;
}

function buildSearchTerms(keyword) {
  const base = String(keyword || '').trim();
  const terms = new Set([base]);
  base.split(/[\s　,、/・()（）\[\]【】]+/).filter(Boolean).forEach(t => terms.add(t));
  const synonyms = {
    '購入注文': ['購入', '注文', '申込', '買付', '登録', '確認', '完了'],
    '定時定額': ['定時', '定額', '積立', '新規', '申込', '金額', '契約', '定時定額（新規）'],
    'ログイン': ['ログイン', 'ID', 'パスワード', '認証'],
    '運用報告書': ['運用報告書', '報告書', '交付', '閲覧'],
    'DB': ['DB', 'テーブル', '登録', '更新', '削除', 'レコード'],
    '非課税口座簡易開設申込日': ['非課税口座簡易開設申込日', '簡易開設', '申込日', '先日付', 'NISA', 'NISA注文チェック', 'エラーメッセージ', '日付'],
    '先日付': ['先日付', '未来日', '日付チェック', 'NISA注文チェック', 'エラー', 'メッセージ'],
    'NISA': ['NISA', '非課税', '非課税口座', '注文チェック', '口座開設']
  };
  Object.entries(synonyms).forEach(([k, vals]) => {
    if (base.includes(k)) vals.forEach(v => terms.add(v));
  });
  return [...terms].filter(Boolean);
}

function scoreText(text, terms) {
  const t = String(text || '').toLowerCase();
  let score = 0;
  terms.forEach(term => {
    const q = String(term).toLowerCase();
    if (!q) return;
    if (t.includes(q)) score += q.length >= 4 ? 5 : 3;
  });
  if (isLikelySpecRow(text)) score += 2;
  return score;
}

function isLikelySpecRow(text) {
  return /(画面|ボタン|リンク|チェック|エラー|入力|表示|遷移|登録|更新|削除|DB|テーブル|確認|項目|条件|必須|上限|下限|正常|異常)/.test(text);
}

function extractDeterministicKnowledge(rows, keyword, notes, documentClassifications) {
  const all = (rows || []).map(r => ({ ...r, text: cleanText(r.text) })).filter(r => r.text);
  const baseTerms = buildSearchTerms(`${keyword} ${notes || ''}`);
  const focusRows = all
    .map((r, idx) => ({ ...r, __idx: idx, score: scoreText(r.text, baseTerms) + (/(エラー|メッセージ|チェック|条件|画面|遷移|入力|項目|日付|NISA|非課税|DB|テーブル|登録|更新|口座|受付|不可|必須|上限|下限|同日|翌|先日付|未来日)/.test(r.text) ? 4 : 0) }))
    .filter(r => r.score > 0 || /(\b\d{4,5}\b|E\d{3,}|エラー|メッセージ|画面|遷移|チェック|必須|NISA|非課税)/i.test(r.text))
    .sort((a,b) => b.score - a.score)
    .slice(0, 260);

  const errorMessages = extractErrorMessageCandidates(all, focusRows).slice(0, 120);
  const screenTransitions = extractScreenTransitionCandidates(all, focusRows).slice(0, 120);
  const validationRules = extractValidationRuleCandidates(all, focusRows).slice(0, 160);
  const inputFields = extractInputFieldCandidates(all, focusRows).slice(0, 120);
  const dbOrInterfaceChecks = focusRows
    .filter(r => /(DB|テーブル|登録|更新|削除|レコード|ステータス|状態|IF|インターフェース|連携|作成|保存)/i.test(r.text))
    .map(r => ({ rule: r.text, source: rowSource(r) }))
    .slice(0, 80);

  return {
    // V2.2: Document Classifier output, passed through so the AI prompt can
    // weight evidence by document type (e.g. trust DB definition documents
    // for dbOrInterfaceChecks, screen design docs for screenTransitions).
    documentClassification: documentClassifications || [],
    focusTerms: baseTerms.slice(0, 50),
    errorMessages,
    screenTransitions,
    validationRules,
    inputFields,
    dbOrInterfaceChecks,
    evidenceRows: focusRows.map(rowRef).slice(0, 140)
  };
}

function rowRef(r) {
  return `[${r.fileName || ''}${r.sheet ? ' / ' + r.sheet : ''} R${r.rowNo || ''}] ${r.text}`;
}

function rowSource(r) {
  return `${r.fileName || ''}${r.sheet ? ' / ' + r.sheet : ''} R${r.rowNo || ''}`.trim();
}

function contextForRow(all, target, radius = 2) {
  return all.filter(r => r.fileName === target.fileName && (r.sheet || '') === (target.sheet || '') && Math.abs((r.rowNo || 0) - (target.rowNo || 0)) <= radius)
    .map(r => r.text)
    .filter(Boolean)
    .join(' / ');
}

function extractErrorMessageCandidates(all, focusRows) {
  const rows = focusRows.filter(r => /(\b\d{4,5}\b|E\d{3,}|エラー|メッセージ|受付不可|受付できません|入力してください|指定してください|表示|不可|警告)/i.test(r.text));
  const out = [];
  for (const r of rows) {
    const ctx = contextForRow(all, r, 2);
    const codes = r.text.match(/\b\d{4,5}\b|E\d{3,}/g) || (ctx.match(/\b\d{4,5}\b|E\d{3,}/g) || []);
    const msg = extractLikelyMessage(ctx || r.text);
    if (codes.length) {
      [...new Set(codes)].forEach(code => out.push({ code, message: msg || '要確認', condition: extractLikelyCondition(ctx || r.text), source: rowSource(r), context: shorten(ctx || r.text, 240) }));
    } else {
      out.push({ code: '要確認', message: msg || shorten(r.text, 120), condition: extractLikelyCondition(ctx || r.text), source: rowSource(r), context: shorten(ctx || r.text, 240) });
    }
  }
  return dedupeBy(out, x => `${x.code}::${x.message}::${x.source}`);
}

function extractLikelyMessage(text) {
  const s = cleanText(text);
  const quoted = s.match(/[「『](.*?)(?:」|』)/);
  if (quoted && /(エラー|不可|ください|ません|表示|受付|指定|入力)/.test(quoted[1])) return quoted[1];
  const parts = s.split(/\s*[|／/。]\s*/).map(cleanText).filter(Boolean);
  const hit = parts.find(p => /(エラー|不可|ください|ません|表示|受付|指定|入力|確認)/.test(p) && p.length <= 140);
  return hit || '';
}

function extractLikelyCondition(text) {
  const s = cleanText(text);
  const parts = s.split(/\s*[|／/。]\s*/).map(cleanText).filter(Boolean);
  const hit = parts.find(p => /(場合|とき|時|条件|チェック|先日付|未来日|過去日|同日|翌|NISA|非課税|未入力|上限|下限|口座)/.test(p) && p.length <= 160);
  return hit || '要確認';
}

function extractScreenTransitionCandidates(all, focusRows) {
  const rows = focusRows.filter(r => /(画面|メニュー|ボタン|リンク|遷移|ログイン|確認|完了|戻る|次へ|登録|押下|クリック|選択)/.test(r.text));
  const out = [];
  for (const r of rows) {
    const ctx = contextForRow(all, r, 2);
    const screens = [...new Set((ctx.match(/[\w一-龠ぁ-んァ-ヶー（）()・／/\- ]{2,40}(?:画面|メニュー|ボタン|リンク|確認|完了|一覧|入力|申込|注文)/g) || []).map(cleanText))].slice(0, 8);
    out.push({ flow: screens.length ? screens.join(' → ') : shorten(r.text, 160), operation: extractOperation(ctx || r.text), source: rowSource(r), context: shorten(ctx || r.text, 260) });
  }
  return dedupeBy(out, x => `${x.flow}::${x.source}`);
}

function extractOperation(text) {
  const s = cleanText(text);
  const parts = s.split(/\s*[|／/。]\s*/).map(cleanText).filter(Boolean);
  const hit = parts.find(p => /(クリック|押下|選択|入力|表示|遷移|ログイン|登録|確認|戻る|次へ)/.test(p) && p.length <= 140);
  return hit || '要確認';
}

function extractValidationRuleCandidates(all, focusRows) {
  const rows = focusRows.filter(r => /(必須|未入力|入力|チェック|上限|下限|日付|先日付|未来日|過去日|同日|翌営業日|営業日|NISA|非課税|口座|フラグ|状態|桁|文字|半角|全角|範囲|不可|可能|エラー)/.test(r.text));
  const out = [];
  for (const r of rows) {
    const ctx = contextForRow(all, r, 1);
    out.push({ rule: shorten(ctx || r.text, 220), field: extractFieldName(ctx || r.text), boundaryValues: extractBoundaryValues(ctx || r.text), expected: extractLikelyMessage(ctx || r.text) || '仕様通りにチェックされること', source: rowSource(r) });
  }
  return dedupeBy(out, x => `${x.field}::${x.rule}`);
}

function extractFieldName(text) {
  const s = cleanText(text);
  const patterns = [
    /([\w一-龠ぁ-んァ-ヶー（）()・／/\-]{2,40}(?:日|日付|区分|番号|コード|金額|数量|フラグ|状態|口座|項目|ID|パスワード))/,
    /([\w一-龠ぁ-んァ-ヶー（）()・／/\-]{2,40})[:：]/
  ];
  for (const p of patterns) { const m = s.match(p); if (m) return cleanText(m[1]); }
  return '要確認';
}

function extractBoundaryValues(text) {
  const hits = String(text).match(/(\d{4}[\/\-年]\d{1,2}[\/\-月]\d{1,2}日?|\d+[,.]?\d*円|\d+桁|\d+文字|0円|未入力|空白|同日|翌営業日|先日付|未来日|過去日|上限|下限|ON|OFF|0|1)/g) || [];
  return [...new Set(hits)].slice(0, 12);
}

function extractInputFieldCandidates(all, focusRows) {
  const out = [];
  for (const r of focusRows) {
    const ctx = contextForRow(all, r, 1);
    const names = ctx.match(/[\w一-龠ぁ-んァ-ヶー（）()・／/\-]{2,35}(?:日|日付|区分|番号|コード|金額|数量|フラグ|状態|口座|ID|パスワード|項目)/g) || [];
    names.forEach(name => out.push({ name: cleanText(name), values: extractBoundaryValues(ctx), source: rowSource(r) }));
  }
  return dedupeBy(out, x => `${x.name}::${x.source}`);
}

function dedupeBy(arr, keyFn) {
  const seen = new Set();
  return arr.filter(x => { const k = keyFn(x); if (seen.has(k)) return false; seen.add(k); return true; });
}

function buildSmartAnalysisPrompt(keyword, functionName, docText, knowledge, sampleText, testType = 'integration') {
  const profile = TEST_TYPE_PROFILES[testType] || TEST_TYPE_PROFILES.integration;
  docText = String(docText || '').slice(0, COPILOT_MAX_DOC_CHARS);
  sampleText = String(sampleText || '').slice(0, COPILOT_MAX_SAMPLE_CHARS);
  const knowledgeText = JSON.stringify(knowledge, null, 2).slice(0, COPILOT_MAX_KNOWLEDGE_CHARS);
  return `あなたは日本の金融系システム(SONAR/投信注文)の上級QAアナリストです。
DAT Copilot Bridge Smart QA Engineとして、アップロード資料を分析し、テストケース生成のための中間分析結果をJSONで作成してください。

【テストケース種別】
${profile.label}
${profile.analysisHint}

【分析範囲の絞り込み — 最重要】
- 対象機能名：${functionName || '指定なし'}
- キーワード：${keyword}
- 上記の対象機能・キーワードに直接関連する内容のみを抽出してください。資料内に他機能（無関係な他画面・他バッチ・他帳票等）の記述があっても、対象機能・キーワードと関係がなければ含めないでください。
- 迷った場合は「対象機能・キーワードの処理フロー、入力項目、業務ルールに実際に登場するか」を基準に判断してください。関連が薄いものは無理に含めないでください。

【カテゴリ — 固定リストではなく、内容に応じて自由に分類してください】
- 資料から読み取れるチェックポイントは、実際に資料に存在するものを漏れなく抽出してください。想定されるカテゴリの例（これに限定しないでください）：
  ビジネスルール／エラーメッセージ／画面遷移／入力チェック・バリデーション／入力項目／DB・IF連携／日付条件／権限チェック／状態遷移／口座状態／排他制御／ログ出力／通知条件／その他、資料に実際に記載がある観点。
- 各チェックポイントに、内容を最もよく表すカテゴリ名を自由に付けてください。固定の分類名に無理に当てはめないでください。同じ観点の繰り返し（同一内容の水増し）は禁止です。
- 推測で作らないでください。資料から読み取れないものは作成しないでください。

【各チェックポイントに必ず含めるフィールド】
- category: チェックポイントの分類（日本語、自由記述。例：「日付条件チェック」「権限チェック」「口座状態遷移」等）
- categoryEn: categoryの英語訳（例：「Date Condition Check」）
- title: 15〜25文字程度の短い日本語ラベル。エンドユーザーが一覧でパッと見て内容が分かる短い見出し（例:「NISA適用フラグの判定」）。detailの文章をそのまま切り取らないでください。要約してください。
- detail: 条件・ルールの詳細情報
- example: 実際のテストデータ例（例: "NISA適用=1, 申込日=注文日" や "UserID=validUser, 試行回数=5"）。具体的な値が資料から読み取れる場合のみ記載し、読み取れない場合は空文字列 "" にしてください（推測で作らない）。
- source: 資料内の参照箇所。必ず「ファイル名 / シート名 R行番号」の形式で、資料抜粋の各行冒頭にある [ファイル名 / シート名 R行番号] の表記をそのまま転記してください（例:「F03投信定時定額伝票_030.xls / チェック仕様_定時定額設定 R916」）。シートの概念がない資料（Word/PDF/テキスト等）はシート名を省略し「ファイル名 R行番号」としてください。行番号だけ（例:「R23」）を書くことは禁止です。

【プログラムが抽出した候補】
${knowledgeText}
上記の documentClassification は各アップロード資料の種別（基本設計書／詳細設計書／DB定義書／画面設計書／エラーメッセージ一覧など）です。種別に応じて根拠の重み付けをしてください（例：DB定義書はDB/IF系の根拠として優先、画面設計書は画面遷移系の根拠として優先、エラーメッセージ一覧はエラーメッセージ系の根拠として優先）。

【過去に生成されたテストケース例（あれば粒度の参考。内容を盲信しない）】
${sampleText || 'なし'}

【資料抜粋】
${docText || 'なし'}

【返却JSON形式】
{
  "targetFunction": "...",
  "checkpoints": [
    {"category": "...", "categoryEn": "...", "title": "...", "detail": "...", "example": "...", "source": "..."}
  ],
  "openQuestions": ["..."]
}`;
}

function buildSmartTestCasePrompt(keyword, notes, docText, knowledge, analysis, sampleText, testType = 'integration') {
  const profile = TEST_TYPE_PROFILES[testType] || TEST_TYPE_PROFILES.integration;
  docText = String(docText || '').slice(0, COPILOT_MAX_DOC_CHARS);
  sampleText = String(sampleText || '').slice(0, COPILOT_MAX_SAMPLE_CHARS);
  const analysisText = JSON.stringify(analysis || {}, null, 2).slice(0, COPILOT_MAX_ANALYSIS_CHARS);
  const knowledgeText = JSON.stringify(knowledge || {}, null, 2).slice(0, COPILOT_MAX_KNOWLEDGE_CHARS);
  return `あなたは日本の金融系システム(SONAR/投信注文)のテスト設計者です。
DAT Copilot Bridge Smart QA Engineとして、資料分析結果をもとにDAT向けの詳細な${profile.label}ケースを作成してください。

【テストケース種別】
${profile.label}
${profile.caseHint}

【対象機能】
${keyword}

【Additional Notes（最優先）】
${notes || '指定なし'}

【中間分析結果】
${analysisText}

【プログラム抽出候補】
${knowledgeText}

【過去生成サンプル（出力粒度の参考。完全コピー禁止）】
${sampleText || 'なし'}

【資料抜粋】
${docText || 'なし'}

【生成方針】
- Additional Notesで指定された内容を必ず中心にしてください。
- ただし、エラーコード/エラーメッセージだけに限定しないでください。関連する screen transition、input validation、date boundary、NISA condition、account status、DB/IF、authority、normal/abnormal cases も、資料に根拠があるものはすべて含めてください。
- Error Message Extractionの結果がある場合、expectedResultにエラーコード・メッセージ・受付可否・遷移可否を具体的に入れてください。
- Screen Transition Analysisの結果がある場合、stepsにログイン → メニュー選択 → 対象画面 → 入力 → 確認/登録 → 結果確認までを具体化してください。
- Validation Analysisの結果がある場合、正常値/異常値/境界値（同日、翌営業日、先日付、未入力、上限超過など）ごとにケース化してください。
- Categoryは固定リストから選ばず、内容に合わせて動的に作成してください。例：NISA申込日チェック、先日付注文チェック、簡易開設申込状態チェック、画面遷移、入力値検証、DB更新確認。
- 件数は固定しないでください。必要十分な件数のみ生成してください。20件未満でも20件以上でも構いません。水増し禁止。
- Genericな「画面遷移」「DB更新」「入力チェック」だけのケースは禁止です。対象機能・条件・項目名が分かる具体的なテスト項目にしてください。
- inputDataは必ず具体的な「項目名 = 値」の形式で書いてください。JSONオブジェクト、[object Object]、Valid Data、Sample Dataは禁止です。
- 値を資料から導ける場合は実際の値/例を作ってください。例：非課税口座簡易開設申込日 = 2026/07/10、注文日 = 2026/07/10、NISAフラグ = 0。
- 値が判断できない場合は「項目名 = 要確認」としてください。
- expectedResultは「どの画面で何が起きるか」「どのエラー/文言が出るか」「DB/状態がどうなるか」を具体的に書いてください。
- 資料にないエラー文言は作らず「該当メッセージは要確認」としてください。
- 資料から判断できない点は推測せず、precondition/inputData/expectedResult内に「要確認」と記載してください。

【出力形式】
必ずJSON配列のみを返してください。Markdownや説明文は禁止です。
各要素は以下のキーのみを持つこと：
no, category, testItem, precondition, steps, inputData, expectedResult, priority

【inputDataの良い例】
非課税口座簡易開設申込日 = 2026/07/10\n注文日 = 2026/07/10\nNISAフラグ = 0\nファンドコード = 要確認\n積立金額 = 10,000円

【expectedResultの良い例】
確認ボタン押下後、NISA注文チェックによりエラー14840が表示され、注文受付不可となること。次画面へ遷移しないこと。
`;
}

// Backward-compatible alias.
function buildPrompt(keyword, notes, docText) {
  const knowledge = extractDeterministicKnowledge([], keyword, notes);
  return buildSmartTestCasePrompt(keyword, notes, docText, knowledge, null, '', 'integration');
}

// --- Analyze → Select → Generate helpers ------------------------------------
// The Analyze step asks the AI for a flat array of checkpoints, each with
// its own freely-chosen category (see buildSmartAnalysisPrompt) —
// buildChecklistFromAnalysis just assigns IDs and normalizes field lengths,
// so the person can tick which checkpoints they actually want covered
// before generating.

// Consumes the new dynamic shape: analysis.checkpoints is a flat array
// where EACH item carries its own category (chosen freely by the AI, not
// restricted to a fixed enum) — see buildSmartAnalysisPrompt. This replaced
// an earlier design that only read from 6 hardcoded keys (businessRules,
// errorMessages, screenTransitions, validationRules, inputFields,
// dbOrInterfaceChecks): anything the AI identified outside those 6 — date
// conditions, authority checks, state transitions, account status, etc. —
// had nowhere to go and was silently dropped before it ever reached the
// checklist UI.
function buildChecklistFromAnalysis(analysis) {
  const items = Array.isArray(analysis && analysis.checkpoints) ? analysis.checkpoints : [];
  return items.map((item, idx) => {
    if (!item || typeof item !== 'object') return null;
    const category = shorten(cleanText(item.category), 30) || 'その他';
    const categoryLabelEn = shorten(cleanText(item.categoryEn), 30) || category;
    const title = shorten(cleanText(item.title), 60) || category;
    const detail = shorten(cleanText(item.detail), 200);
    const example = shorten(cleanText(item.example), 120);
    if (!title && !detail) return null;
    return {
      id: `cp-${idx}`,
      category, categoryLabel: category, categoryLabelEn,
      title, detail, example,
      source: item.source ? cleanText(item.source) : ''
    };
  }).filter(Boolean);
}

// Removes duplicate/near-duplicate checkpoints — e.g. the same rule getting
// extracted twice because it appeared in overlapping document chunks, or
// the AI restating an identical checkpoint under two slightly different
// titles. Keyed on normalized title+detail (not example, which can
// legitimately vary even for the same underlying rule) within the SAME
// category — matching title+detail across category boundaries is treated
// as different checkpoints on purpose, since the same wording can mean
// something different depending on context (e.g. an error message vs. a
// business rule referencing it).
function dedupeChecklist(checklist) {
  const seen = new Set();
  return (checklist || []).filter(cp => {
    const key = `${cp.category}::${normalizeForMatch(`${cp.title}::${cp.detail}`)}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// Analyze-stage coverage: of the individual terms in Keywords + Function
// Name, how many are actually represented somewhere in the extracted
// checklist (title+detail+example). Reuses the same anchor+bigram matching
// as the post-Generate selection coverage, just scored against checklist
// content instead of generated case content, and scoped to the raw search
// terms instead of the checkpoints themselves.
function checklistHaystack(cp) {
  return normalizeForMatch(`${cp.title || ''} ${cp.detail || ''} ${cp.example || ''}`);
}

// Shared by both analyze-stage and generate-stage coverage: scores each
// keyword/function-name term against a set of haystack entries, and splits
// the result by which input (keyword vs function name) each term came
// from — so the person can see those two inputs' contributions separately,
// not just one blended percentage.
function buildTermCoverage(keyword, functionName, haystackEntries) {
  const keywordTerms = buildKeywordCandidates(keyword, '').map(c => ({ ...c, source: 'keyword' }));
  const functionNameTerms = functionName ? buildKeywordCandidates(functionName, '').map(c => ({ ...c, source: 'functionName' })) : [];
  const allTerms = [...keywordTerms, ...functionNameTerms];
  const items = allTerms.map(t => {
    const result = isCandidateCoveredDetailed(t.term, t.term, haystackEntries);
    return {
      term: t.term,
      source: t.source,
      covered: result.covered,
      matchedCheckpointId: result.match ? result.match.id : null,
      matchedCheckpointTitle: result.match ? result.match.title : null,
      matchedCheckpointCategory: result.match ? result.match.categoryLabel : null,
      matchScore: Math.round((result.score || 0) * 100)
    };
  });
  const summarize = (arr) => ({
    total: arr.length,
    covered: arr.filter(i => i.covered).length,
    coveragePercent: arr.length ? Math.round((arr.filter(i => i.covered).length / arr.length) * 100) : null
  });
  return {
    totalTerms: items.length,
    coveredTerms: items.filter(i => i.covered).length,
    coveragePercent: items.length ? Math.round((items.filter(i => i.covered).length / items.length) * 100) : null,
    keyword: summarize(items.filter(i => i.source === 'keyword')),
    functionName: summarize(items.filter(i => i.source === 'functionName')),
    items
  };
}

function buildAnalysisCoverage(keyword, functionName, checklist) {
  const haystackEntries = (checklist || []).map(cp => ({ text: checklistHaystack(cp), ref: cp }));
  return buildTermCoverage(keyword, functionName, haystackEntries);
}

// deriveChecklistTitle/deriveChecklistExample below are still used by the
// deterministic (no-AI) fallback path, which is inherently limited to a
// fixed set of categories it knows how to regex-extract. The AI-driven path
// (buildChecklistFromAnalysis above) reads title/detail/example directly
// off each dynamic checkpoint instead — it no longer needs a per-category
// field-mapping function like this.

// Fallback title when the AI didn't provide one (or we're on the
// document-only fallback path, which never has titles) — short identifying
// label per category, not a truncated sentence.
function deriveChecklistTitle(category, item) {
  switch (category) {
    case 'businessRules': return shorten(item.rule || '', 30) || 'ビジネスルール';
    case 'errorMessages': return item.code && item.code !== '要確認' ? `エラー ${item.code}` : 'エラーメッセージ';
    case 'screenTransitions': return item.flow || [item.from, item.to].filter(Boolean).join('→') || '画面遷移';
    case 'validationRules': return item.field && item.field !== '要確認' ? item.field : '入力チェック';
    case 'inputFields': return item.name && item.name !== '要確認' ? item.name : '入力項目';
    case 'dbOrInterfaceChecks': return shorten(item.rule || '', 30) || 'DB / IF連携';
    default: return '確認事項';
  }
}

// Fallback example (concrete test-data-shaped text) synthesized from
// whatever structured arrays the deterministic extractor already found
// (boundaryValues/validValues/invalidValues/values) — used when the AI
// didn't supply one, or on the document-only fallback path.
function deriveChecklistExample(category, item) {
  switch (category) {
    case 'validationRules': {
      const valid = item.validValues || item.boundaryValues || [];
      const invalid = item.invalidValues || [];
      const parts = [];
      if (valid.length) parts.push(`正常値: ${valid.slice(0, 3).join(', ')}`);
      if (invalid.length) parts.push(`異常値: ${invalid.slice(0, 3).join(', ')}`);
      return parts.join(' / ');
    }
    case 'inputFields': {
      const v = item.values || [];
      return v.length ? v.slice(0, 4).join(', ') : '';
    }
    case 'errorMessages':
      return item.code && item.code !== '要確認' ? `コード: ${item.code}` : '';
    default:
      return '';
  }
}

// When the Copilot Bridge is unavailable during Analyze, build the same
// {checkpoints: [...]} shape directly from the deterministic (regex-based)
// extraction, so the checklist UI still has real, document-sourced content
// to show — just without AI reasoning/synthesis on top. This fallback is
// inherently limited to the fixed categories the regex extractor knows how
// to look for (it can't freely invent new ones the way the AI path can);
// the AI-driven path above is where full dynamic categorization applies.
function deterministicKnowledgeToAnalysis(knowledge, keyword) {
  const checkpoints = [];
  (knowledge.errorMessages || []).forEach(item => checkpoints.push({
    category: 'エラーメッセージ', categoryEn: 'Error Messages',
    title: deriveChecklistTitle('errorMessages', item),
    detail: [item.message, item.condition].filter(Boolean).join(' — ') || item.context || '',
    example: deriveChecklistExample('errorMessages', item),
    source: item.source || ''
  }));
  (knowledge.screenTransitions || []).forEach(item => checkpoints.push({
    category: '画面遷移', categoryEn: 'Screen Transitions',
    title: deriveChecklistTitle('screenTransitions', { flow: item.flow, from: item.from, to: item.to }),
    detail: item.operation || [item.from, item.to].filter(Boolean).join(' → ') || '',
    example: '', source: item.source || ''
  }));
  (knowledge.validationRules || []).forEach(item => checkpoints.push({
    category: '入力チェック / バリデーション', categoryEn: 'Validation Rules',
    title: deriveChecklistTitle('validationRules', { field: item.field }),
    detail: [item.rule, item.expected].filter(Boolean).join(' → ') || '',
    example: deriveChecklistExample('validationRules', { validValues: item.validValues, boundaryValues: item.boundaryValues, invalidValues: item.invalidValues }),
    source: item.source || ''
  }));
  (knowledge.inputFields || []).forEach(item => checkpoints.push({
    category: '入力項目', categoryEn: 'Input Fields',
    title: deriveChecklistTitle('inputFields', item),
    detail: '', example: deriveChecklistExample('inputFields', item),
    source: item.source || ''
  }));
  (knowledge.dbOrInterfaceChecks || []).forEach(item => checkpoints.push({
    category: 'DB / IF連携', categoryEn: 'DB / Interface Checks',
    title: deriveChecklistTitle('dbOrInterfaceChecks', item),
    detail: item.rule || '', example: '', source: item.source || ''
  }));
  return { targetFunction: keyword, checkpoints, openQuestions: [] };
}

// Stage 2 prompt for the NEW workflow: generates cases scoped to EXACTLY the
// checkpoints the person selected (not the full analysis) — this is what
// keeps Generate from re-covering things the user deliberately unchecked.
// Also explicitly invites extra per-case fields beyond the core schema
// (e.g. a flag/permission column) when the content genuinely calls for one,
// which is what makes the result table's columns data-driven downstream.
function buildSelectionTestCasePrompt(keyword, testType, selectedCheckpoints) {
  const profile = TEST_TYPE_PROFILES[testType] || TEST_TYPE_PROFILES.integration;
  const grouped = {};
  (selectedCheckpoints || []).forEach(cp => {
    const cat = cp.categoryLabel || cp.category || 'その他';
    const line = [cp.title, cp.detail, cp.example ? `(テストデータ例: ${cp.example})` : ''].filter(Boolean).join(' — ');
    (grouped[cat] = grouped[cat] || []).push(line || cp.title || cp.detail || '');
  });
  const groupedText = Object.entries(grouped)
    .map(([cat, items]) => `【${cat}】\n` + items.map((t, i) => `${i + 1}. ${t}`).join('\n'))
    .join('\n\n')
    .slice(0, COPILOT_MAX_KNOWLEDGE_CHARS);

  return `あなたは日本の金融系システム(SONAR/投信注文)のテスト設計者です。
DAT Copilot Bridge Smart QA Engineとして、ユーザーが選択した以下のチェックポイントのみを対象に、${profile.label}ケースを作成してください。

【テストケース種別】
${profile.label}
${profile.caseHint}

【対象機能】
${keyword}

【最重要ルール】
- 以下にリストされたチェックポイントを全てカバーするテストケースを作成してください。リストに無い観点を勝手に追加しないでください。
- 1チェックポイントにつき最低1件のテストケースを作成してください（正常系・異常系の両方が読み取れる場合は分けてください）。
- 各テストケースには基本フィールド（no, category, testItem, precondition, steps, inputData, screenName, expectedResult, expectedScreenConfirmation, priority）を必ず含めてください。
- precondition・inputDataは、確認・入力する項目が複数ある場合は文字列ではなく「項目名（意味）」をキー、値を内容とするJSONオブジェクトにしてください。項目が1つだけの単純なケースでは文字列のままで構いません。
  例（複数項目のinputData）：{"ユーザーID":"validUser01", "パスワード（8〜20文字）":"12345678", "取引管理テーブル.タイムオーバーフラグ（時間超過状態）":"1（時間超過あり）"}
  例（単純なprecondition）："ログイン画面を表示していること"
  例（複数項目のprecondition）：{"ログイン状態":"未ログイン", "口座状態":"開設済み・有効"}
- inputData・expectedResultでDB項目名・フラグ値・区分コード・ステータスコードなど技術的な値を扱う場合は、必ずテーブル名・カラム名（またはそれに準ずる項目名）・業務的な意味を併記してください。形式：「テーブル名.カラム名（意味） = 値（意味）」。
  例1（フラグ）：「取引管理テーブル.タイムオーバーフラグ（時間超過状態） = 1（時間超過あり）」
  例2（区分コード）：「口座管理テーブル.口座区分（口座種別） = 2（特定口座）」
  例3（ステータス）：「注文管理テーブル.処理ステータス（注文の処理状態） = 03（承認待ち）」
  これらは一例であり、上記に挙げていない種類の技術的な値（ID、日時コード、権限区分など）についても同じ考え方を適用してください。"timeoverflag=1" や "status=03" のように、業務的な意味が分からない技術名・値だけを書くことは禁止です。判断基準は「エンドユーザー（非エンジニア）がテスト結果を見て、何を確認しているのか理解できるか」です。テーブル名・カラム名が資料から特定できない場合は、無理に作らず「項目名（意味） = 値（意味）」のように、少なくとも意味の説明は必ず添えてください。
- screenName: このテストケースで実際に確認する画面名を記載してください（例:「定時定額設定確認画面」「ログインエラー画面」）。資料に画面名の記載がない場合は、操作の文脈から妥当な画面名を簡潔に記載してください。
- expectedScreenConfirmation: expectedResult（業務的に何が起きるべきか）とは別に、実際に画面上で目視確認すべき内容を具体的に記載してください（例:「エラーメッセージ『14850』がダイアログで表示される」「確認画面に遷移し、入力内容が反映されている」）。expectedResultと同じ内容の繰り返しにしないでください。
- 基本フィールドで表現しきれない重要な情報がある場合に限り、追加のフィールドとして含めてください。その場合もキー名は日本語の意味が分かる名前にするか、値の中に日本語の説明を含めてください（例: キー名"タイムオーバーフラグ"、または値を"1（時間超過あり）"のように）。英語のみの技術的なキー名・値の単体表記（例: "userFlag": "1"）は禁止です。無理に追加フィールドを作る必要はありません。
- 推測で作らないでください。不明な場合は「要確認」としてください。
- priorityは High / Medium / Low のいずれかにしてください。
- 出力はJSON配列のみ。Markdown禁止。

【選択されたチェックポイント】
${groupedText || 'なし'}

【返却JSON形式の例】
[{"no":"TC001","category":"...","testItem":"...","precondition":{"ログイン状態":"..."},"steps":"...","inputData":{"ユーザーID":"...","パスワード":"..."},"screenName":"...","expectedResult":"...","expectedScreenConfirmation":"...","priority":"High"}]`;
}

// Non-AI fallback for the Generate step: builds one straightforward case per
// selected checkpoint directly from its label/source, so the person still
// gets *something* usable to edit even with the Copilot Bridge unavailable.
function buildCasesFromCheckpointsFallback(selectedCheckpoints, keyword) {
  const displayLabel = shortenForFallback(keyword);
  return (selectedCheckpoints || []).map((cp, i) => {
    const summary = [cp.title, cp.detail].filter(Boolean).join(': ') || cp.title || cp.detail || `${displayLabel}の確認`;
    return {
      no: `TC${String(i + 1).padStart(3, '0')}`,
      category: cp.categoryLabel || cp.category || '確認',
      testItem: shorten(summary, 80),
      precondition: `${displayLabel}画面を表示`,
      steps: '1. 対象画面を表示\n2. 該当項目を操作する',
      inputData: cp.example || cp.source || '-',
      screenName: '要確認',
      expectedResult: `${shorten(summary, 100)}どおりに処理されること`,
      expectedScreenConfirmation: '要確認',
      priority: 'Medium'
    };
  });
}
// ---------------------------------------------------------------------------

// --- Copilot Bridge client --------------------------------------------------
// Calls the local "DAT Copilot Bridge" VS Code extension instead of OpenAI.
// The bridge exposes POST /generate accepting { system, prompt, modelFamily }
// and returns { ok, text } where `text` is the raw model completion text
// (same contract the OpenAI code used to get from choices[0].message.content).
// See /copilot-bridge/extension.js for the server implementation.
async function callCopilotBridge(system, prompt) {
  // Final safety net: whatever the individual piece-level caps above add up
  // to, never send more than this many characters to Copilot. Truncating
  // from the end preserves the instructions/rules block (which comes first
  // in both prompt builders) at the cost of the least-recently-added context
  // (doc excerpt), which is the safest place to lose detail.
  if (prompt.length > COPILOT_MAX_TOTAL_PROMPT_CHARS) {
    prompt = prompt.slice(0, COPILOT_MAX_TOTAL_PROMPT_CHARS) + '\n...(文字数上限のため以降省略)';
  }

  try {
    return await sendToCopilotBridge(system, prompt);
  } catch (err) {
    // The bridge's proactive token-count check reports the exact numbers
    // ("Prompt too large for this Copilot model: X tokens > Y max"). Rather
    // than making the person hand-tune .env for every different Copilot
    // account/model, use that real ratio to shrink the prompt precisely and
    // retry once before giving up.
    const m = /Prompt too large for this Copilot model:\s*(\d+)\s*tokens\s*>\s*(\d+)\s*max/i.exec(err.message || '');
    if (m) {
      const actualTokens = Number(m[1]);
      const maxTokens = Number(m[2]);
      const charsPerToken = prompt.length / actualTokens;
      const targetChars = Math.max(1000, Math.floor(maxTokens * charsPerToken * 0.85)); // 15% safety margin
      if (targetChars < prompt.length) {
        const shrunkPrompt = prompt.slice(0, targetChars) + '\n...(モデルのトークン上限のため以降省略)';
        return await sendToCopilotBridge(system, shrunkPrompt);
      }
    }
    throw err;
  }
}

async function sendToCopilotBridge(system, prompt) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), COPILOT_TIMEOUT_MS);
  try {
    let resp;
    try {
      resp = await fetch(COPILOT_BRIDGE_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ system, prompt, modelFamily: COPILOT_MODEL_FAMILY || undefined }),
        signal: controller.signal
      });
    } catch (networkErr) {
      throw new Error(`DAT Copilot Bridge拡張機能に接続できません（${COPILOT_BRIDGE_URL}）。VS Codeを開き、拡張機能が起動していることを確認してください。詳細: ${networkErr.message}`);
    }

    const bodyText = await resp.text();
    if (!resp.ok) {
      throw new Error(`Copilot Bridge HTTP ${resp.status}: ${bodyText.slice(0, 500)}`);
    }
    let data;
    try { data = JSON.parse(bodyText); } catch (_) { throw new Error('Copilot Bridge did not return valid JSON.'); }
    if (!data.ok) throw new Error(data.message || 'Copilot Bridge returned an error.');
    return data.text || '';
  } finally {
    clearTimeout(timer);
  }
}

async function generateByCopilot(prompt) {
  const text = await callCopilotBridge(
    'You are a Japanese financial-system QA test designer. Return only a valid JSON array. Do not use markdown.',
    prompt
  );
  return JSON.parse(extractJson(text));
}

function extractJsonObject(text) {
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start === -1 || end === -1 || end <= start) throw new Error('AI response is not JSON object.');
  return text.slice(start, end + 1);
}

async function generateByCopilotObject(prompt) {
  const text = await callCopilotBridge(
    'You are a Japanese financial-system QA analyst. Return only a valid JSON object. Do not use markdown.',
    prompt
  );
  return JSON.parse(extractJsonObject(text));
}

// --- Result translation (UI language toggle) --------------------------------
// Translates the actual generated test case content (not just static UI
// labels) between Japanese and English when the person switches languages.
// Priority is mapped locally (cheap, deterministic, no AI round-trip needed);
// everything else goes through the Copilot Bridge as one JSON-in/JSON-out
// translation request.
const PRIORITY_JA_TO_EN = { '高': 'High', '中': 'Medium', '低': 'Low' };
const PRIORITY_EN_TO_JA = { High: '高', Medium: '中', Low: '低' };

function translatePriority(value, targetLang) {
  const v = String(value || '').trim();
  if (targetLang === 'en' && PRIORITY_JA_TO_EN[v]) return PRIORITY_JA_TO_EN[v];
  if (targetLang === 'ja' && PRIORITY_EN_TO_JA[v]) return PRIORITY_EN_TO_JA[v];
  return v; // unrecognized value (e.g. already-edited free text) — leave as-is
}

function buildTranslationPrompt(cases, meta, targetLang) {
  const targetLabel = targetLang === 'en' ? 'English' : 'Japanese (日本語)';
  // Cases now have a data-driven key set — send whatever keys are actually
  // present (minus "no"/"priority", which are handled outside translation:
  // "no" must stay untouched, "priority" is mapped locally via a fixed
  // High/Medium/Low ↔ 高/中/低 lookup, not sent to the model).
  const payload = {
    meta: { keyword: meta.keyword || '' },
    cases: (cases || []).map(c => {
      const { no, priority, ...rest } = c || {};
      return { no, ...rest };
    })
  };
  const payloadText = JSON.stringify(payload, null, 2).slice(0, Math.max(2000, COPILOT_MAX_TOTAL_PROMPT_CHARS - 1200));
  return `You are a professional technical translator for QA test case documents in a Japanese financial-system context.
Translate ONLY the natural-language text values in the JSON below into ${targetLabel}.

Rules:
- Keep the JSON structure and keys EXACTLY as given, including any extra/custom fields beyond the usual ones. Keep "no" values exactly as given (e.g. "TC001") — never translate or renumber them.
- precondition and inputData may be either a plain string OR an object of {label: value} pairs. If it's an object, translate BOTH the labels (keys) and the values into ${targetLabel}, keeping the same number of entries — do not flatten it into a string.
- Do not add, remove, or reorder any case or any field within a case.
- Do NOT translate concrete data literals that must stay exact for testing: user IDs, numeric values, dates, currency amounts, error codes (e.g. E9999, 14840). Translate the surrounding sentence around them, not the literal token itself.
- Preserve line breaks (as \\n) inside multi-line fields like steps.
- Return ONLY the JSON object below, fully translated. No markdown, no explanation, no extra keys.

Input JSON:
${payloadText}

Return the exact same JSON shape (same keys per case, same case count/order), translated into ${targetLabel}:
{"meta":{"keyword":"..."},"cases":[{"no":"...", "...": "..."}]}`;
}

async function translateCasesViaCopilot(cases, meta, targetLang) {
  const prompt = buildTranslationPrompt(cases, meta, targetLang);
  const text = await callCopilotBridge(
    'You are a precise technical translator for QA documents. Return only valid JSON. Do not use markdown.',
    prompt
  );
  const parsed = JSON.parse(extractJsonObject(text));
  const translatedByNo = new Map((parsed.cases || []).map(c => [String(c.no), c]));
  const merged = cases.map((rawOrig, i) => {
    const orig = rawOrig || {};
    const tc = translatedByNo.get(String(orig.no)) || (parsed.cases || [])[i] || {};
    const out = { no: orig.no };
    Object.keys(orig).forEach((key) => {
      if (key === 'no') return;
      if (key === 'priority') { out.priority = translatePriority(orig.priority, targetLang); return; }
      if (GROUPED_FIELDS.includes(key) && orig[key] && typeof orig[key] === 'object' && !Array.isArray(orig[key])) {
        // Translated response should mirror the same {label: value} shape —
        // safeText would otherwise flatten it back into one string and undo
        // the sub-column structure entirely.
        const tcVal = tc[key];
        if (tcVal && typeof tcVal === 'object' && !Array.isArray(tcVal) && Object.keys(tcVal).length) {
          const sub = {};
          Object.entries(tcVal).forEach(([sk, sv]) => { sub[safeText(sk)] = safeText(sv); });
          out[key] = sub;
        } else {
          out[key] = orig[key];
        }
        return;
      }
      out[key] = safeText(tc[key]) || orig[key];
    });
    return out;
  });
  const translatedMeta = {
    keyword: safeText(parsed.meta && parsed.meta.keyword) || meta.keyword || ''
  };
  return { cases: merged, meta: translatedMeta };
}
// ---------------------------------------------------------------------------

function extractJson(text) {
  const start = text.indexOf('[');
  const end = text.lastIndexOf(']');
  if (start === -1 || end === -1 || end <= start) throw new Error('AI response is not JSON array.');
  return text.slice(start, end + 1);
}

function buildDocumentBasedCases(keyword, notes, matches, allRows) {
  // `keyword` is the free-form instruction text (up to 2000 chars) — fine
  // for row matching upstream, but too long to repeat inside every
  // template's testItem/precondition/steps text below. Use a short label
  // for display purposes only.
  const displayLabel = shortenForFallback(keyword);
  const rows = (matches && matches.length ? matches : allRows || []).map(r => r.text).filter(Boolean);
  const buttonRows = pickRows(rows, /(ボタン|リンク|メニュー|画面へ|遷移)/, 5);
  const checkRows = pickRows(rows, /(チェック|必須|未入力|上限|下限|エラー|入力)/, 6);
  const dbRows = pickRows(rows, /(DB|テーブル|登録|更新|削除|レコード)/, 3);
  const normalRows = pickRows(rows, /(正常|表示|確認|完了)/, 4);

  const cases = [];
  const add = (category, item, precondition, steps, inputData, expectedResult, priority='High') => {
    cases.push({
      no: `TC${String(cases.length + 1).padStart(3, '0')}`,
      category, testItem: item, precondition, steps, inputData, expectedResult, priority
    });
  };

  add('画面遷移', `${displayLabel}画面を表示する`, 'ログイン済み', `1. メニューを選択\n2. ${displayLabel}メニューまたは対象リンクをクリック`, '-', `${displayLabel}画面が正常に表示されること`, 'High');

  buttonRows.forEach(r => {
    const label = extractActionLabel(r) || '対象ボタン/リンク';
    add('画面遷移', `${label}押下時の遷移確認`, `${displayLabel}画面を表示`, `1. ${displayLabel}画面を表示\n2. 「${label}」をクリック`, '-', summarizeExpected(r, `${label}押下後、仕様通りの画面へ遷移すること`), 'High');
  });

  checkRows.forEach(r => {
    const item = shorten(r, 38);
    add('入力チェック', item, `${displayLabel}画面を表示`, `1. 該当項目に条件に合わない値を入力\n2. 登録/確認ボタンをクリック`, extractInputData(r), summarizeExpected(r, '仕様に従ったエラーメッセージが表示されること'), 'High');
  });

  normalRows.forEach(r => {
    add('正常系', shorten(r, 38), `${displayLabel}画面を表示`, `1. 各項目を正常に入力\n2. 確認/登録ボタンをクリック`, extractInputData(r), summarizeExpected(r, '正常に処理されること'), 'High');
  });

  dbRows.forEach(r => {
    add('DBチェック', shorten(r, 38), '登録/更新処理後', '1. 対象処理を実行する\n2. DBまたは出力結果を確認する', extractInputData(r), summarizeExpected(r, 'DBの対象テーブルに正しく反映されること'), 'High');
  });

  // Only add generic templates when the uploaded documents did not produce enough evidence-based cases.
  // These are not used to force a fixed count; they only prevent an empty/too-small fallback result.
  const templates = [
    ['入力チェック', '必須項目未入力チェック', `${displayLabel}画面を表示`, '1. 必須項目を空白にする\n2. 登録/確認ボタンをクリック', '必須項目: 空白', '必須項目の未入力エラーが表示されること', 'High'],
    ['入力チェック', '入力値の上限チェック', `${displayLabel}画面を表示`, '1. 上限を超える値を入力\n2. 登録/確認ボタンをクリック', '金額/数量: 上限超過', '上限超過エラーが表示されること', 'High'],
    ['異常系', 'キャンセル/戻るボタン確認', `${displayLabel}画面を表示`, '1. 任意項目を入力\n2. キャンセル/戻るボタンをクリック', '-', '前画面へ戻る、または入力内容が仕様通り保持/破棄されること', 'Medium'],
    ['異常系', 'セッションタイムアウト確認', `${displayLabel}画面を表示`, '1. 一定時間操作しない\n2. 画面操作を実施する', '-', 'セッションタイムアウトメッセージが表示され、ログイン画面へ遷移すること', 'Medium'],
    ['正常系', '確認画面で登録処理を行う', '確認画面を表示', '1. 確認画面の内容を確認\n2. 登録ボタンをクリック', '-', '登録が正常に完了し、完了画面が表示されること', 'High']
  ];
  if (cases.length < 3) {
    for (const t of templates) {
      if (cases.length >= 5) break;
      add(...t);
    }
  }
  // Do not force a fixed number of cases. Return all evidence-based cases generated from the uploaded documents.
  return dedupeCases(cases).slice(0, 120);
}

function pickRows(rows, regex, max) {
  return rows.filter(r => regex.test(r)).slice(0, max);
}
function shorten(s, n) { s = cleanText(s); return s.length > n ? s.slice(0, n - 1) + '…' : s; }
function extractActionLabel(s) {
  const m = String(s).match(/[「\[]?([^「」\[\]|]{1,20}?(?:ボタン|リンク|メニュー|アンカー))[」\]]?/);
  return m ? cleanText(m[1]) : '';
}
function extractInputData(s) {
  const hits = String(s).match(/(金額[:：]?[\d,]+|数量[:：]?[\d,]+|ID|パスワード|必須項目|空白|上限|下限|0円|未入力)/g);
  return hits ? [...new Set(hits)].join('\n') : '-';
}
function summarizeExpected(source, fallback) {
  const s = cleanText(source);
  if (!s) return fallback;
  if (/(表示される|遷移する|登録される|更新される|エラー|メッセージ|確認する|出力する)/.test(s)) {
    return `${shorten(s, 70)} こと`;
  }
  return fallback;
}
function dedupeCases(cases) {
  const seen = new Set();
  return cases.filter(c => {
    const k = c.category + c.testItem;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  }).map((c, i) => ({ ...c, no: `TC${String(i + 1).padStart(3, '0')}` }));
}

// Duplicate removal for AI-generated results shown in the Preview. Keys on
// normalized testItem+expectedResult when present (not raw inputData/steps,
// which legitimately vary case-to-case); falls back to comparing the whole
// object when a case doesn't have those fields (e.g. a fully custom shape).
function dedupeGeneratedCasesDynamic(cases) {
  const seen = new Set();
  return cases.filter(c => {
    const hasCoreFields = c.testItem || c.expectedResult;
    const key = hasCoreFields
      ? normalizeForMatch(`${c.testItem || ''}::${c.expectedResult || ''}`)
      : normalizeForMatch(JSON.stringify(Object.keys(c).sort().reduce((o, k) => (k === 'no' ? o : (o[k] = c[k], o)), {})));
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  }).map((c, i) => ({ ...c, no: `TC${String(i + 1).padStart(3, '0')}` }));
}

// Cases now have a data-driven key set (the AI/fallback can include extra
// fields beyond the core schema, e.g. a permission/flag column) — this only
// does the minimum safe cleanup: string-ify every value and guarantee a "no"
// exists, WITHOUT forcing a fixed set of keys or dropping unknown ones.
function normalizeDynamicCases(items) {
  const arr = Array.isArray(items) ? items : [];
  return arr.slice(0, Number(process.env.MAX_TEST_CASES || 120)).map((x, i) => {
    const obj = {};
    if (x && typeof x === 'object') {
      Object.keys(x).forEach(k => {
        if (GROUPED_FIELDS.includes(k) && x[k] && typeof x[k] === 'object' && !Array.isArray(x[k])) {
          // Keep precondition/inputData as an object (sanitizing each
          // sub-value individually) instead of collapsing it through
          // safeText, which would flatten it back into one string and
          // silently undo the whole point of having sub-columns.
          const sub = {};
          Object.entries(x[k]).forEach(([sk, sv]) => { sub[safeText(sk)] = safeText(sv); });
          obj[k] = sub;
        } else {
          obj[k] = safeText(x[k]);
        }
      });
    }
    obj.no = obj.no || `TC${String(i + 1).padStart(3, '0')}`;
    return obj;
  });
}

// `keyword` is now the free-form instruction text (up to 2000 chars), which
// is fine for AI prompts/document matching but far too long to embed in
// fallback filler text (only used when the AI response is missing a field).
// This derives a short, single-line label for that narrow purpose.
function shortenForFallback(text, max = 30) {
  const oneLine = String(text || '').replace(/\s+/g, ' ').trim();
  return oneLine.length > max ? oneLine.slice(0, max) + '…' : (oneLine || '対象機能');
}

function safeText(v) {
  if (v == null) return '';
  if (Array.isArray(v)) return v.map(safeText).filter(Boolean).join('\n');
  if (typeof v === 'object') return Object.entries(v).map(([k,val]) => `${k}: ${safeText(val)}`).join('\n');
  return String(v).replace(/\[object Object\]/g, '').trim();
}

function inferCategory(x, keyword) {
  const t = `${safeText(x.testItem)} ${safeText(x.steps)} ${safeText(x.expectedResult)} ${keyword}`;
  if (/NISA|非課税/.test(t)) return 'NISA注文チェック';
  if (/日付|先日付|未来日/.test(t)) return '日付チェック';
  if (/エラー|メッセージ/.test(t)) return 'エラー表示';
  if (/DB|テーブル/.test(t)) return 'DBチェック';
  if (/遷移|画面/.test(t)) return '画面遷移';
  return '正常系';
}

// --- V3.2: Coverage Engine -------------------------------------------------
// Compares extracted knowledge candidates against generated test cases using
// character bigram overlap — tolerates paraphrasing and particle changes.
//
// V3.2 improvements vs V2.2:
//  - keyword/notes terms are now included as first-class candidates so
//    user-specified topics always contribute to the coverage score.
//  - Anchor threshold for medium-length field/screen names lowered from 0.7
//    to 0.45 — 0.7 was too strict for 6-16 char Japanese field names that
//    the AI naturally restates with minor variation.
//  - Single-pass anchor+full combined check: if full-text overlap >= 0.30
//    alone, the candidate counts as covered even without a strict anchor hit,
//    because a 30% bigram overlap on a domain-specific sentence already
//    implies strong topical relevance.
function normalizeForMatch(s) {
  return cleanText(s).toLowerCase().replace(/[「」『』\s、。，,\.→]/g, '');
}

function caseHaystack(c) {
  const flat = (v) => (v && typeof v === 'object') ? Object.entries(v).map(([k, val]) => `${k} ${val}`).join(' ') : (v || '');
  return normalizeForMatch(`${c.category} ${c.testItem} ${flat(c.precondition)} ${c.steps} ${flat(c.inputData)} ${c.screenName || ''} ${c.expectedResult} ${c.expectedScreenConfirmation || ''}`);
}

function shingles(text, n = 2) {
  const s = normalizeForMatch(text);
  const grams = new Set();
  for (let i = 0; i <= s.length - n; i++) grams.add(s.slice(i, i + n));
  // Error codes as whole tokens
  const codes = s.match(/\d{3,5}|e\d{3,}/g) || [];
  codes.forEach(c => grams.add(c));
  return grams;
}

function overlapRatio(candidateText, haystackText) {
  const candidateGrams = shingles(candidateText, 2);
  if (!candidateGrams.size) return 0;
  const haystackGrams = shingles(haystackText, 2);
  let overlap = 0;
  candidateGrams.forEach(g => { if (haystackGrams.has(g)) overlap++; });
  return overlap / candidateGrams.size;
}

// Like isCandidateCovered, but returns WHICH haystack entry matched (not
// just true/false) so the caller can show the person the actual evidence —
// "this term matched checkpoint X" — instead of an unexplained percentage
// they have to take on faith. haystackEntries: [{ text: normalized string,
// ref: the original object to report back as evidence }].
function isCandidateCoveredDetailed(anchorText, fullText, haystackEntries) {
  const anchor = normalizeForMatch(anchorText);
  const full = normalizeForMatch(fullText || anchorText);
  if (!anchor || anchor === '要確認') return { covered: false, match: null, score: 0 };

  let best = { covered: false, match: null, score: 0 };
  for (const entry of haystackEntries) {
    const h = entry.text;
    const fullOvr = overlapRatio(full, h);
    let matched = false;
    let score = fullOvr;
    if (fullOvr >= 0.30) {
      matched = true;
    } else {
      let anchorMatches;
      if (anchor.length <= 4) anchorMatches = h.includes(anchor);
      else if (anchor.length <= 16) anchorMatches = overlapRatio(anchor, h) >= 0.45;
      else anchorMatches = overlapRatio(anchor, h) >= 0.22;
      if (anchorMatches) {
        if (full === anchor) { matched = true; score = 1; }
        else if (fullOvr >= 0.20) { matched = true; }
      }
    }
    if (matched && score >= best.score) best = { covered: true, match: entry.ref, score };
  }
  return best;
}

function isCandidateCovered(anchorText, fullText, caseHaystacks) {
  const anchor = normalizeForMatch(anchorText);
  const full = normalizeForMatch(fullText || anchorText);
  if (!anchor || anchor === '要確認') return false;

  return caseHaystacks.some(h => {
    // Fast path: full-text bigram overlap alone is sufficient at 0.30
    // (domain sentences share enough distinctive bigrams to avoid false positives)
    const fullOvr = overlapRatio(full, h);
    if (fullOvr >= 0.30) return true;

    // Anchor gate: require the identifying label to appear before counting overlap
    let anchorMatches;
    if (anchor.length <= 4) {
      anchorMatches = h.includes(anchor);           // exact for codes / very short names
    } else if (anchor.length <= 16) {
      anchorMatches = overlapRatio(anchor, h) >= 0.45; // field/screen names (lowered from 0.7)
    } else {
      anchorMatches = overlapRatio(anchor, h) >= 0.22; // long anchor = itself a sentence
    }
    if (!anchorMatches) return false;
    if (full === anchor) return true;               // anchor-only candidate
    return fullOvr >= 0.20;                         // lenient once anchor is confirmed
  });
}

// Coverage for the Analyze → Select → Generate workflow: of the checkpoints
// the person selected, how many actually show up in the generated cases.
// Reuses the same anchor+bigram matching as the document-wide Coverage
// Engine above, just scoped to the selection instead of every candidate the
// Analyze step found.
function buildSelectionCoverage(selectedCheckpoints, cases, keyword, functionName) {
  const haystacks = (cases || []).map(caseHaystack);
  const items = (selectedCheckpoints || []).map(cp => {
    const anchor = cp.title || cp.detail || '';
    const full = [cp.title, cp.detail, cp.example].filter(Boolean).join(' ');
    const covered = isCandidateCovered(anchor, full, haystacks);
    return { id: cp.id, category: cp.categoryLabel || cp.category, title: cp.title || cp.detail || '要確認', covered };
  });
  const totalSelected = items.length;
  const totalCovered = items.filter(i => i.covered).length;

  // Category-level breakdown — shows which categories are well-represented
  // in the generated cases vs which are lagging, instead of only one
  // blended overall percentage that can hide an entire missing category.
  const categoryOrder = [];
  const categoryTally = new Map();
  items.forEach(i => {
    if (!categoryTally.has(i.category)) { categoryTally.set(i.category, { total: 0, covered: 0 }); categoryOrder.push(i.category); }
    const c = categoryTally.get(i.category);
    c.total++; if (i.covered) c.covered++;
  });
  const categoryBreakdown = categoryOrder.map(category => {
    const c = categoryTally.get(category);
    return { category, total: c.total, covered: c.covered, coveragePercent: c.total ? Math.round((c.covered / c.total) * 100) : null };
  });

  // Keyword vs. Function Name term coverage, scored against the actual
  // GENERATED CASES (not the checklist) — confirms the underlying search
  // terms themselves made it into the deliverable, and keeps the two
  // inputs' contributions visibly separate rather than one merged number.
  const caseHaystackEntries = (cases || []).map((c, idx) => ({
    text: caseHaystack(c),
    ref: { id: c.no || `case-${idx}`, title: c.testItem || c.no || '', categoryLabel: c.category || '' }
  }));
  const keywordFunctionCoverage = buildTermCoverage(keyword, functionName, caseHaystackEntries);

  return {
    totalSelected,
    totalCovered,
    coveragePercent: totalSelected ? Math.round((totalCovered / totalSelected) * 100) : null,
    items,
    categoryBreakdown,
    keywordFunctionCoverage
  };
}

// Label helpers — each returns { anchor, full, display }
function errorMessageLabel(x) {
  const hasCode = x.code && x.code !== '要確認';
  const hasMessage = x.message && x.message !== '要確認';
  const anchor = hasCode ? x.code : (hasMessage ? x.message : '');
  const full = hasMessage ? `${hasCode ? x.code + ' ' : ''}${x.message}` : anchor;
  return { anchor, full, display: full || anchor || '要確認' };
}
function screenTransitionLabel(x) {
  const anchor = x.flow && x.flow !== '要確認' ? x.flow : (x.operation || '');
  const full = `${anchor} ${x.operation && x.operation !== '要確認' ? x.operation : ''}`.trim();
  return { anchor, full: full || anchor, display: full || anchor || '要確認' };
}
function validationRuleLabel(x) {
  const anchor = x.field && x.field !== '要確認' ? x.field : (x.rule ? x.rule.slice(0, 20) : '');
  const full = x.rule || anchor;
  return { anchor, full: full || anchor, display: full || anchor || '要確認' };
}
function dbCheckLabel(x) {
  const full = x.rule || '';
  return { anchor: full, full, display: full || '要確認' };
}

// V3.2: build keyword/notes candidates so user-specified topics always show
// up in the coverage denominator.
function buildKeywordCandidates(keyword, notes) {
  const raw = `${keyword} ${notes || ''}`.trim();
  // Split on Japanese/ASCII delimiters and filter to meaningful chunks
  const primary = raw.split(/[、。，,\s\n]+/).map(s => s.trim()).filter(s => s.length >= 2);
  // Long compound phrases (e.g. "非課税口座の簡易開設") are ALSO split into
  // finer sub-terms on particles/connectors. Without this, a long phrase is
  // scored as a single all-or-nothing candidate — if the checklist covers
  // "非課税口座" in one checkpoint and "簡易開設" in another, neither
  // checkpoint alone reaches the match threshold against the full phrase,
  // and a genuinely-covered topic gets reported as missing.
  const subTerms = [];
  primary.forEach(term => {
    if (term.length > 8) {
      term.split(/の|・|\/|／/).map(s => s.trim()).filter(s => s.length >= 2 && s !== term).forEach(s => subTerms.push(s));
    }
  });
  return [...new Set([...primary, ...subTerms])].map(term => ({ term }));
}

// NOTE: the old document-wide buildCoverageReport()/coverageLogLines() pair
// (which scored ALL extracted knowledge, not just what the user selected)
// was removed here — it's fully superseded by buildAnalysisCoverage() (used
// right after Analyze) and buildSelectionCoverage() (used after Generate).
// ---------------------------------------------------------------------------

// Columns are data-driven: computed from the union of keys actually present
// across the generated cases (core fields first in a sensible reading order,
// then any extra AI-added fields like "userFlag" in first-seen order) —
// never a fixed schema, so a case with an unusual extra field just adds a
// column instead of being silently dropped.
const CORE_FIELD_ORDER = ['no', 'category', 'testItem', 'precondition', 'steps', 'inputData', 'screenName', 'expectedResult', 'expectedScreenConfirmation', 'priority'];
const CORE_FIELD_LABELS = { no: 'No', category: 'Category', testItem: 'Test Item', precondition: 'Precondition', steps: 'Steps', inputData: 'Input Data', screenName: 'Screen Name', expectedResult: 'Expected Result', expectedScreenConfirmation: 'Expected Screen Confirmation', priority: 'Priority' };
const CORE_FIELD_WIDTHS = { no: 10, category: 16, testItem: 32, precondition: 20, steps: 48, inputData: 20, screenName: 20, expectedResult: 38, expectedScreenConfirmation: 30, priority: 12 };
// These two core fields can be either a plain string (one simple condition/
// input) or an object of {label: value} pairs (multiple distinct
// preconditions/inputs) — when an object, they render as merged sub-columns
// under one group header instead of one crowded cell.
const GROUPED_FIELDS = ['precondition', 'inputData'];

// Builds the column plan for both the Excel export and (mirrored on the
// frontend) the preview table: each entry is either a plain single column,
// or — for precondition/inputData when at least one case gave an object
// value — a group with its own sub-key list, rendered as merged
// sub-columns under one header instead of one crowded cell holding every
// precondition/input value squashed together.
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
    label: CORE_FIELD_LABELS[key] || humanizeKey(key),
    subKeys: (GROUPED_FIELDS.includes(key) && subKeys[key].length) ? subKeys[key] : null
  }));
}

// Thin flat-list wrapper kept for anything that just needs column keys
// (not the grouping detail) — e.g. dedupe/coverage code that iterates keys.
function computeDynamicColumns(cases) {
  return buildColumnPlan(cases).map(g => g.key);
}

// Converts a camelCase key like "userFlag" into "User Flag" for column
// headers/table headers when there's no explicit label for it.
function humanizeKey(k) {
  const spaced = String(k).replace(/([a-z0-9])([A-Z])/g, '$1 $2');
  return spaced.charAt(0).toUpperCase() + spaced.slice(1);
}

async function writeExcelDynamic(cases, meta, coverage) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Generated Test Cases');
  ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  const metaRows = [
    ['キーワード', meta.keyword || ''],
    ...(meta.functionName ? [['機能名', meta.functionName]] : []),
    ['テストケース種別', meta.testTypeLabel || ''],
    ['作成日', meta.date || new Date().toISOString().slice(0, 10)],
    ['作成者', meta.author || 'AI Test Case Generator (Phase 1)']
  ];
  const plan = buildColumnPlan(cases);
  const flatWidths = [];
  plan.forEach(g => {
    if (g.subKeys) g.subKeys.forEach(() => flatWidths.push(Math.max(16, Math.round((CORE_FIELD_WIDTHS[g.key] || 24) * 1.4 / g.subKeys.length) + 8)));
    else flatWidths.push(CORE_FIELD_WIDTHS[g.key] || 24);
  });
  const flatColumnCount = flatWidths.length;
  const metaValColspan = Math.max(1, flatColumnCount - 1);

  ws.addRows(metaRows);
  ws.addRow([]);

  // Two header rows: row 1 has the group label (merged across its
  // sub-columns for precondition/inputData, or merged DOWN across both
  // header rows for every plain single column so it doesn't leave an
  // awkward empty cell underneath it); row 2 carries the sub-column labels.
  const headerRow1Values = [];
  const headerRow2Values = [];
  plan.forEach(g => {
    if (g.subKeys) {
      headerRow1Values.push(g.label, ...Array(g.subKeys.length - 1).fill(''));
      headerRow2Values.push(...g.subKeys);
    } else {
      headerRow1Values.push(g.label);
      headerRow2Values.push('');
    }
  });
  const headerRow1Number = ws.addRow(headerRow1Values).number;
  const headerRow2Number = ws.addRow(headerRow2Values).number;

  // Helper: get the flat value for one sub-column position of a grouped
  // field. If the case's value isn't an object (e.g. a simple case only
  // ever needed one plain string), the whole string goes in the FIRST
  // sub-column and the rest are left blank, rather than losing the data.
  function groupedCellValue(caseObj, group, subKeyIndex) {
    const raw = caseObj[group.key];
    if (raw && typeof raw === 'object' && !Array.isArray(raw)) return raw[group.subKeys[subKeyIndex]] ?? '';
    return subKeyIndex === 0 ? (raw ?? '') : '';
  }

  cases.forEach(c => {
    const rowValues = [];
    plan.forEach(g => {
      if (g.subKeys) {
        g.subKeys.forEach((sk, i) => rowValues.push(groupedCellValue(c, g, i)));
      } else {
        const v = c[g.key];
        rowValues.push((v && typeof v === 'object') ? Object.entries(v).map(([k, val]) => `${k}: ${val}`).join(' / ') : (v ?? ''));
      }
    });
    ws.addRow(rowValues);
  });
  ws.columns = flatWidths.map(w => ({ width: w }));

  // Merge each meta row's value cell across the remaining columns. Looping
  // over metaRows.length (instead of hardcoding rows 1-4) means adding or
  // removing a meta row (e.g. Function Name being optional) can't silently
  // merge the wrong rows again — same class of bug as the header-row-number
  // fix below.
  for (let i = 1; i <= metaRows.length; i++) {
    ws.mergeCells(i, 2, i, 1 + metaValColspan);
  }

  // Merge the two header rows: grouped fields get a horizontal merge across
  // their sub-columns on row 1; plain single-column fields get a vertical
  // merge spanning both header rows instead of leaving row 2 blank under
  // them.
  let colCursor = 1;
  plan.forEach(g => {
    if (g.subKeys) {
      ws.mergeCells(headerRow1Number, colCursor, headerRow1Number, colCursor + g.subKeys.length - 1);
      colCursor += g.subKeys.length;
    } else {
      ws.mergeCells(headerRow1Number, colCursor, headerRow2Number, colCursor);
      colCursor += 1;
    }
  });

  const fills = ['D9EAD3', 'D9EAD3', 'D9EAD3', 'D9EAD3', 'FFF2CC', 'CFE2F3', 'FCE5CD', 'EADCF8'];
  ws.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.alignment = { vertical: 'middle', wrapText: true };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });
    if (rowNumber === headerRow1Number || rowNumber === headerRow2Number) {
      row.font = { bold: true };
      row.height = 22;
      row.eachCell((cell, idx) => cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fills[(idx - 1) % fills.length] } });
    } else if (rowNumber <= metaRows.length) {
      row.getCell(1).font = { bold: true };
      row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F3F4F6' } };
    }
  });

  // Coverage Report sheet: summarizes how many of the SELECTED checkpoints
  // actually made it into the generated cases (matches buildSelectionCoverage's
  // shape: totalSelected/totalCovered/coveragePercent/items — this used to
  // still reference the old document-wide buildCoverageReport shape
  // (coverage.sections etc), which no longer exists on the object passed in
  // here and crashed every single generate with "Cannot convert undefined
  // or null to object" from Object.entries(coverage.sections)).
  if (coverage && Array.isArray(coverage.items)) {
    const cs = wb.addWorksheet('Coverage Report');
    cs.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
    cs.addRow(['Selected Checkpoint Coverage - Summary']);
    cs.addRow(['Selected Checkpoints', coverage.totalSelected]);
    cs.addRow(['Covered in Generated Cases', coverage.totalCovered]);
    cs.addRow(['Coverage', coverage.coveragePercent == null ? 'N/A' : `${coverage.coveragePercent}%`]);
    cs.addRow([]);

    const byCategory = new Map();
    coverage.items.forEach(i => {
      const key = i.category || '';
      if (!byCategory.has(key)) byCategory.set(key, []);
      byCategory.get(key).push(i);
    });
    const sectionHeaderRowNumbers = [];
    byCategory.forEach((items, cat) => {
      const covered = items.filter(i => i.covered).length;
      const headerRow = cs.addRow([cat || '(none)', `${covered}/${items.length}`]);
      sectionHeaderRowNumbers.push(headerRow.number);
      cs.addRow(['Checkpoint', 'Covered?']);
      items.forEach(i => cs.addRow([i.title, i.covered ? '✓' : '✗']));
      cs.addRow([]);
    });

    cs.columns = [{ width: 70 }, { width: 20 }];
    cs.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
        cell.alignment = { vertical: 'middle', wrapText: true };
      });
      if (rowNumber === 1) row.font = { bold: true, size: 13 };
      if (sectionHeaderRowNumbers.includes(rowNumber)) {
        row.font = { bold: true };
        row.eachCell((cell) => cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9EAD3' } });
      }
    });
  }

  const fileName = `testcases_${Date.now()}.xlsx`;
  const abs = path.join(GENERATED_DIR, fileName);
  await wb.xlsx.writeFile(abs);
  return `/generated/${fileName}`;
}

const server = app.listen(PORT, () => {
  console.log(`DAT AI Test Case Generator running: http://localhost:${PORT}`);
  // Diagnostic banner: prints the ACTUAL effective Copilot prompt-size
  // config this running process is using. If you edited defaults in
  // server.js or .env.example but these numbers don't match what you
  // expect, you have a stale .env file (env vars always win over code
  // defaults) or an old server process that wasn't actually restarted.
  console.log('--- Copilot Bridge config in effect ---');
  console.log(`  COPILOT_BRIDGE_URL           = ${COPILOT_BRIDGE_URL}`);
  console.log(`  COPILOT_MODEL_FAMILY         = ${COPILOT_MODEL_FAMILY || '(default)'}`);
  console.log(`  COPILOT_MAX_DOC_CHARS        = ${COPILOT_MAX_DOC_CHARS}`);
  console.log(`  COPILOT_MAX_SAMPLE_CHARS     = ${COPILOT_MAX_SAMPLE_CHARS}`);
  console.log(`  COPILOT_MAX_KNOWLEDGE_CHARS  = ${COPILOT_MAX_KNOWLEDGE_CHARS}`);
  console.log(`  COPILOT_MAX_ANALYSIS_CHARS   = ${COPILOT_MAX_ANALYSIS_CHARS}`);
  console.log(`  COPILOT_MAX_TOTAL_PROMPT_CHARS = ${COPILOT_MAX_TOTAL_PROMPT_CHARS}`);
  console.log('----------------------------------------');
});

// --- V2.2: Browser Close Cleanup -------------------------------------------
// start_DAT_AI_Tool.bat launches this server and opens it in the default
// browser. When the user closes the browser/terminal window (Ctrl+C on the
// .bat, or the OS sending SIGTERM/SIGINT), make sure we shut down cleanly:
// stop accepting new connections, clear in-memory state (uploadedFiles,
// latestCases, latestCoverage and the Knowledge Base Cache) so nothing
// lingers if the process is ever kept alive by a supervisor/relaunch, and
// only then exit. This avoids orphaned listeners and makes restart-on-close
// behavior predictable.
let shuttingDown = false;
async function cleanupAndExit(signal) {
  if (shuttingDown) return;
  shuttingDown = true;
  console.log(`\n[Browser Close Cleanup] Received ${signal}. Shutting down DAT AI Test Case Generator...`);

  const clearedCacheEntries = clearKnowledgeBaseCache();
  uploadedFiles = [];
  latestCases = [];
  latestSavedCases = [];
  latestMeta = {};
  latestCoverage = null;
  latestAnalysis = null;
  latestKnowledge = null;
  console.log(`[Browser Close Cleanup] Knowledge Base Cache cleared (${clearedCacheEntries} entr${clearedCacheEntries === 1 ? 'y' : 'ies'}). In-memory state reset.`);

  // Fast-path: tell the DAT Copilot Bridge extension to quit VS Code right
  // away instead of waiting for its watchdog to notice this server is gone.
  // Best-effort only — if the bridge isn't running, this just resolves/fails
  // silently and the bridge's own watchdog is the fallback.
  await notifyCopilotBridgeShutdown();

  server.close(() => {
    console.log('[Browser Close Cleanup] HTTP server closed. Goodbye.');
    process.exit(0);
  });

  // Safety net: force-exit if something keeps the event loop alive.
  setTimeout(() => process.exit(0), 3000).unref();
}

async function notifyCopilotBridgeShutdown() {
  try {
    const shutdownUrl = COPILOT_BRIDGE_URL.replace(/\/generate\/?$/, '/shutdown');
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 1500);
    await fetch(shutdownUrl, { method: 'POST', signal: controller.signal }).catch(() => {});
    clearTimeout(timer);
  } catch (_) { /* ignore — bridge not running */ }
}

['SIGINT', 'SIGTERM'].forEach(sig => process.on(sig, () => cleanupAndExit(sig)));
process.on('exit', () => { /* no-op: real cleanup already done in cleanupAndExit */ });

// Global safety net: an uncaught error in ANY route handler (a bug we
// haven't found yet) should never be allowed to silently kill the whole
// Node process — that's what makes a single broken request show up to the
// person as the entire tool going unreachable ("Failed to fetch" on every
// subsequent action, not just the one that triggered it). Log loudly and
// keep the server running; the request that caused it will fail on its own,
// but everything else keeps working.
process.on('unhandledRejection', (reason) => {
  console.error('[UNHANDLED REJECTION] This would previously have crashed the server:', reason);
});
process.on('uncaughtException', (err) => {
  console.error('[UNCAUGHT EXCEPTION] This would previously have crashed the server:', err);
});
// ---------------------------------------------------------------------------
