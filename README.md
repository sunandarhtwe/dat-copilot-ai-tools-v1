# DAT AI Test Case Generator v5.0 - Analyze → Select → Generate

## What's new in v5.0 (complete redesign)

This version replaces the old single-button "Generate" flow with a proper
**Analyze → Select → Generate** workflow, plus folder uploads and a
data-driven result table.

1. **Folder upload (with sub-folders)** — besides picking individual files,
   you can now click "📁 フォルダを選択" and select an entire folder; every
   file inside it, including files in sub-folders, is read in. Folder
   uploads accumulate together with any individually-picked files (use the
   🗑 Clear files link to start over).
2. **Test Case Type selector** — 結合テスト (Integration) / 単体テスト
   (Unit) / 総合テスト (Comprehensive), chosen up front, same as before.
3. **Keyword + Analysis button** — enter a keyword, upload your file(s)/
   folder, then click **🔍 Analysis**. The tool reads the uploaded documents,
   matches them against your keyword, and (via the Copilot Bridge) extracts
   structured candidates — business rules, error messages, screen
   transitions, validation rules, input fields, DB/interface checks.
4. **Selectable checklist** — the analysis result is shown as a checkbox
   list grouped by category, with the *actual extracted text* next to each
   item (e.g. the real validation rule text found in your documents) so you
   can see exactly why each checkpoint was suggested. Everything is checked
   by default; uncheck anything you don't want covered.
5. **Generate from selection** — clicking **✣ Generate Test Cases** builds
   test cases scoped to *only* the checkpoints you left checked — not
   everything the analysis found.
6. **Data-driven result table** — the preview table's columns are no longer
   fixed. They're computed from whatever fields the generated test cases
   actually contain: the usual core fields (No / Category / Test Item /
   Precondition / Steps / Input Data / Expected Result / Priority) always
   appear, and if a case includes something extra and meaningful (e.g. a
   `userFlag` value), that becomes its own column automatically — in the
   on-screen table and in the downloaded Excel file.

Carried over from v4.0 (unchanged): duplicate removal, editable preview,
Save/Cancel with a required-save-before-download guard, add/delete rows with
auto-renumbering "No", and language-aware Excel download (switching 日本語/
English also translates the actual generated content, not just UI labels).

## Architecture

Unchanged from v4.0 — still relies on the **DAT Copilot Bridge** VS Code
extension for GitHub Copilot access. See `copilot-bridge/README.md` for
one-time setup, and the root `start_DAT_Copilot_Tool.bat` for the normal
day-to-day launch (auto-starts VS Code + the bridge if needed, auto-closes
it again when you're done).

```
Browser (public/index.html + script.js)
        │  keyword + docs → Analyze → checklist → select → Generate
        ▼
Node/Express server.js  (this project)
        │  POST /generate  { system, prompt }
        ▼
DAT Copilot Bridge (VS Code extension, /copilot-bridge, port 4321)
        │  vscode.lm.sendRequest(...)
        ▼
GitHub Copilot (signed-in VS Code account)
```

## Setup

1. One-time: run `copilot-bridge/install-extension.bat` (see
   `copilot-bridge/README.md`).
2. `npm install`
3. `cp .env.example .env` (or just run the .bat below, which creates one
   with sensible defaults automatically)
4. Run `start_DAT_Copilot_Tool.bat`, or `npm start` and open
   `http://localhost:3000` manually.

## Fallback behavior

If the DAT Copilot Bridge extension isn't running or GitHub Copilot errors
out, both **Analyze** and **Generate** automatically fall back to
document-based (non-AI) extraction instead of failing outright — you'll
still get a real checklist and real test cases, just without AI reasoning
layered on top. The log panel says so explicitly when this happens.

## Notes

- 会社機密資料・顧客情報を Copilot に送信する前に、必ず会社／DIR 側の利用ルール
  （GitHub Copilot Business/Enterprise の Data retention 設定含む）を確認して
  ください。
- 生成されたテストケース件数は選択したチェックポイント数に応じて変動します。
  固定件数ではありません。
- 根拠がない値やエラーメッセージは推測せず「要確認」と出力します。
