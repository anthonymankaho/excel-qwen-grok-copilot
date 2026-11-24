# Store listing draft · 上架文案草稿 (Marketplace / AppSource)

Status 狀態: SUBMISSION-READY DRAFT 2026-11-24 — D2 decided: **Anthony Man** (must equal the verified Partner Center publisher name exactly). Legal pages live; Worker BYO live (Version 54752ffa). Nothing submitted yet — submission is Anthony's.
App Id: `7f3a9c2e-6b14-4e8d-a1f0-9c5d2e8b6a40` · manifest `manifest.cloudflare.xml` v1.5.0.0 · ProviderName `Anthony Man`.

## Title 標題
- EN: Qwen for Excel and Word — your own AI copilot for workbooks and documents
- 繁中: Qwen for Excel and Word — 自備金鑰的活頁簿與文件 AI 助手

## Short description 簡述
- EN: Ask, review and revise Excel workbooks and Word documents with Qwen, using your own API key. Read-only whole-workbook review, approval-gated edits, Undo, and pane-computed finance checks.
- 繁中: 以您自己的 Qwen 金鑰提問、覆核與修改 Excel 活頁簿及 Word 文件。整本唯讀覆核、審批式修改、Undo、pane 計算之財務檢查。

## Long description 詳述
EN:
Qwen for Excel and Word brings a Copilot-style task pane to Excel and Word. Chat answers from your selection or active sheet and can review the entire workbook read-only (sheet inventory, batched profiles, per-voucher debit/credit balance, export-vs-ledger cross-checks, formula-error scan). Plan proposes changes; Agent executes approved steps with independently computed finance checks; destructive edits always wait for a click and Undo reverses applied changes. A Copilot-style sidebar holds Guide, workbook context, cloud sources, saved skills, history and memory on the local edition; the hosted edition ships core + Guide only. Bring your own QwenCloud / Alibaba Cloud Model Studio key: local mode keeps it on your PC; hosted mode remembers it in Office storage or browser storage when available and sends it per request through the Worker, which never stores it. You pay your provider directly; the add-in is free.

繁中:
Qwen for Excel and Word 為 Excel／Word 帶來 Copilot 式工作窗格。Chat 依選取範圍或作用中工作表回答，并可整本唯讀覆核（工作表目錄、分批摘要、每批次借貸平衡、匯出對總帳交叉核對、公式錯誤掃描）。Plan 提出修改；Agent 執行已批准步驟並附獨立計算之財務檢查；破壞性修改必待點擊，Undo 可復原。Copilot 式側欄於本機版容纳 Guide、活頁簿上下文、雲端來源、已存技能、歷史與記憶；託管版僅核心＋Guide。自備 QwenCloud／阿里雲百煉金鑰：本機存於您的 PC（loopback relay），託管存於瀏覽器、每請求傳送、伺服器絕不儲存。費用直接付予供應商；增益集免費。

## Disclosure lines (required) 披露句（必須）
- EN: "Requires your own QwenCloud / Alibaba Cloud Model Studio API key; model usage is billed to your account. Workbook/attachment text is sent to the relay and QwenCloud only when you press Send with Share on. Scanned PDF pages (≤12) are sent to QwenCloud for transcription with your key."
- 繁中:「需您自己的 QwenCloud／阿里雲百煉 API 金鑰；模型費用計於您的帳戶。只有 Share 開啟並按 Send 時，活頁簿／附件文字才送 relay 與 QwenCloud。掃描 PDF 頁（≤12）以您的 key 送 QwenCloud 轉錄。」

## URLs
- Support: https://excel-qwen-grok.anthonym-354.workers.dev/legal/support.html
- Privacy: https://excel-qwen-grok.anthonym-354.workers.dev/legal/privacy.html
- EULA: https://excel-qwen-grok.anthonym-354.workers.dev/legal/eula.html

## Assets checklist 資產清單 (per Microsoft publishing checklist)
- [x] Store logos generated: `store-assets/logo-96.png` (+64/32); manifest icons 32/64 live (`npm run store-assets` regenerates)
- [x] Pane UI screenshots 1280×800: `store-assets/screenshot-1-sidebar-guide.png`, `screenshot-2-attach-menu.png`, `screenshot-3-settings-byo.png`
- [ ] **Office-in-context screenshots (Excel desktop / Excel web / Word) — Anthony captures** (sideload manifest, real workbook); pane-only shots above are supplementary
- [ ] Optional preview video
- [x] Search keywords: Excel, Word, AI, Qwen, copilot, review, reconciliation, finance, BYO key
- [x] Pricing: Free; external account/fee disclosed (QwenCloud BYO)
- [x] Publisher name decided D2 = Anthony Man; **Partner Center verification still Anthony's**
- [x] Legal pages live at /legal/*.html (publisher final review still recommended)

## Out of scope for v1 (disclose, do not claim) v1 不_claim_
OneDrive search, Saved skills, History, Memory, PDF/XLSX/image extraction are LOCAL-relay features; the hosted listing must not claim them. Grok is not offered.
OneDrive 搜尋、已存技能、歷史、記憶、PDF/XLSX/影像提取為本機 relay 功能；託管上架不得_claim_。不提供 Grok。
