# Store listing draft · 上架文案草稿 (Marketplace / AppSource)

Status 狀態: **HOLD — NOT READY FOR SUBMISSION** (2026-11-25). Preparation only; no offer created or submitted. The BYO model is the intended product behavior (each user enters and pays for their own key; keys are not shared), but the current default Token Plan endpoint is not allowed for custom application backends under Alibaba's current policy, and long-term keys must not be persisted in browser code without an approved design. Publisher identity, full live QA, reviewer access, Office screenshots, and final manifest validation are also outstanding. See [`MARKETPLACE_READINESS.md`](MARKETPLACE_READINESS.md).
App Id: `7f3a9c2e-6b14-4e8d-a1f0-9c5d2e8b6a40` · current manifest `manifest.cloudflare.xml` v1.5.0.0 · provisional `ProviderName` `Anthony Man` (not yet Partner Center verified).

## Title 標題
- EN: Qwen for Excel and Word
- 繁中: Qwen for Excel and Word

## Short description 簡述
- EN: Review and revise Excel workbooks and Word documents with Qwen using your own API key.
- 繁中: 以您自己的 Qwen API key 檢視及修改 Excel 活頁簿和 Word 文件。

## Long description 詳述
EN:
Qwen for Excel and Word adds an AI task pane to Excel and Word. Chat answers from your selection or active sheet and can review the entire workbook read-only (sheet inventory, batched profiles, per-voucher debit/credit balance, export-vs-ledger cross-checks, formula-error scan). Plan proposes changes; Agent executes approved steps with independently computed finance checks; destructive edits always wait for a click and Undo reverses applied changes. The local edition includes Guide, workbook context, cloud sources, saved skills, history and memory; the hosted edition ships core + Guide only. Bring your own QwenCloud / Alibaba Cloud Model Studio key: local mode keeps it on your PC; hosted mode remembers it in Office storage or browser storage when available and sends it per request through the Worker, which never stores it. You pay your provider directly; the add-in is free.

繁中:
Qwen for Excel and Word 為 Excel／Word 提供 AI 工作窗格。Chat 依選取範圍或作用中工作表回答，亦可唯讀檢視整本活頁簿（工作表目錄、分批摘要、每批次借貸平衡、匯出對總帳交叉核對、公式錯誤掃描）。Plan 提出修改；Agent 執行已批准步驟並附獨立計算之財務檢查；破壞性修改必待點擊，Undo 可復原。本機版提供 Guide、活頁簿上下文、雲端來源、已存技能、歷史與記憶；託管版僅提供核心功能及 Guide。自備 QwenCloud／阿里雲百煉 key：本機存於您的 PC；託管版會在 Office／瀏覽器儲存可用時保存，並經 Worker 逐次傳送；Worker 不保存 key。您直接向供應商付費；增益集免費。

## Disclosure lines (required) 披露句（必須）
- EN: "Requires your own QwenCloud / Alibaba Cloud Model Studio API key; model usage is billed to your account. Workbook/attachment text is sent to the relay and QwenCloud only when you press Send with Share on. Scanned PDF pages (≤12) are sent to QwenCloud for transcription with your key."
- 繁中:「需您自己的 QwenCloud／阿里雲百煉 API 金鑰；模型費用計於您的帳戶。只有 Share 開啟並按 Send 時，活頁簿／附件文字才送 relay 與 QwenCloud。掃描 PDF 頁（≤12）以您的 key 送 QwenCloud 轉錄。」

## URLs
- Support (canonical): https://excel-qwen-grok.anthonym-354.workers.dev/legal/support
- Privacy (canonical): https://excel-qwen-grok.anthonym-354.workers.dev/legal/privacy
- EULA (canonical): https://excel-qwen-grok.anthonym-354.workers.dev/legal/eula

## Assets checklist 資產清單 (per Microsoft publishing checklist)
- [x] Store logos generated: `store-assets/logo-96.png` (+64/32); manifest icons 32/64 live (`npm run store-assets` regenerates)
- [x] Pane UI screenshots 1280×800: `store-assets/screenshot-1-sidebar-guide.png`, `screenshot-2-attach-menu.png`, `screenshot-3-settings-byo.png`
- [ ] **Office-in-context screenshots (Excel desktop / Excel web / Word) — Anthony captures** (sideload manifest, real workbook); pane-only shots above are supplementary
- [ ] Optional preview video
- [x] Search keywords draft: Excel, Word, AI, Qwen, review, reconciliation, finance, BYO key (remove competitor-brand comparisons)
- [x] Pricing: Free; external account/fee disclosed (QwenCloud BYO)
- [ ] Publisher legal identity and display name verified in Partner Center; manifest currently says Anthony Man
- [ ] Update the production manifest from redirecting `.html` URLs to the canonical URLs above
- [ ] Decide listing locales: the manifest currently declares `en-US` only; add locale overrides before listing Traditional Chinese, or publish English-only
- [ ] Legal pages return HTTP 200 at canonical URLs; review and update disclosures after API credential/billing design is settled
- [ ] Complete the Microsoft 365 and Copilot program enrollment and manifest/store validation
- [ ] Provide a restricted reviewer test route and credentials through Partner Center certification notes (never in this repository)

## Out of scope for v1 (disclose, do not claim) v1 不_claim_
OneDrive search, Saved skills, History, Memory, PDF/XLSX/image extraction are LOCAL-relay features; the hosted listing must not claim them. Grok is not offered.
OneDrive 搜尋、已存技能、歷史、記憶、PDF/XLSX/影像提取為本機 relay 功能；託管上架不得_claim_。不提供 Grok。
