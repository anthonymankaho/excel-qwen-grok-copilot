# Step-by-step: publish your first Office add-in · 逐步教學：發布您的第一個 Office 增益集

Companion to `SUBMISSION_STEPS.md` (checklist) and `STORE_LISTING.md` (copy to paste). 2026-11-24.
Everything marked **[YOU]** happens in your browser/account; everything marked **[DONE]** is already prepared in this repo or live.
At each 🚦 gate, send me the result (screenshot or text) and I'll handle the next code/doc fix.

---

## Step 0 — Gather 準備 [YOU, ~15 min]
1. A Microsoft account (personal Outlook/Hotmail is fine) you will own as publisher. 一個您擁有的 Microsoft 帳戶。
2. Government ID + phone for identity verification (Microsoft requires it for individuals). 身份證＋電話作個人驗證。
3. Files ready in repo: `manifest.cloudflare.xml`, `store-assets/logo-96.png`, `STORE_LISTING.md`, `STORE_REVIEWER_NOTES.md`. [DONE]
4. **Capture 3 Office-in-context screenshots** (Microsoft wants the add-in running inside Office, ≥1280×800):
   a. Open https://www.office.com → Excel → blank workbook (Excel on the web).
   b. Insert → 插入 → Add-ins/增益集 → the right pane opens → click **其他增益集 (More add-ins)** at its bottom → dialog tab **我的增益集 (My Add-ins)** → **上傳我的增益集 (Upload My Add-in)** → choose `manifest.cloudflare.xml` from this repo. The hosted pane opens. (zh-TW UI labels verified against owner screenshot 2026-11-24; if the upload button is absent, use the desktop-Excel catalog sideload fallback.)
   c. Settings → paste YOUR Qwen key → Use key (hosted = BYO). [DONE code]
   d. Type a question on a small table → wait for answer → resize window to ≥1280×800 → **Win+Shift+S** full-window shot #1 (chat answer).
   e. Ask "Please review the whole file" on a journal-like sheet → shot #2 (review/coverage or validation card).
   f. Open ☰ sidebar → shot #3 (Guide + cards).
   Save as `store-assets/office-1-chat.png`, `office-2-review.png`, `office-3-sidebar.png`. 🚦 send me the 3 files' names/looks; I'll verify dimensions.
   於 Excel web 上傳 manifest、輸入自己的 key、拍 3 張 ≥1280×800 截圖存 store-assets/。

## Step 1 — Partner Center account + program 註冊 [YOU, verification may take days]
1. Go to https://partner.microsoft.com → Sign in → create account if needed (individual).
2. Dashboard → **Enroll** / "Join programs" → choose **Microsoft 365 and Copilot** (this is the program that publishes Office add-ins to AppSource).
3. Complete identity verification; set the **publisher display name exactly `Anthony Man`** (D2). It must match `manifest.cloudflare.xml` `<ProviderName>` character-for-character.
4. Wait for verification approval email. 🚦 send me "verified" + the exact displayed publisher name; if it differs, I'll update the manifest before you upload.
   註冊 Microsoft 365 and Copilot 計劃；驗證後發布名稱须恰為 "Anthony Man"；不同则我先改 manifest。

## Step 2 — Create the offer 建立 offer [YOU, ~5 min]
1. Dashboard → **Marketplace offers** → **+ New offer** → **Office Add-in**.
2. Offer ID: `qwen-for-excel-and-word` (lowercase, hyphens; permanent). Offer alias/name: `Qwen for Excel and Word`.
3. You land in the offer workspace with left nav: Properties / Offer listing / Packages (or "Upload") / Pricing / Review & publish. Names vary slightly by portal version.
   建 Office Add-in offer；ID 永久、用小寫連字。

## Step 3 — Properties 屬性 [YOU, ~10 min]
1. Category: **Productivity**. Keywords: `Excel, Word, AI, Qwen, copilot, review, reconciliation, finance, BYO key`.
2. **Store logo**: upload `store-assets/logo-96.png`. [DONE asset]
3. **Screenshots**: upload your 3 Office shots first, then the 3 pane shots (`screenshot-1..3`) as supplementary. Order matters: first image is the hero.
4. Optional video: skip for v1.
   類別 Productivity；logo＋截圖（Office 截圖排前）。

## Step 4 — Offer listing 上架頁 [YOU, ~15 min]
1. Language: English (United States). Add **Chinese (Traditional)** if the portal offers it; if not, keep the bilingual text inside the English description (our copy is already EN+繁中).
2. Paste from `STORE_LISTING.md`: title, short description, long description. **Keep the disclosure lines verbatim** (BYO key, Share-gated sending, ≤12 scanned pages) — Microsoft reviewers check data-handling claims.
3. URLs: Privacy `https://excel-qwen-grok.anthonym-354.workers.dev/legal/privacy.html`; Terms `.../legal/eula.html`; Support `.../legal/support.html`. [DONE live]
4. Pricing: **Free**. In "external account/fee" notes state: users need their own QwenCloud/Model Studio key and pay their provider directly.
5. Markets: all available (or your choice).
   貼文案＋披露句（勿改）；三個法律 URL；Free；注明自備 key 與費用。

## Step 5 — Upload manifest 上傳 manifest [YOU, ~5 min] 🚦
1. Packages/Upload page → upload `manifest.cloudflare.xml` (v1.4.0.0). [DONE file]
2. Read the validation report. **If any error: copy the exact message to me** — I fix and you re-upload. Common ones: ProviderName mismatch (Step 1), unreachable icon/URL, version format.
3. Success shows the add-in Id `7f3a9c2e-…` and hosts Workbook+Document.
   上傳 manifest；驗證錯誤原文傳我修。

## Step 6 — Reviewer notes 審查員指引 [YOU, ~5 min]
1. Certification/reviewer notes field: paste the whole `STORE_REVIEWER_NOTES.md`. [DONE file]
2. It tells reviewers: BYO key, how to test each mode, known limitations. If you prefer issuing a temporary test key, create it in QwenCloud console yourself and share via secure channel — never paste keys into the portal notes or repo.
   貼審查員指引；測試 key 如需要由您手動發、安全渠道分享。

## Step 7 — Submit & certify 提交與審核 [YOU click; Microsoft reviews]
1. **Review and publish** → pre-flight checklist → **Submit for certification**.
2. Status page shows stages (validation → certification → publishing). Typical: a few days to ~2 weeks. Rejection is normal: 🚦 send me the findings verbatim; I fix code/docs; you resubmit.
3. Do not change the manifest Id between submissions.
   提交；退件原文傳我；修後重交；Id 不變。

## Step 8 — Publish & verify 發布與驗證 [YOU, ~10 min] 🚦
1. On approval, click **Publish** (or it auto-publishes per offer settings).
2. Copy the live AppSource listing URL.
3. Fresh-account check: sign into Office with a DIFFERENT Microsoft account → Excel web → Insert → Add-ins → **Store** → search "Qwen for Excel and Word" → Add → pane opens → enter that account's own Qwen key → ask one question.
4. 🚦 send me the listing URL + version; I stamp README/memory and we close the launch gate.
   發布；以全新帳戶從 Store 安裝驗證；listing URL＋version 傳我結案。

## Step 9 — After launch 發布後 [ongoing]
- Every store update: bump `<Version>` in manifest, re-upload, redeploy Worker (`npx wrangler deploy`) for code changes.
- Optional hygiene: `npx wrangler secret delete RELAY_TOKEN --name excel-qwen-grok` (unused).
- Costs: listing/certification free; users pay their own QwenCloud; your Worker rides Cloudflare plan/free tier.
- Monitor: Cloudflare analytics + your QwenCloud console (your usage only).

---

## Pitfalls 常見陷阱
- Publisher name ≠ manifest ProviderName → validation fail. Fix before upload (🚦 Step 1).
- Screenshots < 1280×800 or pane-only without Office context → listing rejection. Use Step 0 shots.
- Claiming local-only features (OneDrive/skills/history/memory/PDF extraction) on the hosted listing → rejection. Our copy already excludes them; don't add.
- Missing BYO/fee disclosure → rejection. Keep disclosure lines verbatim.
- Putting any API key in portal fields, notes, or repo → security incident. Never.
- Excel web + strict tracking prevention blocks iframe localStorage → hosted "Remember" cannot persist; pane now says so and keeps the key session-only. Reviewers: enter the key once per session, or allow third-party storage for workers.dev.
- Changing offer Id or manifest Id mid-review → restarts certification. Don't.

## What I handle vs you 分工
- Me: manifest/code/doc fixes, validation-error fixes, rejection fixes, asset regeneration (`npm run store-assets`), README/memory stamps.
- You: account, verification, uploads, paste, submit, publish, fresh-account install, any key issuance.