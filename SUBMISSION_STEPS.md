# Partner Center submission steps · 提交步驟 (Anthony executes 由 Anthony 執行)

**Full click-by-click teaching version: `STORE_SUBMISSION_TUTORIAL.md` (Steps 0–9, 🚦 gates, pitfalls).** 完整逐步教學見該檔。

Prepared 2026-11-24. Everything below the line is account/outbound work reserved to Anthony.
Pre-requisites already DONE: BYO Worker live (Version 54752ffa), owner key secret deleted, legal pages live, manifest v1.4.0.0 (ProviderName "Anthony Man"), listing copy + reviewer notes + logos/pane screenshots in repo.

## 1. Enroll 註冊
1. https://partner.microsoft.com → sign in with the Microsoft account you want as publisher.
2. Join the **Microsoft 365 and Copilot** program (Office add-ins). Individual enrollment is free; complete identity verification so the published name equals **Anthony Man** exactly (D2).
   加入 Microsoft 365 and Copilot 計劃；個人免費；完成身份驗證使發布名稱恰為 "Anthony Man"。

## 2. Create the offer 建立 offer
1. New offer → **Office Add-in** (AppSource).
2. Offer ID e.g. `qwen-for-excel-and-word`; alias/name per STORE_LISTING.md title.
3. Upload `manifest.cloudflare.xml` (v1.4.0.0). Validation must pass; fix any schema findings before continuing.
   上傳 manifest；驗證须過。

## 3. Properties / listing 屬性與上架頁
1. Category: Productivity; keywords per STORE_LISTING.md.
2. Paste EN + zh-TW long/short descriptions and the **disclosure lines** (BYO key, Share-gated sending, ≤12 scanned pages).
3. Pricing: **Free**; note external QwenCloud account/fees.
4. Privacy URL / Terms URL / Support URL: the three /legal/*.html URLs (already in manifest).
5. Upload logos (`store-assets/logo-96.png`) and screenshots. Add Office-in-context screenshots you capture; pane-only shots are supplementary.
   上傳 logo／截圖；補您拍的 Office 內截圖。

## 4. Reviewer notes 審查員
Paste `STORE_REVIEWER_NOTES.md`. Reviewer supplies their own Qwen key (BYO); if you prefer to issue a temporary test key, do it manually from the console and share via secure channel — never in the repo.
審查員自備 key；如發臨時測試 key 请手動從 console 發并经安全渠道——不入 repo。

## 5. Submit & certify 提交與審核
1. Submit for certification. Typical Office add-in review: days to ~2 weeks; rejection rounds are normal — fix findings, resubmit.
2. On approval: publish; verify the live AppSource listing URL and a **fresh-account install from Office** (Excel desktop + web). Record listing URL + version in README.
   核准後發布；以全新帳戶從 Office 安裝驗證；README 記錄 listing URL＋version。

## 6. Post-publish 發布後
- Keep `manifest.cloudflare.xml` Version bumped for every store update; redeploy Worker (`npx wrangler deploy`) for code changes.
- Optional: delete the now-unused `RELAY_TOKEN` Worker secret (`npx wrangler secret delete RELAY_TOKEN --name excel-qwen-grok`).
- Monitor: Worker analytics + your QwenCloud console (your own usage only; users pay their own keys).