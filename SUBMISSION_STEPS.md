# Marketplace readiness steps · Marketplace 上架準備步驟

**Current status: HOLD — preparation only.** No Partner Center offer should be created until the publisher and Qwen API/credential decisions in `MARKETPLACE_READINESS.md` are settled.

**目前狀態：暫停提交，僅作準備。** `MARKETPLACE_READINESS.md` 所列發行者身份及 Qwen API／憑證方案未決定前，不應在 Partner Center 建立 offer。

## Gates before submission · 提交前置條件

1. **Publisher.** Complete Partner Center enrollment in the Microsoft 365 and Copilot program with an eligible, verified publisher. Confirm that the legal publisher and display name are authorized and match manifest `ProviderName`. A new enrollment requires a Microsoft Entra work/school account; personal Microsoft 365 Premium by itself is not the publisher program.
2. **API terms and security.** Each user is intended to use their own key. Do not use Token Plan keys with this custom app unless Alibaba explicitly approves the use. Replace the current hosted browser storage for long-term keys or obtain written approval for the design. Choose a supported regional Model Studio API and ensure test access is available to Microsoft reviewers.
3. **Step 5 QA.** Complete live testing of Excel desktop, Excel web, Word, large workbook reviews, finance checks, attachments, Share behavior, approval/Undo, and provider failures on the final hosted build.
4. **Package and listing.** Update manifest publisher/locales/legal URLs; run Microsoft's Office Add-in manifest/store validation; finalize the listing, disclosures, privacy/EULA/support pages, screenshots, and pricing/additional-purchase details.
5. **Reviewer access.** Prepare a usable test account/key/access route and place it only in Partner Center certification notes. Reviewers must not be told to contact the publisher or to supply their own paid API key.
6. **Submit.** Review all offer fields and submit for certification only after every release gate is closed. After approval, publish, then verify a fresh-account install from Office and record the listing URL/version.

## Current package references · 目前套件文件

- Listing draft: `STORE_LISTING.md` — HOLD; not submission-ready.
- Reviewer instructions: `STORE_REVIEWER_NOTES.md` — HOLD; do not paste into certification yet.
- Detailed gates and decisions: `MARKETPLACE_READINESS.md`.
- Production manifest: `manifest.cloudflare.xml` v1.5.0.0 — provisional publisher and locale; legal `.html` URLs currently redirect.

Microsoft references: [Partner Center enrollment](https://learn.microsoft.com/partner-center/create-account), [Marketplace submission guide](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/add-in-submission-guide), [publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist).
