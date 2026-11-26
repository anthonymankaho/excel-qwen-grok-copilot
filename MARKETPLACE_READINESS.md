# Microsoft Marketplace readiness · Microsoft Marketplace 上架準備

**Status as of 2026-11-25: PREPARATION ONLY — DO NOT SUBMIT YET.** No offer has been created or submitted. The current listing and reviewer notes are drafts, not certification-ready materials.

**截至 2026-11-25：僅在準備階段，暫勿提交。** 尚未建立或提交 offer。現有上架文案及審查員指引僅屬草稿，未達認證提交條件。

## Current state · 目前狀況

- Product: **Qwen for Excel and Word**, add-in-only XML manifest, App ID `7f3a9c2e-6b14-4e8d-a1f0-9c5d2e8b6a40`.
- Current hosted manifest: `manifest.cloudflare.xml`, v1.5.0.0, `ProviderName` = `Anthony Man`, default locale `en-US`.
- Production health was read without sending a model request on 2026-11-26: Worker build `2026-11-25.3`, `byo:true`, `configured.qwen:false`. The 2026-11-25 read was build `2026-11-24.10`.
- A provider-terms warning was added to the local pane, support, and EULA sources during this preparation. These source changes have **not** been deployed; the live Worker still serves build `.10` without the new warning.
- The deployed add-in uses BYO key: each user enters their own key; no user key is intentionally shared and no provider key is embedded in the Worker.
- The three `.html` legal URLs in the manifest/listing respond with HTTP 307 and then HTTP 200 at extensionless canonical paths. Use the canonical URLs directly in the submission package.
- Store assets contain the logo and three pane-only screenshots. Excel desktop, Excel web, and Word in-context screenshots are still missing.
- QA evidence: desktop/Word flows were previously reported complete; Anthony confirmed the Excel web key remained saved after reopening and a basic `hi` request worked. A complete live month-end / whole-workbook review after the latest agent changes is still outstanding. This does not close Step 5.

## Release gates · 發行前置條件

### P0 — Qwen API and credential handling

The current default endpoint is Alibaba Cloud Model Studio **Token Plan**. Alibaba's current policy limits Token Plan to interactive use through supported coding/agent tools and expressly excludes custom application backends. That restriction still applies when every user supplies their own key; BYO avoids sharing another person's key, but it does not make this custom add-in an allowed Token Plan client. The successful `hi` smoke test proved connectivity only, not permitted use.

For normal Model Studio API keys, Alibaba's current guidance says not to configure or use long-term keys in client-side code such as browsers or mobile apps. The hosted pane currently saves each user's long-term key in Office/browser storage and sends it through the Worker on each request. This credential flow must be redesigned or confirmed in writing by Alibaba before public release. Do not publish the current Token Plan endpoint as the supported Marketplace setup.

**Recommended decision path:** keep the user-owned-key model, move credentials out of browser storage into an authenticated, encrypted per-user server-side vault, use the standard regional Model Studio API (not Token Plan), and apply per-account limits and revocation. Confirm with Alibaba that the resulting flow is permitted. This requires account sign-in, server-side key lifecycle, privacy disclosures, cost/abuse controls, and a reviewer test account. If choosing a publisher-funded API instead, the publisher pays model charges and must add account sign-in and usage limits. No architecture or billing choice has been made yet.

Current API-key guidance: [Alibaba Cloud Model Studio — Obtain an API key](https://help.aliyun.com/en/model-studio/get-api-key). Singapore regional endpoints and the production-recommended workspace-specific endpoint: [Regional access information](https://help.aliyun.com/en/model-studio/singapore-regional-access-information). Token Plan restrictions and key rules: [Permissions and usage policies](https://help.aliyun.com/en/model-studio/token-plan-permissions-and-usage).

### P0 — Publisher identity and Partner Center

`ProviderName` is set to `Anthony Man`, but the publisher account has not been confirmed or verified in Partner Center. Microsoft requires an enrolled and verified Microsoft 365 and Copilot publisher. New enrollment uses a Microsoft Entra work/school account; a personal Microsoft 365 Premium subscription alone does not establish this enrollment. Before creating an offer, choose the actual legal publisher (Anthony as an individual/sole proprietor, if eligible, or the employer/company with authorization), verify the account and display name, and ensure it matches the manifest. Microsoft says the publisher selected for a new offer cannot be changed later, so **do not create the offer before this is settled**.

References: [Partner Center enrollment](https://learn.microsoft.com/partner-center/create-account), [Open an Office account](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/open-a-developer-account), [submission guide](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/add-in-submission-guide).

### P0 — Full live quality review

Before submission, complete Step 5 against the release candidate using non-sensitive sample data: Excel desktop, Excel web, and Word; workbook inventory and whole-workbook coverage; a realistically large workbook; a finance validation and its evidence; Share-off behavior; attachments; approved edits and Undo; provider interruption/retry; and error messages. The previous full October month-end run ended before completion, and the user stopped further Excel testing at that time. The basic web `hi` test is not a substitute.

### P1 — Package, listing, and certification materials

- Run Microsoft's Office Add-in manifest/store validation against the final manifest and fix every reported issue. Current manifest version is 1.5.0.0; do not treat source XML parsing as Marketplace validation.
- Microsoft requires Marketplace listing languages to be supported by the manifest. It currently declares only `en-US`; either add and validate the Chinese locale overrides before listing zh-TW, or publish English-only initially. [Localization guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/localization).
- Change manifest/listing legal links to the currently working canonical URLs: `https://excel-qwen-grok.anthonym-354.workers.dev/legal/privacy`, `/legal/eula`, and `/legal/support`. These return HTTP 200; the current `.html` links first redirect with HTTP 307.
- Capture real Office-in-context screenshots (Excel desktop, Excel web, Word); keep the existing pane-only screenshots as supplementary. Microsoft requires at least one listing screenshot.
- Rewrite the long description and privacy/EULA/support pages after the credential and billing design is chosen. State the exact Model Studio product/key type, who pays, data sent, provider-side processing, and account requirements.
- Keep `Copilot` comparisons and competitor-brand keywords out of Marketplace copy. The earlier listing draft used "Copilot-style" and `copilot` keywords; these were removed from `STORE_LISTING.md`. Microsoft's policies prohibit comparative marketing using competitor logos or trademarks, so recheck every final listing field and screenshot caption. The English and Traditional Chinese summary drafts are 86 and 44 characters respectively, within the 100-character summary limit; convert the long description to Marketplace HTML before submission. [Microsoft Marketplace policies](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies).
- Replace `STORE_REVIEWER_NOTES.md` before submission. The current note asks reviewers to provide their own paid-provider key or contact the publisher; Microsoft's guide requires complete test instructions and any required paid-service credentials in certification notes, and says reviewers cannot contact an employee for sign-in details. Provide a restricted, working test route through Partner Center; never store credentials in this repository.
- Reconsider whether the app requires an additional paid service in Partner Center. If the reviewer or user must acquire a paid Model Studio account/key, disclose the charges and provide reviewer test access as required.

References: [Microsoft publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist), [step-by-step submission guide](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/add-in-submission-guide), [publish an Office Add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource).

## Order of work · 建議次序

1. Decide publisher identity and the Qwen billing/credential model.
2. Implement and deploy the supported credential flow; update privacy, EULA, support, listing, and reviewer material to describe the actual release.
3. Complete Step 5 live QA on the hosted release candidate and capture the three Office screenshots.
4. Finalize manifest locale/metadata and canonical legal links; run Microsoft's store validator.
5. Complete Partner Center enrollment and publisher verification; only then create the offer under the selected publisher.
6. Complete pricing, disclosures, test access, and reviewer instructions; submit only after every P0 gate is closed.
7. Track certification; after approval, publish and verify a fresh-account install from Office.

## Decisions required before code changes · 程式碼修改前須決定

1. **Publisher:** Anthony as an individual/sole proprietor, or a company publisher (which requires authorization to act for that legal entity)?
2. **API costs:** keep each user's own key and add a secure per-user credential vault, or have the publisher pay for model usage through a managed service with user accounts and quotas?

These choices affect Partner Center identity, authentication, key storage, operating costs, privacy disclosures, listing text, and review access. No keys or account credentials belong in this document or in chat.
