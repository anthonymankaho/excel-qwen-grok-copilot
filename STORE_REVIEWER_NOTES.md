# Reviewer instructions · 審查員指引 (draft — NOT FOR SUBMISSION)

**HOLD as of 2026-11-25.** Do not paste this draft into Partner Center yet. The hosted flow's Token Plan default and browser-stored long-term key need an approved replacement, and a working reviewer test route has not been provisioned. See `MARKETPLACE_READINESS.md`.

**截至 2026-11-25 暫停使用。** 請勿把此草稿貼到 Partner Center。託管版目前的 Token Plan 預設及瀏覽器保存長期 key 尚未改為經確認可用的設計，亦未準備可供審查員使用的測試途徑。詳見 `MARKETPLACE_READINESS.md`。

## Install 安裝
1. After the release gates are cleared, install the final Partner Center package in Excel and Word using the certification instructions supplied with that submission.
2. The current app requires a Qwen API key. If the final release still requires a paid external provider account, the publisher must supply working test access in Partner Center's certification notes. Do not instruct reviewers to contact the publisher or expect them to bring a paid API key. Never commit reviewer credentials to this repository.
   清除發行條件後，依該次 Partner Center 認證指引在 Excel／Word 安裝最終套件。如正式版仍需付費外部服務帳戶，發行者必須在 Partner Center certification notes 提供可用測試權限；不得要求審查員聯絡發行者或自備付費 API key。審查憑證絕不可提交到 repository。

## Core workflows to verify 核心流程
1. **Chat read-only**: open any workbook, select a range, ask "Summarize the selection". Expect an answer citing cells; no writes. Then ask "Please review the whole file" — expect sheet inventory + batched profiles + checks and a coverage statement; still no writes.
2. **Share gate**: untick Share, ask again — expect no workbook data sent (context card shows none).
3. **Plan approval**: switch to Plan, ask "Add a total row". Expect a proposal card; nothing written until Approve/Apply.
4. **Agent + Undo**: Agent mode, Direct off, ask "Format the header bold". Approve, Apply, then Undo last apply — original formatting restored.
5. **Destructive gate**: ask "Delete row 2" — expect the action to wait for a click even with Direct on.
6. **Finance check**: on a journal sheet with Debit/Credit headers, ask "Check debits equal credits" — expect a pane-computed validation card (PASS/FAIL/NEEDS_USER), not model arithmetic.
7. **Error readability**: remove the key and send — expect a 401 message ending with "→" and a bilingual fix.
8. **Rate limit (hosted)**: >60 requests/minute with one key returns 429 with a readable message.

## Sample data 範例資料
Any workbook with headers works; a journal with Voucher/Debit/Credit columns exercises groupBalance. No customer data is required or included.
任何具表頭之活頁簿即可；具 Voucher/Debit/Credit 之日記帳可驗證 groupBalance。不含客戶資料。

## Known limitations (disclosed) 已知限制（已披露）
- Hosted edition: no OneDrive search / skills / history / memory / binary extraction (local-relay features).
- Scanned-PDF transcription needs poppler locally; hosted rejects binary extraction with a clear message.
- Model output may be wrong; finance checks are computed by the pane but are not an audit.
- Rate limits are per Worker isolate (best effort), 60 req/min per key.

## Certification pointers 審核對應
- Data flow: public/legal/privacy.html §1–3. BYO key never stored: worker.mjs `byoKey`/`rateLimited` (no persistence, no logging).
- Destructive defaults: click-gated; Undo present; Chat never writes (tests: tests/taskpane.test.mjs).
- Offline suite: `npm test` (98 tests) + `npm run check`.
