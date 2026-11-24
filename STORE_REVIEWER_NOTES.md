# Reviewer instructions · 審查員指引 (draft 2026-11-24)

## Install 安裝
1. Sideload `manifest.cloudflare.xml` (Excel desktop: shared folder catalog; Excel on the web: Insert > Add-ins > Upload My Add-in). Hosted pane URL: https://excel-qwen-grok.anthonym-354.workers.dev/
2. No publisher credentials are embedded. The reviewer supplies **their own** QwenCloud API key in Settings (BYO). If a temporary test key is preferred, contact the publisher via the Support URL; the publisher issues it manually (never in this repo).
   審查員自備 Qwen key；如需臨時測試 key 经支援頁聯絡發行者手動發出。

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