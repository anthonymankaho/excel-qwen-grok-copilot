# SOP 2 — User Guide 說明書

**Qwen + Grok for Excel and Word.** This document covers how to *use* the add-in. For getting it installed, read `SOP-INSTALL.md`.

The pane is a task pane in desktop Excel or Word, or the hosted pane in Excel on the web. You ask in plain language; the model plans; **the pane writes** through a fixed list of actions. The model cannot run arbitrary Office JavaScript.

---

## 1. The three modes and Direct 三種模式與 Direct

Pick the mode **before** you send. The pane defaults to **Agent**.

| Mode | What it does | Writes? |
|---|---|---|
| **Chat** | Answers from the context you sent | **Never**, even with Direct on |
| **Plan** | Shows goal, steps, assumptions, required inputs, risks and checks | Only after **Approve**, then **Apply edits** |
| **Agent** | Plans, acts, reads the cells back, and checks its own work | Safe actions immediately when **Direct** is on |

**Direct** (next to Fast) decides whether Agent writes as it goes:

- **Direct on** — safe actions are written immediately; deletes, clears, renames, network formulas and irreversible changes still wait for your click.
- **Direct off** — you approve the plan, then every action waits for **Apply edits**.

The Agent loop stops at **4 rounds**, on **Stop**, when a check fails, or when it needs you.

---

## 2. Asking and controlling context 發問與控制上下文

- **Chat in Excel** sends a multi-cell selection. A single cell sends the used range of the active sheet. Formulas appear as `value [=formula]`.
- **Plan and Agent** send workbook *metadata* first — workbook name, sheet names, used-range sizes, table names, selection address. Cell values are pulled only when a step asks for a range.
- **Profile a big sheet.** On a sheet with hundreds of rows the model can request a *profile* instead of raw cells: each column's header, filled/blank counts, distinct count, sample values with counts, and numeric sum/min/max. This is how it chooses pivot fields correctly on a 753-row sheet.
- **Word Chat** sends the selection; an empty selection sends the document body.
- The **context card** shows the exact text being sent. **↻ Refresh** re-reads it.
- **Untick Send context** when the sheet or document must not leave the computer.
- **＋ Whole sheet** forces the active Excel sheet for the next question only.
- A truncation note means the read stopped at 150 rows × 30 columns, or 45,000 characters. Ask about a smaller range.
- **Enter** sends; **Shift+Enter** adds a line. Quick actions: **Summarize**, **Explain**, **Write formula**, **Revise**.
- The send button becomes **Stop** while streaming. **New chat** clears history and the audit trail.
- **Fast** turns Qwen thinking off and sets Grok reasoning to low. Leave it off for a formula audit or a finance plan.
- **Auto** only uses a provider whose key is already loaded, and tries the other loaded provider once if the first call fails.
- Only the latest question carries file context; older turns keep their prose. A request sends at most 24 messages.

---

## 3. What it can change 它可以改什麼

**In Excel:**

- Values and formulas on a cell or a range
- Bold, italic, fill, font colour, number format, alignment, wrap
- Clear a range
- Insert/delete rows, insert columns, add or rename a sheet, autofit, sort
- **Charts — 30 types**: column, bar, line, pie, area, doughnut, scatter, radar, bubble, surface, stock, and their stacked/100%/marker variants
- Tables with a name, and **calculated table columns**
- **Native PivotTables** — row, column and filter fields; value fields with Sum, Count, Average, Max, Min, CountNumbers, Product, StandardDeviation or Variance; Tabular/Compact/Outline layout; grand totals; destination sheet created if missing
- **Filters** — value lists, custom criteria such as `>1000`, top/bottom items or percent, cell or font colour
- **Conditional highlighting** — cell-value rules (greater than, between, equal to, …) with fill, font colour and bold
- **Borders** — any edge, outline or all; style, weight, colour
- **Freeze panes** and unfreeze
- **Dropdowns** (data validation lists)
- **Named ranges**
- **Hide/unhide** rows, columns and sheets
- **Replace every match** on a sheet

**In Word:** replace the selection, insert text at start or end, bold/italic on the selection, the styles Heading 1/2/3, Title and Normal, and **replace every match** in the whole document.

### Example asks 範例提問

> Draft a native PivotTable for Combined Aging (USD), source A5:Q752, rows Entity, filter Currency, values Sum of USD O/S and each ageing bucket plus Credits (USD), on a new sheet.

> Freeze the header row, put thin borders around A1:Q753, and highlight any Total USD over 10000 in red.

> Add a dropdown of MUK / MCA / GPL in column B and name Q2:Q753 as AgingUSD.

---

## 4. What waits for your click 什麼要等你點擊

Even with **Direct on**, these are held for **Apply edits** / **Apply risky changes**:

- Delete rows, clear a range, rename a sheet
- Formulas that can reach the network or run legacy calls: `WEBSERVICE`, `FILTERXML`, `RTD`, `REGISTER.ID`, `CALL`, `DDE`, `HYPERLINK`
- **The four irreversible changes** — replace every match (sheet or document), clear all filters, clear all conditional formats, clear data validation. Their card rows say **"not reversible by Undo"**.

---

## 5. Reading the cards 讀懂卡片

**Revision card** — lists every action written, or every action waiting for a click. `REVISED n CHANGES` means it wrote; `PROPOSED REVISION · APPLY WITH YOUR CLICK` means it did not. An older single-cell reply shows `sheet!cell: old → new`; `(empty)` means the cell was blank, `(sheet missing)` means nothing was written.

**Validation card** — for finance work. Debit = credit uses the Debit and Credit headers, or the columns set under **Finance checks** in Settings. Default tolerance 0.01. If those columns are not identified the card says **NEEDS_USER** and does *not* call the difference a pass or a fail. Formula errors, blank required fields, duplicate keys and a moved amount fail when their columns are known; a changed row count is a warning.

**Audit card** — ends every turn. It records the request, provider, context reads, model decisions, actions, checks and stop reason. **It contains no API key.** `Execute FAILED` now carries the real error text, and rejected actions appear as `N rejected as unusable`. **New chat** clears it.

**If a batch fails**, the pane tells the model the actual error so it can correct itself. Read the audit line rather than guessing.

**↳ Insert last answer into selection** writes the top-left selected cell. One line starting with `=` becomes a formula; anything else is text.

---

## 6. Undo 復原

**Undo last apply** on the card reverses the most recent batch:

- Cell values and formulas return to what was there; a cell that was empty is cleared
- Inserted rows/columns are removed; **deleted rows are restored across the full used width**, up to 200 columns
- A cleared range and a sort are put back
- A new sheet, chart, table, **PivotTable** or named range is deleted; a renamed sheet gets its previous name
- **Autofit is reversed** — every column width and row height is captured first and put back
- Filters, conditional formats, dropdowns, calculated columns, hidden rows/columns/sheets and freeze panes are all restored
- In Word: selection text, bold, italic and style are put back; inserted text is removed when the pane can still find that exact wording

**Not reversible:** the four irreversible changes in §4 — which is exactly why they wait for your click.

Excel's own **Ctrl+Z** still works for edits Excel recorded.

> **Work on a copy for anything substantial.** Undo is per batch, not a full transaction log.
> **重要工作請用副本。** Undo 以批次為單位，不是完整交易記錄。

---

## 7. Keys, URLs and models 金鑰、URL 與模型

Open **Settings**, pick the provider, paste **your own** key, leave **Remember key on this PC** ticked, press **Use key**.

| Key type | URL |
|---|---|
| Qwen **Token Plan** (`sk-sp-…`) | `https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` — the default |
| Regular QwenCloud | `https://<region>.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` |
| Grok | your own xAI key and endpoint |

**Change the URL when you change key type.** A Token Plan key does not work on the pay-as-you-go URL, and vice versa — that mismatch is the usual cause of HTTP 401.

**The model must exist on your plan.** A Token Plan key's catalogue is narrower than the public docs: `qwen-max`, `qwen-turbo` and `qwen-plus` are **not** on it and return `400 Model not exist`. Confirmed working: `qwen4-max` (the pane default), `qwen3.8-max`, `qwen3.7-plus`, `qwen3.6-flash`.

The relay only accepts HTTPS chat-completions on `qwencloudapi.com` or `aliyuncs.com`, and does not follow redirects.

**Remember** stores the key for this Windows user only. **Clear key** removes that copy and the relay copy. A relay restart drops keys held only in memory; with Remember ticked the pane reloads it automatically.

---

## 8. Finance work 財務工作

Journal entries are **drafts**. The pane does not clear a reconciliation difference, does not change source figures to force a total to match, and does not claim debits equal credits unless its own calculated check says so.

When a check fails, **leave the difference in the sheet**, read the audit card, and investigate. Do not ask the model to force the totals together. A check without an identified column is not a pass.

---

## 9. What it cannot do 它做不到什麼

Two limits are properties of the platform, not unfinished work:

- **Python in Excel.** Office.js exposes no API for it — it is a native cloud feature.
- **The native Copilot surface.** A sideloaded task pane cannot become the ribbon-integrated Copilot pane, inline ghost suggestions, or the built-in Analyze Data experience.

Also: message plus context over 50,000 characters is rejected; the Agent loop is capped at 4 rounds; and there is no audit export button yet.

---

## 10. What leaves this computer 什麼會離開這部電腦

The local relay binds to **127.0.0.1** and rejects any other Host or Origin. Nothing is sent before you ask.

A question sends your prompt and the attached context to the provider you picked (or the one Auto selected from loaded keys). Chat context includes cell values and formulas, or the Word selection/document text. Plan and Agent start with metadata and add a range only when a step asks. **From the hosted pane, the same question also passes through Cloudflare.**

A key is never written into the repository or the audit card. Keep the owner's Token Plan key out of chat, files and commits — other people use their own keys.

---

## 11. Usage troubleshooting 使用疑難排解

| What you see | What to do |
|---|---|
| Provider HTTP 401 | Key and URL do not match — §7 |
| `Model not exist` (HTTP 400) | That model is not on your plan. Change it in Settings — §7 |
| Qwen endpoint rejected | Use the Token Plan URL or a regional `…/compatible-mode/v1/chat/completions` |
| Invalid conversation message | Over 50,000 characters. Shorten the message or untick **Send context** |
| Context says truncated | Ask about a smaller range, or let the model profile the sheet |
| `Execute FAILED · …` | Read the reason on the audit line; it now names the actual cause |
| `N rejected as unusable` | The action broke a limit — e.g. a pivot allows 16 value fields and 8 each for rows/columns/filters |
| Stopped: validation failed | Leave the difference. Read the audit card. Do not force the totals |
| Stopped: loop limit reached | The model still wanted to continue after 4 rounds. Ask a narrower question |
| A reply will not finish | Press **Stop**; tick **Fast** if it is slow |
| A change is wrong | **Undo last apply**, or Excel `Ctrl+Z` |
| Auto will not use Grok | No Grok key is loaded. Auto never routes data to a provider without a key |

For relay, sideload and install problems see `SOP-INSTALL.md` §10.

---

## 12. Processing status · 處理狀態

**2026-11-22 — Copilot-parity build verified in real Excel.**

- Native PivotTable creation, sheet profiling and range read-back are **confirmed working** against a live 753-row AR ageing workbook, including an 11-value-field pivot (USD O/S + nine buckets + Credits) with the `GROUP` total row excluded from the source.
- Reaching that took two fix rounds: read-back caps that silently dropped wide ranges, then a pivot value-field cap of 8 that rejected the legitimate 11-field pivot, plus execution errors that were swallowed so `Execute FAILED` carried no reason. Errors now reach both the audit card and the model.
- Tiers 1–3 of the Copilot action list are implemented: pivot, profiling, filters, conditional highlighting, borders, freeze, dropdowns, named ranges, hide, calculated columns, 30 chart types, sheet and Word replace-all. Undo is complete, including autofit and full-width deleted rows.
- Irreversible changes are click-gated and labelled on the card.
- `npm run check` passes and `npm test` is fully green on Windows now that the zip is read in Node rather than through the Unix `unzip` binary.
- Still unreachable: Python in Excel and the native Copilot surface (§9). `GROK_API_KEY` is not set on the hosted Worker, so the hosted pane is Qwen-only.

處理狀態（2026-11-22）：Copilot 看齊建置已在真實 Excel 中驗證。**原生 PivotTable 建立、表剖析與範圍讀回已確認可用**，對象是真實的 753 列 AR ageing 活頁簿，包括一個 11 值欄的 pivot（USD O/S + 九個桶 + Credits），並已把 `GROUP` 總計列排除於來源之外。達成這一點此前經歷兩輪修正：先是靜默丟棄寬範圍的讀回上限，然後是拒絕合理 11 欄 pivot 的 8 欄值上限，以及被吞掉、令 `Execute FAILED` 不帶原因的執行錯誤；現錯誤會同時到達 audit 卡片與模型。Copilot action 清單的 Tier 1–3 均已實作，Undo 亦已完成（包括 autofit 與全寬刪除列）。不可逆改動需點擊批准並在卡片上標明。`npm run check` 通過，且 zip 改以 Node 讀取後 `npm test` 在 Windows 上全綠。仍無法達到：Python in Excel 與原生 Copilot 介面（§9）。hosted Worker 未設定 `GROK_API_KEY`，故 hosted pane 只能用 Qwen。
