# Qwen + Grok for Excel and Word

A local task pane for desktop Excel and Word. You ask in plain language. Qwen or Grok plans the work, and the pane revises the open file through a fixed list of actions. On this PC the pane and the relay both live at `http://127.0.0.1:3000`. Excel on the web uses the HTTPS copy at `https://excel-qwen-grok.anthonym-354.workers.dev`. Company-wide deployment still needs an admin install.

## Processing status · 處理狀態

**2026-11-22 — Qwen Token Plan key rotated across the estate.**

The owner's Qwen Token Plan key was rotated. For this repository's Cloudflare Worker (`excel-qwen-grok`):

- `QWEN_API_KEY` is now **set** as a Worker secret, so the hosted pane can call Qwen with no key entry.
- `GROK_API_KEY` is **still not set** — an xAI key is needed before Grok works in the hosted pane. Until then **Auto** stays on Qwen.
- `RELAY_TOKEN` was left untouched.
- `npx wrangler secret list --name excel-qwen-grok` returns exactly `QWEN_API_KEY` and `RELAY_TOKEN`.

The value was pushed from an environment variable through wrangler's stdin. It was never written into this repository, onto a command line, or into chat.

Verified on the same key, for context: it returns HTTP 200 on both Token Plan endpoints (`token-plan.maas.qwencloudapi.com` and `token-plan.ap-southeast-1.maas.aliyuncs.com`), and its model list does **not** include `qwen-max`, `qwen-turbo`, or `qwen-plus` — those return `400 Model not exist`. `qwen4-max`, this pane's Qwen default, is on the list, so no default-model change was needed.

**Pending owner action 待辦:** set `GROK_API_KEY` if Grok is wanted in the hosted pane.

處理狀態（2026-11-22）：Qwen Token Plan key 已輪換。此 repo 的 Cloudflare Worker `excel-qwen-grok` 上，`QWEN_API_KEY` **已設定**，hosted pane 可直接使用 Qwen；`RELAY_TOKEN` 未動；`GROK_API_KEY` **仍未設定**，需要一把 xAI key，在此之前 **Auto** 只會用 Qwen。key 由環境變數經 wrangler stdin 推送，從未寫入 repo、指令列或對話。同一把 key 已驗證在兩個 Token Plan endpoint 回 HTTP 200，但其模型清單**不含** `qwen-max`／`qwen-turbo`／`qwen-plus`（回 `400 Model not exist`）；本 pane 的 Qwen 預設模型 `qwen4-max` 在清單內，故無需更改。待辦：如需 hosted pane 使用 Grok，請設定 `GROK_API_KEY`。

### 2026-11-22 (later) — local install, and the agent read-back fix

**Installed on the Surface Book.** The note above originally said the add-in was absent there. It is now installed: clone at `C:\Users\Anthony Man\excel-qwen-grok-copilot` (outside OneDrive), sideloaded into desktop Excel under `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` (manifest 1.2.0.0), relay on `127.0.0.1:3000`. Auto-start was deliberately **not** installed. Verified live: the pane loaded the key through Settings, and `POST /api/chat` returned a real `qwen4-max` answer.

**Four defects fixed in the agent read-back.** Agent could not review a workbook or draft a pivot: it asked for a range, the pane logged `Execute SUCCESS · 0 actions`, and the model was told `Observed: SUCCESS` while receiving no cells at all. It then correctly refused to guess columns.

| Defect | Was | Now |
|---|---|---|
| `readGrid` width | 20 rows × **8 columns** — a 17-column ageing sheet lost I:Q, so bucket and amount fields were invisible | full requested width, bounded by `READ_MAX_CELLS` |
| Agent read cap | **160 cells** — `Summary!A1:S40` (760) and `Combined Aging (USD)!A1:Q25` (425) were silently dropped | `READ_MAX_CELLS` = 2000 |
| Rejected-range feedback | fell through to `Observed: SUCCESS`, misleading the model | explicit `No range was read: <reason>` plus a `blocked` audit row |
| `normalizePlan` | dropped a prose-string `plan`, leaving empty plan cards | accepts a string as the goal |

A read that exceeds the context budget now carries an explicit `[truncated: ...]` marker instead of cutting silently.

`npm run check` passes. `npm test` → **52 pass / 2 fail**; both failures are the pre-existing `spawn unzip ENOENT` packaging tests (Windows ships no `unzip` binary) and are unrelated to this change. Three new tests cover the wide read, the honest rejection, and the prose plan.

### 2026-11-22 (Tier 1) — native PivotTable and sheet profiling

Copilot-parity work, stage 1 of 3. Two capabilities added, both verified against the real `qwen4-max` model and the official Office.js typings.

- **`pivot` action — a real native PivotTable.** `worksheet.pivotTables.add(name, source, destination)`, then `rowHierarchies` / `columnHierarchies` / `filterHierarchies` / `dataHierarchies`, `layout.layoutType`, grand totals, and `refresh()`. The destination sheet is created when missing. Undo deletes the pivot, and deletes the sheet too when this action created it. This replaces the earlier "pivots are out of scope" limitation — it was a product choice, not a platform one.
- **`profileRequest` — column statistics instead of raw cells.** The pane reads a whole range in one batch (bounded at 2,000 rows × 60 columns) and returns each column's header, filled/blank counts, distinct count, sample values with counts, and numeric sum/min/max. This is what lets the model choose pivot fields on a sheet far larger than the context window.

**Verified end-to-end with the real model:** asked to draft a pivot for `Combined Aging (USD)`, `qwen4-max` replied that it would build a *native PivotTable* and emitted `profileRequest {sheet:"Combined Aging (USD)", range:"A1:Q753"}` — 12,801 cells that the previous 160-cell cap silently dropped. Its prose `plan` string is now retained rather than discarded.

`npm run check` passes. `npm test` → **55 pass / 2 fail**, the same two pre-existing Windows `spawn unzip ENOENT` packaging failures. Three new tests cover profiling output, native pivot creation with fields and aggregation, and rejection of malformed pivot payloads.

**Not reachable, confirmed against the typings:** Python in Excel (no API — "Python" appears only in error-subtype enums) and the native Copilot ribbon/inline UX (a task pane cannot become that surface).

### 2026-11-22 (Tiers 2 and 3) — the rest of the reachable Copilot action list

**Tier 2 — six actions:** `filter` (value lists, custom criteria, top/bottom items or percent, cell/font colour, and clear), `conditionalFormat` (cell-value rules with fill, font colour and bold, and clear), `borders` (any edge, outline or all; style, weight, colour), `freeze` (rows, columns, unfreeze), `replaceAll` on a sheet, and `dataValidation` dropdowns.

**Tier 3 — depth and trust:** `namedRange`, `hide` for rows, columns and sheets, `tableColumn` for calculated columns, chart types expanded from 4 to 30, Word `replaceAll`, and **Undo completed** — autofit now captures and restores every column width and row height, and deleted rows are restored across the full used width instead of only A–H.

**A real defect found and fixed while completing Undo:** `executeClassified` applied safe actions but discarded the undo records and rendered no card, so **Agent mode with Direct on offered no Undo button at all**, contradicting this README. A new applied-card renderer lists what changed and attaches Undo for every batch.

**Irreversible changes are now click-gated.** A `bulk` flag marks replace-all and the three clear operations; `isDestructive` honours it, so they wait for approval even with Direct on, and the card row says "not reversible by Undo".

**Verified with the real model:** asked to freeze the header, border the range, highlight Total USD over 10000, add an MUK/MCA/GPL dropdown and name Q2:Q753, `qwen4-max` emitted **all five** new actions with correct parameters (`GreaterThan 10000 #FFC7CE`, `all Continuous Thin`, `AgingUSD Q2:Q753`).

`npm run check` passes. `npm test` → **62 pass / 2 fail**, still only the pre-existing Windows `spawn unzip ENOENT` packaging failures. Nine new tests cover the Tier 2 and Tier 3 actions, bulk gating, autofit undo, and rejection of malformed payloads.

處理狀態（2026-11-22，Tier 2 及 3）：**Tier 2 六個 action** — `filter`（值清單、自訂條件、top/bottom 項目或百分比、儲存格／字型顏色、清除）、`conditionalFormat`（儲存格值規則配合填色、字型顏色與粗體、清除）、`borders`（任何邊、外框或全部；樣式、粗幼、顏色）、`freeze`（列、欄、取消凍結）、工作表 `replaceAll`、`dataValidation` 下拉清單。**Tier 3 深度與信任** — `namedRange`、`hide`（列／欄／工作表）、`tableColumn` 計算欄、圖表類型由 4 擴至 30、Word `replaceAll`，以及**完成 Undo**：autofit 現會擷取並還原每個欄寬與列高，刪除的列會按整個已用寬度還原而非只限 A–H。**完成 Undo 時發現並修正一個真實缺陷：** `executeClassified` 套用安全 action 後丟棄 undo 記錄且不渲染卡片，令 **Agent 模式開啟 Direct 時完全沒有 Undo 按鈕**，與本 README 所述不符；新增的「已套用」卡片渲染器會列出改動並為每批附上 Undo。**不可逆改動現需點擊批准：** `bulk` 標記取代全部與三個清除操作，`isDestructive` 會尊重它，故即使開啟 Direct 亦等待批准，卡片該行會註明「not reversible by Undo」。**已用真實模型驗證：** 要求凍結表頭、為範圍加框、標示 Total USD 大於 10000、加入 MUK/MCA/GPL 下拉並把 Q2:Q753 命名時，`qwen4-max` 發出**全部五個**新 action 且參數正確。`npm run check` 通過；`npm test` 為 **62 通過／2 失敗**，仍只是那兩個既有的 Windows `spawn unzip ENOENT` 打包失敗。新增九個測試覆蓋 Tier 2 與 Tier 3 action、bulk 審批、autofit undo 與無效 payload 拒絕。

處理狀態（2026-11-22，較後）：**已安裝於 Surface Book** — 上文原本說本機未安裝，現已更正。clone 位於 `C:\Users\Anthony Man\excel-qwen-grok-copilot`（OneDrive 以外），已 sideload 進桌面 Excel，relay 在 `127.0.0.1:3000`；刻意**未**安裝 auto-start。已實測：pane 經 Settings 載入 key，`POST /api/chat` 回傳真實 `qwen4-max` 答案。

**已修正 agent 讀回的四個缺陷。** 原本 Agent 無法審閱活頁簿或起草 pivot：它要求讀取範圍，pane 記錄 `Execute SUCCESS · 0 actions`，並告訴模型 `Observed: SUCCESS`，但模型完全收不到儲存格。它其後正確地拒絕猜測欄位。四個缺陷為：`readGrid` 只讀 8 欄（17 欄的 ageing 表丟失 I:Q，桶與金額欄位看不見）；agent 讀取上限只有 160 格（`Summary!A1:S40` 共 760 格、`Combined Aging (USD)!A1:Q25` 共 425 格被靜默丟棄）；範圍被拒時回饋誤導的 `Observed: SUCCESS`；`normalizePlan` 丟棄字串形式的 plan 令計劃卡片空白。現已分別改為：保留完整欄寬、上限 `READ_MAX_CELLS` = 2000、明確回報「No range was read: 原因」並記入 audit、接受字串 plan。超出 context 預算的讀取現會附上明確的 `[truncated: ...]` 標記。`npm run check` 通過；`npm test` 為 **52 通過／2 失敗**，兩個失敗是既有的 `spawn unzip ENOENT` 打包測試（Windows 沒有 `unzip`），與是次改動無關。新增三個測試覆蓋寬表讀取、誠實拒絕與字串 plan。

處理狀態（2026-11-22，Tier 1）：Copilot 看齊工程第 1 階段（共 3 階段）。新增兩項能力，均已對真實 `qwen4-max` 模型與官方 Office.js typings 核實。**`pivot` action — 真正的原生 PivotTable**：`pivotTables.add(name, source, destination)` 加上 row／column／filter／data hierarchies、layout、總計與 `refresh()`；目的地工作表不存在時自動建立；Undo 會刪除 pivot，若工作表是由本 action 建立則一併刪除。這取代了先前「pivots 超出範圍」的限制 — 那是產品選擇，非平台限制。**`profileRequest` — 以欄位統計代替原始儲存格**：pane 一次批次讀取整個範圍（上限 2,000 列 × 60 欄），回傳每欄的表頭、已填／空白數、相異值數、樣本值與次數，以及數值 sum／min／max。這令模型能在遠超 context 視窗的工作表上正確挑選 pivot 欄位。**已用真實模型端到端驗證：** 要求為 `Combined Aging (USD)` 起草 pivot 時，`qwen4-max` 回覆會建立*原生 PivotTable*，並發出 `profileRequest {sheet:"Combined Aging (USD)", range:"A1:Q753"}` — 即先前 160 格上限會靜默丟棄的 12,801 格；其字串 `plan` 現已保留而非丟棄。`npm run check` 通過；`npm test` 為 **55 通過／2 失敗**，仍是那兩個既有的 Windows `spawn unzip ENOENT` 打包失敗。**已核實無法達到：** Python in Excel（沒有 API，「Python」只出現在錯誤子類型列舉）與原生 Copilot 功能區／行內介面（任務窗格無法變成該介面）。

## What this update revises

This is the Copilot-style revise. The earlier pane sent one prompt, attached a large dump of the active sheet, and applied one batch of actions. This update adds a mode, a plan, a short agent loop, checks calculated from the workbook, and an audit card.

- **Chat, Plan, and Agent** sit in the composer. Chat answers and leaves the file alone. Plan shows the steps and waits for **Approve**. Agent runs the steps.
- **Metadata first.** Plan and Agent send the workbook name, sheet names, used-range size, table names when Excel exposes them, and the selection address. They ask for one range when they need cell values. Chat still sends the selection, or the used range when the selection is a single cell.
- **The model plans. The pane writes.** Actions stay on the existing list: values, formulas, formatting, rows, columns, sheets, autofit, sort, charts, tables, and the Word selection actions. The model cannot call Office JavaScript on its own.
- **The pane reads the result.** After a write, Agent loads the cells again. It does not treat a reply as success until that read comes back.
- **Finance checks are calculated here.** Debit against credit, formula errors (`#N/A`, `#VALUE!`, `#REF!`, and the same family), blank required fields, duplicate keys, row count, and a total that moved. The model cannot mark these as passed. A failed check stops the loop.
- **A failed check is a workpaper, not a silent fix.** The card states the issue, the evidence, a possible cause, the check, and the recommended action. A journal entry is a draft. The pane does not clear a reconciliation difference on its own.
- **Auto** chooses Qwen or Grok only when that key is already loaded. A finance, Plan, or Agent request prefers Grok when that key is present and **Fast** is off. Otherwise it prefers Qwen. If the chosen provider errors, the pane tries the other loaded key once.
- **An audit card** closes each Chat, Plan, and Agent turn: request, mode, provider, plan, context sent, actions, blocked actions, and check results. API keys are left out. **New chat** clears the trail.
- **Direct, Apply, and Undo stay.** With Direct on, Agent writes safe actions immediately. Deletes, clears, renames, and formulas that can reach the network wait for a click. With Direct off, you approve the plan and then press **Apply edits**. **Undo last apply** is on the card. **Stop** cancels the current round. The loop ends after 4 rounds.

Restart the relay and reopen the pane after you install this update. The manifest is 1.2.0.0. Word still needs `npm run start:word` once.

## What it can change

In Excel:

- Values and formulas on a cell or a range
- Bold, italic, fill, font color, number format, alignment, and wrap
- Clear a range
- Insert or delete rows, insert columns, add or rename a sheet, autofit, and sort
- A chart in 30 types (column, bar, line, pie, area, doughnut, scatter, radar, bubble, surface, stock and their stacked variants), and a table with a name
- A **native PivotTable** on a new or existing sheet: row, column and filter fields, value fields with a chosen aggregation (Sum, Count, Average, Max, Min, CountNumbers, Product, StandardDeviation, Variance), Tabular/Compact/Outline layout, and grand totals
- **Filters** — by value list, custom criteria such as `>1000`, top/bottom items or percent, and cell or font colour
- **Conditional highlighting** — cell-value rules (greater than, between, equal to, and the rest) with fill, font colour and bold
- **Borders** on any edge or all edges, with style, weight and colour
- **Freeze panes**, and unhide
- **Find and replace every match** on a sheet
- **Dropdowns** (data validation lists) on a range
- **Named ranges**, **hidden rows, columns and sheets**, and **calculated table columns**

In Word:

- Replace the selection
- Insert text at the start or end of the document
- Bold and italic on the selection
- The styles Heading 1, Heading 2, Heading 3, Title, and Normal
- **Replace every match** of a phrase in the whole document

**Two things are genuinely out of reach, not merely unbuilt.** It cannot run **Python in Excel** — Office.js exposes no API for it; "Python" appears in the typings only inside error-subtype enums. And it cannot become the **native Copilot surface** — the ribbon-integrated pane, inline ghost suggestions and the built-in Analyze Data experience are not something a sideloaded task pane can turn into. Everything else on the Copilot action list is now implemented.

## First time on a computer

1. Install Node.js 20 or newer.
2. Close Excel and Word.
3. Double-click `Setup and Start.bat`. It installs dependencies, installs the Windows logon auto-start, sideloads the add-in into desktop Excel, and starts the relay.
4. The same steps by hand: `npm install`, `npm run autostart`, then `npm start`.
5. Open the pane from **Home > Add-ins**. The add-in name is **Qwen + Grok for Excel and Word**.
6. Open Settings, pick Qwen or Grok, paste **your own** API key, leave **Remember key on this PC** ticked, and press **Use key**.

To open the same add-in in Word, run `npm run start:word` once with Word closed. `npm start` only launches Excel.

## Every later sign-in

Windows starts the relay at logon with no window. A per-minute task named `ExcelQwenGrokRelayEnsure` starts it again if it stops. Open Excel or Word and the pane is ready. If it is missing, use **Home > Add-ins**.

Stop the relay with `npm run relay-stop`. The minute task will start it again unless you first run `npm run autostart-remove`. `npm run stop` removes the sideload. The next launch then needs `Setup and Start.bat` or `npm start`.

After a code change, restart the relay (`npm run relay-stop`, then `npm run relay-ensure`) and close and reopen the pane. A remembered key stays in the pane until you press **Clear key**. Stopping the relay clears the copy that was only in relay memory.

## Using the pane

The context card shows what the next question will send.

- In Chat, a multi-cell Excel selection is sent. A single cell falls back to the used range of the active sheet. Formulas appear as `value [=formula]`. The read stops at the first 150 rows and 30 columns, then at 45,000 characters. A truncation note says so.
- In Plan and Agent, the next question sends workbook metadata: the workbook name, up to 30 sheet names, each used range’s address and size, table names when Excel exposes them, and the selection address. Tick **＋ Whole sheet** when that question should include the active sheet’s grid.
- In Word Chat, the selection is sent. An empty selection sends the document body, with the same character cap. Plan and Agent send a short document outline and wait to be asked for the selection.
- The card refreshes when the selection changes. **↻ Refresh** reads it again.
- Untick **Send context** to ask with no file data.

Ask in the box and press Enter. Shift+Enter adds a line. **Summarize**, **Explain**, **Write formula**, and **Revise** fill a starting prompt. Replies stream. The send button becomes **Stop** while a reply is running. **New chat** clears the conversation and the audit trail. **Fast** skips deep thinking: Qwen turns thinking off, and Grok uses low reasoning effort.

The footer pill sets Chat, Plan, or Agent, and Qwen, Grok, or Auto. Each provider needs its own key. The dot is green when that key is loaded, amber when the relay is up and the key is not, and red when the relay is down.

**↳ Insert last answer into selection** writes the last reply into the top-left cell of the Excel selection. One line that starts with `=` is written as a formula. Anything else is written as text.

History stays in memory. Reloading the pane clears it. Each request sends at most 24 messages, and only the latest question includes the file context.

## Modes

**Chat** sends the selection and answers. Any action list in the reply is ignored, including when Direct is on.

**Plan** inspects metadata, then returns one plan: goal, steps, assumptions, required inputs, risks, and the checks it wants. Proposed actions sit on the same card. Nothing is written until you press **Approve**. If Direct is off, **Approve** is followed by **Apply edits**.

**Agent** uses the same metadata. It can ask for one range, run the safe actions for that step, read those cells back, and calculate the checks. The loop stops at 4 rounds, when you press **Stop**, when the plan needs you, or when a check fails. With Direct off, a step that would write waits for **Approve**, then **Apply edits**.

**Auto**, next to the mode, picks Qwen or Grok only from keys already loaded in the relay. It will not send the workbook to a provider that has no key.

The audit card lists the request, the provider, the context that was sent, the actions, and the check results.

### Checks the pane calculates

| Check | What the pane compares |
|---|---|
| Debit = credit | Debit and Credit headers, or the columns named in Settings. The default tolerance is 0.01. Without a known pair of columns the result is NEEDS_USER, not a pass or a fail. |
| Formula errors | `#N/A`, `#VALUE!`, `#REF!`, `#DIV/0!`, `#NAME?`, `#NULL!`, `#NUM!` |
| Blank required fields | Empty cells in the columns named in the check or in Settings, on rows that otherwise have data. Unknown columns are NEEDS_USER. |
| Duplicate keys | The same key on more than one row, using the named key columns or one recognized header such as Invoice or Voucher. Unknown keys are NEEDS_USER. |
| Row count | Whether the number of data rows changed. A change is a warning. |
| Total movement | Whether the named amount column moved beyond the tolerance. Several unnamed numeric columns are NEEDS_USER. |

When a check fails, the card keeps the difference and tells you to review it. It does not post the entry or force the totals to match.

## Direct revise

Leave **Direct** ticked, next to Fast, when Agent should write safe changes as it goes. The card lists what changed and offers **Undo last apply**.

Turn **Direct** off to hold the plan until **Approve**, and then hold every action until **Apply edits**.

These formulas are never written automatically, even with Direct on: `WEBSERVICE`, `FILTERXML`, `RTD`, `HYPERLINK`, and the same family of calls. The card marks them and waits for **Apply risky changes**. Deletes, clears, and sheet renames also wait for a click in Agent mode.

**Undo last apply** reverses that batch:

- Cell values and formulas go back to what was there. A cell that was empty is cleared.
- Inserted rows and columns are removed. Deleted rows are restored across the **full used width** of the sheet, up to 200 columns. A cleared range is put back.
- A new sheet, chart, table, PivotTable or named range is removed. A renamed sheet gets its previous name. A sort is put back.
- **Autofit is reversed**: each column width and row height is captured first and put back.
- A filter is cleared on the column it was applied to. A conditional format, a dropdown, a calculated table column, and hidden rows, columns or sheets are all restored. Freeze panes go back to their previous position.
- **Not reversible by Undo**, so these always wait for your click: replacing every match on a sheet or in a document, clearing all filters, clearing all conditional formats, and clearing data validation. The card says so on the row.
- In Word, the selection text, bold, italic, and style are put back. Inserted text is removed when the pane can still find that exact wording.

Excel’s own Undo (`Ctrl+Z`) is still available for edits Excel recorded.

A single-cell reply in the older `edits` shape still shows `sheet!cell: old → new` and waits for **Apply edits**.

## Keys and providers

The key is sent only to the relay on this computer. With **Remember key on this PC** ticked, the pane keeps a copy in its own local storage for this Windows user and loads it into the relay each time the pane opens. Untick Remember for a session that forgets the key when the relay stops. **Clear key** removes both copies. Keys are not written into this repository.

| Provider | Endpoint | Key | Default model |
|---|---|---|---|
| Qwen, Token Plan | `https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` | Token Plan key, `sk-sp-…` | `qwen4-max` |
| Qwen, pay as you go | `https://ap-southeast-1.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` | A regular QwenCloud key | Set in Settings |
| Grok | `https://api.x.ai/v1/responses` | Your xAI key | `grok-4.7` |

A Token Plan key works only on the Token Plan URL. A regular QwenCloud key works only on a regional QwenCloud chat-completions URL. Change the URL in Settings when you change key type. The relay refuses any Qwen URL that is not HTTPS chat-completions on `qwencloudapi.com` or `aliyuncs.com`, and it does not follow redirects.

A Token Plan key also carries its own model list, which is narrower than the public Qwen docs. `qwen-max`, `qwen-turbo`, and `qwen-plus` are not on it and return `400 Model not exist` — that reads like a broken URL but is not one. Set the model in Settings to one the key actually serves. `qwen4-max`, `qwen3.8-max`, `qwen3.7-plus`, and `qwen3.6-flash` were confirmed working on 2026-11-22.

The relay listens on `127.0.0.1` only and rejects any other Host or Origin.

## Cloudflare

The same pane is published at `https://excel-qwen-grok.anthonym-354.workers.dev` so Excel on the web can open it. The local server stays the way to run it on this PC.

The manifest to upload in Excel is `manifest.cloudflare.xml` (version 1.3.0.0). It has its own add-in id, so it can sit beside the local sideload. With Microsoft 365 Premium, add it from **Home → Add-ins → Upload My Add-in**.

The relay token is already stored as the Worker secret `RELAY_TOKEN`. In the hosted pane, open Settings, paste that token, and press **Use key**. The chat route rejects a request that does not carry it. A request from another site is rejected too. The Worker also refuses a pasted provider key and still allows only QwenCloud chat-completions URLs.

Qwen and Grok keys are Worker secrets. They are not typed into the pane and they are not stored in this repository.

`QWEN_API_KEY` is already set on the Worker (2026-11-22), so the hosted pane uses Qwen with no key entry. `GROK_API_KEY` is not set yet. To set or replace either one, from the project folder, after `npx wrangler login`:

```bash
npx wrangler secret put QWEN_API_KEY
npx wrangler secret put GROK_API_KEY
```

`npx wrangler secret list --name excel-qwen-grok` shows which ones exist. It prints names only, never values.

To replace the relay token, run `npx wrangler secret put RELAY_TOKEN` with a passphrase of at least 16 characters, then paste the new value in Settings.

A later code change is published with `npm run deploy:cloudflare` after `npx wrangler login`. Workbook text sent from the hosted pane goes to Cloudflare and then to Qwen or Grok.

## Another computer

Each computer runs its own relay and stores its own keys.

1. Download `download/Qwen-Grok-for-Excel-and-Word.zip` and unzip it. Open the folder inside. It has no API keys. `npm run package` rebuilds that zip after a code change.
2. Install Node.js 20 or newer, close Excel, and double-click `Setup and Start.bat`.
3. Enter your own key. Do not use someone else’s Token Plan key. The plan and its quota belong to that person.

## Commands

| Command | What it does |
|---|---|
| `npm start` | Sideload into desktop Excel and start the relay |
| `npm run start:word` | Sideload into desktop Word and start the relay |
| `npm run dev-server` | Run the relay in the foreground |
| `npm run relay-ensure` | Start the hidden relay if it is not already up |
| `npm run relay-stop` | Stop the relay |
| `npm run autostart` | Install logon start and the minute self-heal task |
| `npm run autostart-remove` | Remove logon start and the self-heal task |
| `npm run stop` | Remove the sideload |
| `npm run check` | Syntax-check the relay, the Worker, the agent core, and the task pane |
| `npm test` | Run the offline pane, agent, relay, and Worker tests |
| `npm run deploy:cloudflare` | Publish the pane and relay to Cloudflare |
| `npm run package` | Rebuild `download/Qwen-Grok-for-Excel-and-Word.zip` |

Auto-start is Windows-only. It uses `%APPDATA%`.

## Tests

`npm run check` and `npm test` need no API key. The pane tests stub the page, Excel, Word, and the stream. They cover context, the click-gated edit card, Direct revise, Undo, risky formulas, Word replace of the selection, Chat, Plan, Agent, a debit-and-credit stop, and Auto staying on a provider that has a key. The relay tests cover the loopback lock, the Qwen URL allowlist, Fast mode, the Agent mode note, and stream errors.

Before trusting a change in the real app, start the relay, load a key, ask one question in Chat, and confirm the answer. For a revise, use Plan or Agent, confirm the card, and use Undo once.

## If something is wrong

| What you see | What to do |
|---|---|
| Relay unavailable | Run `npm run relay-ensure`, or wait up to a minute for the self-heal task |
| Port 3000 is in use | Another relay is running. `npm run relay-stop`, then start it again |
| The pane looks old | Restart the relay, then close and reopen the pane. On another computer, `git pull` first |
| Context says truncated | Select a smaller range, or ask about one section at a time |
| Invalid conversation message | The message plus context is over 50,000 characters. Shorten it, or untick **Send context** |
| Provider HTTP 401 | The key and the URL do not match. `sk-sp-…` needs the Token Plan URL. A regular Qwen key needs the pay-as-you-go URL |
| Model not exist (HTTP 400) | The model is not on that key's list. A Token Plan key does not serve `qwen-max`, `qwen-turbo`, or `qwen-plus`. Pick one it does, in Settings |
| Qwen endpoint rejected | Use the Token Plan URL, or `https://<region>.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` |
| A reply takes too long | Press **Stop**. Tick **Fast** for a shorter reply |
| Stopped. A validation failed | Read the audit card. The difference is still in the sheet. Correct it yourself, then ask again |
| A change looks wrong | Press **Undo last apply**, or Excel `Ctrl+Z` |
| The add-in is missing in Word | Run `npm run start:word` once with Word closed |
