# SOP 2 — User Guide 說明書

**Qwen for Excel and Word.** This document covers how to *use* the add-in. For getting it installed, read `SOP-INSTALL.md`.

The pane is a task pane in desktop Excel or Word, or the hosted pane in Excel on the web. You ask in plain language; the model plans; **the pane writes** through a fixed list of actions. The model cannot run arbitrary Office JavaScript.

---

## 1. The three modes and Direct 三種模式與 Direct

Pick the mode **before** you send. The pane defaults to **Agent**.

| Mode | What it does | Writes? |
|---|---|---|
| **Chat** | Answers from the context you sent; on request it scans, profiles and reads the whole workbook read-only (say "review the whole file") | **Never**, even with Direct on |
| **Plan** | Shows goal, steps, assumptions, required inputs, risks and checks | Only after **Approve**, then **Apply edits** |
| **Agent** | Plans, acts, reads the cells back, and checks its own work | Safe actions immediately when **Direct** is on |

**Direct** (next to Fast) decides whether Agent writes as it goes:

- **Direct on** — safe actions are written immediately; deletes, clears, renames, network formulas and irreversible changes still wait for your click.
- **Direct off** — you approve the plan, then every action waits for **Apply edits**.

The composer shows the current **Agent round out of 25** and what Agent is doing. The AUDIT card stays hidden. If more work is pending at round 25, choose **Continue Agent · 繼續** to give the same task another 25 rounds in the current chat. Review any pending edits or resolve failed checks and missing inputs before sending a follow-up. **Stop** prevents another step from starting. An Office edit already underway may finish and appear on its own card with **Undo last apply** when available.

輸入區會顯示 **Agent 目前輪數（每批最多 25 輪）**及當前進度。AUDIT 卡保持隱藏。第 25 輪後如仍有工作，按 **Continue Agent · 繼續**，即可在同一對話為原任務增加最多 25 輪。若有待審批修改、檢查失敗或資料缺漏，請先處理再發送後續指示。按 **Stop** 會阻止下一步開始。已開始的 Office 修改仍可能完成，並顯示在修改卡片上。可復原的修改會提供 **Undo last apply**。

---

## 2. Asking and controlling context 發問與控制上下文

- **Chat in Excel** sends a multi-cell selection. A single cell sends the used range of the active sheet. Formulas appear as `value [=formula]`.
- **Plan and Agent** send workbook *metadata* first — workbook name, sheet names, used-range sizes, table names, selection address. Agent chooses the relevant sheets and ranges; the user does not set a technical scope.
- **Profile a big sheet.** Agent can scan up to 2 million requested cells locally in 20,000-cell batches and send a compact profile: covered rows/columns, column counts, sample values, numeric sum/min/max and Excel error-value counts with sample addresses. The profile says **INCOMPLETE** if it reaches the cap. A column sum alone does not prove an accounting reconciliation; inspect the relevant rows and formulas.
- **Word Chat** sends the selection; an empty selection sends the document body.
- Press **Workbook** to open the context card. It shows the exact text being sent. **↻ Refresh** re-reads it. Press **Workbook** again to hide the card. 按 **Workbook** 打開上下文卡，它顯示即將送出的文字。**↻ Refresh** 會重讀。再按 **Workbook** 即隱藏。
- **Share Office file data with AI** is a data-sharing switch, not a sheet-scope setting. Turn it off to prevent workbook/document metadata and Agent-requested ranges from being sent to Qwen. The add-in will stop a requested read and show **File sharing off**. Attachments are controlled separately.
- **Attach files** with **＋**. Text types (HTML/HTM, TXT, MD, CSV, TSV, JSON, XML) are up to 5 MB. PDF, DOCX, XLSX/XLSM, and PNG/JPG/WEBP go through the local relay and are up to 20 MB. Choose up to 3 files. The pane keeps up to 18,000 characters per file and at most 26,000 attachment characters in a question. PDF pages are numbered. Excel shows sheet coverage (`COMPLETE` or `INCOMPLETE` plus the next offset). Word keeps headings and table rows.
- HTML is parsed as text and tables. Scripts, CSS, embedded charts/images, and links are not executed or interpreted. The extracted text and selected workbook context are sent to the chosen Qwen provider when you send a question. Open **Workbook** and review the card before sending. Remove files to omit them from the next question; **New chat** clears the conversation and attachments. HTML 會被讀成文字和表格。指令碼、CSS、內嵌圖表和連結不會執行。送出問題前可按 **Workbook** 核對。移除檔案即不帶入下一題；**New chat** 會清除對話和附件。
- **MCP** (desktop/local relay only): press **Connect** for Personal OneDrive; the system browser opens so you can sign in to your own personal Microsoft account and approve read-only access. The app owner configures the public Microsoft OAuth app ID once; colleagues do not enter it in Settings. Search results and snippets stay in the pane. **Review and add to chat** fetches the selected item as text and adds it to the attachment list; it is sent to Qwen only with your next question. The connector does not edit OneDrive files. No account email address is hardcoded or saved by the add-in. OneDrive review supports PDF, DOCX, XLSX/XLSM, HTML, TXT, MD, CSV, TSV, JSON, and XML, with the same coverage line as an attachment. While Share is on, Agent can search and then read one of those files. The hosted Excel-on-the-web pane does not support this connector.
- OAuth access tokens are held in local relay memory, used in HTTPS requests to the selected source, and cleared when the relay stops; reconnect after a restart.
- **Workbook** at the bottom opens WORKBOOK ACCESS on its own. It is not in the sidebar. Press the button again to hide it. The footer **Share** checkbox still controls whether workbook data is sent. 底部 **Workbook** 會獨立打開 WORKBOOK ACCESS，不在側欄。再按一次即隱藏。底部 **Share** 仍決定是否送出活頁簿資料。
- **＋ Active sheet** is a Chat shortcut for the active Excel sheet on the next question only; it does not review every sheet in the workbook. For a whole-workbook review, ask Chat to "review the whole file": the pane scans every sheet for formula errors, profiles the ledgers, and computes per-voucher balance and export-vs-ledger checks, all read-only.
- A truncation note means the read stopped at 150 rows × 30 columns, or 45,000 characters. Ask about a smaller range.
- **Enter** sends; **Shift+Enter** adds a line. Type the question yourself. The Summarize, Explain, Write formula, and Revise shortcuts are no longer on the pane. **Enter** 送出；**Shift+Enter** 換行。請自己輸入問題。這四個快捷按鈕已從窗格移除。
- The send button becomes **Stop** while streaming. **New chat** clears history and the audit trail.
- **Model**, in the composer and in Settings, is `auto`, `qwen4-max` (default), `qwen3.8-max`, or `qwen3.8-flash`. `auto` lets Qwen choose the model for that question. The choice is remembered on this PC. 作曲列與 Settings 的模型只此四個，預設 `qwen4-max`。`auto` 由 Qwen 為該題選擇模型。選擇會記在這部電腦。
- **Effort** is Off, Low, Medium, High (default), or Extra. The relay sends `reasoning_effort` and does not send `thinking_budget`. Off skips thinking. Effort 為 Off、Low、Medium、High（預設）或 Extra。relay 傳送 `reasoning_effort`，不傳送 `thinking_budget`。Off 關閉思考。
- **Fast** still skips thinking for that question, even when Effort is High or Extra. Leave Fast off and Effort at High for a formula audit or a finance plan. Fast 仍可對該題關閉思考，即使 Effort 是 High 或 Extra。公式覆核或財務計劃請關掉 Fast，Effort 保持 High。
- The pane runs on **Qwen only** (Grok removed 2026-11-24); it sends nothing until the Qwen key is loaded.
- Only the latest question carries workbook and attachment context; older turns keep their prose. A request sends at most 24 messages.

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

**Finance evidence card** — appears when a prompt or its supplied context is finance-related. It lists the workbook ranges and report/document lines sent to the model, then checks bracketed citations such as `[Sheet1!A1:C8]` or `[report.html:L0001-L0008]` against those locations. A match confirms only that the location was supplied; it does not prove that the location supports the claim. Agent validation cards show independently calculated checks and their source range. Chat prose arithmetic is not independently recalculated; when Chat requests finance checks, the pane calculates them over the full ranges and shows the same validation cards.

**財務證據卡** — 當問題或所提供上下文涉及財務時顯示。卡片會列出傳給模型的工作簿範圍及報表／文件行數，並把 `[Sheet1!A1:C8]` 或 `[report.html:L0001-L0008]` 等方括號引用與已提供位置比對。比對成功只代表該位置已提供，不能證明該來源支持答案。Agent 驗證卡會顯示增益集獨立計算的檢查及其來源範圍；Chat 文字中的算術不會由增益集獨立重算，但 Chat 要求財務檢查時，增益集會於完整範圍計算並顯示相同驗證卡。

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

Each user is intended to enter and pay for their own key; keys are not shared. **Current provider-terms hold (2026-11-25): do not use an Alibaba Token Plan key with this custom add-in.** Alibaba limits Token Plan to supported interactive coding/agent tools and excludes custom application backends. The previous HTTP 200/`hi` test proved connectivity only, not permission to use that plan with this add-in. The current default URL below is a legacy test configuration, not an approved Marketplace setup.

The hosted pane also stores the user's long-term key in Office/browser storage. Alibaba's current API-key guidance says not to configure or use long-term keys in browser/mobile client code. The hosted key flow needs to be redesigned or explicitly confirmed by Alibaba before public release. See `MARKETPLACE_READINESS.md`.

| Legacy test key type | URL (reference only; not approved for this custom add-in) |
|---|---|
| Qwen **Token Plan** (`sk-sp-…`) | `https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` |

The supported Model Studio API endpoint depends on the user's region and workspace; use the current [Alibaba Model Studio API-key and regional endpoint guidance](https://help.aliyun.com/en/model-studio/get-api-key) only after the add-in's secure credential flow has been approved and implemented.

The relay only accepts HTTPS chat-completions on `qwencloudapi.com` or `aliyuncs.com`, and does not follow redirects.

**Remember** stores the key for this Windows user only. **Clear key** removes that copy and the relay copy. A relay restart drops keys held only in memory; with Remember ticked the pane reloads it automatically.

---

## 8. Finance work 財務工作

Journal entries are **drafts**. The pane does not clear a reconciliation difference, does not change source figures to force a total to match, and does not claim debits equal credits unless its own calculated check says so.

When a check fails, **leave the difference in the sheet** and investigate. Do not ask the model to force the totals together. A check without an identified column is not a pass.

---

## 9. What it cannot do 它做不到什麼

Two limits are properties of the platform, not unfinished work:

- **Python in Excel.** Office.js exposes no API for it — it is a native cloud feature.
- **The native Copilot surface.** A sideloaded task pane cannot become the ribbon-integrated Copilot pane, inline ghost suggestions, or the built-in Analyze Data experience.
- **GitHub skill install.** In Skills, paste `owner/repo` or a `https://github.com/owner/repo` link to one `SKILL.md`. A public repo can omit the token. Remember stores the token in this browser only. Tick Replace to update an existing skill. The file is saved in the OneDrive skills folder. The pane does not follow a redirect and does not send the token to Qwen. 在 Skills 貼上 `owner/repo` 或指向一個 `SKILL.md` 的 GitHub 連結。公開 repo 可以不填 token。Remember 只把 token 留在這個瀏覽器。勾選 Replace 才更新已有技能。檔案存入 OneDrive 技能資料夾。窗格不跟隨重新導向，也不把 token 送給 Qwen。
- **Attachment formats.** The picker reads HTML/HTM, TXT, MD, CSV, TSV, JSON, XML, PDF, DOCX, XLSX/XLSM, PNG, JPG, and WEBP. PDF, Word, Excel, and images need the local relay. HTML scripts and CSS are ignored. A scanned PDF is transcribed one page at a time, up to 12 pages, with your own Qwen key. OneDrive review uses the same PDF, Word, and Excel reader and states coverage. Agent can search OneDrive and read one chosen file while Share is on. Skills are chosen from the description and the `triggers` line; `qwen-chat-assistant` is not attached to every question. The sidebar **Guide** starts collapsed.

Also: message plus context over 50,000 characters is rejected; the Agent loop is capped at 25 rounds (**Continue Agent · 繼續** resumes); raw range reads page with `offset` and profiles/finance checks cover the full requested range; and there is no audit export button yet.

---

## 10. What leaves this computer 什麼會離開這部電腦

The local relay binds to **127.0.0.1** and rejects any other Host or Origin. Workbook context and source contents are not sent to Qwen until you submit a question. Searching OneDrive sends the query to Microsoft Graph; previewing a result fetches its text to the local pane. The connector only searches and reads; it does not change OneDrive files.

A question sends your prompt, the extracted text from selected attachments and reviewed cloud results, and (when enabled) workbook/document context to Qwen (the only provider). Original cloud files are not uploaded to Qwen; only the text you selected is included. Chat context includes cell values and formulas, or the Word selection/document text. In Excel, Plan and Agent start with workbook metadata and add a range only when a step asks. In Word, Plan and Agent send the open document. A quoted passage stays in that text when the document is long. **From the hosted pane, the same question also passes through Cloudflare.**

A provider key or signed-in account email is never written into the repository or the audit card. Provider API keys are sent to the local relay. Cloud OAuth tokens are held in local relay memory and used in HTTPS requests to the selected source. These credentials are not sent to Qwen. Keep provider keys and OAuth client secrets out of chat, files and commits; `cloud-oauth.json` holds only public app client IDs.

本機 Relay 綁定 **127.0.0.1**，並拒絕其他 Host 或 Origin。工作簿內容和來源檔案文字不會在提交問題前傳給 Qwen。搜尋 OneDrive 會把搜尋字詞傳給 Microsoft Graph；預覽結果會把文字讀取到本機窗格。本機連接器只會搜尋／讀取，不會修改 OneDrive 檔案。

提交問題後，提示、已選附件及雲端來源文字，以及（如已啟用）工作簿／文件上下文會傳給所選 Qwen provider。原始雲端檔案不會上傳到 Qwen；只有使用者選取的文字會加入問題。Excel Chat 上下文包括儲存格值及公式；Word 上下文包括選取文字或文件文字。在 Excel，Plan 和 Agent 先讀取工作簿中繼資料，只在步驟需要時再讀取範圍。在 Word，Plan 和 Agent 會送出開啟中的文件。文件很長時，引號裡的段落仍會保留。**使用網頁版增益集時，同一問題也會經 Cloudflare。**

Provider key 或登入帳戶電郵地址不會寫入 repository 或 audit card。Provider API key 會傳至本機 Relay。Cloud OAuth 權杖會保留於本機 Relay 記憶體，並透過 HTTPS 用於所選來源請求；不會傳給 Qwen。不要把 provider key 或 OAuth client secret 放入聊天、檔案或 commit；`cloud-oauth.json` 只存公開的 app client ID。

---

## 11. Usage troubleshooting 使用疑難排解

| What you see | What to do |
|---|---|
| Provider HTTP 401 | Key and URL do not match — §7 |
| `Model not exist` (HTTP 400) | That model is not on your plan. Change it in Settings — §7 |
| Qwen endpoint rejected | Do not switch to a Token Plan key; use only a supported API/credential flow after the Marketplace update |
| Invalid conversation message | Over 50,000 characters. Shorten the message, remove an attachment, or untick **Send workbook context** |
| Context says truncated | Ask about a smaller range, or let the model profile the sheet |
| `Execute FAILED · …` | Read the reason on the audit line; it now names the actual cause |
| `N rejected as unusable` | The action broke a limit — e.g. a pivot allows 16 value fields and 8 each for rows/columns/filters |
| Stopped: validation failed | Leave the difference. Do not force the totals |
| Agent reached its 25-round limit | The model requested another step. Choose **Continue Agent · 繼續** |
| `Could not finish` with no answer | The status line names the reason. A common cause is a relay process still running an older build: the pane detects the mismatch and says so — restart the relay (`npm run relay-stop`, then `npm run relay-ensure` for a hidden relay, or `npm start` to also re-sideload) and send again. The pane JS refreshes from disk on load; the relay process only picks up new code when restarted |
| A reply will not finish | Press **Stop**. Set **Effort** to Off, or tick **Fast**, for a shorter reply |
| A change is wrong | **Undo last apply**, or Excel `Ctrl+Z` |
| Pane is Qwen-only | Grok was removed on 2026-11-24. Only the Qwen key is used, and nothing is sent until it is loaded |

For relay, sideload and install problems see `SOP-INSTALL.md` §10.

---

## 12. Processing status · 處理狀態

**2026-11-26 — Audit card hidden.** The pane no longer shows an AUDIT card after a question. The trail is still saved with the session. Close and reopen the task pane.

**2026-11-26 — 隱藏 AUDIT 卡。** 提問之後，窗格不再顯示 AUDIT 卡。紀錄仍保存在該次對話。請關閉並重新開啟工作窗格。

**2026-11-26 — GitHub repository in the chat box.** The composer has a **No Repository** button, like Qwen Coder. Open it to choose **No Repository** or one recent repository from the GitHub login on this PC. The next question uses that repository for GitHub reads and writes. Share must be on. The hosted site does not include this. Close and reopen the pane.

**2026-11-26 — 聊天框的 GitHub repo。** 輸入框有 **No Repository** 按鈕，跟 Qwen Coder 一樣。打開後可選 **No Repository**，或這部電腦 GitHub 登入的最近 repo。下一則問題會用那個 repo 做 GitHub 讀寫。Share 要開著。託管網站沒有這項。請關閉並重新開啟窗格。

**2026-11-26 — Connect to Github.** Local build `2026-11-26.3`. Open **MCP** and click **Connect to Github**. This uses the GitHub CLI login already on this PC. You do not paste a token or a password. The token stays in the GitHub CLI credential store. **Disconnect** only stops this pane from using that login. Restart the local relay and reopen the pane. The hosted site does not include this.

**2026-11-26 — Connect to Github。** 本機版本 `2026-11-26.3`。打開 **MCP**，按 **Connect to Github**。這用的是這部電腦 GitHub CLI 已有的登入。不用貼上 token，也不用輸入密碼。Token 留在 GitHub CLI 憑證庫。**Disconnect** 只是這個窗格不再使用該登入。請重啟本機 relay 並重新開啟窗格。託管網站沒有這項。

**2026-11-26 — Word desk, commit, and push.** Save the repo folder in **MCP**. In Word, a finished Qwen answer is added at the end of the document. `open path/file.js` opens that file. `git pull` refreshes it. `git commit` writes the document back. `git push` pushes the current branch and does not use `--force`. `git commit and push` does both. The command must start the message. Remove Qwen notes before commit. Restart the local relay and reopen Word.

**2026-11-26 — Word 桌面、提交與 push。** 在 **MCP** 儲存 repo 資料夾。在 Word，完成的 Qwen 回答會加在文件結尾。`open path/file.js` 打開該檔。`git pull` 刷新。`git commit` 把文件寫回。`git push` 推送目前分支，不用 `--force`。`git commit and push` 兩者都做。指令必須是訊息的開頭。提交前刪去 Qwen 註記。請重啟本機 relay 並重新開啟 Word。

**2026-11-26 — GitHub MCP.** Local build `2026-11-26.1`. In **MCP**, save a GitHub token if the repo is private. Agent can list repositories and read one file while Share is on. **Write to GitHub** opens a pull request on a new branch. **Write to default branch** is a second click. The hosted site does not include this connector. Restart the local relay and reopen the pane.

**2026-11-26 — GitHub MCP。** 本機版本 `2026-11-26.1`。私人 repo 請在 **MCP** 儲存 GitHub token。Share 開啟時，Agent 可列出 repo 和讀一個檔案。**Write to GitHub** 在新分支開 pull request。**Write to default branch** 是另一次點擊。託管網站沒有這條連接。請重啟本機 relay 並重新開啟窗格。

**2026-11-26 — Hosted pane deployed.** https://excel-qwen-grok.anthonym-354.workers.dev now serves build `2026-11-25.3` (`byo:true`). The web pane has Model, Effort, and the Workbook button. MCP, Skills, History, Memory, and PDF/Word/Excel extraction stay on the local relay. Close and reopen the Excel web task pane.

**2026-11-26 — 託管窗格已部署。** https://excel-qwen-grok.anthonym-354.workers.dev 現在是版本 `2026-11-25.3`（`byo:true`）。網頁版已有 Model、Effort 和 Workbook 按鈕。MCP、Skills、History、Memory，以及 PDF／Word／Excel 擷取仍只在本機 relay。請關閉並重新開啟 Excel 網頁版工作窗格。

**2026-11-25 — Model and Effort.** Build `2026-11-25.2`. Composer and Settings list `auto`, `qwen4-max` (default), `qwen3.8-max`, and `qwen3.8-flash`. Effort is Off, Low, Medium, High, or Extra. The relay sends `reasoning_effort` and never `thinking_budget`. Fast still forces thinking off. Restart the local relay after this bump. Hosted Worker was not deployed. `npm test` 123/123. Excel and Word were not opened for a live check.

**2026-11-25 — 模型與 Effort。** 版本 `2026-11-25.2`。作曲列與 Settings 可選 `auto`、`qwen4-max`（預設）、`qwen3.8-max`、`qwen3.8-flash`。Effort 為 Off、Low、Medium、High 或 Extra。relay 傳送 `reasoning_effort`，不傳送 `thinking_budget`。Fast 仍強制關閉思考。此版本需重啟本機 relay。託管 Worker 未部署。`npm test` 123/123。未開啟 Excel 或 Word 做實機檢查。

**2026-11-22 — Copilot-parity build verified in real Excel.**

- Native PivotTable creation, sheet profiling and range read-back are **confirmed working** against a live 753-row AR ageing workbook, including an 11-value-field pivot (USD O/S + nine buckets + Credits) with the `GROUP` total row excluded from the source.
- Reaching that took two fix rounds: read-back caps that silently dropped wide ranges, then a pivot value-field cap of 8 that rejected the legitimate 11-field pivot, plus execution errors that were swallowed so `Execute FAILED` carried no reason. Errors now reach both the audit card and the model.
- Tiers 1–3 of the Copilot action list are implemented: pivot, profiling, filters, conditional highlighting, borders, freeze, dropdowns, named ranges, hide, calculated columns, 30 chart types, sheet and Word replace-all. Undo is complete, including autofit and full-width deleted rows.
- Irreversible changes are click-gated and labelled on the card.
- `npm run check` passes and `npm test` is fully green on Windows now that the zip is read in Node rather than through the Unix `unzip` binary.
- Still unreachable: Python in Excel and the native Copilot surface (§9). The pane is Qwen-only since 2026-11-24 (Grok removed).

**2026-11-24 — scope-free Qwen in Excel.** Grok removed. `contextRequest` accepts `offset` so any range is walkable in 2,000-cell batches; profiles and finance checks scan the full requested range (160-cell check cap and 2M-cell/60-column profile caps removed); a check that cannot fully read its range returns NEEDS_USER, never PASS; round ceiling 25 with Continue Agent. `npm test` 75/75; live Excel retest pending pane reload.
**2026-11-24 (later) — Chat whole-file review, Copilot parity.** Chat now runs a read-only tool loop: `scanRequest` scans every sheet for formula errors and sizes, `groupBalance` balances each voucher, `crossCheck` compares an export against the ledger (rows + totals), plus profiles and paged reads; the report shape is conclusion + passed-checks table + exceptions by priority + suggested formulas + next steps (fixes only after confirmation). Chat still never writes. `npm test` 78/78.
**2026-11-24 (later still) — build handshake + visible failures.** Live screenshots showed a new pane against an old relay process (health fingerprint) and a silent round-1 exit. Pane and relay now handshake on `RELAY_BUILD`/`PANE_BUILD` (2026-11-24.3); a mismatch blocks the send with restart instructions. Every early failure and exception now writes its reason into the status line and the audit card instead of a bare "Could not finish". `npm test` 79/79.
**2026-11-24 (root cause) — silent round-1 death fixed.** The audit card (Request/Context present, no Model/error entry) pinpointed `renderMessage('assistant')` reading the deleted `activeProvider` select → TypeError while creating the answer bubble, and the catch's error bubble threw again → silent exit. Label is now 'QWEN'; a DOM-contract test pins every `$('id')` in taskpane.js to index.html. `npm test` 80/80.
**2026-11-24 (skills) — Copilot OneDrive skills store integrated.** The relay reads/writes `<OneDrive>/Documents(文件)/Copilot/Microsoft Excel/skills` locally (verified against `Desktop\OneDrive\文件\…`, 25 valid skills); with Share on, up to 3 relevant skills (≤12k chars) attach to each question and `@name` forces one; the pane "Saved skills" card lists them and "Save current task as skill" writes a **new** folder (Copilot format and naming rules, never overwrites). No Graph write scope needed; the OneDrive client syncs. `npm test` 83/83.
**2026-11-24 (Guide + PDF fix) —** a "Guide · 說明" SOP card (open by default) joins the sidebar: modes, Share consent gate, key model, data flow, local-only feature list. The hosted pane hides every local-data card (Cloud sources included) and saves no sessions = Marketplace v1 is core + Guide only. PDF/XLSX attachment cap raised to 20 MB (only extracted text enters context); attachment failures now show their reason inline in the attachment list. Relay `/api/extract` verified live with a PDF (HTTP 200). `npm test` 93/93; build .7.
**2026-11-24 (readable, problem-solving errors, his order) —** every error message now ends with "→" plus a bilingual cause + next step (central explainError map: build mismatch / relay down / 401 / 429 / model not found / MB cap / no text layer / poppler missing / Share off / session missing); applied to toasts, inline attachment notes, bubbles and card statuses; Guide gains a troubleshooting block. Pane-only, no relay restart needed. `npm test` 94/94. **Live-verified same day by Anthony: PDF attach + read and write-to-Excel flows completed.**
**2026-11-24 (Copilot-style attach UX, his screenshot order) —** the composer 📎 becomes a **＋ menu**: Upload images & files, Attach cloud files (opens sidebar Cloud sources), Record skill (opens sidebar skill save form), Change data source (opens sidebar); attached files render as **icon + name + × chips** above the composer (removable); PNG/JPG/WEBP image attachments added = vision transcription with your own Qwen key (clear error without a key). `npm test` 94/94; build .9.
**2026-11-24 (scanned-PDF vision fallback, his choice) —** when no text layer exists: pdftotext `-layout` then `-raw`; if both empty, pdftoppm renders up to 12 pages (150 dpi) and **your own Qwen key** (qwen4-max vision) transcribes them; the text carries a "[Vision transcription …]" provenance prefix. Without a key, a clear error remains. Images go only to QwenCloud (disclosed in the Guide); not available on hosted. Build .8.
**2026-11-24 (sidebar) — Copilot-style ☰ drawer.** The topbar ☰ toggles a left sidebar (hidden by default, open/closed state remembered) holding the five cards — Workbook context, Cloud sources, Saved skills, History, Memory — plus "＋ New chat · 新聊天"; main keeps only the conversation and composer. A Share quick toggle in the footer mirrors the sidebar consent gate both ways. `npm test` 88/88.
**2026-11-24 (memory + history) — CLI-style /resume.** Every run end saves a session (messages, audit, checkpoint, workbook, outcome) to `%LOCALAPPDATA%\QwenExcelCopilot\sessions` (local, per Windows user, never repo/OneDrive; newest 100 kept); the pane "History" card lists them and Resume restores the conversation plus the Agent checkpoint. The "Memory" card adds durable notes to `memory.md`, auto-attached (≤4k chars) while Share is on. PCs without OneDrive/Copilot get a fallback skills store at `%LOCALAPPDATA%\QwenExcelCopilot\skills` (auto-created on save) or a repo-local `skills/` folder. `npm test` 87/87.
**2026-11-24 — 無範圍限制 Qwen in Excel。** Grok 移除。`contextRequest` 接受 `offset`，任何範圍以每批 2,000 格分頁讀完；摘要與財務檢查掃描完整請求範圍（160 格檢查上限、200 萬格／60 欄摘要上限移除）；無法完整讀取的檢查回傳 NEEDS_USER，絕不 PASS；輪數上限 25，保留續行。`npm test` 75/75；實機重測待重新載入 pane。
**2026-11-24（較後）— Chat 整本覆核，看齊 Copilot。** Chat 現具唯讀工具迴圈：`scanRequest` 掃描全部工作表的公式錯誤與大小、`groupBalance` 每批次借貸平衡、`crossCheck` 匯出對總帳（列數＋加總）、profile／分頁讀取；報告格式為結論＋已通過表＋分優先級例外＋建議公式＋下一步（修改須先確認）。Chat 仍絕不寫入。`npm test` 78/78。
**2026-11-24（再後）— 版本握手＋失敗可見。** 實機截圖顯示 pane 已新、relay 進程仍舊（health 指紋），且第一輪無聲結束。現 pane 與 relay 以 `RELAY_BUILD`/`PANE_BUILD`（2026-11-24.3）握手，不符即擋下並指示重啟；所有早期失敗與例外會寫入狀態列與 audit 卡，不再只顯示「Could not finish」。`npm test` 79/79。
**2026-11-24（根因）— 無聲第一輪死亡已修。** audit 卡（有 Request／Context、無 Model／錯誤）指向 `renderMessage('assistant')` 讀取已刪除的 `activeProvider` select → 建立回覆气泡時 TypeError、catch 再拋 → 無聲結束。標籤改 'QWEN'；新增 DOM 契約測試（taskpane.js 每個 `$('id')` 必須存在於 index.html）。`npm test` 80/80。
**2026-11-24（技能整合）— Copilot OneDrive skills 庫。** relay 本機讀寫 `<OneDrive>/Documents(文件)/Copilot/Microsoft Excel/skills`（實測解析到 `Desktop\OneDrive\文件\…`、25 個有效技能）；Share 開啟時每次提問自動帶入最多 3 個相關技能（≤12k 字）、`@name` 強制选用；pane「Saved skills」卡清單＋「Save current task as skill」寫入**新**資料夾（Copilot 格式、名稱規則、絕不覆蓋）。无需 Graph 寫入權限；OneDrive 客戶端負責同步。`npm test` 83/83。
**2026-11-24（Guide＋PDF 修正）—** 側欄新增「Guide · 說明」SOP 卡（預設開）：模式、Share 同意閘、金鑰模型、資料流、本機限定功能清單。hosted pane 隱藏所有本機資料卡（含 Cloud sources）且不存 session＝Marketplace v1 僅核心＋Guide。PDF/XLSX 附件上限升至 20 MB（僅提取文字入上下文）；附件失敗原因 inline 顯示於附件清單。relay `/api/extract` 實測 PDF 200。`npm test` 93/93；build .7。
**2026-11-24（可讀錯誤＝解法，他指示）—** 所有錯誤訊息結尾加「→」雙語原因＋下一步（explainError 中央對照：build 不符／relay down／401／429／模型不存在／MB 上限／無文字層／缺 poppler／Share off／session 不存在）；toast、inline 附件註記、气泡、卡狀態皆套用；Guide 加疑難排解區塊。pane-only，无需重啟 relay。`npm test` 94/94。**實機驗證（同日，Anthony）：PDF 附件＋讀取與寫入 Excel 流程測試完成。**
**2026-11-24（Copilot 式附件 UX，他截圖指示）—** composer 的 📎 改為 **＋ 選單**：上傳影像和檔案、附加雲端檔案（開側欄 Cloud sources）、錄製技能（開側欄 skills 儲存表單）、變更資料來源（開側欄）；已附檔案以**圖示＋名稱＋× chips** 顯示於 composer 上方（可移除）；新增 PNG/JPG/WEBP 影像附件＝经自有 Qwen key 視覺轉錄（無 key 報明確錯誤）。`npm test` 94/94；build .9。
**2026-11-24（掃描 PDF 視覺後備，他選定）—** 無文字層時：pdftotext `-layout`→`-raw` 皆空→pdftoppm 轉最多 12 頁 PNG（150 dpi）→经**您自己的 Qwen key**（qwen4-max 視覺）轉錄，回文字附「[Vision transcription …]」來源前綴；無 key 則回明確錯誤。圖片只送 QwenCloud（Guide 已披露）；hosted 不提供。build .8。
**2026-11-24（側欄）— Copilot 式 ☰ 抽。** 顶欄 ☰ 開關左側欄（預設隱藏、記住開關狀態），內含 Workbook 上下文、Cloud sources、Saved skills、History、Memory 五卡＋「＋ New chat · 新聊天」；main 只留對話與composer。footer 新增 Share 快速開關，與側欄內同意閘門雙向同步。`npm test` 88/88。
**2026-11-24（記憶＋歷史）— 如 CLI 的 /resume。** 每次執行結束自動存 session（對話、audit、checkpoint、活頁簿名、結果）到 `%LOCALAPPDATA%\QwenExcelCopilot\sessions`（本機、每 Windows 用戶、不入 repo／OneDrive；保留最新 100 個）；pane「History」卡清單＋Resume 恢復對話與 Agent checkpoint 續做。「Memory」卡手動加註記，存 `memory.md`，Share 開啟時自動帶入（≤4k 字）。無 OneDrive／Copilot 的電腦：skills 後備庫 `%LOCALAPPDATA%\QwenExcelCopilot\skills`（Save 時自動建立）或 repo 內 `skills/`。`npm test` 87/87。

處理狀態（2026-11-22）：Copilot 看齊建置已在真實 Excel 中驗證。**原生 PivotTable 建立、表剖析與範圍讀回已確認可用**，對象是真實的 753 列 AR ageing 活頁簿，包括一個 11 值欄的 pivot（USD O/S + 九個桶 + Credits），並已把 `GROUP` 總計列排除於來源之外。達成這一點此前經歷兩輪修正：先是靜默丟棄寬範圍的讀回上限，然後是拒絕合理 11 欄 pivot 的 8 欄值上限，以及被吞掉、令 `Execute FAILED` 不帶原因的執行錯誤；現錯誤會同時到達 audit 卡片與模型。Copilot action 清單的 Tier 1–3 均已實作，Undo 亦已完成（包括 autofit 與全寬刪除列）。不可逆改動需點擊批准並在卡片上標明。`npm run check` 通過，且 zip 改以 Node 讀取後 `npm test` 在 Windows 上全綠。仍無法達到：Python in Excel 與原生 Copilot 介面（§9）。hosted Worker 未設定 `GROK_API_KEY`，故 hosted pane 只能用 Qwen。
