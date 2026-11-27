# Qwen for Excel and Word

A local task pane for desktop Excel and Word. You ask in plain language. Qwen plans the work, and the pane revises the open file through a fixed list of actions. On this PC the pane and the relay both live at `http://127.0.0.1:3000`. Excel on the web uses the HTTPS copy at `https://excel-qwen-grok.anthonym-354.workers.dev`. Company-wide deployment still needs an admin install.

**Folder on this PC:** `C:\Users\User\Desktop\excel-qwen-grok-copilot`. Run npm from that folder. Installation steps are in [SOP-INSTALL.md](SOP-INSTALL.md). The older Desktop folder `Excel Qwen Grok Copilot` is archived. Do not install from it.

**這部電腦的資料夾：** `C:\Users\User\Desktop\excel-qwen-grok-copilot`。請在該資料夾執行 npm。安裝步驟見 [SOP-INSTALL.md](SOP-INSTALL.md)。桌面上較舊的 `Excel Qwen Grok Copilot` 已封存，不要用它安裝。

---

## Release roadmap 發行計劃

| Step | English | 繁體中文 |
|---|---|---|
| 1 | Review attached HTML and text reports. | 檢視附加的 HTML 與文字報表。 |
| 2 | Connect personal OneDrive for read-only search and review. | 連接個人版 OneDrive，提供唯讀搜尋及檢視。 |
| 3 | Trace finance answers to workbook cells or report lines and show the checks performed. | 把財務答案追溯至工作簿儲存格或報表行數，並顯示已執行的檢查。 |
| 4 | Show Agent progress and let users continue at the round limit. Keep edit approval and Undo controls. | 顯示 Agent 進度及輪數上限續做按鈕，保留修改審批及復原功能。 |
| 5 | Finish live testing of realistic workflows with month-end workbooks and HTML reports, including reconciliation, large inputs and provider interruptions. | 完成實機測試：用實際月結工作簿及 HTML 報表測試對帳、大量資料及模型服務中斷。 |
| 6 | Prepare and publish the production add-in in Microsoft Marketplace after the release gates are cleared. | 清除發行前置條件後，準備並上架正式版至 Microsoft Marketplace。 |

### Step 5 progress — 2026-11-24 / 第 5 步進度

The October close trial showed that asking the user to pick technical read ranges is not a usable workbook-review flow. Agent now receives a workbook inventory and chooses the sheets and ranges itself. A large `profileRequest` scans up to 2 million cells locally in 20,000-cell batches, then sends compact column statistics, Excel error-value counts and explicit coverage to the model. A failed range read is reported as unread evidence so Agent can retry. **Share Office file data with AI** controls whether metadata and requested workbook reads can be sent; it is a data-sharing choice, not a sheet-scope setting. Raw range reads are delivered in 2,000-cell batches that Agent pages through with `contextRequest.offset` until the whole range is read; profiles and finance checks scan the full requested range with no cell or column cap; Agent pauses at 25 model rounds with **Continue Agent · 繼續**. Anthony confirmed Excel web key persistence after reopening and a basic `hi` response; the full month-end workbook review after these changes still needs a live retest. Step 5 remains in progress.

10 月月結試用顯示，不應要求使用者自行設定技術性的讀取範圍。現在 Agent 先取得活頁簿清單，再自行選擇工作表及範圍。大型 `profileRequest` 會在本機以每批 20,000 格掃描最多 200 萬格，向模型回傳精簡欄位統計、Excel 錯誤值數目及明確覆蓋範圍。範圍讀取失敗會標示為未讀取，供 Agent 重試。**Share Office file data with AI** 是資料分享選項，不是工作表範圍設定。原始範圍讀取以每批 2,000 格交付，Agent 用 `contextRequest.offset` 分頁直至讀完整個範圍；摘要與財務檢查掃描完整請求範圍，無儲存格或欄數上限；Agent 每 25 輪暫停，可按 **Continue Agent · 繼續**。Anthony 已確認 Excel 網頁版重開後仍保存 key，並成功收到基本 `hi` 回覆；但上述變更後的完整月結活頁簿覆核仍須實機重測。第 5 步仍進行中。

**2026-11-24 — Qwen in Excel, scope-free.** The Grok section was removed (UI, routing, relay, docs); the pane is Qwen-only. `contextRequest` accepts `offset` so any range is walkable; the 160-cell finance-check cap and the 2M-cell/60-column profile caps are gone — checks and profiles cover the full requested range, and a check that cannot fully read its range returns NEEDS_USER, never PASS. Round ceiling raised 10 → 25. `npm run check` and `npm test` pass 75/75; live Excel retest pending the owner reloading the pane.
**2026-11-24 (later) — Chat whole-file review (Copilot parity).** Chat runs a read-only tool loop: `scanRequest` (every-sheet formula-error scan), `groupBalance` (per-voucher DR=CR), `crossCheck` (export vs ledger rows + totals), profiles and paged reads; report shape = passed-checks table, exceptions by priority, suggested formulas, next steps with confirmation-gated fixes. Chat never writes. `npm test` 78/78.
**2026-11-24 (live verification) —** Anthony completed live tests: PDF attach + read (text-layer and scanned-vision paths) and write-to-Excel flows. The post-`101bbeb` batches (build .9: Guide card, hosted data-feature stripping, 20 MB binary caps, vision fallback, Copilot attach menu/chips, readable problem-solving errors) are tested offline 94/94 and remain uncommitted pending his push order.
**2026-11-24（實機驗證）—** Anthony 完成實機測試：PDF 附件＋讀取（文字層與掃描視覺路徑）及寫入 Excel 流程。`101bbeb` 之後各批（build .9：Guide 卡、hosted 資料功能剝離、20 MB 二进制上限、視覺後備、Copilot 附件選單／chips、可讀附解法錯誤）離線 94/94 通過，仍未 commit，待其 push 指示。
**2026-11-24 (Marketplace Phase 1 + 2 draft) —** the Worker is now a BYO-key pass-through: the caller's own key per request, no owner key or shared relay token in the serving path, 60 req/min per key, nothing stored or logged; the hosted pane keeps the key in-browser only. Store pack drafted: manifest v1.4.0.0 with Privacy/Terms/Support URLs, bilingual legal pages in `public/legal/`, `STORE_LISTING.md` + `STORE_REVIEWER_NOTES.md`. `wrangler deploy`, owner-secret cleanup, Partner Center enrollment and submission remain Anthony's. 98/98 tests.
**2026-11-24 (manifest 1.5.0.0) —** both manifests bumped to 1.5.0.0 (cloud + local branding/ProviderName synced, local legal URLs added). Reminder: `wrangler deploy` never updates the manifest; Excel web caches the uploaded copy — remove the old 我的增益集 entry and re-upload after manifest changes.
**2026-11-24（manifest 1.5.0.0）—** 兩個 manifest 升 1.5.0.0（cloud＋本機品牌／ProviderName 同步、本機法律 URL）。提醒：`wrangler deploy` 不更新 manifest；Excel web 快取上傳副本——manifest 變更後须移除舊我的增益集項目再上傳。
**2026-11-24 (deploy live) —** Anthony ran `npx wrangler deploy` (Version `54752ffa-e235-4dfd-99e6-9c85758165f7`); live-verified: health `byo:true` + `configured.qwen:false`, legal pages 200, served pane .9, no-key chat 401 with readable fix. G1 closed live. On Anthony's order the owner `QWEN_API_KEY` Worker secret was then deleted (`wrangler secret delete`); only the unused `RELAY_TOKEN` remains. Remaining: D2 publisher name, Partner Center + submission.
**2026-11-24（deploy 上線）—** Anthony 執行 `npx wrangler deploy`（Version `54752ffa-e235-4dfd-99e6-9c85758165f7`）；實機驗證：health `byo:true`＋`configured.qwen:false`、法律頁 200、pane .9、無 key chat 401 附解法。G1 上線關閉。依 Anthony 指示已刪除 owner `QWEN_API_KEY` Worker secret（`wrangler secret delete`）；僅餘未使用之 `RELAY_TOKEN`。餘：D2 發行者名稱、Partner Center＋提交。
**2026-11-24 (Excel web key + HTTP 500 repair) —** Build .10 prefers `OfficeRuntime.storage` for the hosted BYO key, then browser storage; Save verifies the stored value and reports session-only use if neither works. Clear removes saved copies. A stream request retries one HTTP 500/502/503/504 before showing a bilingual error with the status. The earlier HTTP 500 was intermittent; its exact upstream cause was not captured. Local checks pass 106/106. After deployment, Anthony confirmed that the key remained saved after reopening the Excel web pane and `hi` succeeded.
**2026-11-24（Excel 網頁版 key 與 HTTP 500 修正）—** .10 版託管增益集優先使用 `OfficeRuntime.storage` 記住自備 key，其次是瀏覽器儲存；Save 會讀回確認，兩者均不可用則清楚標示僅限本次使用。Clear 會清除已保存的副本。串流遇到 HTTP 500/502/503/504 時重試一次，仍失敗則顯示雙語狀態碼。先前 500 屬間歇性；尚未取得確切上游原因。本機檢查 106/106 通過。部署後 Anthony 確認 Excel 網頁版重新開啟仍保存 key，傳送 `hi` 成功。
**2026-11-24（Marketplace Phase 1＋2 草稿）—** Worker 改 BYO-key pass-through：每請求自帶用户 key、服務路徑無 owner key／共用 relay token、每 key 60 req/min、不儲存不記錄；hosted pane 金鑰只存瀏覽器。上架包草稿：manifest v1.4.0.0 含 Privacy/Terms/Support URLs、`public/legal/` 雙語法律頁、`STORE_LISTING.md`＋`STORE_REVIEWER_NOTES.md`。`wrangler deploy`、owner-secret 清理、Partner Center 註冊與提交仍屬 Anthony。98/98 測試。
**2026-11-24 — Qwen in Excel，無範圍限制。** Grok 區塊已移除（UI、路由、relay、文件），pane 僅用 Qwen。`contextRequest` 接受 `offset`，任何範圍皆可分頁讀完；160 格財務檢查上限與 200 萬格／60 欄摘要上限已移除——檢查與摘要覆蓋完整請求範圍，無法完整讀取的檢查回傳 NEEDS_USER，絕不 PASS。輪數上限 10 → 25。`npm run check` 與 `npm test` 75/75 通過；實機 Excel 重測待擁有者重新載入 pane。
**2026-11-24（較後）— Chat 整本覆核（Copilot 看齊）。** Chat 具唯讀工具迴圈：`scanRequest` 全表公式錯誤掃描、`groupBalance` 每批次借貸、`crossCheck` 匯出對總帳、profile／分頁讀取；報告含已通過表、分優先級例外、建議公式與下一步（修改須確認）。Chat 絕不寫入。`npm test` 78/78。

### Step 6 — Microsoft Marketplace launch 正式上架

**Preparation started 2026-11-25 · 準備已開始；目前不得提交** — see [`MARKETPLACE_READINESS.md`](MARKETPLACE_READINESS.md) for the release gates. The existing BYO key setup is not ready for public release: the current default is a Token Plan endpoint, and the hosted pane stores long-term keys in Office/browser storage. Publisher identity and full live workbook QA are also outstanding.

**2026-11-25 開始準備；目前不得提交** — 發行前置條件見 [`MARKETPLACE_READINESS.md`](MARKETPLACE_READINESS.md)。現有自備 key 流程尚未達到公開發行條件：目前預設為 Token Plan endpoint，託管 pane 會在 Office／瀏覽器儲存長期 key；發行者身份及完整活頁簿實機 QA 亦未完成。

1. **Production service.** Keep the HTTPS pane and Worker, then settle a supported Qwen API and credential flow. Every user will use their own key; no key is shared. The current Token Plan route and hosted browser storage must be replaced or explicitly approved by Alibaba before release. If moving credentials server-side, add per-user authentication and usage controls. Decide whether OneDrive will ship in the hosted release. [Office Add-ins manifest guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests)

   **正式服務。** 以 HTTPS 部署工作窗格及 Relay，改用個別使用者或組織登入，服務憑證存於伺服器，並確認 Qwen 的使用方式與收費模式。決定 OneDrive 是否納入正式託管版本。
2. **Manifest and installation.** Create and validate the production manifest with the publisher identity, app ID, version, icons and support URL. Test clean installation and core workflows on the Excel and Word clients the listing will support. [Microsoft publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist)

   **正式發行檔與安裝。** 建立及驗證正式 manifest，確認發行者名稱、應用程式 ID、版本、圖示和支援網址。以市集標示支援的 Excel 及 Word 用戶端測試全新安裝及主要功能。
3. **Listing and review.** Enroll the publisher in the Microsoft 365 and Copilot program in Partner Center. Prepare the listing, screenshots, HTTPS links for support, privacy and the end user license agreement (EULA), and reviewer instructions with sample files and test access. Explain when workbook, document or attachment text is sent to the relay and selected model provider, and disclose any external account or fee required. [Microsoft submission guide](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/add-in-submission-guide)

   **市集頁面與審查。** 在 Partner Center 將發行者註冊至 Microsoft 365 and Copilot 計劃。準備上架文案、截圖、支援／私隱／使用者授權合約（EULA）的 HTTPS 網址，以及審查用的操作步驟、範例檔案和測試權限。說明工作簿、文件或附件文字何時會傳到 Relay 及所選模型服務，並列明所需外部帳戶及費用。
4. **Publish and verify.** Submit for Microsoft certification, resolve any review findings, then verify the public listing and complete an install from Office with a fresh account. Record the live listing URL and first published version. [Publish an Office Add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource)

   **上架與確認。** 提交 Microsoft 審核及處理審核意見。上架後確認市集頁面已公開，並以新帳戶在 Office 內實際安裝。記錄正式連結及首個發行版本。

---

## Install 安裝

**Teammates — this is the whole install.** Each person uses their own API key; none is bundled, and the owner's key is never shared.

### From a terminal (recommended) 用終端機（建議）

**Close Excel and Word first**, then run these three lines:

```powershell
cd "C:\Users\Your Name\Qwen-for-Excel-and-Word"
npm run setup
npm start
```

- **`npm run setup`** — checks Node.js 20+, runs `npm install`, prints the next steps. Installs no auto-start, writes no key.
- **`npm start`** — sideloads the add-in into desktop Excel and opens Excel. **Leave that terminal window open: it *is* the relay.**

Then in Excel: **Home → Add-ins → Qwen for Excel and Word** → **Settings** → pick Qwen → paste **your own** key → leave **Remember** ticked → **Use key**. The dot turns green.

Plain npm is identical: `npm install` then `npm start`. Requires **Node.js 20+** from https://nodejs.org.

> **Quote the path.** If your user folder contains a space — e.g. `C:\Users\Anthony Man` — PowerShell splits an unquoted path and `cd` fails.
> **路徑要加引號。** 使用者資料夾若含空格，未加引號的路徑會令 PowerShell 的 `cd` 失敗。

### From the zip 用 zip

Get `Qwen-for-Excel-and-Word.zip` from the owner, or download it from **[`download/`](download/Qwen-for-Excel-and-Word.zip)** in this repository. Extract it and run the same commands in that folder. `START-HERE.txt` inside is the one-page quick start. The zip carries **no keys**.

The zip is a build artifact: rebuild it with `npm run package` after any code or doc change, and commit the result so the download never serves stale code. If you have repository access, cloning and running `npm run setup` is always the most current route.

### Or double-click 或雙擊

`Setup and Start.bat` does the same install **plus** a logon auto-start and a per-minute keep-alive task. Skip it if you do not want background automation — `npm start` alone is enough.

### Word, and Excel on the web

- **Word:** close Word, then `npm run start:word`.
- **Excel on the web:** **Home → Add-ins → Upload My Add-in** → `manifest.cloudflare.xml`. The hosted pane currently asks each user for their own Qwen API key and sends it per request through the Worker; no provider key is embedded in the Worker. See [`MARKETPLACE_READINESS.md`](MARKETPLACE_READINESS.md) before using Token Plan credentials or preparing a public release.

### Read-only cloud source review (desktop local relay only)

The desktop pane can search and review the signed-in user's personal OneDrive. Search results stay in the pane until the user chooses **Review and add to chat**; that selected text is included with the next question. The connector is read-only and does not change OneDrive files.

Each colleague clicks **Connect**, signs in through the system browser, and approves access to their own personal Microsoft account. Account email addresses and OAuth tokens are not shared between colleagues; tokens stay in that PC's local relay memory and are cleared when the relay stops. The app's public Microsoft OAuth client ID is set once by the owner in the root `cloud-oauth.json` file. That file contains no client secret or user email. The hosted Excel-on-the-web pane does not use this local connector.

Before colleagues connect:

1. Register a Microsoft public client that supports personal Microsoft accounts. Add the **Mobile and desktop applications** platform with `http://localhost` redirect support and delegated `Files.Read` permission. The sign-in uses the system browser and PKCE; no Microsoft client secret is used.
2. The app owner fills `microsoftClientId` once in the root `cloud-oauth.json` before packaging or sharing the add-in. Restart the local relay after changing this file. This is a public app ID, not a user credential. Colleagues then open **MCP → Connect** and approve their own account in the browser; no per-user Settings entry is needed.

OneDrive text review supports XLSX/XLSM, DOCX, HTML, TXT, MD, CSV, TSV, JSON, and XML. OneDrive uses Microsoft Graph's [personal-account file permissions](https://learn.microsoft.com/en-us/graph/permissions-reference).

#### 雲端來源唯讀檢視（只支援桌面本機 Relay）

桌面版可搜尋及檢視登入者自己的個人 OneDrive。搜尋結果只會留在增益集窗格；使用者按下 **Review and add to chat** 後，所選文字才會加入下一個問題。連接器只有讀取權限，不會修改 OneDrive 檔案。

每位同事按 **Connect**，在系統瀏覽器登入並授權自己的個人 Microsoft 帳戶。帳戶電郵地址及 OAuth 權杖不會在同事之間共用；權杖只保留在該電腦本機 Relay 記憶體，Relay 停止後即清除。應用程式的公開 Microsoft OAuth 用戶端 ID 由管理者一次性填入根目錄 `cloud-oauth.json`；檔案不含 client secret 或使用者電郵。Excel 網頁版不使用這個本機連接器。

同事連接前，管理者需：

1. 註冊支援個人 Microsoft 帳戶的 public client，加入 **Mobile and desktop applications** 平台、`http://localhost` redirect 支援及委派 `Files.Read` 權限。登入透過系統瀏覽器及 PKCE 完成；不使用 Microsoft client secret。
2. 管理者在分享或打包增益集前，只需把 `microsoftClientId` 填入根目錄 `cloud-oauth.json` 一次；更改後要重新啟動本機 Relay。這是公開應用程式 ID，不是使用者憑證。同事之後按 **MCP → Connect**，在瀏覽器登入並授權自己的帳戶；無須逐人設定 Settings。

OneDrive 文字檢視支援 XLSX/XLSM、DOCX、HTML、TXT、MD、CSV、TSV、JSON 和 XML。OneDrive 使用 Microsoft Graph 的[個人帳戶檔案權限](https://learn.microsoft.com/en-us/graph/permissions-reference)。

### Documentation 文件

| File | Purpose |
|---|---|
| **[SOP-INSTALL.md](SOP-INSTALL.md)** | **Installation 安裝程序** — prerequisites, CLI steps, verification, daily running, uninstall, troubleshooting, teammate handover |
| **[SOP-USER-GUIDE.md](SOP-USER-GUIDE.md)** | **User guide 說明書** — modes and Direct, context control, every action, cards, Undo, keys/URLs/models, finance rules, limits |
| [SOP.md](SOP.md) | Index of the two |

**Stop the relay:** `npm run relay-stop` · **Remove the sideload:** `npm run stop` · **Remove auto-start:** `npm run autostart-remove`

---

## Processing status · 處理狀態

**2026-11-28 — Fixed: creating a sheet and writing to it in one batch.** A model batch like `addSheet "Sales Invoice"` followed by `set`/`format` on that new sheet failed with the Office.js error "property 'isNullObject' is not available… call load… context.sync()", rolling back the whole batch (real case: 39 demo invoice actions, nothing written). Cause: the missing-sheet guard read `isNullObject` on a freshly added worksheet object, which real Office.js forbids before `load()`. The guard now skips sheets created in the same batch. The test mock reproduces the real throw, and a regression test (addSheet→set→format) plus a missing-sheet fail-closed test were added; the regression test was proven to fail on the old code. `npm test` **151/151**. Local relay restarted on build `2026-11-26.3`; reopen the pane and retry. Hosted Worker not deployed. / **修正：同一批次新增工作表並寫入。** addSheet 後對新表 set/format 會觸發 Office.js「isNullObject 未 load」錯誤，整批回滾（真實案例：39 個 demo 發票動作全沒寫入）。原因：缺表檢查讀取了刚新增、未 load 的工作表物件。現同一批次新增的表略過該檢查；mock 模擬真实拋錯，新增回归測試（已證明旧代码會失敗）與缺表 fail-closed 測試。151/151 全綠；本機 relay 已重啟，重開 pane 即可重試。託管 Worker 未部署。

**2026-11-26 — Audit card hidden.** The pane no longer shows an AUDIT card after a question. The trail is still saved with the session. Close and reopen the task pane.

**2026-11-26 — 隱藏 AUDIT 卡。** 提問之後，窗格不再顯示 AUDIT 卡。紀錄仍保存在該次對話。請關閉並重新開啟工作窗格。

**2026-11-26 — GitHub repository in the chat box.** The composer **No Repository** button lists recent repositories from the GitHub login on this PC. Choosing one attaches that `owner/repo` to the next question. **No Repository** leaves the question on a blank workspace. Share must be on. The hosted site does not include this. Close and reopen the pane.

**2026-11-26 — 聊天框的 GitHub repo。** 輸入框的 **No Repository** 按鈕列出這部電腦 GitHub 登入的最近 repo。選了一個之後，下一則問題會帶上那個 `owner/repo`。選 **No Repository** 就是空白工作區。Share 要開著。託管網站沒有這項。請關閉並重新開啟窗格。

**2026-11-26 — Connect to Github.** Local build `2026-11-26.3`. In **MCP**, **Connect to Github** uses the GitHub CLI login already on this PC. The token stays in that CLI credential store and is not sent to Qwen or saved in the pane. **Disconnect** stops this pane from using the account and does not sign the PC out. The hosted site does not include this. Restart the local relay and reopen the pane.

**2026-11-26 — Connect to Github。** 本機版本 `2026-11-26.3`。在 **MCP** 按 **Connect to Github**，使用這部電腦 GitHub CLI 已有的登入。Token 留在 CLI 憑證庫，不會送給 Qwen，也不會存在窗格。**Disconnect** 只是這個窗格不再使用該帳戶，不會把電腦登出。託管網站沒有這項。請重啟本機 relay 並重新開啟窗格。

**2026-11-26 — Word desk, commit, and push.** In Word, Agent and Plan send the document text, the same way Excel sends the sheet. A selection does not replace that text. A quoted passage stays in that text when the document is long. A finished Qwen answer is added at the end of the document. It does not replace the file. `open src/app.js` loads one local Git file. `git pull` fast-forwards and refreshes it. `git commit` writes the document back to that file. `git push` pushes the current branch and does not use `--force`. `git commit and push` does both. Remove Qwen notes before commit. Set the repo folder under **MCP**. The hosted site does not include this. Restart the local relay and reopen Word.

**2026-11-26 — Word 桌面、提交與 push。** 在 Word，Agent 和 Plan 會送出文件文字，跟 Excel 送出工作表一樣。選取範圍不會取代那段文字。文件很長時，引號裡的段落仍會保留。完成的 Qwen 回答會加在文件結尾，不會覆蓋檔案。`open src/app.js` 載入一個本機 Git 檔。`git pull` 以 fast-forward 更新並刷新。`git commit` 把文件寫回那個檔。`git push` 推送目前分支，不用 `--force`。`git commit and push` 兩者都做。提交前刪去 Qwen 註記。Repo 資料夾在 **MCP** 設定。託管網站沒有這項。請重啟本機 relay 並重新開啟 Word。

**2026-11-26 — GitHub MCP.** Local build `2026-11-26.1`. The MCP card can list repositories and read one text file. A write is sent only after **Write to GitHub**, which opens a pull request on a new branch. **Write to default branch** is a separate click and does not force-push. The token stays in this browser and goes only to `api.github.com`. Share must be on. The hosted site does not include this connector. Restart the local relay and reopen the pane.

**2026-11-26 — GitHub MCP。** 本機版本 `2026-11-26.1`。MCP 卡可列出 repo 和讀一個文字檔。寫入只在按 **Write to GitHub** 之後送出，並在新分支開 pull request。**Write to default branch** 是另一次點擊，不會 force-push。Token 留在這個瀏覽器，只送到 `api.github.com`。Share 要開著。託管網站沒有這條連接。請重啟本機 relay 並重新開啟窗格。

**2026-11-26 — Hosted pane deployed.** `npx wrangler deploy` published `main` commit `6534f67` to https://excel-qwen-grok.anthonym-354.workers.dev. Live health is `build: 2026-11-25.3`, `byo:true`, `configured.qwen:false`. Worker version `ca5916b0-4f70-485b-a186-20e5bd769b23`. The served pane includes Model, Effort, and the Workbook button. MCP, Skills, History, Memory, and PDF/Word/Excel extraction stay on the local relay. Close and reopen the Excel web task pane. The email-and-password team version is cancelled. It was never deployed, and a Token Plan key cannot back that shared login.

**2026-11-26 — 託管窗格已部署。** `npx wrangler deploy` 把 `main` commit `6534f67` 發佈到 https://excel-qwen-grok.anthonym-354.workers.dev。線上 health 為 `build: 2026-11-25.3`、`byo:true`、`configured.qwen:false`。Worker 版本 `ca5916b0-4f70-485b-a186-20e5bd769b23`。網站已有 Model、Effort 和 Workbook 按鈕。MCP、Skills、History、Memory，以及 PDF／Word／Excel 擷取仍只在本機 relay。請關閉並重新開啟 Excel 網頁版工作窗格。email 加密碼的團隊版已取消。它從未部署，Token Plan key 也不能用於那種共用登入。

**2026-11-25 — Pane names and GitHub skill install.** Commit `6534f67`. Local build `2026-11-25.3`. **Skills** installs one `SKILL.md` from a GitHub repository into the OneDrive skills folder. A public repo can omit the token. Remember keeps the token in this browser only. **Workbook** opens WORKBOOK ACCESS above the button. **Guide** starts collapsed. **MCP** is the OneDrive card. The four shortcut buttons and the bottom remark are gone. The hosted Worker was not deployed.

**2026-11-25 — 窗格名稱與從 GitHub 安裝技能。** Commit `6534f67`。本機版本 `2026-11-25.3`。**Skills** 可把一個 GitHub `SKILL.md` 安裝到 OneDrive 技能資料夾。公開 repo 可以不填 token。Remember 只把 token 留在這個瀏覽器。**Workbook** 在按鈕上方打開 WORKBOOK ACCESS。**Guide** 預設收合。**MCP** 是 OneDrive 卡。四個快捷按鈕和底部說明已移除。託管 Worker 未部署。

**2026-11-25 — Model and Effort.** Local build `2026-11-25.2`. The composer and Settings offer `auto`, `qwen4-max` (default), `qwen3.8-max`, and `qwen3.8-flash`. Effort is Off, Low, Medium, High (default), or Extra. The relay sends `reasoning_effort` and does not send `thinking_budget`. Fast still turns thinking off for that question. Restart the local relay so the pane and relay builds match. The hosted Worker was not deployed. `npm test` 123/123.

**2026-11-25 — 模型與 Effort。** 本機版本 `2026-11-25.2`。作曲列與 Settings 可選 `auto`、`qwen4-max`（預設）、`qwen3.8-max`、`qwen3.8-flash`。Effort 為 Off、Low、Medium、High（預設）或 Extra。relay 傳送 `reasoning_effort`，不傳送 `thinking_budget`。Fast 仍可對該題關閉思考。請重啟本機 relay，讓窗格與 relay 版本一致。託管 Worker 未部署。`npm test` 123/123。

**2026-11-25 — Shared file reader, OneDrive tool, skill triggers.** Local build `2026-11-25.1`. Attach and OneDrive use one reader for PDF, DOCX, and XLSX/XLSM. Excel and PDF results say `Coverage: COMPLETE` or `Coverage: INCOMPLETE` with the next row or page. Scanned PDF pages are transcribed one at a time, up to 12. Agent may search personal OneDrive and read one file while Share is on; the hosted pane does not. Skills match a `triggers` line and are no longer forced to include `qwen-chat-assistant`. Restart the local relay so the pane and relay builds match. `npm test` 115/115.

**2026-11-25 — 共用檔案讀取、OneDrive 工具、技能觸發。** 本機版本 `2026-11-25.1`。附加與 OneDrive 共用 PDF、DOCX、XLSX/XLSM 讀取器。Excel 與 PDF 會標示 `Coverage: COMPLETE` 或 `Coverage: INCOMPLETE`，並給出下一列或下一頁。掃描 PDF 逐頁轉錄，最多 12 頁。Share 開啟時，Agent 可搜尋個人 OneDrive 並讀取一個檔案；網頁版沒有此功能。Skills 依 `triggers` 列配對，不再每次強制帶入 `qwen-chat-assistant`。請重啟本機 relay，讓窗格與 relay 版本一致。`npm test` 115/115。

**2026-11-25 — Marketplace preparation.** The hosted Worker is live as BYO-key pass-through (`build: 2026-11-25.3`, `byo:true`, `configured.qwen:false` at the 2026-11-26 health check). Each user enters their own key; the add-in does not use an owner key. Do not use Token Plan keys with this custom add-in unless Alibaba confirms this use is permitted. The Marketplace release is on hold pending a supported API/credential design, publisher enrollment, and full live workbook testing. See [`MARKETPLACE_READINESS.md`](MARKETPLACE_READINESS.md).

**2026-11-25 — Marketplace 準備。** 託管 Worker 以 BYO-key 方式運作（2026-11-26 health check：`build: 2026-11-25.3`、`byo:true`、`configured.qwen:false`）。每位使用者輸入自己的 key；增益集不使用作者 key。除非阿里雲確認允許，請勿在此自訂增益集使用 Token Plan key。Marketplace 發行暫停，待確立合規 API／憑證設計、完成發行者註冊及完整活頁簿實機測試。詳見 [`MARKETPLACE_READINESS.md`](MARKETPLACE_READINESS.md)。

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

The 2,000-row profile limit above describes the 2026-11-22 version; the 2026-11-24 Step 5 update replaced it with batched scanning. 上述 2,000 列限制屬 2026-11-22 版本；2026-11-24 的第 5 步更新已改為分批掃描。

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

### 2026-11-22 (first real Excel run) — failures were invisible, and an ageing pivot was silently dropped

The first run against a real workbook (753-row AR ageing) proved profiling and range reads work, but the pivot still failed and **the pane never said why**. Four defects fixed:

- **Execution errors were swallowed.** `executeClassified` only raised a toast; the audit showed `Execute FAILED · 1 actions` with no reason, and the model received a bare `Observed: FAILED`. The real Office.js message is now recorded in the audit **and** fed back to the model so it can correct itself.
- **A pivot value-field cap of 8 silently rejected an ageing pivot.** This workbook legitimately needs 11 value fields (USD O/S + nine buckets + Credits). The cap is now 16, and rejected actions are reported to the model and shown in the audit as `N rejected as unusable` instead of vanishing.
- **Agent shows its round and progress.** If work is pending at round 10, **Continue Agent · 繼續** resumes the same task in the current chat. Review requests and check failures have their own stop messages.
- **Agent 顯示輪數和工作進度。** 第 10 輪仍有工作時，可按 **Continue Agent · 繼續**，在目前對話接續原任務。待審批修改及檢查失敗會各自顯示停止原因。
- **Pivot field names are validated against the real headers.** A mismatch now names the offending field and lists the available headers, and removes the partial pivot and any sheet it created, instead of failing opaquely later. The destination sheet is also synced before use.

`npm run check` passes. `npm test` → **64 pass / 2 fail**, still only the Windows `spawn unzip ENOENT` packaging failures. Two new tests: an 11-value-field ageing pivot is accepted, and an unknown field produces a named error with no partial pivot left behind.

處理狀態（2026-11-22，首次真實 Excel 執行）：對真實活頁簿（753 列 AR ageing）的首次執行證明剖析與範圍讀取正常，但 pivot 仍失敗，而**pane 從未說明原因**。已修四個缺陷：**執行錯誤被吞掉** — `executeClassified` 只彈 toast，audit 顯示 `Execute FAILED · 1 actions` 而無原因，模型只收到 `Observed: FAILED`；現已把真實 Office.js 訊息記入 audit **並**回饋給模型讓其自行修正。**pivot 值欄上限 8 會靜默拒絕 ageing pivot** — 此活頁簿合理需要 11 個值欄（USD O/S + 九個桶 + Credits），上限現為 16，被拒 action 會告知模型並在 audit 顯示 `N rejected as unusable`。**「Stopped · validation failed」誤導** — `decide()` 在迴圈上限亦回 FAILED，令只是用完輪數的執行被標為驗證失敗；現已區分兩者。**pivot 欄位名稱會對真實表頭驗證** — 不符時會指出問題欄位並列出可用表頭，並移除已建立的部分 pivot 與工作表；目的地工作表亦會先 sync。`npm run check` 通過；`npm test` 為 **64 通過／2 失敗**。新增兩個測試。

## What this update revises

This is the Copilot-style revise. The earlier pane sent one prompt, attached a large dump of the active sheet, and applied one batch of actions. This update adds a mode, a plan, a short agent loop, checks calculated from the workbook, and an audit card.

- **Chat, Plan, and Agent** sit in the composer. Chat answers and leaves the file alone. Plan shows the steps and waits for **Approve**. Agent runs the steps.
- **Metadata first.** Plan and Agent send the workbook name, sheet names, used-range size, table names when Excel exposes them, and the selection address. They ask for one range when they need cell values. Chat still sends the selection, or the used range when the selection is a single cell.
- **The model plans. The pane writes.** Actions stay on the existing list: values, formulas, formatting, rows, columns, sheets, autofit, sort, charts, tables, and the Word selection actions. The model cannot call Office JavaScript on its own.
- **The pane reads the result.** After a write, Agent loads the cells again. It does not treat a reply as success until that read comes back.
- **Finance checks are calculated here.** Debit against credit, formula errors (`#N/A`, `#VALUE!`, `#REF!`, and the same family), blank required fields, duplicate keys, row count, and a total that moved. The model cannot mark these as passed. A failed check stops the loop.
- **Finance answers are traceable.** Finance prompts request bilingual fact, calculation, and assumption sections with exact sheet/cell or report-line citations. The evidence card lists the sources actually sent and flags citations outside them; a location match does not prove that the source supports a claim. Agent check results include the checked sheet and range; Chat arithmetic is not independently recalculated.
- **財務答案可追溯。** 財務問題會要求以雙語列出事實、計算及假設，並引用確切工作表／儲存格或報表行數。證據卡列出實際傳送的來源，並標示超出來源範圍的引用；位置比對不代表來源一定支持結論。Agent 檢查結果會列明工作表及範圍；Chat 算術不會由增益集獨立重算。
- **A failed check is a workpaper, not a silent fix.** The card states the issue, the evidence, a possible cause, the check, and the recommended action. A journal entry is a draft. The pane does not clear a reconciliation difference on its own.
- The pane runs on **Qwen only** (Grok removed 2026-11-24). It sends workbook data only when the Qwen key is already loaded in the relay.
- The AUDIT card is hidden. The same trail is saved with the session: request, mode, provider, plan, context sent, actions, blocked actions, and check results. API keys are left out. **New chat** clears the trail.
- **Direct, Apply, and Undo.** With Direct on, Agent writes safe actions immediately. Deletes, clears, renames, and formulas that can reach the network wait for a click. With Direct off, you approve the plan and then press **Apply edits**. **Undo last apply** is on the card. **Stop** cancels model generation and prevents the next step. An Office edit already in progress may finish.

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

**Two things are genuinely out of reach, not merely unbuilt.** It cannot run **Python in Excel** — Office.js exposes no API for it; "Python" appears in the typings only inside error-subtype enums. And it cannot become the **native Copilot surface** — the ribbon-integrated pane, inline ghost suggestions and the built-in Analyze Data experience are not something a sideloaded task pane can turn into. The attachment picker currently supports HTML and text-based files only; it does not parse PDF, DOCX, or XLSX.

## First time on a computer

See **[Install](#install-安裝)** at the top for the short version. In full:

1. Install Node.js 20 or newer from https://nodejs.org.
2. Close Excel and Word.
3. From the project folder run `npm run setup` (Node check + `npm install`), then `npm start`. This sideloads the add-in and starts the relay, and installs **no** auto-start.
4. Prefer one double-click? `Setup and Start.bat` does the same **and** installs the Windows logon auto-start. By hand that is `npm install`, `npm run autostart`, then `npm start`.
5. Open the pane from **Home > Add-ins**. The add-in name is **Qwen for Excel and Word**.
6. **Provider-terms hold:** every user is intended to enter their own key, but the current add-in defaults to Alibaba Token Plan. Alibaba's current policy excludes Token Plan from custom app backends. Do not use a Token Plan key with this add-in unless Alibaba confirms permission. The supported standard API/credential flow is being prepared; see [`MARKETPLACE_READINESS.md`](MARKETPLACE_READINESS.md).

To open the same add-in in Word, run `npm run start:word` once with Word closed. `npm start` only launches Excel.

Full installation detail, verification and uninstall: **[SOP-INSTALL.md](SOP-INSTALL.md)**.

## Every later sign-in

Windows starts the relay at logon with no window. A per-minute task named `ExcelQwenGrokRelayEnsure` starts it again if it stops. Open Excel or Word and the pane is ready. If it is missing, use **Home > Add-ins**.

Stop the relay with `npm run relay-stop`. The minute task will start it again unless you first run `npm run autostart-remove`. `npm run stop` removes the sideload. The next launch then needs `Setup and Start.bat` or `npm start`.

After a code change, restart the relay (`npm run relay-stop`, then `npm run relay-ensure`) and close and reopen the pane. A remembered key stays in the pane until you press **Clear key**. Stopping the relay clears the copy that was only in relay memory.

## Using the pane

Press **Workbook** at the bottom to open WORKBOOK ACCESS. It is not in the sidebar. Press **Workbook** again to hide it. The card shows what the next question will send.

- In Chat, a multi-cell Excel selection is sent. A single cell falls back to the used range of the active sheet. Formulas appear as `value [=formula]`. The read stops at the first 150 rows and 30 columns, then at 45,000 characters. A truncation note says so.
- In Plan and Agent, the next question sends workbook metadata: the workbook name, sheet names within an 8,000-character metadata cap (up to 30 with used-range sizes and table names), and the selection address. Agent chooses which sheet to inspect and can profile a large sheet in local batches. **＋ Active sheet** is a Chat shortcut and does not mean the whole workbook.
- In Word Chat, the selection is sent. An empty selection sends the document body, with the same character cap. Plan and Agent send a short document outline and wait to be asked for the selection.
- The card refreshes when the selection changes. **↻ Refresh** reads it again.
- Untick **Share Office file data with AI** to prevent open workbook/document data and Agent-requested ranges from being sent. Attached files are controlled separately.
- Use **＋** to upload up to three files. HTML/HTM, TXT, MD, CSV, TSV, JSON, and XML are up to 5 MB. PDF, DOCX, XLSX/XLSM, and PNG/JPG/WEBP go through the local relay and are up to 20 MB. HTML tables and readable text are extracted locally; scripts, CSS, chart images, and links are not run or interpreted. The pane shows a preview of extracted text before sending. It sends at most 18,000 characters per file and 26,000 characters from attachments per question. **New chat** clears attachments and conversation history.

Ask in the box and press Enter. Shift+Enter adds a line. Replies stream. The send button becomes **Stop** while a reply is running. **New chat** clears the conversation and the audit trail. **Model** is `auto`, `qwen4-max` (default), `qwen3.8-max`, or `qwen3.8-flash`. `auto` lets Qwen choose. **Effort** is Off, Low, Medium, High (default), or Extra; the relay sends `reasoning_effort` and does not send `thinking_budget`. **Fast** still skips thinking for that question, even when Effort is High or Extra. The choice is remembered on this PC.

The footer sets Chat, Plan, or Agent. The pane runs on Qwen. The dot is green when the Qwen key is loaded, amber when the relay is up and the key is not, and red when the relay is down. The sidebar ☰ holds **Guide · 說明** (collapsed until opened), **MCP**, **Skills**, History, and Memory. There is no bottom remark under the question box.

**↳ Insert last answer into selection** writes the last reply into the top-left cell of the Excel selection. One line that starts with `=` is written as a formula. Anything else is written as text.

History stays in memory. Reloading the pane clears it. Each request sends at most 24 messages. Only the latest question includes workbook context and selected attachments, so old context is not repeatedly sent.

## Modes

**Chat** sends the selection and answers. Any action list in the reply is ignored, including when Direct is on.

**Plan** inspects metadata, then returns one plan: goal, steps, assumptions, required inputs, risks, and the checks it wants. Proposed actions sit on the same card. Nothing is written until you press **Approve**. If Direct is off, **Approve** is followed by **Apply edits**.

**Agent** uses the same metadata. It can ask for one range, run the safe actions for that step, read those cells back, and calculate the checks. The composer shows `Agent 1/25` and the current activity. At the round limit, use **Continue Agent · 繼續** if offered. Use the review card for pending edits. With Direct off, a step that would write waits for **Approve**, then **Apply edits**.

The pane runs on Qwen only. The current Token Plan endpoint is under provider-terms review; do not use a Token Plan key with this custom app. The supported standard API/credential flow is pending.

The AUDIT card is hidden. The saved session still lists the request, the provider, the context that was sent, the actions, and the check results.

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
| Legacy Qwen Token Plan test configuration — do not use with this custom add-in | `https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` | Token Plan key, `sk-sp-…` | `qwen4-max` |

A Token Plan key must not be used with this custom application backend unless Alibaba confirms the use is permitted. A normal Model Studio key must use its matching regional/workspace endpoint; the application's standard API and secure credential flow are not yet finalized. See the current [Alibaba API-key guidance](https://help.aliyun.com/en/model-studio/get-api-key) before using a supported release.

A Token Plan key also carries its own model list, which is narrower than the public Qwen docs. `qwen-max`, `qwen-turbo`, and `qwen-plus` are not on it and return `400 Model not exist` — that reads like a broken URL but is not one. The picker offers `auto`, `qwen4-max`, `qwen3.8-max`, and `qwen3.8-flash`. Any other id is rejected before the request is sent. `qwen4-max`, `qwen3.8-max`, `qwen3.7-plus`, and `qwen3.6-flash` were confirmed working on 2026-11-22; `qwen3.7-plus` and `qwen3.6-flash` are no longer in the picker.

The relay listens on `127.0.0.1` only and rejects any other Host or Origin.

## Cloudflare

The same pane is published at `https://excel-qwen-grok.anthonym-354.workers.dev` so Excel on the web can open it. The local server stays the way to run it on this PC.

The manifest to upload in Excel is `manifest.cloudflare.xml` (version 1.5.0.0). It has its own add-in id, so it can sit beside the local sideload. With Microsoft 365 Premium, add it from **Home → Add-ins → Upload My Add-in**.

The current hosted test build accepts each user's own Qwen API key, stores it in Office/browser storage, and forwards it through the Worker. Keys are not shared. **Do not use Alibaba Token Plan keys with this custom app unless Alibaba confirms permission; the current endpoint is a legacy test configuration.** The hosted long-term-key flow also needs an approved design before public release. See `MARKETPLACE_READINESS.md`.

目前託管測試版接受各使用者自己的 Qwen API key，存於 Office／瀏覽器儲存並經 Worker 轉送；key 不會共用。**除非阿里雲確認允許，請勿在此自訂應用程式使用 Alibaba Token Plan key；目前 endpoint 僅屬舊測試設定。** 託管版長期 key 流程亦須在公開發行前改為經確認的設計。詳見 `MARKETPLACE_READINESS.md`。

A later code change is published with `npm run deploy:cloudflare` after `npx wrangler login`. Workbook text sent from the hosted pane goes to Cloudflare and then to Qwen.

## Another computer

Each computer runs its own relay and stores its own keys.

1. Download `download/Qwen-for-Excel-and-Word.zip` and unzip it. Open the folder inside. It has no API keys. `npm run package` rebuilds that zip after a code change.
2. Install Node.js 20 or newer, close Excel, and double-click `Setup and Start.bat`.
3. Each person has their own key; no key is shared. Do not use a Token Plan key with this custom app until Alibaba confirms permission and the add-in is updated.

## Commands

| Command | What it does |
|---|---|
| `npm run setup` | **CLI install:** check Node 20+, run `npm install`, print the next steps. No auto-start, no key |
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
| `npm run package` | Rebuild `download/Qwen-for-Excel-and-Word.zip` |

Auto-start is Windows-only. It uses `%APPDATA%`.

## Tests

`npm run check` and `npm test` need no API key. The pane tests stub the page, Excel, Word, and the stream. They cover context, the click-gated edit card, Direct revise, Undo, risky formulas, Word replace of the selection, Chat, Plan, Agent, a debit-and-credit stop, and Qwen-only routing that never sends without a key. The relay tests cover the loopback lock, the Qwen URL allowlist, Fast mode, the Agent mode note, and stream errors. The download-zip tests read the pack with a pure-Node reader (`readZipEntry`, `listZipEntries`) rather than the Unix `unzip` binary, so **the whole suite passes on Windows**; they assert the zip carries both SOPs, both manifests, a `package.json` with a `setup` script, and no `.env`, `node_modules` or `.git` content.

Before trusting a change in the real app, start the relay, load a key, ask one question in Chat, and confirm the answer. For a revise, use Plan or Agent, confirm the card, and use Undo once.

## If something is wrong

| What you see | What to do |
|---|---|
| Relay unavailable | Run `npm run relay-ensure`, or wait up to a minute for the self-heal task |
| Port 3000 is in use | Another relay is running. `npm run relay-stop`, then start it again |
| The pane looks old | Restart the relay, then close and reopen the pane. On another computer, `git pull` first |
| Context says truncated | Select a smaller range, or ask about one section at a time |
| Invalid conversation message | The message plus context is over 50,000 characters. Shorten it, remove an attachment, or untick **Send workbook context** |
| Provider HTTP 401 | The key and endpoint may not match. Do not switch to a Token Plan key for this custom app; use only the supported API flow after it is released |
| Model not exist (HTTP 400) | The model is not on that key's list. A Token Plan key does not serve `qwen-max`, `qwen-turbo`, or `qwen-plus`. Pick one it does, in Settings |
| Qwen endpoint rejected | The API route is being updated for Marketplace readiness; see `MARKETPLACE_READINESS.md` |
| A reply takes too long | Press **Stop**. Tick **Fast** for a shorter reply |
| Stopped. A validation failed | The difference is still in the sheet. Correct it yourself, then ask again |
| A change looks wrong | Press **Undo last apply**, or Excel `Ctrl+Z` |
| The add-in is missing in Word | Run `npm run start:word` once with Word closed |
