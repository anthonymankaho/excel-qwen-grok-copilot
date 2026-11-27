# SOP 1 — Installation 安裝程序

**Qwen for Excel and Word.** This document covers getting the add-in onto a PC and keeping it running. For how to *use* it, read `SOP-USER-GUIDE.md` (說明書).

Each person uses **their own API key**. The owner's Token Plan key stays on the owner's computer and is never shared, never committed, and never pasted into chat.

---

## 1. What you need 需要什麼

| Requirement | Detail |
|---|---|
| Node.js | **20 or newer** — https://nodejs.org |
| OS | Windows for the desktop sideload and for auto-start. The relay itself runs anywhere Node runs |
| Office | Desktop **Excel** and/or **Word** (Microsoft 365 / 2019+). Excel on the web uses the hosted route in §7 |
| Folder | The folder that holds `manifest.xml`. Every command below runs from there |
| Key | Your own Qwen key. Nothing is bundled |

The pane and the local relay both live at `http://127.0.0.1:3000`. Both manifests are version **1.5.0.0**. The local build to use is **2026-11-26.4**.

**This PC’s folder is `C:\Users\User\Desktop\excel-qwen-grok-copilot`.** Run every `npm` command from that folder. PowerShell opened in `C:\WINDOWS\System32` cannot see `package.json` and fails with `ENOENT`. The older folder `C:\Users\User\Desktop\Excel Qwen Grok Copilot` is an archived checkout. Do not install or edit it.

**這部電腦的資料夾是 `C:\Users\User\Desktop\excel-qwen-grok-copilot`。** 所有 `npm` 指令都在該資料夾執行。在 `C:\WINDOWS\System32` 執行會找不到 `package.json`。較舊的 `C:\Users\User\Desktop\Excel Qwen Grok Copilot` 是已封存的副本，不要在那裡安裝或修改。

> **Quote every path.** If your user folder contains a space — e.g. `C:\Users\Anthony Man` — an unquoted path breaks in PowerShell. Always write `cd "C:\Users\Your Name\..."`.
> **每個路徑都要加引號。** 使用者資料夾若含空格（例如 `C:\Users\Anthony Man`），未加引號的路徑在 PowerShell 會失敗。

---

## 2. Get the files 取得檔案

**Option A — zip (normal route for teammates).** The owner sends `Qwen-for-Excel-and-Word.zip`. Extract it anywhere outside a cloud-synced folder if you can. The zip contains **no keys**.

**Option B — git (only if you have repository access).** The repository is private.

```powershell
git clone https://github.com/anthonymankaho/excel-qwen-grok-copilot.git
cd "excel-qwen-grok-copilot"
```

To rebuild the zip yourself, from the project folder:

```powershell
npm run package
```

That writes `download/Qwen-for-Excel-and-Word.zip`, including a generated `START-HERE.txt`. It skips `node_modules`, `.git`, `.wrangler`, `download`, and any `.env*` file.

---

## 3. Install from a terminal (recommended) 用終端機安裝（建議）

**Close Excel and Word first.** The sideload cannot register while they are open.

```powershell
Set-Location "C:\Users\User\Desktop\excel-qwen-grok-copilot"
npm run setup
npm start
```

On another PC, use the folder where you extracted the zip or cloned the repo. Quote the path when it contains a space.

- **`npm run setup`** checks your Node version, runs `npm install`, and prints the next steps. It installs **no** auto-start and writes **no** key.
- **`npm start`** sideloads `manifest.xml` into desktop Excel and opens Excel. **Leave that terminal window open — it *is* the relay.** Closing it stops the relay.

Prefer plain npm? `npm install` then `npm start` is exactly equivalent; `setup` only adds the version check and the guidance.

Then in Excel: **Home → Add-ins → Qwen for Excel and Word**.

### First key entry 首次輸入 key

**Provider-terms hold:** the current build defaults to Alibaba Token Plan, which Alibaba does not allow through custom application backends. Do not enter a Token Plan key in this add-in unless Alibaba confirms permission. The supported Model Studio API/credential flow is still being prepared. See `MARKETPLACE_READINESS.md` before entering a key.

### Optional cloud sources（選用雲端來源）

OneDrive search is desktop/local-relay only. Each colleague signs in to their own personal Microsoft account; no personal email address or user token is shared. The app owner configures the public Microsoft OAuth app ID once in `cloud-oauth.json` before packaging or sharing the add-in.

1. Create a Microsoft public app registration that supports personal Microsoft accounts. Add the **Mobile and desktop applications** platform with `http://localhost` redirect support and delegated `Files.Read`. Sign-in uses the system browser and PKCE; no Microsoft client secret is used.
2. Fill `microsoftClientId` once in the root `cloud-oauth.json`, then include the file in the shared package. Restart the local relay after changing it. This is a public application ID, not a user credential. Do not put account emails, access tokens, or client secrets in the file.
3. Each colleague opens **MCP → Connect** and signs in through the system browser. Search results are read-only. A selected result reaches Qwen only after **Review and add to chat** and sending a question.

Each person uses their own account and grants consent separately; the pane no longer asks colleagues to enter OAuth app settings. These connectors are unavailable in the hosted Excel-on-the-web pane. See `README.md` and `SOP-USER-GUIDE.md` for details.

#### 選用雲端來源

OneDrive 搜尋只支援桌面版／本機 Relay。每位同事登入自己的個人 Microsoft 帳戶；不會共用個人電郵地址或使用者權杖。管理者在分享或打包增益集前，只需在 `cloud-oauth.json` 設定一次公開 Microsoft OAuth 應用程式 ID。

1. **Microsoft：** 建立支援個人 Microsoft 帳戶的 public app registration，加入 **Mobile and desktop applications** 平台、`http://localhost` redirect 支援和委派 `Files.Read` 權限。登入使用系統瀏覽器和 PKCE；不使用 Microsoft client secret。
2. 在根目錄 `cloud-oauth.json` 一次填入 `microsoftClientId`，並在分享套件時附上此檔案；更改後要重新啟動本機 Relay。這是公開應用程式 ID，不是使用者憑證。不可在檔案放入帳戶電郵、存取權杖或 client secret。
3. 每位同事開啟 **MCP → Connect**，在系統瀏覽器登入。搜尋結果只讀；只有按 **Review and add to chat** 並傳送問題後，所選結果才會傳給 Qwen。

每位使用者各自登入及授權；窗格不再要求同事輸入 OAuth 應用程式設定。Excel 網頁版不支援這些連接器。詳情見 `README.md` 和 `SOP-USER-GUIDE.md`。

---

## 4. Double-click alternative 雙擊替代方案

`Setup and Start.bat` runs `npm install`, then `npm run autostart`, then `npm start`.

**Understand what `autostart` adds before you use it.** It writes two items for the current Windows user:

- Startup item `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\ExcelQwenGrokRelay.bat`
- Scheduled task `ExcelQwenGrokRelayEnsure`, running **every minute**, which restarts the relay if it is down

That is convenient but it is persistent background automation. **`npm start` alone is enough to use the add-in.** Choose the bat only if you want the relay alive at every sign-in. Auto-start needs Windows; on another OS the script stops and says so.

Remove it later with `npm run autostart-remove`.

---

## 5. Word

Word needs its own sideload. **Close Word**, then from the project folder:

```powershell
npm run start:word
```

The pane name is the same. Excel and Word can both be sideloaded; each host loads the pane into its own window.

---

## 6. Verify the install 驗證安裝

From the project folder:

```powershell
npm run check
npm test
```

Both must pass and neither needs an API key. `check` syntax-checks the six JavaScript files; `test` runs the offline suite covering the pane, the relay, Agent behaviour, Undo, finance checks and the download zip.

Then confirm the relay is alive:

```powershell
curl http://127.0.0.1:3000/api/health
```

Expect `{"ok":true,"configured":{"qwen":false}}` before you enter a key, and `qwen:true` after. Finally ask one real question in **Chat** mode — a reply proves the whole chain.

---

## 7. Excel on the web / hosted pane 網頁版 Excel

Use this when the pane must open from an `https://` address. No local install is needed.

1. In Excel: **Home → Add-ins → Upload My Add-in** → choose `manifest.cloudflare.xml`.
2. The pane address is `https://excel-qwen-grok.anthonym-354.workers.dev`.
3. The current hosted test build asks each user to enter their own Qwen key in Settings; it saves the key in Office/browser storage when available and sends it per request through the Worker. Each user's key is separate; no key is shared, and no provider key is embedded in the Worker.
4. **Marketplace hold:** the current default is an Alibaba Token Plan endpoint. Alibaba's current policy does not allow Token Plan keys through custom application backends, even when the key belongs to that user. Do not use Token Plan credentials with this add-in unless Alibaba confirms the use is permitted. The hosted long-term-key storage also needs an approved design before public release. See `MARKETPLACE_READINESS.md`.

The Worker exposes the pane and forwards user-authorized requests; it has no provider key configured. Do not add a provider key or relay secret to this repository. After a code change, redeploy with `npm run deploy:cloudflare` only when the release is ready and authorized.

---

## 8. Daily running 日常運作

**With auto-start installed:** sign in to Windows and the relay starts with no window. Open Excel; if the pane is missing use **Home → Add-ins**. If Settings says the relay is down, wait up to a minute or run `npm run relay-ensure`.

**Without auto-start (the `npm start` route):** run `npm start` from the project folder each session and leave the window open. `npm run dev-server` starts the relay alone, without re-sideloading or opening Excel — useful when the add-in is already registered.

The sideload persists across Excel restarts until `npm run stop`. It is registered per user at `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`.

### After a code change 程式碼變更後

```powershell
Set-Location "C:\Users\User\Desktop\excel-qwen-grok-copilot"
npm run relay-stop
npm run relay-ensure
```

Then close the task pane and open it again. `npm start` is only needed when you also want to re-sideload and open Excel. Run these commands inside the project folder, not from `C:\WINDOWS\System32`.

**The relay loads its code at process start**, so a code change needs a restart. Pane files are served with `no-store`, so reopening the pane loads the new screen. On another computer, `git pull` first. The pane and relay builds must match. The current local build is **2026-11-26.4**. It includes the model list (`auto`, `qwen4-max`, `qwen3.8-max`, `qwen3.8-flash`), Effort, GitHub skill install, GitHub MCP, **Connect to Github** through the GitHub CLI login on this PC, Word `open` / `git pull` / `git commit` / `git push` for one local file, appending each finished Qwen answer in Word, and the pane names **Workbook**, **MCP**, and **Skills**. `git push` does not use `--force`. **Disconnect** in the pane does not sign the PC out of GitHub. See `SOP-USER-GUIDE.md`.

A restart clears keys held only in relay memory. If **Remember key** was ticked, the pane reloads it automatically on reconnect; otherwise re-enter it in Settings.

---

## 9. Stop and uninstall 停止與移除

| Command | Effect |
|---|---|
| `npm run relay-stop` | Stops the relay. Keys held only in relay memory are dropped. The sideload stays |
| `npm run stop` | Removes the sideload from Excel. Next launch needs `npm start` |
| `npm run autostart-remove` | Removes the Startup item and the per-minute task |
| **Clear key** in Settings | Removes the remembered key from this Windows user and from the relay |

If auto-start is installed, `relay-stop` alone is not enough — the minute task restarts it. Run `autostart-remove` first.

---

## 10. Install troubleshooting 安裝疑難排解

| What you see | What to do |
|---|---|
| `Could not read package.json` / `ENOENT` | PowerShell is not in the project folder. Run `Set-Location "C:\Users\User\Desktop\excel-qwen-grok-copilot"` first. On this PC that is the only live folder |
| `Set-Location : 找不到接受引數 ... 的位置參數` | The path has a space and was not quoted. Use `Set-Location "C:\Users\Your Name\..."` |
| `npm start` fails to sideload | Excel or Word is still open. Close both and retry |
| Relay unavailable in Settings | The terminal running the relay was closed. Run `npm start` or `npm run dev-server` again |
| Port 3000 in use | `npm run relay-stop`, then start again |
| `No relay listening on port 3000.` | Nothing to stop — that is a normal message, not an error |
| Pane offline just after sign-in | Auto-start needs up to a minute. Or run `npm run relay-ensure` |
| Pane shows an old screen | Restart the relay and reopen the pane. On another PC, `git pull` first |
| Word does not list the add-in | `npm run start:word` with Word closed |
| `npm test` shows 2 failures about `unzip` | Older packs only. Current versions read the zip in Node and pass on Windows |
| Provider HTTP 401 | Key and URL do not match — see `SOP-USER-GUIDE.md` §7 |
| `Model not exist` or `Unsupported model` | Use `auto`, `qwen4-max`, `qwen3.8-max`, or `qwen3.8-flash` in the composer or Settings |

---

## 11. Handing it to a teammate 交給隊友

1. Send the zip (**no keys inside**) or grant repository access.
2. They install Node 20+, close Excel, and run `npm run setup` then `npm start` — or double-click `Setup and Start.bat` if they want auto-start.
3. They enter **their own** key with Remember ticked.
4. Each computer has its own relay, its own optional startup item, and its own remembered keys. Nothing is shared between machines.
5. The hosted pane currently asks each user for their own key, but it is not ready for public use with Token Plan credentials. Follow `MARKETPLACE_READINESS.md` before relying on the hosted service.

---

## 12. Processing status · 處理狀態

**2026-11-26 — Review fixes.** Local build `2026-11-26.4`. In Word, only a file path after `open` loads a repo file. A failed `git commit` restores that file. Git stops after 90 seconds if it is waiting. The download zip matches this tree. Restart the local relay and reopen the pane. The hosted Worker was not redeployed. `npm test` 152/152.

**2026-11-26 — 覆核修正。** 本機版本 `2026-11-26.4`。在 Word，只有 `open` 後面是檔案路徑才會載入 repo 檔。`git commit` 失敗會還原該檔。Git 若一直等候，90 秒後會停止。下載 zip 與目前程式一致。請重啟本機 relay 並重新開啟窗格。託管 Worker 未重新部署。

**2026-11-26 — Hosted pane deployed.** `npx wrangler deploy` from this folder published commit `6534f67`. Live health is `build: 2026-11-25.3`, `byo:true`, `configured.qwen:false`. Close and reopen the Excel web task pane. MCP, Skills, History, Memory, and PDF/Word/Excel extraction remain local-only. The email-and-password team version is cancelled and was never deployed. A Token Plan key cannot back that shared login.

**2026-11-26 — 託管窗格已部署。** 在此資料夾執行 `npx wrangler deploy`，發佈 commit `6534f67`。線上 health 為 `build: 2026-11-25.3`、`byo:true`、`configured.qwen:false`。請關閉並重新開啟 Excel 網頁版工作窗格。MCP、Skills、History、Memory，以及 PDF／Word／Excel 擷取仍只在本機。email 加密碼的團隊版已取消，且從未部署。Token Plan key 不能用於那種共用登入。

**2026-11-25 — Pane names and GitHub skill install.** Commit `6534f67`. Local build `2026-11-25.3`. **Skills** can install one `SKILL.md` from `https://github.com/owner/repo`. The token stays in this browser when Remember is ticked. **Guide** starts collapsed. **MCP** is the OneDrive card. **Workbook** opens access above the button. The hosted Worker was not redeployed.

**2026-11-25 — 窗格名稱與從 GitHub 安裝技能。** Commit `6534f67`。本機版本 `2026-11-25.3`。**Skills** 可從 `https://github.com/owner/repo` 安裝一個 `SKILL.md`。勾選 Remember 時，token 只留在這個瀏覽器。**Guide** 預設收合。**MCP** 是 OneDrive 卡。**Workbook** 在按鈕上方打開存取預覽。託管 Worker 未重新部署。

**2026-11-25 — Folder, build, and restart.** This PC uses `C:\Users\User\Desktop\excel-qwen-grok-copilot` only. Local build `2026-11-25.2`. Both manifests are `1.5.0.0`. After a code change, run `npm run relay-stop` then `npm run relay-ensure` from that folder and reopen the pane. The hosted Worker was not redeployed. `npm test` 123/123.

**2026-11-25 — 資料夾、版本與重啟。** 這部電腦只用 `C:\Users\User\Desktop\excel-qwen-grok-copilot`。本機版本 `2026-11-25.2`。兩個 manifest 都是 `1.5.0.0`。程式變更後，在該資料夾執行 `npm run relay-stop`，再執行 `npm run relay-ensure`，然後重新開啟窗格。託管 Worker 未重新部署。`npm test` 123/123。

**2026-11-22 — CLI install path added; docs split in two.**

- `npm run setup` is the new CLI entry point: Node version check, `npm install`, then printed next steps. It installs no auto-start and writes no key.
- `START-HERE.txt` in the zip now leads with the terminal route and quotes every example path.
- This installation SOP and `SOP-USER-GUIDE.md` replace the single combined runbook. `SOP.md` is now an index.
- The download zip is verified by a pure-Node reader (`readZipEntry`, `listZipEntries`) instead of the Unix `unzip` binary, so `npm test` passes on Windows. The pack is asserted to contain both SOPs, the manifests, `package.json` with a `setup` script, and no `.env`, `node_modules` or `.git` content.
- The local add-in **is** installed on the owner's Surface Book at `C:\Users\Anthony Man\excel-qwen-grok-copilot`, sideloaded, with auto-start deliberately **not** installed. An earlier note here saying it was absent is superseded.

處理狀態（2026-11-22）：新增 CLI 安裝路徑，文件一分為二。**`npm run setup`** 是新的 CLI 入口：檢查 Node 版本、執行 `npm install`、然後印出下一步；它**不**安裝 auto-start，亦**不**寫入任何 key。zip 內的 `START-HERE.txt` 現以終端機步驟為主，並為每個範例路徑加引號。本安裝 SOP 與 `SOP-USER-GUIDE.md` 取代原本單一的 runbook，`SOP.md` 現為索引。下載 zip 改以純 Node 讀取器（`readZipEntry`、`listZipEntries`）驗證，不再依賴 Unix 的 `unzip`，因此 `npm test` 在 Windows 上全部通過；並斷言 zip 內含兩份 SOP、兩個 manifest、含 `setup` 腳本的 `package.json`，且不含 `.env`、`node_modules` 或 `.git`。本地 add-in **已**安裝於擁有人的 Surface Book（`C:\Users\Anthony Man\excel-qwen-grok-copilot`），已 sideload，且刻意**未**安裝 auto-start；先前此處指其未安裝的說法已被取代。
