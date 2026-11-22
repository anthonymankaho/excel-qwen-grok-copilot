# SOP 1 — Installation 安裝程序

**Qwen + Grok for Excel and Word.** This document covers getting the add-in onto a PC and keeping it running. For how to *use* it, read `SOP-USER-GUIDE.md` (說明書).

Each person uses **their own API key**. The owner's Token Plan key stays on the owner's computer and is never shared, never committed, and never pasted into chat.

---

## 1. What you need 需要什麼

| Requirement | Detail |
|---|---|
| Node.js | **20 or newer** — https://nodejs.org |
| OS | Windows for the desktop sideload and for auto-start. The relay itself runs anywhere Node runs |
| Office | Desktop **Excel** and/or **Word** (Microsoft 365 / 2019+). Excel on the web uses the hosted route in §7 |
| Folder | The folder that holds `manifest.xml`. Every command below runs from there |
| Key | Your own Qwen or Grok key. Nothing is bundled |

The pane and the local relay both live at `http://127.0.0.1:3000`. The local manifest is version **1.2.0.0**; the hosted manifest is **1.3.0.0**.

> **Quote every path.** If your user folder contains a space — e.g. `C:\Users\Anthony Man` — an unquoted path breaks in PowerShell. Always write `cd "C:\Users\Your Name\..."`.
> **每個路徑都要加引號。** 使用者資料夾若含空格（例如 `C:\Users\Anthony Man`），未加引號的路徑在 PowerShell 會失敗。

---

## 2. Get the files 取得檔案

**Option A — zip (normal route for teammates).** The owner sends `Qwen-Grok-for-Excel-and-Word.zip`. Extract it anywhere outside a cloud-synced folder if you can. The zip contains **no keys**.

**Option B — git (only if you have repository access).** The repository is private.

```powershell
git clone https://github.com/anthonymankaho/excel-qwen-grok-copilot.git
cd "excel-qwen-grok-copilot"
```

To rebuild the zip yourself, from the project folder:

```powershell
npm run package
```

That writes `download/Qwen-Grok-for-Excel-and-Word.zip`, including a generated `START-HERE.txt`. It skips `node_modules`, `.git`, `.wrangler`, `download`, and any `.env*` file.

---

## 3. Install from a terminal (recommended) 用終端機安裝（建議）

**Close Excel and Word first.** The sideload cannot register while they are open.

```powershell
cd "C:\Users\Your Name\Qwen-Grok-for-Excel-and-Word"
npm run setup
npm start
```

- **`npm run setup`** checks your Node version, runs `npm install`, and prints the next steps. It installs **no** auto-start and writes **no** key.
- **`npm start`** sideloads `manifest.xml` into desktop Excel and opens Excel. **Leave that terminal window open — it *is* the relay.** Closing it stops the relay.

Prefer plain npm? `npm install` then `npm start` is exactly equivalent; `setup` only adds the version check and the guidance.

Then in Excel: **Home → Add-ins → Qwen + Grok for Excel and Word**.

### First key entry 首次輸入 key

1. In the pane open **Settings**.
2. Pick **Qwen** or **Grok**.
3. Paste **your own** key.
4. Leave **Remember key on this PC** ticked.
5. Press **Use key**. The provider dot turns green.

Pair the key with the right URL — see `SOP-USER-GUIDE.md` §7. A `sk-sp-…` Token Plan key uses the Token Plan URL, which is already the default.

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

Both must pass and neither needs an API key. `check` syntax-checks the five JavaScript files; `test` runs the offline suite covering the pane, the relay, Agent behaviour, Undo, finance checks and the download zip.

Then confirm the relay is alive:

```powershell
curl http://127.0.0.1:3000/api/health
```

Expect `{"ok":true,"configured":{"qwen":false,"grok":false}}` before you enter a key, and `qwen:true` after. Finally ask one real question in **Chat** mode — a reply proves the whole chain.

---

## 7. Excel on the web / hosted pane 網頁版 Excel

Use this when the pane must open from an `https://` address. No local install is needed.

1. In Excel: **Home → Add-ins → Upload My Add-in** → choose `manifest.cloudflare.xml`.
2. The pane address is `https://excel-qwen-grok.anthonym-354.workers.dev`.
3. Open **Settings**, paste the **relay token**, press **Use key**. The matching `RELAY_TOKEN` secret is already on the Worker. Ask the owner for the token; it is not in the zip.
4. **No provider key is needed** — `QWEN_API_KEY` is already a Worker secret (set 2026-11-22). `GROK_API_KEY` is **not** set, so the hosted pane runs Qwen only and **Auto** stays on Qwen.

Owner-only, to set or replace a provider secret:

```powershell
npx wrangler login
npx wrangler secret put QWEN_API_KEY
npx wrangler secret put GROK_API_KEY
npx wrangler secret list --name excel-qwen-grok
```

`secret list` prints **names only, never values**. Never put a key in the repository or in chat. After a code change, redeploy with `npm run deploy:cloudflare`. To rotate the relay token, `npx wrangler secret put RELAY_TOKEN` with a passphrase of at least 16 characters, then paste the new value into Settings.

The hosted route rejects a missing token, a wrong token, a call from another site, and a pasted provider key.

---

## 8. Daily running 日常運作

**With auto-start installed:** sign in to Windows and the relay starts with no window. Open Excel; if the pane is missing use **Home → Add-ins**. If Settings says the relay is down, wait up to a minute or run `npm run relay-ensure`.

**Without auto-start (the `npm start` route):** run `npm start` from the project folder each session and leave the window open. `npm run dev-server` starts the relay alone, without re-sideloading or opening Excel — useful when the add-in is already registered.

The sideload persists across Excel restarts until `npm run stop`. It is registered per user at `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`.

### After a code change 程式碼變更後

```powershell
npm run relay-stop
npm start
```

**The relay caches its prompt at process start**, so a prompt change needs a restart. Pane files are served with `no-store`, so reopening the pane is enough for those. On another computer, `git pull` first.

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
| `Set-Location : 找不到接受引數 ... 的位置參數` | The path has a space and was not quoted. Use `cd "C:\Users\Your Name\..."` |
| `npm start` fails to sideload | Excel or Word is still open. Close both and retry |
| Relay unavailable in Settings | The terminal running the relay was closed. Run `npm start` or `npm run dev-server` again |
| Port 3000 in use | `npm run relay-stop`, then start again |
| `No relay listening on port 3000.` | Nothing to stop — that is a normal message, not an error |
| Pane offline just after sign-in | Auto-start needs up to a minute. Or run `npm run relay-ensure` |
| Pane shows an old screen | Restart the relay and reopen the pane. On another PC, `git pull` first |
| Word does not list the add-in | `npm run start:word` with Word closed |
| `npm test` shows 2 failures about `unzip` | Older packs only. Current versions read the zip in Node and pass on Windows |
| Provider HTTP 401 | Key and URL do not match — see `SOP-USER-GUIDE.md` §7 |
| `Model not exist` (HTTP 400) | That model is not on your key's plan. Change the model in Settings |

---

## 11. Handing it to a teammate 交給隊友

1. Send the zip (**no keys inside**) or grant repository access.
2. They install Node 20+, close Excel, and run `npm run setup` then `npm start` — or double-click `Setup and Start.bat` if they want auto-start.
3. They enter **their own** key with Remember ticked.
4. Each computer has its own relay, its own optional startup item, and its own remembered keys. Nothing is shared between machines.
5. For the hosted pane instead, they need only the relay token from you.

---

## 12. Processing status · 處理狀態

**2026-11-22 — CLI install path added; docs split in two.**

- `npm run setup` is the new CLI entry point: Node version check, `npm install`, then printed next steps. It installs no auto-start and writes no key.
- `START-HERE.txt` in the zip now leads with the terminal route and quotes every example path.
- This installation SOP and `SOP-USER-GUIDE.md` replace the single combined runbook. `SOP.md` is now an index.
- The download zip is verified by a pure-Node reader (`readZipEntry`, `listZipEntries`) instead of the Unix `unzip` binary, so `npm test` passes on Windows. The pack is asserted to contain both SOPs, the manifests, `package.json` with a `setup` script, and no `.env`, `node_modules` or `.git` content.
- The local add-in **is** installed on the owner's Surface Book at `C:\Users\Anthony Man\excel-qwen-grok-copilot`, sideloaded, with auto-start deliberately **not** installed. An earlier note here saying it was absent is superseded.

處理狀態（2026-11-22）：新增 CLI 安裝路徑，文件一分為二。**`npm run setup`** 是新的 CLI 入口：檢查 Node 版本、執行 `npm install`、然後印出下一步；它**不**安裝 auto-start，亦**不**寫入任何 key。zip 內的 `START-HERE.txt` 現以終端機步驟為主，並為每個範例路徑加引號。本安裝 SOP 與 `SOP-USER-GUIDE.md` 取代原本單一的 runbook，`SOP.md` 現為索引。下載 zip 改以純 Node 讀取器（`readZipEntry`、`listZipEntries`）驗證，不再依賴 Unix 的 `unzip`，因此 `npm test` 在 Windows 上全部通過；並斷言 zip 內含兩份 SOP、兩個 manifest、含 `setup` 腳本的 `package.json`，且不含 `.env`、`node_modules` 或 `.git`。本地 add-in **已**安裝於擁有人的 Surface Book（`C:\Users\Anthony Man\excel-qwen-grok-copilot`），已 sideload，且刻意**未**安裝 auto-start；先前此處指其未安裝的說法已被取代。
