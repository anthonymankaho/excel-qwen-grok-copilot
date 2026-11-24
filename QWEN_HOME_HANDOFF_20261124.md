# Qwen home handoff / 回家續做交接 — 2026-11-24

This file is for continuing **Qwen for Excel and Word** on another computer. Start from the GitHub `main` branch. Do not rely on the old `CODEX_HANDOFF_20261124_1646.md`; it describes an earlier read-only review at commit `3f9ce1d`.

本文件供另一部電腦上的 Qwen Code 續做 **Qwen for Excel and Word**。以 GitHub `main` 為準；舊 `CODEX_HANDOFF_20261124_1646.md` 只涵蓋較早提交 `3f9ce1d` 的唯讀審查。

## Source of truth / 版本來源

- Private repository / 私有儲存庫: `https://github.com/anthonymankaho/excel-qwen-grok-copilot`
- Branch / 分支: `main`
- Verified HEAD before this handoff / 本交接前已核對提交: `4627819` (`Record successful Excel web key retest`)
- Work computer project / 工作電腦專案: `C:\Users\User\Desktop\excel-qwen-grok-copilot`
- Hosted Excel web pane / 網頁版: `https://excel-qwen-grok.anthonym-354.workers.dev`
- Hosted build / 部署版本: `2026-11-24.10`; Cloudflare Worker version `922f0de0-193f-42f1-b79b-b36e224863ac`
- Local working tree was clean and matched `origin/main` before this handoff. GitHub has the code; it does **not** include the user's month-end finance workbook or API key. / 本交接前本機與遠端一致且無未提交修改。GitHub 有程式碼，沒有用戶的月結財務工作簿或 API key。

## Verified now / 已核實

1. `npm run check` passed; `npm test` passed **106/106** on the work computer after build `.10`. / 工作電腦上語法檢查及 106/106 測試通過。
2. Cloudflare `/api/health` returned HTTP 200, `build: 2026-11-24.10`, `hosted: true`, `byo: true`, `configured.qwen: false`. The hosted `/taskpane.js` also served build `.10`. / 雲端健康檢查與網頁程式均為 `.10`。
3. After deployment, Anthony confirmed the Qwen key remained saved after closing and reopening the **Excel web** pane and `hi` succeeded. / Anthony 實測 Excel 網頁版重新開啟仍保存 key，傳送 `hi` 成功。
4. Anthony reported local desktop Excel and Word testing complete. This is owner feedback; the specific scenarios and screenshots are not reproduced in this handoff. / Anthony 表示桌面 Excel 與 Word 測試已完成；本文件沒有重列每個案例及截圖。
5. The earlier intermittent HTTP 500 was **not** traced to a definite upstream cause. Build `.10` retries one HTTP 500/502/503/504 and reports the status if it persists. A later `hi` succeeded. / 先前間歇性 HTTP 500 未確定根因；`.10` 會重試一次並在持續失敗時顯示狀態碼，之後的 `hi` 已成功。

## At-home setup / 在家安裝

1. Sign in to GitHub with an account allowed to access the private repository, then clone `main`. Do not copy a possibly older Desktop folder as the source of truth. / 用有權限的 GitHub 帳戶複製 `main`；不要以較舊桌面資料夾為準。
2. Install Node.js 20 or newer and Microsoft 365 desktop Excel/Word if testing the local add-in. / 測試桌面版需 Node.js 20+ 及 Microsoft 365 Excel/Word。
3. In PowerShell, from the cloned project folder, run:

   ```powershell
   npm run setup
   npm run check
   npm test
   npm start
   ```

4. In Excel: **Home → Add-ins → Qwen for Excel and Word → Settings**. Enter **your own QwenCloud API key locally in the pane**. For Word, close Word first, then run `npm run start:word`. / 在 Excel 增益集 Settings 本機輸入自己的 QwenCloud API key；測試 Word 前先關閉 Word，再執行 `npm run start:word`。
5. For Excel on the web, use `manifest.cloudflare.xml` (version `1.5.0.0`) if the add-in is not already installed in that browser. The hosted pane needs no local Node relay; each user enters their own key in Settings. / Excel 網頁版如未安裝，使用 `manifest.cloudflare.xml`；託管版不需要本機 Node relay，每位使用者自行輸入 key。

No key, password, OAuth secret, personal login address, or confidential workbook data belongs in GitHub, this Markdown file, Qwen chat, logs, or screenshots. A key remembered in the web pane may be in Office storage or browser storage on that user's device; the Worker forwards the key per request and does not store it. / 不要把 key、密碼、OAuth secret、登入電郵或機密工作簿資料放進 GitHub、本文件、Qwen 對話、log 或截圖。網頁版記住的 key 可能在該用戶的 Office 或瀏覽器儲存；Worker 每次請求轉發，不會保存。

## Where to look / 主要檔案

- `public/taskpane.js`: pane UI, Office context, BYO key storage, Agent/Chat execution. / 工作窗格、Office 內容、金鑰保存、Agent/Chat。
- `public/agent-core.js`: routing, action parsing, checks, audit summary. / 路由、動作解析、檢查與審計摘要。
- `worker.mjs` and `wrangler.jsonc`: hosted Cloudflare API and deployment. / 雲端 API 與部署。
- `server.mjs`, `relay-core.mjs`: local relay and shared provider logic. / 本機 relay 與共用模型邏輯。
- `manifest.xml` (desktop) and `manifest.cloudflare.xml` (web). / 桌面與網頁 manifest。
- `README.md`, `SOP-INSTALL.md`, `SOP-USER-GUIDE.md`, `STORE_LISTING.md`, `STORE_REVIEWER_NOTES.md`: usage and Marketplace preparation. / 使用與上架文件。

## Next work / 下一步

1. On the home computer, confirm `git status`, `git log -1`, `npm run check`, and `npm test` before changing code. / 在家先確認 Git 狀態、版本及測試。
2. Test a **non-confidential sample workbook** in Excel web beyond `hi`: Chat read-only whole-workbook review, Agent's requested range reads, Share consent toggle, proposed edits, approval, and Undo. Record exact pass/fail evidence. The web `hi` test alone does not prove these flows. / 用非機密範例活頁簿測試 Excel 網頁版完整流程；`hi` 成功不等於整本覆核及修改流程已驗證。
3. If HTTP 500 recurs, record time (Hong Kong time), mode, model, HTTP status, and whether retry succeeded. Inspect Cloudflare Worker exceptions without printing request bodies or keys. / 若再出現 500，記錄香港時間、模式、模型、狀態碼及重試結果；檢查 Cloudflare 錯誤但不要輸出請求內容或 key。
4. Continue Step 6 Marketplace readiness: confirm publisher identity, Partner Center enrollment, manifest validation, listing/legal text, reviewer instructions, and clean-install QA. Do not mark the add-in Marketplace-published until an actual public listing and fresh-account install are verified. / 繼續第 6 步市集準備；實際公開上架及新帳戶安裝驗證前，不要宣稱已正式上架。
5. Keep local-only features and hosted limitations distinct: personal OneDrive connector currently belongs to the local relay; Google Drive and Gmail were removed from scope. / 分清本機與託管功能：個人 OneDrive 連接器目前只在本機 relay；Google Drive 與 Gmail 已取消。

## Paste into Qwen Code / 可貼到 Qwen Code 的續做指示

> Continue the private `anthonymankaho/excel-qwen-grok-copilot` repository from current `main`. Read `QWEN_HOME_HANDOFF_20261124.md` and the current `README.md`, then verify `git status`, HEAD, `npm run check`, and `npm test`. Build `.10` is deployed, and the owner confirmed Excel web key persistence plus a successful `hi`. Do not treat that as proof of full web workbook QA. Start with a non-confidential sample workbook and complete the web Chat/Agent/Share/approval/Undo checks, recording exact evidence. If a defect appears, fix it with focused tests and explain what remains unverified. Continue Marketplace readiness only from current files. Keep all keys, login details, and finance workbooks out of source control and chat. Reply in English and Traditional Chinese.

> 從私有儲存庫 `anthonymankaho/excel-qwen-grok-copilot` 的最新 `main` 續做。先讀本文件與最新 `README.md`，確認 Git 狀態、提交及測試。`.10` 已部署，擁有者確認 Excel 網頁版 key 可跨重新開啟保存，`hi` 成功；這並不代表整本活頁簿網頁版 QA 已完成。先用非機密範例檔完成 Chat／Agent／Share／審批／Undo 測試並記錄證據。發現缺陷才作針對性修正與測試，說明未驗證項目。金鑰、登入資料及財務工作簿不得進入 Git 或對話。請以英文及繁體中文回覆。
