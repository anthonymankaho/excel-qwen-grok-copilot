# Marketplace preparation guide · Marketplace 上架準備指引

**Status 2026-11-25: HOLD — do not create an offer or submit yet.** This guide replaces the earlier step-by-step draft, which incorrectly said a personal Microsoft account was sufficient and asked reviewers to bring their own paid Qwen key or contact the publisher.

**截至 2026-11-25：暫停——請勿建立 offer 或提交。** 本指引取代舊版逐步草稿；舊稿錯誤表示個人 Microsoft 帳戶已足夠，並要求審查員自備付費 Qwen key 或聯絡發行者。

## Gate 1 — Decide the publisher · 確認發行者

Choose the legal publisher before creating a Partner Center offer: Anthony as an eligible individual/sole proprietor, or the company with its authorization. Confirm who can accept Microsoft's publisher agreement. A new Microsoft 365 and Copilot enrollment uses a Microsoft Entra work/school account and requires verification; a personal Microsoft 365 Premium account alone is not enough. The publisher selected on an offer cannot be changed later. The manifest currently says `ProviderName` = `Anthony Man`, which remains provisional until verified.

建立 offer 前先選擇法律發行者：由符合資格的 Anthony 個人／獨資經營者發行，或由已授權的公司發行。確認誰有權接受 Microsoft 發行者協議。新的 Microsoft 365 and Copilot 註冊需使用 Microsoft Entra 工作／學校帳戶並完成驗證；單有個人版 Microsoft 365 Premium 並不足夠。Offer 建立後不能更換發行者。目前 manifest 的 `ProviderName` 為 `Anthony Man`，在驗證前仍屬暫定。

## Gate 2 — Decide the Qwen API and key flow · 確認 Qwen API 與金鑰流程

Each colleague is expected to enter and pay for their own key; no user key is shared. However, the current default is Token Plan, and Alibaba's current policy excludes Token Plan from custom application backends. Alibaba also advises against using long-term keys in browser code. The hosted pane currently stores each user's key in Office/browser storage and forwards it through the Worker. Resolve this with an approved secure credential flow and a supported Model Studio API before continuing. See `MARKETPLACE_READINESS.md` for the recommended options and official references.

每位同事自行輸入及支付自己的 key，不會共用。惟目前預設為 Token Plan，而阿里雲現行政策不允許 Token Plan 用於自訂應用程式後端；阿里雲亦建議不要在瀏覽器程式碼使用長期 key。託管 pane 現時把各自的 key 存在 Office／瀏覽器儲存，並經 Worker 轉送。繼續前須改為經確認的安全憑證流程及受支援的 Model Studio API。建議選項與官方連結見 `MARKETPLACE_READINESS.md`。

## Gate 3 — Finish product QA · 完成產品 QA

After the API flow is settled, complete the Step 5 live matrix on the exact release candidate: Excel desktop, Excel web, Word; realistic whole-workbook review and coverage; finance checks; attachments; Share-off; approval/Undo; and provider failures. The previous web `hi` smoke test passed, but the full month-end workbook review after the latest agent changes has not been completed. Capture Office-in-context screenshots for Excel desktop, Excel web, and Word using non-sensitive sample content.

API 流程決定後，以最終候選版本完成第 5 步實機矩陣：Excel 桌面版、Excel 網頁版、Word；實際整本活頁簿覆核及覆蓋說明；財務檢查；附件；Share 關閉；批准／Undo；Provider 故障。先前網頁版 `hi` 基本測試成功，但最新 Agent 版本尚未完成完整月結活頁簿覆核。使用非敏感範例製作 Excel 桌面版、網頁版及 Word 的 Office 內截圖。

## Gate 4 — Finalize package and reviewer access · 完成套件及審查員權限

1. Update `manifest.cloudflare.xml` with the verified publisher, supported listing locales, and canonical legal URLs. The current `.html` URLs redirect to extensionless paths.
2. Run Microsoft's Office Add-in manifest/store validator and resolve all findings.
3. Rewrite `STORE_LISTING.md`, privacy/EULA/support disclosures, and `STORE_REVIEWER_NOTES.md` to match the actual credential flow, data path, provider charges, and supported features.
4. If Microsoft reviewers need an external paid service, create a restricted test route and place its working credentials in Partner Center certification notes only. Microsoft says not to rely on contacting the publisher for credentials. Never place secrets in this repo or in this Markdown guide.

1. 依已驗證發行者、支援的上架語言及正式法律頁網址更新 `manifest.cloudflare.xml`。目前 `.html` 網址會轉址至不含副檔名的正式路徑。
2. 執行 Microsoft Office Add-in manifest／商店驗證，處理所有報告。
3. 按實際憑證流程、資料路徑、供應商收費及支援功能，重寫上架文案、私隱／EULA／支援披露及 `STORE_REVIEWER_NOTES.md`。
4. 若 Microsoft 審查員需用外部付費服務，建立受限測試途徑，並只在 Partner Center certification notes 提供有效憑證。Microsoft 不接受依賴聯絡發行者補交帳密。絕不可將密碼放入此 repo 或 Markdown 指引。

## Gate 5 — Partner Center and submission · Partner Center 與提交

After the preceding gates are closed, complete verified publisher enrollment in the Microsoft 365 and Copilot program, create the offer under the correct publisher, upload the validated manifest, complete properties/listings/pricing, and provide certification instructions. Submit for certification only after reviewing the complete draft. Following approval, publish and verify a fresh-account install from Office; record the live listing URL and version.

以上前置條件完成後，完成 Microsoft 365 and Copilot 計劃的發行者驗證，以正確發行者建立 offer、上傳已驗證 manifest、完成屬性／文案／定價及認證指引。檢查完整草稿後才提交審核。核准後發布，並以新帳戶從 Office 實際安裝；記錄正式上架網址及版本。

### Official references · 官方參考

- [Partner Center enrollment](https://learn.microsoft.com/partner-center/create-account)
- [Microsoft Marketplace submission guide](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/add-in-submission-guide)
- [Microsoft 365 app publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist)
- [Alibaba Token Plan permissions and usage policies](https://help.aliyun.com/en/model-studio/token-plan-permissions-and-usage)
- [Alibaba Model Studio API key guidance](https://help.aliyun.com/en/model-studio/get-api-key)
