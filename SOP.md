# SOP — Qwen + Grok for Excel and Word

This is the runbook for the local add-in. Each person uses their own API key. The owner’s Token Plan key stays on the owner’s computer and is never shared.

The add-in name in Excel and Word is **Qwen + Grok for Excel and Word**. The pane and the relay are `http://127.0.0.1:3000`. Run every command from the folder that holds `manifest.xml`. The manifest for this update is 1.2.0.0.

## What this update revises

The earlier pane was one prompt, one sheet dump, and one batch of actions. This update is the Copilot-style revise. Operators should treat the list below as the change to brief, and the later sections as the daily runbook.

1. The composer has **Chat**, **Plan**, and **Agent**. Chat never writes. Plan waits for **Approve**. Agent can write.
2. Plan and Agent send workbook metadata first: workbook name, sheet names, used-range size, table names when Excel exposes them, and the selection address. Cell values go only when the step asks for a range, or when the user ticks **＋ Whole sheet**.
3. The model returns a plan and a controlled action list. The pane executes that list. It does not accept arbitrary Office JavaScript.
4. After a material write, Agent reads the cells back. A check is calculated from those cells.
5. The calculated checks are debit = credit, formula errors, blank required fields, duplicate keys, row count, and total movement. Debit, credit, amount, key, and required columns come from headers or from Finance checks in Settings. A missing column is NEEDS_USER and stops the loop. A failed check also stops the loop. The card records issue, evidence, possible cause, the check, and the recommended action.
6. Journal entries are drafts. The pane does not clear a reconciliation difference and does not claim that debits equal credits unless its own check says so.
7. **Auto** uses only a provider whose key is already in the relay. It tries the other loaded provider once if the first call fails.
8. Each Chat, Plan, and Agent turn ends with an audit card. The card has no API key. **New chat** clears it. There is no export button yet. The trail is a list, so export can be added later without a new design.
9. **Direct**, **Apply edits**, **Undo last apply**, the risky-formula click, **Stop**, and the Windows relay auto-start are unchanged.
10. With Direct on, Agent writes safe actions and holds deletes, clears, renames, and network formulas for a click. With Direct off, the user approves the plan and then applies the actions.
11. The loop stops at 4 rounds, on **Stop**, when the plan needs the user, or when a check fails.

Restart the relay and reopen the pane after installing this update. On another computer, `git pull` first.

## 1. First start on a PC

1. Install Node.js 20 or newer. Close Excel and Word.
2. Double-click `Setup and Start.bat`. It runs `npm install`, `npm run autostart`, then `npm start`.
3. `npm start` sideloads the manifest into desktop Excel. For Word, run `npm run start:word` once with Word closed.
4. In the pane, open Settings, choose the provider, enter your own key, leave **Remember key on this PC** ticked, and press **Use key**.

`npm run autostart` writes two Windows items for the current user:

- Startup item `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\ExcelQwenGrokRelay.bat`
- Scheduled task `ExcelQwenGrokRelayEnsure`, every minute, which starts the relay if it is down

Auto-start needs Windows. On any other system the install script stops and tells you so.

## 2. A normal day

1. Sign in to Windows. The startup item starts the relay with no window.
2. Open Excel or Word. If the pane is not there, use **Home > Add-ins**.
3. If Settings says the relay is down, wait up to a minute, or run `npm run relay-ensure`.
4. After you change code, run `npm run relay-stop`, then `npm run relay-ensure`, and reopen the pane. `Start Local Relay.bat` is the foreground fallback.

The sideload stays until `npm run stop`. Auto-start stays until `npm run autostart-remove`. Both are per user.

## 3. Asking

- Pick **Chat**, **Plan**, or **Agent** before you send. The default in the pane is Agent.
- Chat in Excel sends a multi-cell selection. A single cell sends the used range of the active sheet. Formulas are included as `value [=formula]`.
- Plan and Agent send metadata first. They pull one range when the step needs values.
- Word Chat sends the selection. An empty selection sends the document body. Plan and Agent send a short outline until the selection is requested.
- The context card shows the exact text. **↻ Refresh** reads it again.
- Untick **Send context** when the sheet or document must not leave the computer.
- **＋ Whole sheet** forces the active Excel sheet for the next question only.
- A truncation note means the read stopped at 150 rows by 30 columns, or at 45,000 characters. Select a smaller range.
- Enter sends. Shift+Enter adds a line. Quick actions are **Summarize**, **Explain**, **Write formula**, and **Revise**.
- The send button is **Stop** while a reply is streaming. **New chat** clears history and the audit trail.
- **Fast** turns Qwen thinking off and sets Grok reasoning to low. Leave it off for a formula audit or a finance plan.
- **Auto** may be left on. It will not switch to a provider that has no key.
- Only the latest question carries file context. Older turns keep their prose. A request sends at most 24 messages.

## 4. Revising the file

1. Read the mode before you send.
   - **Chat** reads context and does not write.
   - **Plan** shows the goal, steps, assumptions, required inputs, risks, and checks. **Approve** is required before any write.
   - **Agent** may write safe actions when **Direct** is on. It reads the cells back after a write. A failed check stops the loop. The limit is 4 rounds.
2. Read the card before you move on. It lists every action that was written, or every action waiting for a click.
3. Safe actions that can write immediately: values, formulas, bold, italic, fill, font color, number format, alignment, wrap, insert rows, insert columns, add a sheet, autofit, sort, charts, tables, and in Word replace the selection, insert text, bold, italic, and a named style.
4. These wait for a click even when Direct is on: delete rows, clear a range, rename a sheet, and formulas that can reach the network (`WEBSERVICE`, `FILTERXML`, `RTD`, `HYPERLINK`, and similar). The card waits for **Apply edits** or **Apply risky changes**.
5. Turn **Direct** off when you want the plan approved and then every action held for **Apply edits**.
6. An older single-cell reply still shows `sheet!cell: old → new`. `(empty)` means the cell is blank. `(sheet missing)` means that sheet is not in the workbook, and Apply writes nothing.
7. For a finance task, read the validation card. Debit = credit uses Debit and Credit headers, or the columns set under Finance checks. The default tolerance is 0.01. If those columns are not identified, the card says NEEDS_USER and does not call the difference a pass or a fail. Formula errors, blank required fields, duplicate keys, and a moved amount fail when their columns are known. A changed row count is a warning. Treat the journal as a draft and leave the difference in the sheet.
8. Press **Undo last apply** on the card if the batch is wrong. Cell values return to the previous value, and a cell that was empty is cleared. Inserted rows and columns are removed. Deleted rows come back for columns A through H. A cleared range and a sort are restored. A new sheet, chart, or table is deleted. A renamed sheet gets its previous name. Autofit is not undone. In Word, Undo restores the selection text, bold, italic, and style, and removes inserted text when it can still find that wording.
9. Excel `Ctrl+Z` still works for changes Excel recorded.
10. **↳ Insert last answer into selection** writes the top-left selected cell. One line starting with `=` is a formula. Anything else is text.
11. Read the audit card at the end of the turn. It is the session record of the request, the provider, the context, the actions, and the checks.

Borders, conditional formatting, filters, freeze panes, pivots, and replacing every match in the Word document are not actions in this update.

## 5. Stop

- `npm run relay-stop` stops the relay and drops keys held only in relay memory. The minute task starts the relay again unless you run `npm run autostart-remove` first. A remembered key stays in the pane until **Clear key**. The sideload stays.
- `npm run stop` removes the sideload. The next open needs `Setup and Start.bat` or `npm start`.

## 6. After a code change

1. From the project folder, run `npm run check && npm test`. Both must pass. They need no API key.
2. Start the relay, load a key, and ask one real question in Chat. For a revise, use Plan or Agent, confirm the card, and press Undo once.
3. Commit on a branch. The owner pushes.
4. Restart the relay and reopen the pane. On another computer, `git pull` first.

## 7. Another computer

1. The owner grants repository access or hands over a zip. The zip must not contain keys.
2. The other person installs Node.js 20 or newer, closes Excel, runs `Setup and Start.bat`, and enters their own key with Remember ticked.
3. Pair the key with the URL. `sk-sp-…` uses the Token Plan URL, which is the default. A regular QwenCloud key uses the pay-as-you-go URL, changed in Settings. Grok uses that person’s own xAI key.
4. Each computer has its own relay, its own startup item, and its own remembered keys.

## 8. When it breaks

| What you see | What to do |
|---|---|
| Relay unavailable | Section 2, step 3 |
| Pane offline just after sign-in | Wait up to a minute, or run `npm run relay-ensure`. Confirm the Startup item and the task from section 1 |
| Port 3000 in use | `npm run relay-stop`, then start the relay again |
| Pane shows an old screen | Restart the relay and reopen the pane. On another computer, `git pull` first. This update is manifest 1.2.0.0 |
| Provider HTTP 401 | Key and URL do not match. Token Plan keys use the Token Plan URL. Regular Qwen keys use the pay-as-you-go URL |
| Qwen endpoint rejected | Use the Token Plan URL or `https://<region>.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` |
| Invalid conversation message | Message plus context is over 50,000 characters. Shorten the message or untick **Send context** |
| Context says truncated | Ask about a smaller range |
| Stopped. A validation failed | Leave the difference. Read the audit card. Do not ask the model to force the totals together |
| A reply will not finish | Press **Stop**. Tick **Fast** if the reply is taking too long |
| A change is wrong | **Undo last apply**, or Excel `Ctrl+Z` |
| Word does not list the add-in | `npm run start:word` with Word closed |

## 9. Cloudflare

Use this when the pane should open from an `https://` address, including Excel on the web. The published pane is `https://excel-qwen-grok.anthonym-354.workers.dev`.

1. In Excel, use **Home → Add-ins → Upload My Add-in** and choose `manifest.cloudflare.xml`.
2. Open Settings in the hosted pane, paste the relay token, and press **Use key**. Remember keeps it in that pane only. The matching `RELAY_TOKEN` secret is already on the Worker.
3. From the project folder, run `npx wrangler login`, then `npx wrangler secret put QWEN_API_KEY` and `npx wrangler secret put GROK_API_KEY`. Do not put those keys in the repository or in chat.
4. After a later code change, run `npm run deploy:cloudflare`. To replace the relay token, run `npx wrangler secret put RELAY_TOKEN` with a passphrase of at least 16 characters and paste the new value in Settings.

The local `127.0.0.1` relay is unchanged. The hosted route rejects a missing token, a token that does not match, a call from another site, and a pasted provider key.

## 10. What leaves this computer

The relay binds to `127.0.0.1` and rejects any other Host or Origin. A Qwen call must be HTTPS chat-completions on `qwencloudapi.com` or `aliyuncs.com`. Redirects are not followed.

A key is not written into the repository, and it is not written into the audit card. With Remember ticked, the pane stores it for this Windows user only. **Clear key** removes that copy and the relay copy.

A question sends the prompt and the attached context to the provider you picked, or to the provider Auto selected from keys already loaded. Chat context includes cell values and formulas, or the Word selection or document text. Plan and Agent start with metadata and add a range only when a step asks for it. Nothing is sent before you ask. From the hosted pane, that same question also passes through Cloudflare.

Keep the Token Plan key out of chat, files, and commits. Other people use their own keys.
