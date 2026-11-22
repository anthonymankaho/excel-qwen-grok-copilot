# SOP — Qwen + Grok for Excel and Word

This is the runbook for the local add-in. Each person uses their own API key. The owner’s Token Plan key stays on the owner’s computer and is never shared.

The add-in name in Excel and Word is **Qwen + Grok for Excel and Word**. The pane and the relay are `http://127.0.0.1:3000`. Run every command from the folder that holds `manifest.xml`.

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

- Excel sends a multi-cell selection. A single cell sends the used range of the active sheet. Formulas are included as `value [=formula]`.
- Word sends the selection. An empty selection sends the document body.
- The context card shows the exact text. **↻ Refresh** reads it again.
- Untick **Send context** when the sheet or document must not leave the computer.
- **＋ Whole sheet** forces the active Excel sheet for the next question only.
- A truncation note means the read stopped at 150 rows by 30 columns, or at 45,000 characters. Select a smaller range.
- Enter sends. Shift+Enter adds a line. Quick actions are **Summarize**, **Explain**, **Write formula**, and **Revise**.
- The send button is **Stop** while a reply is streaming. **New chat** clears history in the pane.
- **Fast** turns Qwen thinking off and sets Grok reasoning to low. Leave it off for a formula audit.
- Only the latest question carries file context. Older turns keep their prose. A request sends at most 24 messages.

## 4. Revising the file

**Direct** is on by default.

1. Read the card before you move on. It lists every action that was written, or every action waiting for a click.
2. Safe actions write immediately: values, formulas, formatting, borders, conditional formatting, filters, freeze, rows, columns, sheets, autofit, sort, charts, tables, pivots, and in Word the selection, highlight, style, and find-and-replace.
3. An amber note means a formula can reach the network or run a legacy call (`WEBSERVICE`, `FILTERXML`, `RTD`, `HYPERLINK`, and similar). Those wait for **Apply risky changes**.
4. Turn **Direct** off when you want every action to wait for **Apply edits**.
5. An older single-cell reply still shows `sheet!cell: old → new`. `(empty)` means the cell is blank. `(sheet missing)` means that sheet is not in the workbook, and Apply writes nothing.
6. Press **Undo last apply** on the card if the batch is wrong. Cell values return to the previous value, and a cell that was empty is cleared. Inserted rows and columns are removed. Deleted rows come back. Deleted columns come back for the first 30 rows. A new sheet, chart, table, or pivot is deleted. A filter is removed. Freeze returns to the previous state. Autofit is not undone. In Word, Undo restores the selection and puts find-and-replace text back when it can still find the new wording.
7. Excel `Ctrl+Z` still works for changes Excel recorded.
8. **↳ Insert last answer into selection** writes the top-left selected cell. One line starting with `=` is a formula. Anything else is text.

## 5. Stop

- `npm run relay-stop` stops the relay and drops keys held only in relay memory. The minute task starts the relay again unless you run `npm run autostart-remove` first. A remembered key stays in the pane until **Clear key**. The sideload stays.
- `npm run stop` removes the sideload. The next open needs `Setup and Start.bat` or `npm start`.

## 6. After a code change

1. From the project folder, run `npm run check && npm test`. Both must pass. They need no API key.
2. Start the relay, load a key, and ask one real question. For a revise, confirm the card and press Undo once.
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
| Pane shows an old screen | Restart the relay and reopen the pane. On another computer, `git pull` first |
| Provider HTTP 401 | Key and URL do not match. Token Plan keys use the Token Plan URL. Regular Qwen keys use the pay-as-you-go URL |
| Qwen endpoint rejected | Use the Token Plan URL or `https://<region>.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` |
| Invalid conversation message | Message plus context is over 50,000 characters. Shorten the message or untick **Send context** |
| Context says truncated | Ask about a smaller range |
| A reply will not finish | Press **Stop**. Tick **Fast** if Grok is taking too long |
| A change is wrong | **Undo last apply**, or Excel `Ctrl+Z` |
| Word does not list the add-in | `npm run start:word` with Word closed |

## 9. What leaves this computer

The relay binds to `127.0.0.1` and rejects any other Host or Origin. A Qwen call must be HTTPS chat-completions on `qwencloudapi.com` or `aliyuncs.com`. Redirects are not followed.

A key is not written into the repository. With Remember ticked, the pane stores it for this Windows user only. **Clear key** removes that copy and the relay copy.

A question sends the prompt and the attached context to the provider you picked. Context includes cell values and formulas, or the Word selection or document text. Nothing is sent before you ask.

Keep the Token Plan key out of chat, files, and commits. Other people use their own keys.
