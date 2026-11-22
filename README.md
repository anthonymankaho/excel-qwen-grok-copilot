# Qwen + Grok for Excel and Word

A local task pane for desktop Excel and Word. It reads the selection, chats with Qwen or Grok, and revises the open file. **Direct** is on by default: safe changes are written immediately, and the card has **Undo last apply**. Formulas that can reach the network stay on the card until you click.

The pane and the relay both live on this computer at `http://127.0.0.1:3000`. Excel for the web, Word for the web, and company-wide deployment need an HTTPS host and an admin install. This copy does not.

## What it can change

In Excel:

- Values and formulas on a cell or a range
- Bold, italic, fill, font color, number format, alignment, wrap, and borders
- Conditional formatting: a value rule, a color scale, duplicates, unique values, blanks, or above average
- An AutoFilter, including keeping only chosen values
- Frozen rows and columns
- Insert or delete rows and columns, add or rename a sheet, autofit, and sort
- A column, bar, line, or pie chart, a table, and a pivot built from header names in the source range

In Word:

- Replace the selection, or replace every match of a phrase
- Insert text at the start or end of the document
- Bold, italic, highlight, and the styles Heading 1, Heading 2, Heading 3, Title, and Normal

It does not run Python in Excel. A pivot only uses the header names you can see in the source range.

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

- In Excel, a multi-cell selection is sent. A single cell falls back to the used range of the active sheet. Formulas appear as `value [=formula]`. The read stops at the first 150 rows and 30 columns, then at 45,000 characters. A truncation note says so.
- In Word, the selection is sent. An empty selection sends the document body, with the same character cap.
- The card refreshes when the selection changes. **↻ Refresh** reads it again.
- Untick **Send context** to ask with no file data.
- **＋ Whole sheet** forces the active Excel sheet for the next question only.

Ask in the box and press Enter. Shift+Enter adds a line. **Summarize**, **Explain**, **Write formula**, and **Revise** fill a starting prompt. Replies stream. The send button becomes **Stop** while a reply is running. **New chat** clears the conversation. **Fast** skips deep thinking: Qwen turns thinking off, and Grok uses low reasoning effort.

The footer pill switches Qwen and Grok. Each provider needs its own key. The dot is green when that key is loaded, amber when the relay is up and the key is not, and red when the relay is down.

**↳ Insert last answer into selection** writes the last reply into the top-left cell of the Excel selection. One line that starts with `=` is written as a formula. Anything else is written as text.

History stays in memory. Reloading the pane clears it. Each request sends at most 24 messages, and only the latest question includes the file context.

## Direct revise

Leave **Direct** ticked, next to Fast, when you want the file changed. Ask for the change in plain language. The reply explains it and returns a list of actions. Safe actions are written at once. The card lists what changed and offers **Undo last apply**.

Turn **Direct** off to hold every action until you press **Apply edits**.

These formulas are never written automatically, even with Direct on: `WEBSERVICE`, `FILTERXML`, `RTD`, `HYPERLINK`, and the same family of calls. The card marks them and waits for **Apply risky changes**.

**Undo last apply** reverses that batch:

- Cell values and formulas go back to what was there. A cell that was empty is cleared.
- Inserted rows and columns are removed. Deleted rows are restored. Deleted columns are restored for the first 30 rows.
- A new sheet, chart, table, or pivot is removed. A renamed sheet gets its previous name.
- A filter is removed. Freeze goes back to the previous freeze, or unfreezes.
- Autofit is not reversed.
- In Word, the selection is put back, and a find-and-replace puts the previous wording back when the new text can be found.

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

The relay listens on `127.0.0.1` only and rejects any other Host or Origin.

## Another computer

Each computer runs its own relay and stores its own keys.

1. Get the project from the private repository, or from a zip that does not contain keys.
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
| `npm run check` | Syntax-check `server.mjs` and `public/taskpane.js` |
| `npm test` | Run the offline pane and relay tests |

Auto-start is Windows-only. It uses `%APPDATA%`.

## Tests

`npm run check` and `npm test` need no API key. The pane tests stub the page, Excel, Word, and the stream. They cover context, the click-gated edit card, Direct revise, Undo, risky formulas, and Word replace. The relay tests cover the loopback lock, the Qwen URL allowlist, Fast mode, and stream errors.

Before trusting a change in the real app, start the relay, load a key, ask one question, and confirm the answer. For a revise, confirm the card and use Undo once.

## If something is wrong

| What you see | What to do |
|---|---|
| Relay unavailable | Run `npm run relay-ensure`, or wait up to a minute for the self-heal task |
| Port 3000 is in use | Another relay is running. `npm run relay-stop`, then start it again |
| The pane looks old | Restart the relay, then close and reopen the pane. On another computer, `git pull` first |
| Context says truncated | Select a smaller range, or ask about one section at a time |
| Invalid conversation message | The message plus context is over 50,000 characters. Shorten it, or untick **Send context** |
| Provider HTTP 401 | The key and the URL do not match. `sk-sp-…` needs the Token Plan URL. A regular Qwen key needs the pay-as-you-go URL |
| Qwen endpoint rejected | Use the Token Plan URL, or `https://<region>.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` |
| A reply takes too long | Press **Stop**. Tick **Fast** for a shorter Grok reply |
| A change looks wrong | Press **Undo last apply**, or Excel `Ctrl+Z` |
| The add-in is missing in Word | Run `npm run start:word` once with Word closed |
