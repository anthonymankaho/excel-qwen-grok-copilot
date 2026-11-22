# Qwen + Grok for Excel and Word

A local task pane for desktop Excel and Word. You ask in plain language. Qwen or Grok plans the work, and the pane revises the open file through a fixed list of actions. On this PC the pane and the relay both live at `http://127.0.0.1:3000`. Excel on the web uses the HTTPS copy at `https://excel-qwen-grok.anthonym-354.workers.dev`. Company-wide deployment still needs an admin install.

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
- A column, bar, line, or pie chart, and a table with a name

In Word:

- Replace the selection
- Insert text at the start or end of the document
- Bold and italic on the selection
- The styles Heading 1, Heading 2, Heading 3, Title, and Normal

It does not run Python in Excel. Borders, conditional formatting, filters, freeze panes, pivots, and replacing every match of a phrase in Word are outside this update.

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
| Debit = credit | The first two numeric columns in the requested range. A difference beyond 0.005 fails. |
| Formula errors | `#N/A`, `#VALUE!`, `#REF!`, `#DIV/0!`, `#NAME?`, `#NULL!`, `#NUM!` |
| Blank required fields | Empty cells in the requested columns, on rows that otherwise have data |
| Duplicate keys | The same key on more than one row |
| Row count | Whether the number of rows changed. A change is a warning. |
| Total movement | Whether the sum moved beyond the tolerance. A move fails. |

When a check fails, the card keeps the difference and tells you to review it. It does not post the entry or force the totals to match.

## Direct revise

Leave **Direct** ticked, next to Fast, when Agent should write safe changes as it goes. The card lists what changed and offers **Undo last apply**.

Turn **Direct** off to hold the plan until **Approve**, and then hold every action until **Apply edits**.

These formulas are never written automatically, even with Direct on: `WEBSERVICE`, `FILTERXML`, `RTD`, `HYPERLINK`, and the same family of calls. The card marks them and waits for **Apply risky changes**. Deletes, clears, and sheet renames also wait for a click in Agent mode.

**Undo last apply** reverses that batch:

- Cell values and formulas go back to what was there. A cell that was empty is cleared.
- Inserted rows and columns are removed. Deleted rows are restored for columns A through H. A cleared range is put back.
- A new sheet, chart, or table is removed. A renamed sheet gets its previous name. A sort is put back.
- Autofit is not reversed.
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

The relay listens on `127.0.0.1` only and rejects any other Host or Origin.

## Cloudflare

The same pane is published at `https://excel-qwen-grok.anthonym-354.workers.dev` so Excel on the web can open it. The local server stays the way to run it on this PC.

The manifest to upload in Excel is `manifest.cloudflare.xml` (version 1.3.0.0). It has its own add-in id, so it can sit beside the local sideload. With Microsoft 365 Premium, add it from **Home → Add-ins → Upload My Add-in**.

The relay token is already stored as the Worker secret `RELAY_TOKEN`. In the hosted pane, open Settings, paste that token, and press **Use key**. The chat route rejects a request that does not carry it. A request from another site is rejected too. The Worker also refuses a pasted provider key and still allows only QwenCloud chat-completions URLs.

Qwen and Grok keys are Worker secrets. They are not typed into the pane and they are not stored in this repository. From the project folder, after `npx wrangler login`:

```bash
npx wrangler secret put QWEN_API_KEY
npx wrangler secret put GROK_API_KEY
```

To replace the relay token, run `npx wrangler secret put RELAY_TOKEN` with a passphrase of at least 16 characters, then paste the new value in Settings.

A later code change is published with `npm run deploy:cloudflare` after `npx wrangler login`. Workbook text sent from the hosted pane goes to Cloudflare and then to Qwen or Grok.

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
| `npm run check` | Syntax-check the relay, the Worker, the agent core, and the task pane |
| `npm test` | Run the offline pane, agent, relay, and Worker tests |
| `npm run deploy:cloudflare` | Publish the pane and relay to Cloudflare |

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
| Qwen endpoint rejected | Use the Token Plan URL, or `https://<region>.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` |
| A reply takes too long | Press **Stop**. Tick **Fast** for a shorter reply |
| Stopped. A validation failed | Read the audit card. The difference is still in the sheet. Correct it yourself, then ask again |
| A change looks wrong | Press **Undo last apply**, or Excel `Ctrl+Z` |
| The add-in is missing in Word | Run `npm run start:word` once with Word closed |
