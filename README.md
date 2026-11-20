# Qwen + Grok for Excel

A local Microsoft Excel task pane inspired by Copilot and Claude for Microsoft 365. It automatically reads your selection (or the active sheet) as context with formulas, chats about the workbook, summarizes or explains spreadsheet content, helps write formulas, proposes cell edits that are applied only with your click, and can insert the latest answer into the selected cell. Choose Qwen or Grok as the provider.

## Run locally

1. Install Node.js 20 or newer.
2. First time on a computer: close Excel and double-click `Setup and Start.bat` — it installs dependencies, installs the logon auto-start, sideloads the manifest into desktop Excel and starts the relay. (Manual equivalent: `npm install`, `npm run autostart`, then `npm start`.)
3. Open Settings, choose a provider, enter **your own** API key once and press **Use key** with **Remember key on this PC** ticked — like a Claude sign-in, it persists for this Windows user and is re-loaded into the relay automatically on every pane load.
4. Every later logon the relay starts automatically, **hidden (no window)**. A per-minute self-heal scheduled task (`ExcelQwenGrokRelayEnsure`, current user) restarts it within a minute if it ever dies. Open Excel and the pane is ready (if it does not open: **Home > Add-ins**). No bat double-clicks needed.
5. Stop the relay with `npm run relay-stop` (the self-heal task brings it back within a minute; for a permanent stop run `npm run autostart-remove` first). Manual foreground fallback / restart after a code change: `npm run relay-stop` then `npm run relay-ensure`, or `Start Local Relay.bat`. Relay memory clears on stop (a remembered key copy stays in pane storage until **Clear key**).
6. `npm run stop` unregisters the sideload; `npm run autostart-remove` removes the logon auto-start.

## Using the pane

1. **Workbook context (automatic).** The `WORKBOOK CONTEXT · AUTO` card shows what the next question will send: your multi-cell selection, or the active sheet's used range when the selection is a single cell. Formulas are included as `value [=formula]`. **↻ Refresh** re-reads it; untick **Send context** to ask without workbook data; **＋ Whole sheet** forces the full active sheet for the next question only.
2. **Ask.** Type in the composer and press Enter (Shift+Enter = new line), or use the quick actions **Summarize**, **Explain**, **Write formula** (finish the prompted sentence, then send). Replies stream in token by token; the bubble shows **Thinking…** until the first text arrives (measured ~1.6–1.9 s for small questions; longer while a large sheet is analyzed). Tick **Fast** next to the provider pill to skip deep thinking (Qwen only) for quicker, shallower replies.
3. **Answers.** Replies appear in the conversation with a `via <model>` tag. **↳ Insert last answer into selection** writes the answer text into the top-left cell of the current selection.
4. **Cell edits (approval-gated).** When you ask for changes, the model replies with a **PROPOSED EDITS** card listing `sheet!cell ← value` rows. Review them, then click **Apply edits** to write (formulas as formulas, other values as values) or **Dismiss** to discard. Nothing is written without your click; Excel's own Undo still works afterwards.
5. **Provider.** The pill at the composer footer switches Qwen/Grok; each provider needs its own key loaded in Settings.
6. **History.** Conversation lives in memory only — reloading the pane clears it. At most 24 messages are sent per request.

## Other computers and colleagues (sharing)

1. Get the code: the owner adds you as a collaborator on the private repository, or hands you a zip of the project. Keys are never in the repo or the zip.
2. Install Node.js 20+, close Excel, double-click `Setup and Start.bat`.
3. Enter **your own** key once (Remember ticked). Key/endpoint pairing: a Token Plan key (`sk-sp-…`) works only with the default Token Plan URL; a regular QwenCloud key works only after switching the URL in Settings to `https://ap-southeast-1.maas.qwencloudapi.com/compatible-mode/v1/chat/completions`; Grok needs your own xAI key.
4. Every computer runs its own loopback relay and stores its own keys; nothing is shared between users. Never use someone else's personal Token Plan key — the plan's terms and quota are per person.

The task pane and API relay are served from `http://127.0.0.1:3000`. This is a local development add-in for desktop Excel. Excel for the web and organizational deployment require an HTTPS-hosted manifest/app and suitable admin sideloading or deployment.

## Providers

- **Qwen** uses the owner's QwenCloud International Token Plan through the OpenAI-compatible endpoint `https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` with a Token Plan key (`sk-sp-…`). Default model: `qwen4-max` (verified 2026-11-20; `qwen4-max-preview` and `qwen3.8-max` also work on this endpoint). For pay-as-you-go instead, change the URL in Settings to `https://ap-southeast-1.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` and use a regular Singapore QwenCloud API key — Token Plan keys do not work on that endpoint, and regular keys do not work on the Token Plan endpoint.
- **Grok** uses xAI's Responses API at `https://api.x.ai/v1/responses`. Default model: `grok-4.7`.

## Key handling

API keys are posted to the loopback relay and held in process memory; stopping the Node process clears them. With **Remember key on this PC** ticked, the pane also persists the key in its local browser storage (this PC, this Windows user only) and re-loads it into the relay on every pane load; untick for memory-only sessions, and **Clear key** removes both copies. Keys are never written to the repo or any shared file. Like Claude for Microsoft 365, the pane automatically sends workbook context with each question: your multi-cell selection when you have one, otherwise the active sheet's used range, with formulas as `value [=formula]`. Untick **Send context** to ask without workbook data; **＋ Whole sheet** forces the full active sheet for the next question. Context is capped at 45,000 characters per message with an explicit truncation note; review sensitive data before sharing. Only the latest question carries the workbook context — earlier turns keep their prose only, so input tokens do not multiply with conversation length. Cell changes are proposed as an edit list and written only when you click **Apply edits**. The relay binds only to `127.0.0.1`. The Qwen section is configured for the owner's QwenCloud International Token Plan key and was verified end-to-end through this relay on 2026-11-20. Token Plan is marketed for interactive coding/agent tools; this personal loopback add-in uses it at the owner's explicit direction. Keep the key private — never commit it or paste it into shared files.

## Tests

- `npm run check` — syntax check of `server.mjs` and `public/taskpane.js`.
- `npm test` — offline suite `tests/taskpane.test.mjs` (`node --test`, zero dependencies, stub DOM + Office.js + SSE): auto-context rules, Send-context opt-out, Whole-sheet override, truncation cap, apply-only-on-click, Dismiss, invalid edit payloads, oversize-prompt guard, remembered-key auto-load and persistence, history context dedup, Fast flag. 13 tests.
- Live check (manual): start the relay, load a key in Settings, ask one question and confirm a real answer; for edit proposals confirm the **PROPOSED EDITS** card appears and nothing is written before **Apply edits**.

## Troubleshooting

- **Relay unavailable** in Settings → run `npm run relay-ensure`, or wait up to a minute for the self-heal task.
- **Port 3000 already in use** → another relay instance is running; stop it (`Ctrl+C` in its terminal) before starting a new one.
- **Pane shows old UI after a code change** → restart the relay, then close and reopen the task pane (**Home > Add-ins**); on another computer, `git pull` first.
- **Context hint says truncated** → the sheet exceeds 45,000 characters; select a smaller range or ask section by section for full detail.
- **Invalid conversation message.** → message plus context exceeds 50,000 characters; shorten the message or untick **Send context**.
- **Provider HTTP 401** → key/endpoint mismatch: Token Plan keys (`sk-sp-…`) only work on the Token Plan URL; regular keys only on the pay-as-you-go URL.
- **Excel for the web** → not supported; desktop Excel only.

## Current scope

This is a starter app, not Microsoft's Copilot product. It includes chat, automatic workbook context (selection or active sheet, formulas included), quick prompts, approval-gated cell edits (**Apply edits**), and writing the last answer into the active selection. It does not execute arbitrary model-generated Excel commands or edit cells without your click.

## Progress

- **2026-11-20** — Initial app by Codex (`299e0e4`).
- **2026-11-20** — Qwen switched to the owner's QwenCloud International Token Plan endpoint with default model `qwen4-max`; live-verified end-to-end through the relay (`c98ac6c`, pushed by the owner). Codex's earlier "Token Plan cannot serve this add-in" claim disproved by live probe.
- **2026-11-20** — Claude-for-Microsoft-365-style behavior: automatic workbook context per question and approval-gated cell edits; shared `SYSTEM_PROMPT` for both providers; offline test suite added, 9/9 pass; live edit-proposal verified with `qwen4-max` (`f1427f7`).
- **2026-11-20** — README usage/tests/troubleshooting sections and `SOP.md` runbook added.
- **2026-11-20** — Consolidated to one local folder: sideload re-pointed to `Excel Qwen Grok Copilot`; the redundant clone `excel-qwen-grok-copilot` (identical commit) awaits retirement once its file lock is released.
- **2026-11-20** — Zero-touch redesign (Claude-style): remembered keys with one-time entry, per-user logon auto-start (`Setup and Start.bat`, `npm run autostart`), minimized relay console; sharing model documented — colleagues on their own PCs with their own keys.
- **2026-11-20** — Relay made invisible and self-healing: hidden process (no console to close accidentally) + per-minute `ExcelQwenGrokRelayEnsure` task; kill-and-respawn verified live; `npm run relay-stop` / `relay-ensure` added.
- **2026-11-20** — Latency review + fixes: SSE streaming through the relay (`/api/chat-stream`, normalized deltas for both providers, Thinking… indicator), history context dedup, **Fast** toggle (`enable_thinking:false`, gateway-verified). Measured: first text ~1.6–1.9 s small / ~6.4 s with 10k-char context (thinking), then incremental.
- **Pending** — Grok section verification (xAI endpoint, `grok-4.7`) handed to Grok via `Desktop\Qwen Code\GROK_HANDOFF_20261120_0833.md`.
