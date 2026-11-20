# SOP — Qwen + Grok for Excel (local add-in)

Owner/operator: Anthony. Each user enters **their own** key once; the owner's personal Token Plan key is never shared. Agents: Qwen Code owns the Qwen provider + shared context/edit layer; Grok owns the Grok section (see `Desktop\Qwen Code\GROK_HANDOFF_20261120_0833.md`); Codex authored the initial app.

## 0. Folder on this PC (consolidated 2026-11-20)

Single active copy: `C:\Users\User\Desktop\Excel Qwen Grok Copilot` — both the working repo and the sideloaded runtime (registry `HKCU\Software\Microsoft\Office\16.0\Wef\Developer\{c42d1d4e-bce1-4efe-9f70-66f64889f4cc}` → its `manifest.xml`, re-pointed 2026-11-20). Run every bat / npm command here.

The former second clone `Desktop\excel-qwen-grok-copilot` (identical commit `18c9838`) is redundant and nothing references it any more. Its retirement move was blocked by an open file handle on 2026-11-20: close whatever holds it (Explorer window, terminal, or another agent session with that folder open), then move or delete it.

## 1. Daily start (zero-touch since 2026-11-20)

1. First time on a PC: close Excel → double-click `Setup and Start.bat` (npm install + logon auto-start + sideload + relay). Manual equivalent: `npm install`, `npm run autostart`, `npm start`.
2. Settings (⚙) → provider → enter **your own** key once → **Use key** with **Remember key on this PC** ticked. The pane re-loads it into the relay on every load (Claude-style sign-in once).
3. Every later logon: the Startup item `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\ExcelQwenGrokRelay.bat` starts the relay **hidden (no window)**. The per-minute task `ExcelQwenGrokRelayEnsure` restarts it within a minute if it dies (kill-and-respawn verified 2026-11-20). Open Excel → pane ready (else **Home > Add-ins > "Qwen + Grok for Excel"**). No clicks needed.
4. Manual fallback / restart after a code change: `npm run relay-stop` then `npm run relay-ensure` (or wait ≤1 min for self-heal), or foreground `Start Local Relay.bat`.
5. The sideload persists until `npm run stop`; the auto-start persists until `npm run autostart-remove`. Both are per-user.

## 2. Asking questions

- Context is automatic: multi-cell selection, else active sheet used range; formulas as `value [=formula]`. The `WORKBOOK CONTEXT · AUTO` card shows exactly what will be sent.
- **Send context** unticked = no workbook data leaves the machine (use for sensitive sheets).
- **＋ Whole sheet** = force the full active sheet for the next question only.
- Truncation note in the hint = sheet over 45,000 chars; select a smaller range for full detail.
- Enter sends; Shift+Enter new line; quick actions **Summarize / Explain / Write formula**.
- Replies stream (Thinking… until first text). **Fast** tick = skip deep thinking (Qwen only) for quicker, shallower answers; leave unticked for formula audits.
- Only the latest question carries workbook context; older turns keep prose only (token/latency control).

## 3. Approving cell edits

1. Read every row of the **PROPOSED EDITS** card (`sheet!cell ← value`) against the workbook.
2. Click **Apply edits** only when correct; **Dismiss** otherwise. Nothing is written without the click.
3. After Apply, spot-check the cells in the sheet; Excel Undo (`Ctrl+Z`) reverses a bad apply.
4. **↳ Insert last answer into selection** writes prose into the top-left selected cell — also click-gated.

## 4. Stop

- `npm run relay-stop`: kills the relay (keys clear from relay memory). The self-heal task restarts it within a minute — for a **permanent** stop run `npm run autostart-remove` first (removes Startup item + task). Remembered pane copies persist until **Clear key**. Sideload stays registered.
- `npm run stop`: also unregisters the sideload (next start needs `Setup and Start.bat` / `npm start`).

## 5. After any code change (agent or owner)

1. In the working repo: `npm run check && npm test` (11/11 required).
2. One live question through the relay before relying on the change; for edit features confirm the proposal card and click-gating.
3. Commit locally. **Only Anthony pushes** (`git push origin main`).
4. On this PC the working repo is the runtime copy: close the minimized relay window, run `Start Local Relay.bat`, reopen the pane. On other computers: `git pull` first, then restart relay and reopen pane.

## 6. Other computers / colleagues (sharing model)

1. Owner grants repo collaborator access or hands over a zip (never containing keys). Outbound sharing is the owner's action.
2. Colleague: Node.js 20+, close Excel, `Setup and Start.bat`, enter **their own** key once (Remember ticked).
3. Key/endpoint pairing: `sk-sp-…` ↔ Token Plan URL (default); regular QwenCloud `sk-` ↔ PAYG URL (change in Settings); Grok ↔ own xAI key.
4. Each PC: own loopback relay, own Startup item, own remembered keys. Never share the owner's personal Token Plan key (plan ToS + quota are per person).

## 7. Incident quick table

| Symptom | Action |
|---|---|
| Settings shows "Relay unavailable" | Start the relay (section 1.4) |
| Pane offline right after logon | Wait ≤1 min (self-heal) or `npm run relay-ensure`; check Startup item + task exist (section 1.3) |
| Port 3000 in use | Stop the other relay terminal first |
| Pane shows old UI | Restart relay → reopen pane (other computers: `git pull` first) |
| Provider 401 | Key/endpoint mismatch: `sk-sp-…` ↔ Token Plan URL; regular `sk-` ↔ PAYG URL |
| "Invalid conversation message." | Message + context > 50,000 chars; shorten or untick Send context |
| Hint says truncated | Ask per range/section instead of whole sheet |
| Edits applied wrongly | Excel `Ctrl+Z`; Dismiss future proposals until reviewed |

## 8. Privacy boundaries

- Relay binds `127.0.0.1` only; keys never written to the repo or shared files. With Remember ticked, a key persists in the pane's local browser storage (this PC, this Windows user only); untick for memory-only; **Clear key** removes both copies.
- Leaves the machine: your prompts + the attached context (selection/used range, formulas included) to the chosen provider — nothing else, and nothing before a question is sent.
- Token Plan ToS caveat (personal plan marketed for interactive coding/agent tools) is accepted by the owner for this personal loopback tool; keep the key out of chat, files and commits. Colleagues use their own keys, never the owner's.
