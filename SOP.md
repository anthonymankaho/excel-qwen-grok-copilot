# SOP — Qwen + Grok for Excel (local add-in)

Owner/operator: Anthony. Keys are entered by the owner only. Agents: Qwen Code owns the Qwen provider + shared context/edit layer; Grok owns the Grok section (see `Desktop\Qwen Code\GROK_HANDOFF_20261120_0833.md`); Codex authored the initial app.

## 0. Folder on this PC (consolidated 2026-11-20)

Single active copy: `C:\Users\User\Desktop\Excel Qwen Grok Copilot` — both the working repo and the sideloaded runtime (registry `HKCU\Software\Microsoft\Office\16.0\Wef\Developer\{c42d1d4e-bce1-4efe-9f70-66f64889f4cc}` → its `manifest.xml`, re-pointed 2026-11-20). Run every bat / npm command here.

The former second clone `Desktop\excel-qwen-grok-copilot` (identical commit `18c9838`) is redundant and nothing references it any more. Its retirement move was blocked by an open file handle on 2026-11-20: close whatever holds it (Explorer window, terminal, or another agent session with that folder open), then move or delete it.

## 1. Daily start

1. The Excel sideload persists across restarts (verified 2026-11-20 in the registry). The full `Start Excel Qwen Grok Copilot.bat` / `npm start` is needed only **first time on a PC** or after `npm run stop`.
2. Every session the **relay must be running** (pane, API and keys live in that process): double-click `Start Local Relay.bat` in `Excel Qwen Grok Copilot` (or `npm run dev-server`). Keep the terminal open.
3. Open Excel; the pane opens automatically, else **Home > Add-ins > "Qwen + Grok for Excel"**.
4. Settings (⚙) → provider → paste the key → **Use key**. Keys live in relay memory only and are re-entered after every relay start — by design.
5. Optional auto-start at logon (Startup-folder shortcut to `Start Local Relay.bat`) is a persistence change: implement only on the owner's explicit yes.

## 2. Asking questions

- Context is automatic: multi-cell selection, else active sheet used range; formulas as `value [=formula]`. The `WORKBOOK CONTEXT · AUTO` card shows exactly what will be sent.
- **Send context** unticked = no workbook data leaves the machine (use for sensitive sheets).
- **＋ Whole sheet** = force the full active sheet for the next question only.
- Truncation note in the hint = sheet over 45,000 chars; select a smaller range for full detail.
- Enter sends; Shift+Enter new line; quick actions **Summarize / Explain / Write formula**.

## 3. Approving cell edits

1. Read every row of the **PROPOSED EDITS** card (`sheet!cell ← value`) against the workbook.
2. Click **Apply edits** only when correct; **Dismiss** otherwise. Nothing is written without the click.
3. After Apply, spot-check the cells in the sheet; Excel Undo (`Ctrl+Z`) reverses a bad apply.
4. **↳ Insert last answer into selection** writes prose into the top-left selected cell — also click-gated.

## 4. Stop

- `Ctrl+C` in the relay terminal: stops relay, clears keys from memory. Sideload stays registered.
- `npm run stop`: also unregisters the sideload (next start needs the full bat / `npm start`).

## 5. After any code change (agent or owner)

1. In the working repo: `npm run check && npm test` (9/9 required).
2. One live question through the relay before relying on the change; for edit features confirm the proposal card and click-gating.
3. Commit locally. **Only Anthony pushes** (`git push origin main`).
4. On this PC the working repo is the runtime copy: restart the relay (`Ctrl+C` then `Start Local Relay.bat`) and reopen the pane. On other computers: `git pull` first, then restart relay and reopen pane.

## 6. Second computer (home)

Clone the private repo `anthonymankaho/excel-qwen-grok-copilot`, Node.js 20+, `npm install`, close Excel, `npm start` (first time sideloads), enter keys in Settings. Later sessions: `Start Local Relay.bat` only.

## 7. Incident quick table

| Symptom | Action |
|---|---|
| Settings shows "Relay unavailable" | Start the relay (section 1.2) |
| Port 3000 in use | Stop the other relay terminal first |
| Pane shows old UI | Restart relay → reopen pane (other computers: `git pull` first) |
| Provider 401 | Key/endpoint mismatch: `sk-sp-…` ↔ Token Plan URL; regular `sk-` ↔ PAYG URL |
| "Invalid conversation message." | Message + context > 50,000 chars; shorten or untick Send context |
| Hint says truncated | Ask per range/section instead of whole sheet |
| Edits applied wrongly | Excel `Ctrl+Z`; Dismiss future proposals until reviewed |

## 8. Privacy boundaries

- Relay binds `127.0.0.1` only; keys never written to disk, repo, or browser storage.
- Leaves the machine: your prompts + the attached context (selection/used range, formulas included) to the chosen provider — nothing else, and nothing before a question is sent.
- Token Plan ToS caveat (personal plan marketed for interactive coding/agent tools) is accepted by the owner for this personal loopback tool; keep the key out of chat, files and commits.
