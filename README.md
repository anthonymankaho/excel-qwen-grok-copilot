# Qwen + Grok for Excel

A local Microsoft Excel task pane inspired by Copilot. It can read the current cell selection, attach selected values to a chat, summarize or explain spreadsheet content, help write formulas, and insert the latest answer into the selected cell. Choose Qwen or Grok as the provider.

## Run locally

1. Install Node.js 20 or newer.
2. Install project dependencies once by opening PowerShell in this folder and running `npm install`.
3. Close Excel, then double-click `Start Excel Qwen Grok Copilot.bat` (or run `npm start`). The command starts the local relay and sideloads the XML manifest into desktop Excel.
4. Keep the terminal window open while using the add-in. If its pane does not open automatically, select it from **Home > Add-ins**. To restart only the local relay after a code change, use `Start Local Relay.bat`.
5. Open Settings, choose a provider, enter its API key, and press **Use key**.
6. When finished, press `Ctrl+C` in the terminal to stop the session and clear the in-memory keys. You can also run `npm run stop` to unregister the development sideload.

## Continue on another computer

1. Clone the private GitHub repository to the computer.
2. Install Node.js 20 or newer and run `npm install` in the project folder.
3. Close Excel and run `npm start` (or double-click `Start Excel Qwen Grok Copilot.bat`).
4. Enter the provider API key again in the add-in Settings. Keys are intentionally not synced or committed.

The task pane and API relay are served from `http://127.0.0.1:3000`. This is a local development add-in for desktop Excel. Excel for the web and organizational deployment require an HTTPS-hosted manifest/app and suitable admin sideloading or deployment.

## Providers

- **Qwen** uses the owner's QwenCloud International Token Plan through the OpenAI-compatible endpoint `https://token-plan.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` with a Token Plan key (`sk-sp-…`). Default model: `qwen4-max` (verified 2026-11-20; `qwen4-max-preview` and `qwen3.8-max` also work on this endpoint). For pay-as-you-go instead, change the URL in Settings to `https://ap-southeast-1.maas.qwencloudapi.com/compatible-mode/v1/chat/completions` and use a regular Singapore QwenCloud API key — Token Plan keys do not work on that endpoint, and regular keys do not work on the Token Plan endpoint.
- **Grok** uses xAI's Responses API at `https://api.x.ai/v1/responses`. Default model: `grok-4.7`.

## Key handling

API keys are posted to the loopback relay and kept only in process memory. They are not saved to a file or browser storage. Stopping the Node process clears them. Workbook values are sent to the selected AI provider only after you choose **Add selection to chat**; review sensitive data before sharing. The relay binds only to `127.0.0.1`. The Qwen section is configured for the owner's QwenCloud International Token Plan key and was verified end-to-end through this relay on 2026-11-20. Token Plan is marketed for interactive coding/agent tools; this personal loopback add-in uses it at the owner's explicit direction. Keep the key private — never commit it or paste it into shared files.

## Current scope

This is a starter app, not Microsoft's Copilot product. It includes chat, selected-range context, quick prompts, and writing the last answer into the active selection. It does not execute arbitrary model-generated Excel commands or edit cells without your click.
