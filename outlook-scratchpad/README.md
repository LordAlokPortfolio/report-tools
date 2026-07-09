# Boss Replies Scratchpad (Outlook Add-in scaffold)

A task-pane add-in for Outlook that lets you capture the email you're reading
into a persistent "Boss Replies" list, with notes, a checklist, a priority,
a due date, a status, and a draft-reply field.

## Current scope

This is a **scaffold**, not the finished cross-device product described in
the original spec:

- **Persistence is local-only.** Tasks are stored in the browser's
  `localStorage`, scoped to one device/profile. They do **not** sync to
  Microsoft To Do, Outlook Tasks, or other devices yet.
- **Microsoft Graph / Microsoft To Do sync is stubbed.** See `graph.js` —
  every method is a documented TODO with the exact Graph endpoint and the
  Azure AD app registration steps needed to make it real. Wiring this up
  requires an Azure AD (Microsoft Entra) app registration with a client ID,
  which needs to be created by you or your IT admin — an AI agent can't
  provision that on your behalf.
- **No native drag-and-drop.** Office.js (the add-in API) doesn't expose a
  drop target for dragging a message out of the message list into a custom
  task pane — that's a built-in Outlook feature (My Day / To Do pane), not
  something third-party add-ins can hook into. Instead, this add-in adds a
  **"Capture this email" button** that appears when you have a message open:
  click it while reading an email and it creates a task with the subject,
  sender, and a link back to the source message (the link itself needs the
  Graph wiring above to be a durable, cross-device URL — see the TODO in
  `graph.js#getMessageWebLink`).

## Files

| File | Purpose |
| --- | --- |
| `manifest.xml` | Outlook add-in manifest (classic XML manifest). Registers a "Boss Replies" button on the message-read ribbon that opens the task pane. |
| `taskpane.html` / `taskpane.css` / `taskpane.js` | The task pane UI: capture button, filters, task cards with note/checklist/priority/due date/status. |
| `graph.js` | Stubbed Microsoft Graph client. Every method is a TODO — fill in `GRAPH_CONFIG.clientId` and implement sign-in/Graph calls once you have an Azure AD app registration. |
| `commands.js` / `commands.html` | Empty function file required by the manifest's command surface (no `ExecuteFunction` commands are registered today). |
| `assets/icon-*.png` | Placeholder ribbon icons (solid color). Swap for real branded icons before shipping. |

## Try it locally (sideload into Outlook on the web)

Office Add-ins must be served over HTTPS, and the manifest above points at
`https://localhost:3000`. To test:

1. Install a trusted local dev certificate (one-time):
   ```
   npx office-addin-dev-certs install
   ```
2. From this folder, serve the files over HTTPS on port 3000 with any static
   server that can use that certificate, e.g.:
   ```
   npx http-server . -p 3000 --ssl \
     --cert "$(npx office-addin-dev-certs verify-store 2>/dev/null || echo ~/.office-addin-dev-certs/localhost.crt)" \
     --key ~/.office-addin-dev-certs/localhost.key
   ```
   (Exact cert paths vary by OS — `office-addin-dev-certs` prints them after
   `install`. Any HTTPS static server pointed at this folder works.)
3. Go to Outlook on the web → **Settings → Get Add-ins → My add-ins → Add a
   custom add-in → Add from file**, and upload `manifest.xml`.
4. Open any email. A **Boss Replies** button appears on the read ribbon —
   click it to open the task pane, then click **Capture this email**.

## Next steps toward the full product spec

1. Register the Azure AD app and fill in `GRAPH_CONFIG` in `graph.js`.
2. Add the MSAL.js browser SDK and implement `signIn()` / `getAccessToken()`.
3. Implement `createTask` / `updateTask` / `deleteTask` / `getMessageWebLink`
   against Microsoft Graph, and call them from `taskpane.js` alongside the
   existing localStorage writes (local state stays as an optimistic cache).
4. Replace the placeholder icons in `assets/` with branded artwork.
5. Consider moving to the unified JSON manifest + Teams Toolkit if you also
   want this to run as a Microsoft 365 / Teams app, not just Outlook.
