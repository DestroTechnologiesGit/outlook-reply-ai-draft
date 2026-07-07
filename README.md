# Outlook AI Draft Reply Add-in

Adds a "Draft with AI" button to Outlook's reading pane. Clicking it sends the
open email's subject/body to your own backend, which retrieves relevant
**Purlfrost knowledge-base context (RAG)** and calls the OpenAI API, then opens
Outlook's reply window pre-filled with the AI draft. **Nothing is sent
automatically** — the user reviews and hits Send themselves.

## How RAG works here

The backend does not answer from the model's memory alone. For each email it:

1. Sends the email's subject + body to the **n8n "KB Search API" webhook**
   (`/webhook/kb-search`), authenticated with a header token.
2. That webhook embeds the text and searches the same **`purlfrost-rag` Qdrant
   collection** used by the Purlfrost chatbot and n8n draft workflow, returning
   the top matching KB chunks (products, prices, policies, delivery, returns).
3. The backend injects those chunks into the draft prompt so the reply is
   grounded in real Purlfrost content.

If the KB is unset or unreachable, the backend **falls back to a plain
(ungrounded) draft** — replies never break.

### One-time n8n setup (the KB Search webhook)

1. In n8n: **Workflows → Import from File** → select
   `n8n/KB_Search_API.workflow.json` from this repo.
2. Create a **Header Auth** credential named `KB Search Header Auth`:
   Name = `X-KB-Token`, Value = a long random token (keep it secret).
   Select it on the **Webhook** node.
3. Confirm the **KB Search** node uses the `Qdrant account` credential and
   **KB Embeddings** uses the `OpenAI account` credential (they auto-map if the
   ids already exist; otherwise pick them).
4. Click **Activate**. The endpoint is then live at
   `https://<your-n8n-host>/webhook/kb-search`.
5. Put the **same token** in the backend's `.env` as `KB_SEARCH_TOKEN`, and set
   `KB_SEARCH_URL` to the webhook URL (see `.env.example`).

## ⚠️ First: rotate your API key

You pasted a live API key in chat at some point. Go to your provider's key
dashboard, revoke it, and generate a new one. Never put API keys directly in
add-in front-end code (commands.js/taskpane.js) — that code is fully visible
to anyone who inspects the add-in, so the key would be stolen instantly. This
project avoids that by keeping the key only in the backend's `.env` file.

## Project structure

```
outlook-ai-draft/
├── manifest.xml          # Add-in manifest (sideload this into Outlook)
├── src/
│   ├── commands.html      # Function-file host for the ribbon button
│   ├── commands.js        # Ribbon button logic
│   ├── taskpane.html       # Fallback UI (opens if user clicks the app icon)
│   ├── taskpane.js
│   └── icons/              # Placeholder icons — swap these for your branding
└── server/
    ├── server.js           # Express proxy: Outlook add-in → this → OpenAI
    ├── package.json
    └── .env.example
```

## 1. Backend setup

```bash
cd server
npm install
cp .env.example .env
# edit .env, paste your NEW OpenAI key into OPENAI_API_KEY
```

Outlook add-ins require HTTPS, even for local testing. The easiest way to get
a trusted local cert:

```bash
npx office-addin-dev-certs install
```

This installs a cert Office trusts, at `~/.office-addin-dev-certs/`. Point
your `.env` at it:

```
SSL_CERT_PATH=/home/YOU/.office-addin-dev-certs/localhost.crt
SSL_KEY_PATH=/home/YOU/.office-addin-dev-certs/localhost.key
```

Then run the server:

```bash
npm start
```

You should see `HTTPS server running on https://localhost:3000`. Visit
`https://localhost:3000/taskpane.html` in a browser to confirm it loads
without a cert warning.

## 2. Sideload the add-in into Outlook

**Outlook on the web / new Outlook (Windows):**
1. Open Outlook → Settings (gear) → **Manage add-ins** → **My add-ins**
2. Scroll to **Custom add-ins** → **Add a custom add-in** → **Add from file**
3. Select `manifest.xml`
4. Open any email — you'll see a **Draft with AI** button in the ribbon

**Outlook Desktop (classic, Windows/Mac):**
1. File → Manage Add-ins (or Get Add-ins → My add-ins → Add Custom Add-in →
   From File)
2. Select `manifest.xml`

If Outlook rejects the manifest, double check the `IconUrl`,
`AppDomain`, and `SourceLocation` URLs in `manifest.xml` all match
where your server is actually running.

## 3. Using it

- Open any email
- Click **Draft with AI**
- A notification shows "Generating AI draft…"
- Outlook's reply window opens, pre-filled with the AI-written reply
- Review, edit, and send — the add-in never sends anything itself

## Going to production

For real users (not just your own testing), you'll need:

1. **Host the front-end + backend somewhere with a real HTTPS domain**
   (e.g. Render, Railway, Fly.io, a small VPS, or Azure App Service — the
   `server/server.js` Express app can serve both the static files and the
   `/api/generate-draft` endpoint from one deployment).
2. **Update every `https://localhost:3000` reference in `manifest.xml`,
   `commands.js`, and `taskpane.js`** to your real domain.
3. **Set `OPENAI_API_KEY` as an environment variable** in your hosting
   provider's dashboard — never bake it into the deployed code or a Docker
   image layer.
4. If you want this listed for other people to install (not just sideload),
   submit it through **Microsoft AppSource** / the Microsoft 365 admin center
   for org-wide deployment — that's a separate submission process from
   sideloading.

## Notes on "user permission"

The manifest requests `<Permissions>ReadWriteMailbox</Permissions>`, which is
what triggers Outlook's install-time consent prompt (read/write the current
item). The add-in reads the currently open email's subject/body and opens a
reply form — it doesn't read anything else in the mailbox, and it doesn't
send email on its own.

## Model choice

Default model is `gpt-4o-mini`. Change `OPENAI_MODEL` in `.env` to any other
OpenAI chat-completions model ID, e.g. `gpt-4o` or `gpt-4.1`, depending on
cost/quality tradeoffs you want.
