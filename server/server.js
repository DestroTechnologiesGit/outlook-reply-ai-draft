const express = require("express");
const cors = require("cors");
const fs = require("fs");
const https = require("https");
require("dotenv").config();

const app = express();
app.use(cors());
app.use(express.json({ limit: "2mb" }));

// Serve the add-in's static front-end files (manifest points at this same origin).
app.use(express.static(require("path").join(__dirname, "..", "src")));

const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";

// RAG knowledge base (optional). When set, the server retrieves relevant
// Purlfrost KB context from the n8n "KB Search API" webhook and grounds the
// draft in it. If unset or unreachable, the server falls back to a plain,
// ungrounded draft so replies never break.
const KB_SEARCH_URL = process.env.KB_SEARCH_URL;
const KB_SEARCH_TOKEN = process.env.KB_SEARCH_TOKEN;
const KB_TIMEOUT_MS = parseInt(process.env.KB_TIMEOUT_MS || "8000", 10);

if (!OPENAI_API_KEY) {
  console.warn(
    "WARNING: OPENAI_API_KEY is not set. Create a .env file (see .env.example)."
  );
}
if (!KB_SEARCH_URL) {
  console.warn(
    "NOTE: KB_SEARCH_URL is not set — running without RAG (plain drafts). " +
      "Set KB_SEARCH_URL + KB_SEARCH_TOKEN to ground drafts in the Purlfrost KB."
  );
}

// Fetch grounding context from the KB Search webhook. Never throws — on any
// error (unset, timeout, non-200) it returns "" so drafting continues.
async function fetchKbContext(query) {
  if (!KB_SEARCH_URL || !query || !query.trim()) return "";

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), KB_TIMEOUT_MS);
  try {
    const resp = await fetch(KB_SEARCH_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...(KB_SEARCH_TOKEN ? { "X-KB-Token": KB_SEARCH_TOKEN } : {}),
      },
      body: JSON.stringify({ text: query }),
      signal: controller.signal,
    });
    if (!resp.ok) {
      console.warn("KB search non-200:", resp.status, (await resp.text()).slice(0, 200));
      return "";
    }
    const data = await resp.json();
    const context = (data && data.context) || "";
    console.log(`KB search: ${data && data.hits ? data.hits : 0} hits`);
    return typeof context === "string" ? context : "";
  } catch (err) {
    console.warn("KB search failed (using ungrounded draft):", err.name || err.message);
    return "";
  } finally {
    clearTimeout(timer);
  }
}

app.post("/api/generate-draft", async (req, res) => {
  try {
    const { subject, from, body } = req.body || {};

    if (!body) {
      return res.status(400).json({ error: "Missing email body." });
    }

    // Retrieve grounding context from the Purlfrost knowledge base (RAG).
    const kbQuery = `${subject || ""}\n\n${body}`.trim();
    const kbContext = await fetchKbContext(kbQuery);

    const kbBlock = kbContext
      ? `\nUse ONLY the Purlfrost knowledge base below for facts about products, prices, policies, delivery, returns and procedures. Do not invent details. If the knowledge base doesn't cover something, keep the reply general and suggest contacting customer service rather than guessing.

--- PURLFROST KNOWLEDGE BASE ---
${kbContext}
--- END KNOWLEDGE BASE ---
`
      : "";

    const prompt = `You are drafting a reply to the email below on behalf of Purlfrost (a UK window-film retailer). Write only the reply body text — no subject line, no "Dear/Hi" boilerplate unless natural, no explanations, no markdown. Keep it concise and professional, matching the tone of the original.
${kbBlock}
From: ${from || "unknown"}
Subject: ${subject || "(no subject)"}
---
${body}
---

Reply:`;

    const aiResponse = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: OPENAI_MODEL,
        messages: [{ role: "user", content: prompt }],
        max_tokens: 500,
      }),
    });

    if (!aiResponse.ok) {
      const errText = await aiResponse.text();
      console.error("OpenAI error:", aiResponse.status, errText);
      return res.status(502).json({ error: `OpenAI error: ${errText}` });
    }

    const data = await aiResponse.json();
    const draft = data?.choices?.[0]?.message?.content || "";

    res.json({ draft, grounded: Boolean(kbContext) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Internal error" });
  }
});

const PORT = process.env.PORT || 3000;

// Office Add-ins require HTTPS. For local dev, use office-addin-dev-certs
// (see README) to generate a trusted localhost cert.
const certPath = process.env.SSL_CERT_PATH;
const keyPath = process.env.SSL_KEY_PATH;

if (certPath && keyPath && fs.existsSync(certPath) && fs.existsSync(keyPath)) {
  https
    .createServer(
      { cert: fs.readFileSync(certPath), key: fs.readFileSync(keyPath) },
      app
    )
    .listen(PORT, () => console.log(`HTTPS server running on https://localhost:${PORT}`));
} else {
  console.warn(
    "No SSL cert configured — starting plain HTTP. Outlook add-ins need HTTPS for anything but this local fallback; see README."
  );
  app.listen(PORT, () => console.log(`HTTP server running on http://localhost:${PORT}`));
}
