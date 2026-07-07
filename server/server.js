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

const OPENROUTER_API_KEY = process.env.OPENROUTER_API_KEY;
const OPENROUTER_MODEL = process.env.OPENROUTER_MODEL || "meta-llama/llama-3.1-8b-instruct:free";

if (!OPENROUTER_API_KEY) {
  console.warn(
    "WARNING: OPENROUTER_API_KEY is not set. Create a .env file (see .env.example)."
  );
}

app.post("/api/generate-draft", async (req, res) => {
  try {
    const { subject, from, body } = req.body || {};

    if (!body) {
      return res.status(400).json({ error: "Missing email body." });
    }

    const prompt = `You are drafting a reply to the email below. Write only the reply body text — no subject line, no "Dear/Hi" boilerplate unless natural, no explanations, no markdown. Keep it concise and professional, matching the tone of the original.

From: ${from || "unknown"}
Subject: ${subject || "(no subject)"}
---
${body}
---

Reply:`;

    const orResponse = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${OPENROUTER_API_KEY}`,
        "Content-Type": "application/json",
        "HTTP-Referer": "https://destrotechnologies.com",
        "X-Title": "Outlook AI Draft Add-in",
      },
      body: JSON.stringify({
        model: OPENROUTER_MODEL,
        messages: [{ role: "user", content: prompt }],
        max_tokens: 500,
      }),
    });

    if (!orResponse.ok) {
      const errText = await orResponse.text();
      console.error("OpenRouter error:", orResponse.status, errText);
      return res.status(502).json({ error: `OpenRouter error: ${errText}` });
    }

    const data = await orResponse.json();
    const draft = data?.choices?.[0]?.message?.content || "";

    res.json({ draft });
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
