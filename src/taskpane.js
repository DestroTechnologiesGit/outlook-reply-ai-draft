/* global Office */

const BACKEND_URL = "https://outlook-reply-ai-draft.onrender.com/api/generate-draft";

Office.onReady((info) => {
  const btn = document.getElementById("draftBtn");
  // Outside Outlook (e.g. opened directly in a browser) there is no mailbox.
  if (!info.host || !Office.context.mailbox) {
    btn.disabled = true;
    setStatus(
      "Server is reachable ✓ — but this page only works inside Outlook. " +
        "Open an email in Outlook and launch the add-in from there."
    );
    return;
  }
  btn.addEventListener("click", handleClick);
});

function setStatus(text) {
  document.getElementById("status").textContent = text;
}

function getBodyText(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

async function handleClick() {
  const btn = document.getElementById("draftBtn");
  const item = Office.context.mailbox.item;

  btn.disabled = true;
  setStatus("Generating AI draft…");

  try {
    const bodyText = await getBodyText(item);
    const subject = item.subject || "";
    const sender =
      (item.from && (item.from.displayName || item.from.emailAddress)) || "";

    const response = await fetch(BACKEND_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ subject, from: sender, body: bodyText }),
    });

    if (!response.ok) {
      throw new Error(`Backend error ${response.status}: ${await response.text()}`);
    }

    const data = await response.json();
    const draftText = (data.draft || "").trim();
    if (!draftText) throw new Error("No draft text returned.");

    item.displayReplyForm(draftText);
    setStatus("Draft inserted into reply window.");
  } catch (err) {
    console.error(err);
    setStatus("Error: " + (err && err.message ? err.message : err));
  } finally {
    btn.disabled = false;
  }
}
