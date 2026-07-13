/* global Office */

// Point this at your backend proxy (never call OpenAI directly from here —
// that would expose your API key to anyone who views the add-in source).
const BACKEND_URL = "https://outlook-ai.srv1747149.hstgr.cloud/api/generate-draft";
const NOTIFICATION_KEY = "aiDraftStatus";

Office.onReady(() => {
  // Office.js is ready; Office.actions.associate below wires up the button.
});

function showNotification(item, message, isError) {
  try {
    item.notificationMessages.replaceAsync(NOTIFICATION_KEY, {
      type: isError
        ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
        : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message,
      icon: "icon-16",
      persistent: false,
    });
  } catch (e) {
    // notificationMessages isn't available on every host; ignore silently.
  }
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

// displayReplyForm treats its string as HTML, so plain-text newlines collapse.
// Escape the draft and convert paragraphs/line breaks to real HTML.
function draftToHtml(text) {
  const escaped = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  const paragraphs = escaped
    .split(/\n{2,}/)
    .map((p) => "<p>" + p.trim().replace(/\n/g, "<br>") + "</p>")
    .join("");
  return (
    '<div style="font-family: Calibri, Arial, sans-serif; font-size: 11pt;">' +
    paragraphs +
    "</div>"
  );
}

async function generateDraft(event) {
  const item = Office.context.mailbox.item;

  try {
    showNotification(item, "Generating AI draft…", false);

    const bodyText = await getBodyText(item);
    const subject = item.subject || "";
    const sender =
      (item.from && (item.from.displayName || item.from.emailAddress)) || "";

    const response = await fetch(BACKEND_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        subject: subject,
        from: sender,
        body: bodyText,
      }),
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Backend error ${response.status}: ${errText}`);
    }

    const data = await response.json();
    const draftText = (data.draft || "").trim();

    if (!draftText) {
      throw new Error("No draft text returned from backend.");
    }

    // Opens Outlook's native reply window, pre-filled with the AI draft.
    // The user still reviews and hits Send themselves — nothing is sent automatically.
    item.displayReplyForm({ htmlBody: draftToHtml(draftText) });

    showNotification(item, "AI draft inserted into reply.", false);
  } catch (err) {
    console.error(err);
    showNotification(
      item,
      "Couldn't generate draft: " + (err && err.message ? err.message : err),
      true
    );
  } finally {
    event.completed();
  }
}

// Required: makes generateDraft callable from the manifest's ExecuteFunction action.
Office.actions.associate("generateDraft", generateDraft);
