// Ensure Office is initialized
Office.initialize = () => { /* ready */ };

// ---- Config ----
const CONFIG = {
  internalDomain: "innobothealth.com",
  riskyExtensions: [
    ".jpg",".jpeg",".png",".gif",".bmp",".tiff",
    ".mp4",".mov",".avi",".mkv",".webm",
    ".zip",".rar",".7z",".gz",".tgz"
  ]
};

// ---- HIPAA-ish patterns ----
const HIPAA_PATTERNS = [
  /\b\d{3}-\d{2}-\d{4}\b/,                               // SSN
  /\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b/,                   // DOB / dates
  /\b[A-TV-Z][0-9][0-9AB](\.[0-9A-TV-Z]{1,4})?\b/,       // ICD-10
  /(MRN|Medical\s*Record|Patient\s*Name|Diagnosis|DOB|Chart|Encounter)/i
];
const containsHIPAA = (t="") => HIPAA_PATTERNS.some(rx => rx.test(t));

// ---- ItemSend handler (must be global) ----
function onMessageSend(event) {
  const item = Office.context.mailbox.item;

  Promise.all([
    getBodyText(item),       // plain text
    getBodyHtml(item),       // html for links
    Promise.resolve(item.subject || ""),
    Promise.resolve(item.attachments || [])
  ]).then(([bodyText, bodyHtml, subject, attachments]) => {
    // Build warnings
    const warnings = [];

    if (containsHIPAA(subject) || containsHIPAA(bodyText)) {
      warnings.push("⚠️ Possible HIPAA-sensitive text detected in subject/body.");
    }

    const extLinks = extractExternalLinks(bodyHtml, bodyText, CONFIG.internalDomain);
    if (extLinks.length > 0) {
      warnings.push(`🔗 External links found:\n• ${extLinks.slice(0,8).join("\n• ")}${extLinks.length>8?`\n…and ${extLinks.length-8} more`:``}\nPlease confirm linked content is PHI-free.`);
    }

    const riskyNames = (attachments || [])
      .map(a => a.name || a.id || "")
      .filter(name => CONFIG.riskyExtensions.includes((name.toLowerCase().match(/\.[a-z0-9]+$/)||[""])[0]));
    if ((attachments || []).length > 0) {
      warnings.push(`📎 ${attachments.length} attachment(s) detected.`);
    }
    if (riskyNames.length > 0) {
      warnings.push(`🖼️ Unscannable/risky file types:\n• ${riskyNames.join("\n• ")}\nPlease confirm these are PHI-free.`);
    }

    if (warnings.length === 0) {
      // No issues -> allow send
      event.completed({ allowEvent: true });
      return;
    }

    // Show dialog and wait for user
    const dialogUrl = `${getBaseUrl()}/modal.html?w=${encodeURIComponent(warnings.join("\n\n"))}`;
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 50, width: 50, displayInIframe: true },
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          // If dialog failed, block send (fail-safe)
          event.completed({ allowEvent: false });
          return;
        }
        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          // Message from modal: "allow" or "block"
          try { dialog.close(); } catch (e) {}
          if (arg.message === "allow") {
            event.completed({ allowEvent: true });
          } else {
            event.completed({ allowEvent: false });
          }
        });
      }
    );
  }).catch(err => {
    // Any error -> fail-safe block
    console.error("HIPAA Add-in error:", err);
    event.completed({ allowEvent: false });
  });
}

// ---- Helpers ----
function getBodyText(item) {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? (res.value || "") : "");
    });
  });
}

function getBodyHtml(item) {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Html, (res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? (res.value || "") : "");
    });
  });
}

function extractExternalLinks(html, text, internalDomain) {
  const links = new Set();

  // From HTML <a href="">
  if (html) {
    const hrefRegex = /href="([^"]+)"/gi;
    let m;
    while ((m = hrefRegex.exec(html)) !== null) {
      links.add(m[1]);
    }
  }

  // From raw text
  if (text) {
    const urlRegex = /(https?:\/\/[^\s<]+)/g;
    let m;
    while ((m = urlRegex.exec(text)) !== null) {
      links.add(m[1]);
    }
  }

  // Filter internal domain + mailto
  return [...links]
    .filter(u => !u.startsWith("mailto:"))
    .filter(u => !u.includes(internalDomain));
}

// Compute base URL where this file is hosted
function getBaseUrl() {
  const scripts = document.getElementsByTagName("script");
  for (let s of scripts) {
    if (s.src && s.src.includes("functions.js")) {
      const u = new URL(s.src);
      return `${u.protocol}//${u.host}${u.pathname.substring(0, u.pathname.lastIndexOf("/"))}`;
    }
  }
  // Fallback (edit if needed)
  return "https://YOURDOMAIN.com/hipaa-outlook-addin";
}

// Expose globally for the manifest
window.onMessageSend = onMessageSend;
