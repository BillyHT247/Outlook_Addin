// Email Coding – subject + body header (no BCC)
// Uses a <div id="emailCodingHeader"> wrapper to update in-place.

const HEADER_DIV_ID = "emailCodingHeader";

Office.onReady(() => {
  const btn = document.getElementById("applyCodeButton");
  if (btn) {
    btn.addEventListener("click", applyEmailCode);
  }

  // Show/hide custom time row based on mode
  const timeMode = document.getElementById("dueTimeMode");
  const customRow = document.getElementById("dueCustomTimeRow");
  if (timeMode && customRow) {
    const updateCustomVisibility = () => {
      customRow.style.display =
        timeMode.value === "CUSTOM" ? "block" : "none";
    };
    timeMode.addEventListener("change", updateCustomVisibility);
    updateCustomVisibility();
  }
});

function applyEmailCode() {
  const whenCode = getSelectValue("whenSelect");
  const typeCode = getSelectValue("typeSelect");
  const timeCode = getSelectValue("timeSelect");

  if (!whenCode || !typeCode || !timeCode) {
    setStatus("Please select WHEN, TYPE, and TIME.");
    return;
  }

  const item = Office.context.mailbox.item;

  // Must be a compose message with editable subject
  if (
    !item ||
    !item.subject ||
    item.itemType !== Office.MailboxEnums.ItemType.Message
  ) {
    setStatus("Open Email Coding while composing an email.");
    return;
  }

  const prefix = `${whenCode} - ${typeCode} - ${timeCode} - `;

  // Read & update subject first
  item.subject.getAsync((subjectResult) => {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Could not read the subject.");
      return;
    }

    const currentSubject = subjectResult.value || "";

    // Strip any existing code prefix to avoid stacking
    const prefixRegex =
      /^(HP|QH|TWP|TWM|TWE|NTW) - (A|I|D|Q|R) - (1m|3m|5m|10m|15m|30m|60m|1h\+) - /;

    const strippedSubject = currentSubject.replace(prefixRegex, "");
    const newSubject = prefix + strippedSubject;

    item.subject.setAsync(newSubject, (setSubjectResult) => {
      if (setSubjectResult.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Could not set the subject.");
        return;
      }

      // After subject is updated, update the body header
      updateEmailBodyHeader(item, typeCode, timeCode);
    });
  });
}

// Insert / update the header block at the top of the email body
function updateEmailBodyHeader(item, typeCode, timeCode) {
  const headerInputEl = document.getElementById("headerInput");
  const headerText = headerInputEl
    ? headerInputEl.value.trim()
    : "";

  const showEffort =
    typeCode === "A" || typeCode === "I" || typeCode === "D";
  const showDue = typeCode === "A";

  const dueInfo = showDue ? buildDueInfo() : "";

  const lines = [];

  // Header: user-written line (if any)
  if (headerText) {
    lines.push("Header: " + headerText);
  }

  // Effort: TIME code, only for A/I/D
  if (showEffort && timeCode) {
    lines.push("Effort: " + timeCode);
  }

  // Due: only for TYPE = A and if we have enough info
  if (dueInfo) {
    lines.push("Due: " + dueInfo);
  }

  item.body.getAsync(
    Office.CoercionType.Html,
    (bodyResult) => {
      if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Could not read the email body.");
        return;
      }

      let body = bodyResult.value || "";

      // Remove any existing header <div id="emailCodingHeader">...</div>
      body = removeExistingHeaderDiv(body);

      // If no header/effort/due lines to insert, we're done
      if (lines.length === 0) {
        item.body.setAsync(
          body,
          { coercionType: Office.CoercionType.Html },
          (setBodyResult) => {
            if (
              setBodyResult.status !==
              Office.AsyncResultStatus.Succeeded
            ) {
              setStatus("Subject updated; body header unchanged.");
            } else {
              setStatus("Email Coding applied (subject only).");
            }
          }
        );
        return;
      }

      const escapedLines = lines.map(escapeHtml);
      const headerHtml =
        `<div id="${HEADER_DIV_ID}"><p>` +
        escapedLines.join("<br>") +
        "</p></div><br>";

      const newBody = headerHtml + body;

      item.body.setAsync(
        newBody,
        { coercionType: Office.CoercionType.Html },
        (setBodyResult) => {
          if (
            setBodyResult.status !==
            Office.AsyncResultStatus.Succeeded
          ) {
            setStatus(
              "Subject updated, but could not update body header."
            );
          } else {
            setStatus("Email Coding applied.");
          }
        }
      );
    }
  );
}

// Remove all existing header <div id="emailCodingHeader">...</div> blocks
// plus a leading <br> (or <p><br></p>) after the last one.
function removeExistingHeaderDiv(html) {
  // Match any <div ... id="emailCodingHeader" ...>...</div>, case-insensitive,
  // allowing extra attributes and any content inside.
  const divRegex = new RegExp(
    "<div[^>]*id=([\"'])" +
      HEADER_DIV_ID +
      "\\1[^>]*>[\\s\\S]*?<\\/div>",
    "ig"
  );

  let newHtml = html.replace(divRegex, "");

  // After removing, also strip one leading <br> or <p><br></p> at the very start.
  newHtml = newHtml.replace(
    /^(\s*<br\s*\/?>|\s*<p>\s*<br\s*\/?>\s*<\/p>)/i,
    ""
  );

  return newHtml;
}


// Build the Due: line (for TYPE = A only)
function buildDueInfo() {
  const dateEl = document.getElementById("dueDate");
  const modeEl = document.getElementById("dueTimeMode");

  if (!dateEl || !modeEl) {
    return "";
  }

  const dateValue = dateEl.value; // yyyy-mm-dd
  const mode = modeEl.value;

  if (!dateValue) {
    // No date selected -> no Due line
    return "";
  }

  const [yyyy, mm, dd] = dateValue.split("-");
  if (!yyyy || !mm || !dd) {
    return "";
  }

  // Format as MM/DD/YY to match examples
  const formattedDate = `${mm}/${dd}/${yyyy.slice(2)}`;

  let timePart = "";

  if (mode === "EOB" || mode === "EOD") {
    timePart = mode;
  } else if (mode === "CUSTOM") {
    const hourEl = document.getElementById("dueHour");
    const minuteEl = document.getElementById("dueMinute");
    const ampmEl = document.getElementById("dueAmPm");

    const hourRaw = hourEl ? hourEl.value.trim() : "";
    const minuteRaw = minuteEl ? minuteEl.value.trim() : "";
    const ampm = ampmEl ? ampmEl.value : "";

    if (!hourRaw || !ampm) {
      // Require at least hour + AM/PM, otherwise just date
      return formattedDate;
    }

    let hour = parseInt(hourRaw, 10);
    if (Number.isNaN(hour) || hour < 1 || hour > 12) {
      return formattedDate;
    }

    let minute = minuteRaw || "00";
    let minuteNum = parseInt(minute, 10);
    if (Number.isNaN(minuteNum) || minuteNum < 0 || minuteNum > 59) {
      minuteNum = 0;
    }
    minute = String(minuteNum).padStart(2, "0");

    timePart = `${hour}:${minute}${ampm.toLowerCase()} ET`;
  } else {
    // Mode is empty / unknown -> just the date
    return formattedDate;
  }

  return `${formattedDate} ${timePart}`;
}

function getSelectValue(id) {
  const el = document.getElementById(id);
  return el ? el.value : "";
}

function setStatus(message) {
  const el = document.getElementById("status");
  if (el) {
    el.textContent = message;
  }
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}