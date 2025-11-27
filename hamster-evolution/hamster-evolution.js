const LOG_ADDRESS =
  "email-log-appsRuWy6BC5D4NuW.2aca-wtrgt8WWx3ZdCuIaY.2a98@automations.airtableemail.com";

Office.onReady(() => {
  const btn = document.getElementById("applyCodeButton");
  if (btn) {
    btn.addEventListener("click", applyEmailCode);
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

  // Must be a compose message with subject + BCC access.
  if (
    !item ||
    !item.subject ||
    !item.bcc ||
    item.itemType !== Office.MailboxEnums.ItemType.Message
  ) {
    setStatus("Open Hamster Evolution while composing an email.");
    return;
  }

  const prefix = `${whenCode} - ${typeCode} - ${timeCode} - `;

  // 1) Update the subject
  item.subject.getAsync((subjectResult) => {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Could not read the subject.");
      return;
    }

    const currentSubject = subjectResult.value || "";

    // Remove any existing prefix to avoid stacking.
    const prefixRegex =
      /^(HP|QH|TWP|TWM|TWE|NTW) - (A|I|D|Q|R) - (1m|3m|5m|10m|15m|30m|60m|1\+h) - /;

    const strippedSubject = currentSubject.replace(prefixRegex, "");

    item.subject.setAsync(prefix + strippedSubject, (setSubjectResult) => {
      if (setSubjectResult.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Could not set the subject.");
        return;
      }

      // 2) Ensure logging BCC is present
      ensureLoggingBcc(item, (bccOk) => {
        if (bccOk) {
          setStatus("Hamster Evolution applied.");
        } else {
          setStatus("Hamster Evolution applied, but BCC could not be updated.");
        }
      });
    });
  });
}

function ensureLoggingBcc(item, callback) {
  item.bcc.getAsync((bccResult) => {
    if (bccResult.status !== Office.AsyncResultStatus.Succeeded) {
      callback(false);
      return;
    }

    let recipients = bccResult.value || [];

    const exists = recipients.some(
      (r) =>
        r &&
        r.emailAddress &&
        r.emailAddress.toLowerCase() === LOG_ADDRESS.toLowerCase()
    );

    if (exists) {
      callback(true);
      return;
    }

    // Copy and append our logging address.
    recipients = recipients.slice();
    recipients.push({ emailAddress: LOG_ADDRESS, displayName: "" });

    item.bcc.setAsync(recipients, (setResult) => {
      callback(setResult.status === Office.AsyncResultStatus.Succeeded);
    });
  });
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
