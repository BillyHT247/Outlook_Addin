// Email Coding

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

  // Must be a compose message with a subject we can edit.
  if (
    !item ||
    !item.subject ||
    item.itemType !== Office.MailboxEnums.ItemType.Message
  ) {
    setStatus("Open Email Coding while composing an email.");
    return;
  }

  const prefix = `${whenCode} - ${typeCode} - ${timeCode} - `;

  // 1) Read the current subject
  item.subject.getAsync((subjectResult) => {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Could not read the subject.");
      return;
    }

    const currentSubject = subjectResult.value || "";

    // Remove any existing Email Coding prefix to avoid stacking.
    const prefixRegex =
      /^(HP|QH|TWP|TWM|TWE|NTW) - (A|I|D|Q|R) - (1m|3m|5m|10m|15m|30m|60m|1\+h) - /;

    const strippedSubject = currentSubject.replace(prefixRegex, "");

    // 2) Write the new subject
    item.subject.setAsync(prefix + strippedSubject, (setSubjectResult) => {
      if (setSubjectResult.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Could not set the subject.");
        return;
      }

      setStatus("Email Coding applied");
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
